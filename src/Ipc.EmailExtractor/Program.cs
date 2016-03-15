using System.Text.RegularExpressions;
using System;
using System.Configuration;
using System.Collections.Generic;
using ImapX;
using System.IO;
using CsvHelper;
using System.Text;
using System.Net;
using System.Globalization;

namespace Ipc.EmailExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            var cmdline = new Arguments(args);
                       
            var emailAddress = cmdline["emailAddress"] ?? ConfigurationManager.AppSettings["ipc.username"].ToString();
            var password = cmdline["password"] ?? ConfigurationManager.AppSettings["ipc.password"].ToString();

            Console.WriteLine("Using email address: " + emailAddress);
            if (string.IsNullOrEmpty(password)) {
                Console.WriteLine("Please enter your password:");
                password = Console.ReadLine();
                Console.WriteLine("");
            }


            ReadEmail(emailAddress, password);
        }

        static void ReadEmail(string emailAddress, string password)
        {
            var mailServer = ConfigurationManager.AppSettings["ipc.mailserver"].ToString();
            var from = ConfigurationManager.AppSettings["ipc.from"].ToString();
            var startDate = Convert.ToDateTime(ConfigurationManager.AppSettings["ipc.startdate"].ToString());
            var lastUpdated = ReadTimestampMarker(startDate);

            using (ImapClient imap = new ImapClient(mailServer, true))
            {
                imap.IsDebug = true;
              
                    Console.WriteLine("Connecting to {0}....", mailServer);
                if (imap.Connect())
                {
                    Console.WriteLine("Logging in to IMAP server...");
                    bool Islogin = imap.Login(emailAddress, password);
                    if (Islogin)
                    {

                        try
                        {
                            Console.WriteLine("Connected to IMAP server.");
                            Console.WriteLine("Reading emails...");
                            var messages = imap.Folders.Inbox.SubFolders["Autoniq"].Search(string.Format("SINCE {0} HEADER FROM \"{1}\"", lastUpdated.ToString("dd-MMM-yyyy"), from), ImapX.Enums.MessageFetchMode.Basic);
                            Console.WriteLine("Found {0} emails from {1}", messages.Length, from);

                            var emails = new List<EmailMessage>();
                            foreach (var message in messages)
                            {
                                var emailMessage = ParseEmailForCar(message);
                                if (emailMessage != null)
                                {
                                    var car = emailMessage.Car;
                                    Console.WriteLine("\n------ Found car ----- \n{0}\n{1}\n{2}\n{3}\n{4}\n", car.Make, car.VIN, car.Mileage, car.Color, car.AutoniqLink);
                                    emails.Add(emailMessage);
                                }
                            }

                            WriteCsv(EmailType.Bought, emails);
                            WriteCsv(EmailType.Sold, emails);
                            WriteCsv(EmailType.BackInStock, emails);

                            WriteTimestampMarker();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error while parsing messages. Error: " + ex.Message);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Error while authenticating to IMAP account. ");
                    }
                }
                else
                {
                    Console.WriteLine("Error while connecting to IMAP server. ");
                }

                
                Console.WriteLine("Process completed.");
                Console.WriteLine("Press any key to close.");
                Console.ReadLine();
            }
        }

        static void WriteCsv(EmailType emailType, List<EmailMessage> emails)
        {
            var fileName = "cars." + emailType.ToString().ToLower() + ".csv";
            var fileExists = File.Exists(fileName);
            using (var writer = new StreamWriter(fileName, true))
            {
                using (var csv = new CsvWriter(writer))
                {
                    csv.Configuration.Encoding = Encoding.UTF8;
                    if (!fileExists)
                    {
                        csv.WriteHeader(typeof(Car));
                    }
                    foreach (var emailMessage in emails)
                    {
                        if (emailMessage.EmailType == emailType)
                        {
                            csv.WriteRecord<Car>(emailMessage.Car);
                        }
                    }
                }
            }
        }

        static void WriteTimestampMarker(DateTime? startDate = null)
        {
            var dateToUse = startDate ?? DateTime.UtcNow;
            File.WriteAllText("LastUpdated.txt", dateToUse.ToString("MM/dd/yyyy HH:mm:ss"));
        }

        static DateTime ReadTimestampMarker(DateTime startDate)
        {
            if (File.Exists("LastUpdated.txt") == false)
            {
                WriteTimestampMarker(startDate);
            }

            var lastUpdated = File.ReadAllText("LastUpdated.txt");
            DateTime dt;
            if (DateTime.TryParseExact(lastUpdated, "MM/dd/yyyy HH:mm:ss", CultureInfo.GetCultureInfo("en-US"),
                                       DateTimeStyles.None, out dt))
            {
                return dt;
            }

            throw new ArgumentOutOfRangeException("Unable to load date from LastUpdated.txt");

        }

        static EmailMessage ParseEmailForCar(Message message)
        {
            var body = message.Body.HasText ? message.Body.Text : message.Body.Html;
            Match carMatches;

            var emailMessage = new EmailMessage();

            var car = new Car();
            if (message.Body.HasText) {
                carMatches  = Regex.Match(message.Body.Text, "Vehicle:(.*?\n).*?\nVIN:(.*?\n)Mileage:(.*?\n)Color:(.*?\n)", RegexOptions.IgnoreCase);
            //    car.Description = carMatches.Groups[1].Value.Trim();
                
            }
            else
            {
                //carMatches = Regex.Match(message.Body.Html, @"<tr>(?:.?\n)<td(?:.*?)>Vehicle:<\/td>(?:.?\n)<td>(.*?)<\/td>(?:.?\n)<\/tr>(?:.?\n)<tr>(?:.?\n)<td(?:.*?)>VIN:<\/td>(?:.?\n)<td>(.*?)<\/td>(?:.?\n)<\/tr>(?:.?\n)<tr>(?:.?\n)<td(?:.*?)>Mileage:<\/td>(?:.?\n)<td>(.*?)<\/td>(?:.?\n)<\/tr>(?:.?\n)<tr>(?:.?\n)<td(?:.*?)>Color:<\/td>(?:.?\n)<td>(.*?)<\/td>(?:.?\n)<\/tr>(?:.?\n)", RegexOptions.IgnoreCase);
                carMatches = Regex.Match(message.Body.Html,
                    @"<tr>\s*<td(?:.*?)Vehicle:<\/td>\s*<td>\s*(.*?)<\/td>\s*<\/tr>\s*<tr>\s*<td(?:.*?)VIN:<\/td>\s*<td>\s*(.*?)<\/td>\s*<\/tr>\s*<tr>\s*<td(?:.*?)Mileage:<\/td>\s*<td>\s*(.*?)<\/td>\s*<\/tr>\s*<tr>\s*<td(?:.*?)Color:<\/td>\s*<td>\s*(.*?)<\/td>\s*<\/tr>",
                    RegexOptions.IgnoreCase);
                
                var description = carMatches.Groups[1].Value.Trim();
                if (description.Contains("<a"))
                {
                    var linkMatches = Regex.Match(description, "<a href=\"(.*?)\">(.*?)<\\/a>", RegexOptions.IgnoreCase);
                    if (linkMatches.Groups.Count == 3)
                    {
                        ParseDescription(car, linkMatches.Groups[2].Value.Trim());
                        car.AutoniqLink = linkMatches.Groups[1].Value.Trim();
                    }
                    else
                    {
                        ParseDescription(car, carMatches.Groups[1].Value.Trim());    // no go, just do the default
                    }
                }
                else
                {
                    ParseDescription(car, carMatches.Groups[1].Value.Trim());
                }
                

                // determine email type
                var firstLineRegex = Regex.Match(message.Body.Html, "<div style=\"(?:.*\\n)(.*?)(?:.?\\n)<\\/div>");
                if (firstLineRegex.Groups.Count == 2)
                {
                    var firstLine = firstLineRegex.Groups[1].Value.Trim();
                    if (Regex.IsMatch(firstLine, "Sold"))
                    {
                        emailMessage.EmailType = EmailType.Sold;
                    }
                    else if (Regex.IsMatch(firstLine, "Bought"))
                    {
                        emailMessage.EmailType = EmailType.Bought;
                    }
                    else
                    {
                        emailMessage.EmailType = EmailType.BackInStock;
                    }
                }
            }

            if (carMatches.Groups.Count != 5)
            {
                return null; 
            }

            car.VIN = carMatches.Groups[2].Value.Trim();
            car.Mileage = Convert.ToDouble(carMatches.Groups[3].Value.Trim());
            car.Color = carMatches.Groups[4].Value.Trim();
            emailMessage.Car = car;

            // Determine the type of email


            return emailMessage;
            
           
        }

        protected static void ParseDescription(Car car, string description)
        {
            description = WebUtility.HtmlDecode(description);
            var descriptionMatches = Regex.Match(description, @"(\d*)\s(\w+)\s(.*)", RegexOptions.IgnoreCase);
            if (descriptionMatches.Groups.Count == 4)
            {
                car.Year = descriptionMatches.Groups[1].Value.Trim();
                car.Make = descriptionMatches.Groups[2].Value.Trim();
                car.Model = descriptionMatches.Groups[3].Value.Trim();
            }
        }
    }
}
