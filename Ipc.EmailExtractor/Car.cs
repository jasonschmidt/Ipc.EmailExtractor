namespace Ipc.EmailExtractor
{
    public class Car
    {
        public Car()
        {
            this.Mileage = 0.0;
        }

        public string Year { get; set; }
        public string Make { get; set; }
        public string Model { get; set; }
        public string VIN { get; set; }
        public string Color { get; set; }
        public double Mileage { get; set; }

        public string AutoniqLink { get; set; }
    }
}
