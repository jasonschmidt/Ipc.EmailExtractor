namespace Ipc.EmailExtractor
{
    public class Car
    {
        public Car()
        {
            this.Mileage = 0.0;
        }

        public string VIN { get; set; }
        public string Description { get; set; }
        public string Color { get; set; }
        public double Mileage { get; set; }

        public string AutoniqLink { get; set; }
    }
}
