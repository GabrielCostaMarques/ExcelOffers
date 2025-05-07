namespace ExcelOffers.ExcelOffers.Domain.Entities
{
    internal class Product
    {
        public string ShipName { get; set; }
        public string ProductName { get; set; }
        public string CabinCategory { get; set; }
        public string CabinClass { get; set; }
        public Fares Fares { get; set; } = new Fares();
        public Localization Localization { get; set; } = new Localization();

        public Product() { }
        public Product(string shipName, string productName, string cabinCategory, string cabinClass, Fares fares, Localization localization)
        {
            ShipName = shipName;
            ProductName = productName;
            CabinCategory = cabinCategory;
            CabinClass = cabinClass;
            Fares = fares;
            Localization = localization;
        }
    }
}
