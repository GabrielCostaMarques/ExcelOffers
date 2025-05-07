namespace ExcelOffers.ExcelOffers.Domain.Entities
{
    internal class Localization
    {
        public string PortEmbarkDesembark { get; set; }
        public string DestinyProduct { get; set; }
        public DateTime EmbarkDate { get; set; }

        public Localization() { }
        public Localization(string portEmbarkDesembark, string destinyProduct, DateTime embarkDate)
        {
            PortEmbarkDesembark = portEmbarkDesembark;
            DestinyProduct = destinyProduct;
            EmbarkDate = embarkDate;
        }
    }
}
