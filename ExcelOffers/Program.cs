using OfficeOpenXml;


namespace ExcelOffers
{

    class Program
    {
        public static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            OfferGenerate offerGenerate = new();
            offerGenerate.GenerateOffers();
            
        }


    }
}


