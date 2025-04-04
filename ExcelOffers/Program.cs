using System;
using System.Data;
using System.Security.Cryptography;
using ExcelOffers.Entities;
using ExcelOffers.Factory;
using ExcelOffers.Services;
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


