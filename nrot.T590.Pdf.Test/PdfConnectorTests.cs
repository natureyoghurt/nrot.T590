using System;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using nrot.T590.Models;

namespace nrot.T590.Pdf.Test
{
    [TestClass]
    public class PdfConnectorTests
    {
        public static readonly string PdfFilePath = @"Files/Rechnungsformular_KM_d-v2.3.18.pdf";
        public static readonly string PdfFilePath2 = @"Files/Rechnungsformular_KM_d-v2.3.18_2.pdf";
        public static readonly string PdfOutputPath = @"Files/Output.pdf";

        [TestMethod]
        public void GetPdfFieldsTest()
        {
            //var fields = PdfConnector.GetPdfFields(PdfFilePath2, PdfOutputPath);
            var fieldsValuesList = PdfConnector.GetPdfFields(PdfFilePath2, PdfOutputPath);

            Thread.Sleep(10000);
        }

        [TestMethod]
        public void GenerateBillTest()
        {
            var patient = new Patient
            {
                Id = 1,
                Name = "Muster",
                Vorname = "Hans",
                Strasse = "Musterstrasse 33",
                Plz = 3333,
                Ort = "Musterdorf",
                Geburtsdatum = new DateTime(1999, 9, 9),
                Geschlecht = GeschlechtType.M,
                PatientenNr = "PatientNr_100001",
                AhvNr = "AhvNr_100001",
                VekaNr = "VekaNr_100001",
                VersichertenNr = "VersichertenNr_100001",
                Kanton = "SG",
                Kopie = true,
                VerguetungsArt = VerguetungsartType.Tg,
                VertragsNr = "VertragsNr_100001"
            };

            PdfConnector.GenerateBill(patient);
        }
    }
}
