using System;
using System.Runtime.CompilerServices;
using System.Runtime.Remoting.Messaging;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using nrot.T590.Excel.Models;

namespace nrot.T590.Excel.Test
{
    [TestClass]
    public class ExcelConnectorTests
    {
        //private static readonly string ExcelFilePath = @"Files/Patientenliste.xlsx";
        private static readonly string ExcelFilePath = $"{Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)}/nrot.T590/Patientenliste.xlsx";

        [TestMethod]
        public async Task ReadAllPatientsFromExcelAsyncTest()
        {
            var excel = new ExcelConnector(ExcelFilePath);

            var patients = await excel.ReadAllPatientsFromExcelAsync();

        }

        [TestMethod]
        public async Task StorePatientRecordInExcelAsyncTest()
        {
            var patient = new Patient
            {
                Id = 0,
                Name = "Muster",
                Vorname = "Hans",
                Strasse = "Teststrasse 99",
                Plz = 9999,
                Ort = "Testort",
                Geburtsdatum = new DateTime(1901, 1, 1),
                Geschlecht = GeschlechtType.M,
                PatientenNr = "PatientenNr002",
                AhvNr = "AhvNr002",
                VekaNr = "VekaNr002",
                VersichertenNr = "VersNr002",
                Kanton = "TE",
                Kopie = false,
                VerguetungsArt = VerguetungsartType.Tp,
                VertragsNr = "VertrNr002"
            };

            var excel = new ExcelConnector(ExcelFilePath);

            var res = await excel.StorePatientRecordInExcelAsync(patient);
        }
    }
}
