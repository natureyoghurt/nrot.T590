using Microsoft.VisualStudio.TestTools.UnitTesting;

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
            PdfConnector.GetPdfFields(PdfFilePath2, PdfOutputPath);
        }
    }
}
