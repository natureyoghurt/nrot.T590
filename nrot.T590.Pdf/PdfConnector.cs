using Spire.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using nrot.T590.Models;
using Spire.Pdf.Widget;

namespace nrot.T590.Pdf
{
    public static class PdfConnector
    {
        public static void GetPdfFields(string pdfTemplatePath, string pdfTargetPath)
        {
            // https://www.e-iceblue.com/Knowledgebase/Spire.PDF/Spire.PDF-Program-Guide/FormField/How-to-Fill-XFA-Form-Fields-in-C-VB.NET.html

            var timeStampString = $"{DateTime.Now:yyyyMMdd-HHmmss}_";
            var timedPdfTargetPath = Path.Combine(Path.GetDirectoryName(pdfTargetPath), $"{timeStampString}_{Path.GetFileName(pdfTargetPath)}");

            var doc = new PdfDocument();
            doc.LoadFromFile(pdfTemplatePath);

            var formWidget = doc.Form as PdfFormWidget;
            var xfaFields = formWidget.XFAForm.XfaFields;

            var xfaFieldsList = new List<string>();

            foreach (var xfaField in xfaFields)
            {
                if (xfaField is XfaTextField)
                {
                    var xtf = xfaField as XfaTextField;
                    xfaFieldsList.Add($"{xtf.Name}({xtf.FieldType}): '{xtf.Value}'");
                }
            }

            doc.SaveToFile(timedPdfTargetPath, FileFormat.PDF);
        }

        public static void GenerateBill(Patient patient)
        {
            throw new NotImplementedException();
        }
    }
}
