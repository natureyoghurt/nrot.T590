using Spire.Pdf;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using nrot.T590.Models;
using Spire.Pdf.Widget;

namespace nrot.T590.Pdf
{
    public static class PdfConnector
    {
        public static string PdfTemplateFilePath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "nrot.T590", Settings.Default.PdfTemplateFileName);

        public static IDictionary<string, string> GetPdfFields(string pdfTemplatePath, string pdfTargetPath)
        {
            var fieldsTypesList = new Dictionary<string, string>();
            var fieldsValuesList = new Dictionary<string, string>();
            // https://www.e-iceblue.com/Knowledgebase/Spire.PDF/Spire.PDF-Program-Guide/FormField/How-to-Fill-XFA-Form-Fields-in-C-VB.NET.html

            var timeStampString = $"{DateTime.Now:yyyyMMdd-HHmmss}_";
            var timedPdfTargetPath = Path.Combine(Path.GetDirectoryName(pdfTargetPath), $"{timeStampString}_{Path.GetFileName(pdfTargetPath)}");

            var doc = new PdfDocument();
            doc.LoadFromFile(pdfTemplatePath);

            var formWidget = doc.Form as PdfFormWidget;
            var xfaFields = formWidget.XFAForm.XfaFields;

            foreach (var xfaField in xfaFields)
            {
                if (xfaField is XfaTextField)
                {
                    var xtf = xfaField as XfaTextField;
                    fieldsTypesList.Add(xtf.Name, xtf.FieldType);
                    fieldsValuesList.Add(xtf.Name, xtf.Value);
                }
            }

            doc.SaveToFile(timedPdfTargetPath, FileFormat.PDF);

            return fieldsValuesList;
        }

        public static void GenerateBill(Patient patient)
        {
            var timeStampString = $"{DateTime.Now:yyMMdd-hhmmss}";
            var targetPdfFileName = $"{timeStampString}_Rechnung_{patient.Name}{patient.Vorname}.pdf";
            var targetPdfPath = Path.Combine(Settings.Default.CustomerRootDir, $"{patient.Name} {patient.Vorname}", targetPdfFileName);

            var doc = new PdfDocument();
            doc.LoadFromFile(PdfTemplateFilePath);
            var formWidget = doc.Form as PdfFormWidget;
            var xfaFields = formWidget.XFAForm.XfaFields;

            foreach (var xfaField in xfaFields)
            {
                if (xfaField is XfaTextField)
                {
                    switch (((XfaTextField)xfaField).Name)
                    {
                        #region Überschrift

                        case "request[0].Physio[0].Ueberschrift[0].TP_TG[0]":
                        {
                            ((XfaTextField)xfaField).Value = patient.VerguetungsArt.ToString().ToUpper();
                            break;
                        }

                        #endregion

                        #region Physio Kopf Rechnungssteller
                        case "request[0].Physio[0].Kopf[0].Rechnungssteller_EAN[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsGlnNr;
                                break;
                            }
                        case "request[0].Physio[0].Kopf[0].Rechnungssteller_Namen[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsName;
                                break;
                            }
                        case "request[0].Physio[0].Kopf[0].Rechnungssteller_Email[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsEMail;
                                break;
                            }
                        case "request[0].Physio[0].Kopf[0].Rechnungssteller_Telefon[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsTel;
                                break;
                            }
                        case "request[0].Physio[0].Kopf[0].Rechnungssteller_ZSR[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsZsrNr;
                                break;
                            }
                        case "request[0].Physio[0].Kopf[0].Rechnungssteller_Strasse[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsStrasse;
                                break;
                            }
                        case "request[0].Physio[0].Kopf[0].Rechnungssteller_PLZ[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsPlz;
                                break;
                            }
                        case "request[0].Physio[0].Kopf[0].Rechnungssteller_Ort[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsOrt;
                                break;
                            }
                        #endregion

                        #region Physio Kopf Leistungserbringer
                        case "request[0].Physio[0].Kopf[0].Leistungserbringer_EAN[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsGlnNr;
                                break;
                            }
                        case "request[0].Physio[0].Kopf[0].Leistungserbringer_Namen[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsName;
                                break;
                            }
                        case "request[0].Physio[0].Kopf[0].Leistungserbringer_Email[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsEMail;
                                break;
                            }
                        case "request[0].Physio[0].Kopf[0].Leistungserbringer_Telefon[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsTel;
                                break;
                            }
                        case "request[0].Physio[0].Kopf[0].Leistungserbringer_ZSR[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsZsrNr;
                                break;
                            }
                        case "request[0].Physio[0].Kopf[0].Leistungserbringer_Strasse[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsStrasse;
                                break;
                            }
                        case "request[0].Physio[0].Kopf[0].Leistungserbringer_PLZ[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsPlz;
                                break;
                            }
                        case "request[0].Physio[0].Kopf[0].Leistungserbringer_Ort[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsOrt;
                                break;
                            }
                        #endregion

                        #region Physio Patient
                        case "request[0].Physio[0].Patient[0].Name[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.Name;
                                break;
                            }
                        case "request[0].Physio[0].Patient[0].Vorname[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.Vorname;
                                break;
                            }
                        case "request[0].Physio[0].Patient[0].Strasse[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.Strasse;
                                break;
                            }
                        case "request[0].Physio[0].Patient[0].PLZ[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.Plz.ToString();
                                break;
                            }
                        case "request[0].Physio[0].Patient[0].Ort[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.Ort;
                                break;
                            }
                        case "request[0].Physio[0].Patient[0].Unfallnummer[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.PatientenNr;
                                break;
                            }
                        case "request[0].Physio[0].Patient[0].AHV_Nr[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.AhvNr;
                                break;
                            }
                        case "request[0].Physio[0].Patient[0].VeKa_Nr[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.VekaNr;
                                break;
                            }
                        case "request[0].Physio[0].Patient[0].Versicherten_Nr[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.VersichertenNr;
                                break;
                            }
                        case "request[0].Physio[0].Patient[0].Vertrags_Nr[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.VertragsNr;
                                break;
                            }
                        #endregion

                        #region KoTr Kopf Rechnungssteller
                        case "request[0].KoTr[0].Kopf[0].Rechnungssteller_EAN[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsGlnNr;
                                break;
                            }
                        case "request[0].KoTr[0].Kopf[0].Rechnungssteller_Namen[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsName;
                                break;
                            }
                        case "request[0].KoTr[0].Kopf[0].Rechnungssteller_Email[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsEMail;
                                break;
                            }
                        case "request[0].KoTr[0].Kopf[0].Rechnungssteller_Telefon[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsTel;
                                break;
                            }
                        case "request[0].KoTr[0].Kopf[0].Rechnungssteller_ZSR[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsZsrNr;
                                break;
                            }
                        case "request[0].KoTr[0].Kopf[0].Rechnungssteller_Adresse[0]":
                            {
                                ((XfaTextField)xfaField).Value = $"{Settings.Default.RsStrasse} {Settings.Default.RsPlz} {Settings.Default.RsOrt}";
                                break;
                            }
                        #endregion

                        #region KoTr Kopf Leistungserbringer
                        case "request[0].KoTr[0].Kopf[0].Leistungserbringer_EAN[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsGlnNr;
                                break;
                            }
                        case "request[0].KoTr[0].Kopf[0].Leistungserbringer_Namen[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsName;
                                break;
                            }
                        case "request[0].KoTr[0].Kopf[0].Leistungserbringer_Email[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsEMail;
                                break;
                            }
                        case "request[0].KoTr[0].Kopf[0].Leistungserbringer_Telefon[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsTel;
                                break;
                            }
                        case "request[0].KoTr[0].Kopf[0].Leistungserbringer_ZSR[0]":
                            {
                                ((XfaTextField)xfaField).Value = Settings.Default.RsZsrNr;
                                break;
                            }
                        case "request[0].KoTr[0].Kopf[0].Leistungserbringer_Adresse[0]":
                            {
                                ((XfaTextField)xfaField).Value = $"{Settings.Default.RsStrasse} {Settings.Default.RsPlz} {Settings.Default.RsOrt}";
                                break;
                            }
                        #endregion

                        #region KoTr Patient
                        case "request[0].KoTr[0].Patient[0].Name[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.Name;
                                break;
                            }
                        case "request[0].KoTr[0].Patient[0].Vorname[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.Vorname;
                                break;
                            }
                        case "request[0].KoTr[0].Patient[0].Strasse[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.Strasse;
                                break;
                            }
                        case "request[0].KoTr[0].Patient[0].PLZ[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.Plz.ToString();
                                break;
                            }
                        case "request[0].KoTr[0].Patient[0].Ort[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.Ort;
                                break;
                            }
                        case "request[0].KoTr[0].Patient[0].Geschlecht[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.Geschlecht.ToString();
                                break;
                            }
                        case "request[0].KoTr[0].Patient[0].Unfallnummer_KoTr[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.PatientenNr;
                                break;
                            }
                        case "request[0].KoTr[0].Patient[0].AHV_Nr[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.AhvNr;
                                break;
                            }
                        case "request[0].KoTr[0].Patient[0].VeKa_Nr[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.VekaNr;
                                break;
                            }
                        case "request[0].KoTr[0].Patient[0].Versicherten_Nr_KoTr[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.VersichertenNr;
                                break;
                            }
                        case "request[0].KoTr[0].Patient[0].Kanton[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.Kanton;
                                break;
                            }
                        case "request[0].KoTr[0].Patient[0].Rechnungskopie[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.Kopie ? "Ja" : "Nein";
                                break;
                            }
                        case "request[0].KoTr[0].Patient[0].Vergütungsart[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.VerguetungsArt.ToString().ToUpper();
                                break;
                            }
                        case "request[0].KoTr[0].Patient[0].Vertragsnummer[0]":
                            {
                                ((XfaTextField)xfaField).Value = patient.VertragsNr;
                                break;
                            }
                        #endregion

                        default:
                        {
                            break;
                        }
                    }
                }
            }

            doc.SaveToFile(targetPdfPath, FileFormat.PDF);
            doc.Close();
        }
    }
}
