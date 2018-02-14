using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Win32;
using nrot.T590.Excel;
using nrot.T590.Excel.Models;

namespace nrot.T590.Gui
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        // TODO: 2config!
        //private const string ExcelFilePath = @"Files/Patientenliste.xlsx";
        private readonly string ExcelFilePath;
        public static IEnumerable<GeschlechtType> GetGeschlechtEnumValues => Enum.GetValues(typeof(GeschlechtType)).Cast<GeschlechtType>();
        public static IEnumerable<VerguetungsartType> GetVerguetungsArtEnumValues => Enum.GetValues(typeof(VerguetungsartType)).Cast<VerguetungsartType>();

        ExcelConnector _excel;
        Patient _patient = new Patient();

        public MainWindow()
        {
            ExcelFilePath = $"{Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)}/nrot.T590/Patientenliste.xlsx";
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            GetPatientRecords();
        }

        private void btnRefreshRecord_Click(object sender, RoutedEventArgs e)
        {
            GetPatientRecords();
        }

        /// <summary>  
        /// Getting Data of each cell  
        /// </summary>  
        /// <param name="sender"></param>  
        /// <param name="e"></param>  
        private void DataGridPatient_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            try
            {
                Mouse.OverrideCursor = Cursors.Wait;

                for (var i = 0; i < 16; i++)
                {
                    object o = DataGridPatient.Columns[i].GetCellContent(e.Row);
                    switch (i)
                    {
                        case 0: // Id
                        {
                            if (o.GetType() == typeof(TextBox))
                            {
                                _patient.Id = Convert.ToInt32(((TextBox) o).Text);
                            }

                            break;
                        }
                        case 1: // Name
                        {
                            if (o.GetType() == typeof(TextBox))
                            {
                                _patient.Name = ((TextBox) o).Text;
                            }

                            break;
                        }
                        case 2: // Vorname
                        {
                            if (o.GetType() == typeof(TextBox))
                            {
                                _patient.Vorname = ((TextBox) o).Text;
                            }

                            break;
                        }
                        case 3: // Strasse
                        {
                            if (o.GetType() == typeof(TextBox))
                            {
                                _patient.Strasse = ((TextBox) o).Text;
                            }

                            break;
                        }
                        case 4: // Plz
                        {
                            if (o.GetType() == typeof(TextBox))
                            {
                                _patient.Plz = Convert.ToInt32(((TextBox) o).Text);
                            }
                            break;
                        }
                        case 5: // Ort
                        {
                            if (o.GetType() == typeof(TextBox))
                            {
                                _patient.Ort = ((TextBox) o).Text;
                            }

                            break;
                        }
                        case 6: // Geburtsdatum
                        {
                            if (o.GetType() == typeof(TextBox))
                            {
                                var content = ((TextBox) o).Text;

                                if (DateTime.TryParse(content, out DateTime contentAsDateTime))
                                {
                                    _patient.Geburtsdatum = contentAsDateTime;
                                }
                                else
                                {
                                    _patient.Geburtsdatum = null;
                                }
                            }

                            break;
                        }
                        case 7: // Geschlecht
                        {
                            if (o.GetType().BaseType == typeof(ComboBox))
                            {
                                var content = ((ComboBox) o).Text.ToUpper();

                                switch (content)
                                {
                                    case "M":
                                    {
                                        _patient.Geschlecht = GeschlechtType.M;
                                        break;
                                    }
                                    case "W":
                                    {
                                        _patient.Geschlecht = GeschlechtType.W;
                                        break;
                                    }
                                    default:
                                    {
                                        throw new Exception($"No matching GeschlechtType '{content}' found.");
                                    }
                                }
                            }

                            break;
                        }
                        case 8: // PatientenNr
                        {
                            if (o.GetType() == typeof(TextBox))
                            {
                                _patient.PatientenNr = ((TextBox) o).Text;
                            }

                            break;
                        }
                        case 9: // AhvNr
                        {
                            if (o.GetType() == typeof(TextBox))
                            {
                                _patient.AhvNr = ((TextBox) o).Text;
                            }

                            break;
                        }
                        case 10: // VekaNr
                        {
                            if (o.GetType() == typeof(TextBox))
                            {
                                _patient.VekaNr = ((TextBox) o).Text;
                            }

                            break;
                        }
                        case 11: // VersichertenNr
                        {
                            if (o.GetType() == typeof(TextBox))
                            {
                                _patient.VersichertenNr = ((TextBox) o).Text;
                            }

                            break;
                        }
                        case 12: // Kanton
                        {
                            if (o.GetType() == typeof(TextBox))
                            {
                                _patient.Kanton = ((TextBox) o).Text;
                            }

                            break;
                        }
                        case 13: // Kopie
                        {
                            if (o.GetType() == typeof(CheckBox))
                            {
                                _patient.Kopie = Convert.ToBoolean(((CheckBox) o).IsChecked);
                            }

                            break;
                        }
                        case 14: // VerguetungsArt
                        {
                            if (o.GetType().BaseType == typeof(ComboBox))
                            {
                                var content = ((ComboBox) o).Text.ToUpper();

                                switch (content)
                                {
                                    case "TG":
                                    {
                                        _patient.VerguetungsArt = VerguetungsartType.Tg;
                                        break;
                                    }
                                    case "TP":
                                    {
                                        _patient.VerguetungsArt = VerguetungsartType.Tp;
                                        break;
                                    }
                                    default:
                                    {
                                        throw new Exception($"No matching VerguetungsartType '{content}' found.");
                                    }
                                }
                            }

                            break;
                        }
                        case 15: // VertragsNr
                        {
                            if (o.GetType() == typeof(TextBox))
                            {
                                _patient.VertragsNr = ((TextBox) o).Text;
                            }
                            break;
                        }
                        default:
                        {
                            throw new Exception($"Column number '{i}' does not exist in DataGridPatient!");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        /// <summary>  
        /// Get entire Row  
        /// </summary>  
        /// <param name="sender"></param>  
        /// <param name="e"></param>  
        private void DataGridPatient_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            try
            {
                Mouse.OverrideCursor = Cursors.Wait;

                var isSave = _excel.StorePatientRecordInExcelAsync(_patient).Result;

                if (isSave)
                {
                    //MessageBox.Show("Patient record saved successfully.");
                    GetPatientRecords();
                }
                else
                {
                    throw new Exception("Error while saving patient record!");
                    //MessageBox.Show("Error problem occured.");
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace, ex.Message);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        /// <summary>  
        /// Get Record info to update  
        /// </summary>  
        /// <param name="sender"></param>  
        /// <param name="e"></param>  
        private void DataGridPatient_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                Mouse.OverrideCursor = Cursors.Wait;
                _patient = DataGridPatient.SelectedItem as Patient;
            }
            catch (Exception)
            {
                // TODO: ExceptionHandling?
                throw;
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void GetPatientRecords()
        {
            try
            {
                Mouse.OverrideCursor = Cursors.Wait;

                _excel = new ExcelConnector(ExcelFilePath);
                DataGridPatient.ItemsSource = _excel.ReadAllPatientsFromExcelAsync().Result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }
    }
}
