using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using nrot.T590.Excel;
using nrot.T590.Excel.Models;

namespace nrot.T590.Gui
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
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
        private void dataGridStudent_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            try
            {
                //FrameworkElement stud_ID = dataGridStudent.Columns[0].GetCellContent(e.Row);
                //if (stud_ID.GetType() == typeof(TextBox))
                //{
                //    _stud.StudentID = Convert.ToInt32(((TextBox)stud_ID).Text);
                //}

                //FrameworkElement stud_Name = dataGridStudent.Columns[1].GetCellContent(e.Row);
                //if (stud_Name.GetType() == typeof(TextBox))
                //{
                //    _stud.Name = ((TextBox)stud_Name).Text;
                //}

                //FrameworkElement stud_Email = dataGridStudent.Columns[2].GetCellContent(e.Row);
                //if (stud_Email.GetType() == typeof(TextBox))
                //{
                //    _stud.Email = ((TextBox)stud_Email).Text;
                //}

                //FrameworkElement stud_Class = dataGridStudent.Columns[3].GetCellContent(e.Row);
                //if (stud_Class.GetType() == typeof(TextBox))
                //{
                //    _stud.Class = ((TextBox)stud_Class).Text;
                //}

                //FrameworkElement stud_Address = dataGridStudent.Columns[4].GetCellContent(e.Row);
                //if (stud_Address.GetType() == typeof(TextBox))
                //{
                //    _stud.Address = ((TextBox)stud_Address).Text;
                //}

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>  
        /// Get entire Row  
        /// </summary>  
        /// <param name="sender"></param>  
        /// <param name="e"></param>  
        private void dataGridStudent_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            try
            {
                var isSave = _excel.StorePatientRecordInExcelAsync(_patient).Result;

                if (isSave)
                {
                    MessageBox.Show("Patient record saved successfully.");
                }
                else
                {
                    MessageBox.Show("Error problem occured.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace, ex.Message);
            }

        }

        /// <summary>  
        /// Get Record info to update  
        /// </summary>  
        /// <param name="sender"></param>  
        /// <param name="e"></param>  
        private void dataGridStudent_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _patient = DataGridPatient.SelectedItem as Patient;
        }

        private void GetPatientRecords()
        {
            _excel = new ExcelConnector(ExcelFilePath);

            try
            {
                DataGridPatient.ItemsSource = _excel.ReadAllPatientsFromExcelAsync().Result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

    }
}
