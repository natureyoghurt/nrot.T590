using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.OleDb;
using System.Text;
using System.Threading.Tasks;
using nrot.T590.Excel.Models;

namespace nrot.T590.Excel
{
    // http://www.c-sharpcorner.com/UploadFile/rahul4_saxena/read-write-and-update-an-excel-file-in-wpf/

    public class ExcelConnector
    {
        private readonly string _excelFilePath;
        private OleDbConnection conn;

        public ExcelConnector(string excelFilePath)
        {
            _excelFilePath = excelFilePath ?? throw new ArgumentNullException(nameof(excelFilePath));
            conn = new OleDbConnection(GetConnectionString());
        }
        private string GetConnectionString()
        {
            var props = new Dictionary<string, string>
            {
                ["Provider"] = "Microsoft.ACE.OLEDB.12.0;",
                ["Extended Properties"] = "Excel 12.0 XML",
                ["Data Source"] = _excelFilePath
            };

            // XLS - Excel 2003 and Older
            //props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
            //props["Extended Properties"] = "Excel 8.0";
            //props["Data Source"] = "C:\\MyExcel.xls";

            var sb = new StringBuilder();

            foreach (var prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        public async Task<ObservableCollection<Patient>> ReadAllPatientsFromExcelAsync()
        {
            var patients = new ObservableCollection<Patient>();

            await conn.OpenAsync();

            var cmd = new OleDbCommand();
            cmd.Connection = conn;
            cmd.CommandText = "SELECT * FROM [Patienten$]";

            var reader = await cmd.ExecuteReaderAsync();

            while (reader.Read())
            {
                patients.Add(new Patient
                {
                    Id = Convert.ToInt32(reader["Id"]),
                    Name = reader["Name"].ToString(),
                    Vorname = reader["Vorname"].ToString(),
                    Strasse = reader["Strasse"].ToString(),
                    Plz = Convert.ToInt32(reader["Plz"]),
                    Ort = reader["Ort"].ToString(),
                    Geburtsdatum = Convert.ToDateTime(reader["Geburtsdatum"]),
                    Geschlecht = Convert.ToChar(reader["Geschlecht"]).Equals('m') ? GeschlechtType.M : GeschlechtType.W,
                    PatientenNr = reader["PatientenNr"].ToString(),
                    AhvNr = reader["AhvNr"].ToString(),
                    VekaNr = reader["VekaNr"].ToString(),
                    VersichertenNr = reader["VersichertenNr"].ToString(),
                    Kanton = reader["Kanton"].ToString(),
                    Kopie = Convert.ToBoolean(reader["Kopie"]),
                    VerguetungsArt = CastVerguetungsArt(reader["VerguetungsArt"]),
                    VertragsNr = reader["VertragsNr"].ToString()
                });
            }

            reader.Close();
            conn.Close();

            return patients;
        }

        public async Task<bool> StorePatientRecordInExcelAsync(Patient patient)
        {
            var isSave = false;

            if (patient.Id == 0)
            {
                patient.Id = await GetNextPatientId();
            }

            var cmd = new OleDbCommand
            {
                Connection = conn,
                CommandText = !IsPatientRecordExistingAsync(patient).Result
                    ? "INSERT INTO [Patienten$] VALUES (@Id, @Name, @Vorname, @Strasse, @Plz, @Ort, @Geburtsdatum, @Geschlecht, @PatientenNr, @AhvNr, @VekaNr, @VersichertenNr, @Kanton, @Kopie, @VerguetungsArt, @VertragsNr)"
                    : "UPDATE [Patienten$] SET Name=@Name, Vorname=@Vorname, Strasse=@Strasse, Plz=@Plz, Ort=@Ort, Geburtsdatum=@Geburtsdatum, Geschlecht=@Geschlecht, PatientenNr=@PatientenNr, AhvNr=@AhvNr, VekaNr=@VekaNr, VersichertenNr=@VersichertenNr, Kanton=@Kanton, Kopie=@Kopie, VerguetungsArt=@VerguetungsArt, VertragsNr=@VertragsNr WHERE Id=@Id"
            };

            cmd.Parameters.AddWithValue("@Id", patient.Id);
            cmd.Parameters.AddWithValue("@Name", patient.Name);
            cmd.Parameters.AddWithValue("@Vorname", patient.Vorname);
            cmd.Parameters.AddWithValue("@Strasse", patient.Strasse);
            cmd.Parameters.AddWithValue("@Plz", patient.Plz);
            cmd.Parameters.AddWithValue("@Ort", patient.Ort);
            cmd.Parameters.AddWithValue("@Geburtsdatum", patient.Geburtsdatum);
            cmd.Parameters.AddWithValue("@Geschlecht", patient.Geschlecht);
            cmd.Parameters.AddWithValue("@PatientenNr", patient.PatientenNr);
            cmd.Parameters.AddWithValue("@AhvNr", patient.AhvNr);
            cmd.Parameters.AddWithValue("@VekaNr", patient.VekaNr);
            cmd.Parameters.AddWithValue("@VersichertenNr", patient.VersichertenNr);
            cmd.Parameters.AddWithValue("@Kanton", patient.Kanton);
            cmd.Parameters.AddWithValue("@Kopie", patient.Kopie);
            cmd.Parameters.AddWithValue("@VerguetungsArt", patient.VerguetungsArt);
            cmd.Parameters.AddWithValue("@VertragsNr", patient.VertragsNr);

            await conn.OpenAsync();

            if (await cmd.ExecuteNonQueryAsync() > 0)
            {
                isSave = true;
            }

            conn.Close();

            return isSave;
        }

        private async Task<int> GetNextPatientId()
        {
            await conn.OpenAsync();

            var cmd = new OleDbCommand
            {
                Connection = conn,
                CommandText = "SELECT MAX(Id) FROM [Patienten$]"
            };

            var maxId = Convert.ToInt32(await cmd.ExecuteScalarAsync());

            conn.Close();

            return maxId + 1;
        }

        private async Task<bool> IsPatientRecordExistingAsync(Patient patient)
        {
            var isRecordExisting = false;

            await conn.OpenAsync();

            var cmd = new OleDbCommand
            {
                Connection = conn,
                CommandText = $"SELECT *  FROM [Patienten$] WHERE Id = {patient.Id}"
            };

            var reader = await cmd.ExecuteReaderAsync();

            if (reader.HasRows)
            {
                isRecordExisting = true;
            }

            reader.Close();
            conn.Close();

            return isRecordExisting;
        }

        private static VerguetungsartType CastVerguetungsArt(object o)
        {
            if (!(o is string))
            {
                // TODO: exception message
                throw new Exception($"'{o}' is not a valid VerguetungsArt!");
            }

            var va = o.ToString();

            switch (va)
            {
                case "TG":
                    return VerguetungsartType.Tg;
                case "TP":
                    return VerguetungsartType.Tp;
                default:
                    // TODO: exception message
                    throw new ArgumentOutOfRangeException();
            }
        }
    }
}
