﻿using System;
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
            try
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

                return patients;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error occured while reading patient records from database.", ex);
            }
            finally
            {
                conn.Close();
            }
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
                    //? "INSERT INTO [Patienten$] VALUES (@Id, @Name, @Vorname, @Strasse, @Plz, @Ort, @Geburtsdatum, @Geschlecht, @PatientenNr, @AhvNr, @VekaNr, @VersichertenNr, @Kanton, @Kopie, @VerguetungsArt, @VertragsNr)"
                    //? "INSERT INTO [Patienten$](Id, Name, Vorname, Strasse, Plz, Ort, Geburtsdatum, Geschlecht, PatientenNr, AhvNr, VekaNr, VersichertenNr, Kanton, Kopie, VerguetungsArt, VertragsNr) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                    ? "INSERT INTO [Patienten$](Id, Name, Vorname, Strasse, Plz, Ort, Geschlecht, PatientenNr, AhvNr, VekaNr, VersichertenNr, Kanton, Kopie, VerguetungsArt, VertragsNr) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                    //: "UPDATE [Patienten$] SET Name=@Name, Vorname=@Vorname, Strasse=@Strasse, Plz=@Plz, Ort=@Ort, Geburtsdatum=@Geburtsdatum, Geschlecht=@Geschlecht, PatientenNr=@PatientenNr, AhvNr=@AhvNr, VekaNr=@VekaNr, VersichertenNr=@VersichertenNr, Kanton=@Kanton, Kopie=@Kopie, VerguetungsArt=@VerguetungsArt, VertragsNr=@VertragsNr WHERE Id=@Id"
                    //: "UPDATE [Patienten$] SET Name=?, Vorname=?, Strasse=?, Plz=?, Ort=?, Geburtsdatum=?, Geschlecht=?, PatientenNr=?, AhvNr=?, VekaNr=?, VersichertenNr=?, Kanton=?, Kopie=?, VerguetungsArt=?, VertragsNr=? WHERE Id=?"
                    : "UPDATE [Patienten$] SET Name=?, Vorname=?, Strasse=?, Plz=?, Ort=?, Geschlecht=?, PatientenNr=?, AhvNr=?, VekaNr=?, VersichertenNr=?, Kanton=?, Kopie=?, VerguetungsArt=?, VertragsNr=? WHERE Id=?"
            };

            //cmd.Parameters.AddWithValue("Id", patient.Id);
            cmd.Parameters.Add("Id", OleDbType.Integer).Value = patient.Id;
            cmd.Parameters["Id"].IsNullable = false;
            //cmd.Parameters.AddWithValue("Name", patient.Name);
            cmd.Parameters.Add("Name", OleDbType.VarChar).Value = patient.Name;
            cmd.Parameters["Name"].IsNullable = false;
            //cmd.Parameters.AddWithValue("Vorname", patient.Vorname);
            cmd.Parameters.Add("Vorname", OleDbType.VarChar).Value = patient.Vorname;
            cmd.Parameters["Vorname"].IsNullable = false;
            //cmd.Parameters.AddWithValue("Strasse", patient.Strasse);
            cmd.Parameters.Add("Strasse", OleDbType.VarChar).Value = patient.Strasse;
            cmd.Parameters["Strasse"].IsNullable = false;
            //cmd.Parameters.AddWithValue("Plz", patient.Plz);
            cmd.Parameters.Add("Plz", OleDbType.Integer).Value = patient.Plz;
            cmd.Parameters["Plz"].IsNullable = false;
            //cmd.Parameters.AddWithValue("Ort", patient.Ort);
            cmd.Parameters.Add("Ort", OleDbType.VarChar).Value = patient.Ort;
            cmd.Parameters["Ort"].IsNullable = false;
            //cmd.Parameters.AddWithValue("Geburtsdatum", patient.Geburtsdatum.ToString("mm/dd/yyyy"));
            //cmd.Parameters.Add("Geburtsdatum", OleDbType.VarChar).Value = patient.Geburtsdatum.ToString("mm/dd/yyyy");
            //cmd.Parameters["Geburtsdatum"].IsNullable = true;
            //cmd.Parameters.AddWithValue("Geschlecht", patient.Geschlecht.ToString());
            cmd.Parameters.Add("Geschlecht", OleDbType.VarChar).Value = patient.Geschlecht.ToString();
            cmd.Parameters["Geschlecht"].IsNullable = false;
            //cmd.Parameters.AddWithValue("PatientenNr", patient.PatientenNr);
            cmd.Parameters.Add("PatientenNr", OleDbType.VarChar).Value = patient.PatientenNr;
            cmd.Parameters["PatientenNr"].IsNullable = true;
            //cmd.Parameters.AddWithValue("AhvNr", patient.AhvNr);
            cmd.Parameters.Add("AhvNr", OleDbType.VarChar).Value = patient.AhvNr;
            cmd.Parameters["AhvNr"].IsNullable = true;
            //cmd.Parameters.AddWithValue("VekaNr", patient.VekaNr);
            cmd.Parameters.Add("VekaNr", OleDbType.VarChar).Value = patient.VekaNr;
            cmd.Parameters["VekaNr"].IsNullable = true;
            //cmd.Parameters.AddWithValue("VersichertenNr", patient.VersichertenNr);
            cmd.Parameters.Add("VersichertenNr", OleDbType.VarChar).Value = patient.VersichertenNr;
            cmd.Parameters["VersichertenNr"].IsNullable = true;
            //cmd.Parameters.AddWithValue("Kanton", patient.Kanton);
            cmd.Parameters.Add("Kanton", OleDbType.VarChar).Value = patient.Kanton;
            cmd.Parameters["Kanton"].IsNullable = true;
            //cmd.Parameters.AddWithValue("Kopie", patient.Kopie);
            cmd.Parameters.Add("Kopie", OleDbType.Boolean).Value = patient.Kopie;
            cmd.Parameters["Kopie"].IsNullable = false;
            //cmd.Parameters.AddWithValue("VerguetungsArt", patient.VerguetungsArt.ToString());
            cmd.Parameters.Add("VerguetungsArt", OleDbType.VarChar).Value = patient.VerguetungsArt.ToString();
            cmd.Parameters["VerguetungsArt"].IsNullable = true;
            //cmd.Parameters.AddWithValue("VertragsNr", patient.VertragsNr);
            cmd.Parameters.Add("VertragsNr", OleDbType.VarChar).Value = patient.VertragsNr;
            cmd.Parameters["VertragsNr"].IsNullable = true;

            try
            {
                await conn.OpenAsync();

                if (await cmd.ExecuteNonQueryAsync() > 0)
                {
                    isSave = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error occured while inserting/updating patient record.", ex);
            }
            finally
            {
                if (conn.State == System.Data.ConnectionState.Open)
                {
                    conn.Close();
                }
            }

            return isSave;
        }

        private async Task<int> GetNextPatientId()
        {
            try
            {
                await conn.OpenAsync();

                var cmd = new OleDbCommand
                {
                    Connection = conn,
                    CommandText = "SELECT MAX(Id) FROM [Patienten$]"
                };

                var res = await cmd.ExecuteScalarAsync();
                var maxId = res == DBNull.Value ? 0 : Convert.ToInt32(res);

                return maxId + 1;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error occured while searching next patients id.", ex);
            }
            finally
            {
                if (conn.State == System.Data.ConnectionState.Open)
                {
                    conn.Close();
                }
            }
        }

        private async Task<bool> IsPatientRecordExistingAsync(Patient patient)
        {
            var isRecordExisting = false;

            try
            {
                await conn.OpenAsync();

                var cmd = new OleDbCommand
                {
                    Connection = conn,
                    CommandText = $"SELECT * FROM [Patienten$] WHERE Id=?"
                };

                cmd.Parameters.Add("Id", OleDbType.Integer).Value = patient.Id;

                var reader = await cmd.ExecuteReaderAsync();

                if (reader.HasRows)
                {
                    isRecordExisting = true;
                }

                reader.Close();

                return isRecordExisting;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error occured while checking existence of patient with id = '{patient.Id}'.", ex);
            }
            finally
            {
                if (conn.State == System.Data.ConnectionState.Open)
                {
                    conn.Close();
                }
            }
        }

        private static VerguetungsartType CastVerguetungsArt(object o)
        {
            if (!(o is string))
            {
                // TODO: exception message
                throw new Exception($"'{o}' is not a valid VerguetungsArt!");
            }

            var va = o.ToString();

            switch (va.ToUpper())
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
