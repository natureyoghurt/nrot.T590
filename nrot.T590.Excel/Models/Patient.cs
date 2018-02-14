using System;

namespace nrot.T590.Excel.Models
{
    public class Patient
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Vorname { get; set; }
        public string Strasse { get; set; }
        public int Plz { get; set; }
        public string Ort { get; set; }
        public DateTime? Geburtsdatum { get; set; }
        public GeschlechtType Geschlecht { get; set; }
        public string PatientenNr { get; set; }
        public string AhvNr { get; set; }
        public string VekaNr { get; set; }
        public string VersichertenNr { get; set; }
        public string Kanton { get; set; }
        public bool Kopie { get; set; }
        public VerguetungsartType VerguetungsArt { get; set; }
        public string VertragsNr { get; set; }
    }
}