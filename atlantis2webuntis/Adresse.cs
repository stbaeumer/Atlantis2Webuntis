using System.Collections.Generic;

namespace atlantis2webuntis
{
    public class Adresse
    {
        private string art;

        public Adresse()
        {
            Telefons = new List<string>();
        }

        public int IdAtlantis { get; set; }
        /// <summary>
        /// Vater, Mutter , Festnetznummer, Mobilnummer, Schüler
        /// </summary>
        public string Art
        {
            get
            {
                return this.art;
            }
            set
            {
                this.art = value;
                if (value == "V")
                    this.art = "Vater";
                if (value == "M")
                    this.art = "Mutter";
                if (value == "E")
                    this.art = "Eltern";
                if (value == "0")
                    this.art = "Schüler";
                if (value == "T")
                    this.art = "Sonstige";
            }
        }
        public string Nachname { get; internal set; }
        public string Vorname { get; internal set; }
        public List<string> Telefons { get; internal set; }
        public string Email { get; internal set; }
        public string Strasse { get; internal set; }
        public string Plz { get; internal set; }
        public string Ort { get; internal set; }
        public string Anrede { get; internal set; }
        public string SorgeberechtigtJn { get; internal set; }
    }
}