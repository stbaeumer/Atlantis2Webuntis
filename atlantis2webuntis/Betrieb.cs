using System;

namespace atlantis2webuntis
{
    public class Betrieb
    {
        public string Adressart { get; internal set; }
        public string Name { get; internal set; }
        public string Straße { get; internal set; }
        public string PLZ { get; internal set; }
        public string Ort { get; internal set; }
        public string Telefon { get; internal set; }
        public string EMail { get; internal set; }
        public string WWW { get; internal set; }
        public string Postfach { get; internal set; }
        public string LandKürzel { get; internal set; }
        public string PostfachPlz { get; internal set; }
        public int SchuelerIdAtlantis { get; internal set; }
        public DateTime AusbildungsVertragsBeginn { get; internal set; }
        public DateTime AusbildungsVertragsEnde { get; internal set; }
    }
}