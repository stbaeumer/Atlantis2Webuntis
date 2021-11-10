using System;

namespace atlantis2webuntis
{
    public class Bemerkung
    {
        public string Kürzel { get; internal set; } // 'B' steht für beurlaubt. Beurlaubte Schüler werden nicht nach Untis exportiert
        public DateTime Von { get; internal set; }
        public DateTime Bis { get; internal set; }
        public int AtlantisSchuelerId { get; internal set; }
    }
}