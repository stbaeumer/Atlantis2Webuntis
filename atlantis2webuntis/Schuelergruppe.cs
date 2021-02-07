using System;
using System.Collections.Generic;

namespace atlantis2webuntis
{
    public class Schuelergruppe
    {
        public string Name { get; internal set; }
        public int FachId { get; internal set; }
        public string KlasseIds { get; internal set; }
        public List<Schueler> Schuelers { get; internal set; }
        public int UnterrichtsId { get; internal set; }
        public string Fachkürzel { get; internal set; }        
    }
}