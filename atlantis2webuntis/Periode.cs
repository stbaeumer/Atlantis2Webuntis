using System;

namespace atlantis2webuntis
{
    public class Periode
    {
        public int IdUntis { get; internal set; }
        public string Name { get; internal set; }
        public string Langname { get; internal set; }
        public DateTime Von { get; internal set; }
        public DateTime Bis { get; internal set; }
    }
}