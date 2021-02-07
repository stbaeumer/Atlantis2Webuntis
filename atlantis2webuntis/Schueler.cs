using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace atlantis2webuntis
{
    public class Schueler
    {
        public string Status { get; internal set; }
        public int Bezugsjahr { get; internal set; }

        public int Id { get; set; }
        public string Name { get; set; }
        public string Firstname { get; set; }
        public string Grade { get; set; }
        public DateTime Birthday { get; set; }
        public string ImagePath { get; set; }
        public int MyProperty { get; set; }
        public string Telefon { get; set; }
        public string Mail { get; set; }
        public string Kurzname { get; set; }
        public string Geburtsdatum { get; set; }
        public DateTime Eintrittsdatum { get; set; }
        public DateTime Austrittsdatum { get; set; }
        public string Geschlecht { get; set; }
        public string Mobil { get; set; }
        public string Strasse { get; set; }
        public string Plz { get; set; }
        public string Ort { get; set; }
        public string ErzMobil { get; set; }
        public string ErzTelefon { get; set; }
        public bool Volljährig { get; set; }
        public string ErzName { get; set; }
        public string BetriebName { get; set; }
        public string BetriebStrasse { get; set; }
        public string BetriebPlz { get; set; }
        public string BetriebOrt { get; set; }
        public string BetriebTelefon { get; set; }
        public string Geschlecht34 { get; internal set; }
        public string AktuellJN { get; internal set; }
        public Betrieb Betrieb { get; internal set; }
        public List<Adresse> Adressen { get; internal set; }
        public DateTime Relianmeldung { get; internal set; }
        public DateTime Reliabmeldung { get; internal set; }

        private string istVolljährig(string pDatum)
        {
            if (pDatum != "")
            {
                DateTime birthday = DateTime.ParseExact(pDatum, "dd.MM.yyyy", CultureInfo.InvariantCulture);

                int years = DateTime.Now.Year - birthday.Year;
                birthday = birthday.AddYears(years);
                if (DateTime.Now.CompareTo(birthday) < 0) { years--; }
                if (years < 18)
                {
                    return "0";
                }
            }

            return "1";
        }

        private string datumUmwandeln(string pDatum)
        {
            try
            {
                pDatum = pDatum.Replace("\"", "");

                DateTime dt = DateTime.ParseExact(pDatum, "yyyy-MM-dd hh:mm:ss", CultureInfo.InvariantCulture);
                return dt.ToString("dd.MM.yyyy");
            }
            catch (Exception)
            {
                return "";
            }
        }

        internal string SchreibeDatensatz()
        {
            return string.Format(
                            Id + ";" +
                            Kurzname + "@students.berufskolleg-borken.de" + ";" +
                            Name + ";" +
                            Firstname + ";" +
                            Grade + ";" +
                            Kurzname + ";" +
                            Geschlecht.ToUpper() + ";" +
                            Birthday.ToString("dd.MM.yyyy") + ";" +
                            /* Wenn das Das Eintrittsdatum leer ist, wird der Wert aus dem Importdialog genommen.*/
                            "" + ";" +
                            Austrittsdatum.ToString("dd.MM.yyyy") + ";" +
                            Telefon + ";" +
                            (Adressen == null ? "" : Adressen.Count >= 1 ? Adressen[0].Telefons.Count >= 1 ? Adressen[0].Telefons[0] : "" : "") + ";" +
                            (Adressen == null ? "" : Adressen.Count >= 1 ? Adressen[0].Strasse : "") + ";" +
                            (Adressen == null ? "" : Adressen.Count >= 1 ? Adressen[0].Plz : "") + ";" +
                            (Adressen == null ? "" : Adressen.Count >= 1 ? Adressen[0].Ort : "") + ";" +
                            (Adressen == null ? "" : Adressen.Count >= 2 ? Adressen[1].Vorname + " " + Adressen[1].Nachname : "") + ";" +
                            (Adressen == null ? "" : Adressen.Count >= 2 ? Adressen[1].Telefons.Count >= 1 ? Adressen[1].Telefons[0] : "" : "") + ";" +
                            (Adressen == null ? "" : Adressen.Count >= 2 ? Adressen[1].Telefons.Count >= 2 ? Adressen[1].Telefons[1] : "" : "") + ";" +
                            (Volljährig ? "1" : "0") + ";" +
                            (Betrieb != null ? Betrieb.Name : "") + ";" +
                            (Betrieb != null ? Betrieb.Straße : "") + ";" +
                            (Betrieb != null ? Betrieb.PLZ : "") + ";" +
                            (Betrieb != null ? Betrieb.Ort : "") + ";" +
                            (Betrieb != null ? Betrieb.Telefon : "") + ";" +
                            Kurzname + "@students.berufskolleg-borken.de" + ";" +
                            Kurzname
                            );
        }

        string FBereinigen(string Textinput)
        {
            string Text = Textinput;

            Text = Text.ToLower();                          // Nur Kleinbuchstaben
            Text = FUmlauteBehandeln(Text);                 // Umlaute ersetzen


            Text = Regex.Replace(Text, "-", "_");           //  kein Minus-Zeichen
            Text = Regex.Replace(Text, ",", "_");           //  kein Komma            
            Text = Regex.Replace(Text, " ", "_");           //  kein Leerzeichen
            // Text = Regex.Replace(Text, @"[^\w]", string.Empty);   // nur Buchstaben

            Text = Regex.Replace(Text, "[^a-z]", string.Empty);   // nur Buchstaben

            Text = Text.Substring(0, 1);  // Auf maximal 6 Zeichen begrenzen
            return Text;
        }

        string FUmlauteBehandeln(string Textinput)
        {
            string Text = Textinput;

            // deutsche Sonderzeichen
            Text = Regex.Replace(Text, "[æ|ä]", "ae");
            Text = Regex.Replace(Text, "[Æ|Ä]", "Ae");
            Text = Regex.Replace(Text, "[œ|ö]", "oe");
            Text = Regex.Replace(Text, "[Œ|Ö]", "Oe");
            Text = Regex.Replace(Text, "[ü]", "ue");
            Text = Regex.Replace(Text, "[Ü]", "Ue");
            Text = Regex.Replace(Text, "ß", "ss");

            // Sonderzeichen aus anderen Sprachen
            Text = Regex.Replace(Text, "[ã|à|â|á|å]", "a");
            Text = Regex.Replace(Text, "[Ã|À|Â|Á|Å]", "A");
            Text = Regex.Replace(Text, "[é|è|ê|ë]", "e");
            Text = Regex.Replace(Text, "[É|È|Ê|Ë]", "E");
            Text = Regex.Replace(Text, "[í|ì|î|ï]", "i");
            Text = Regex.Replace(Text, "[Í|Ì|Î|Ï]", "I");
            Text = Regex.Replace(Text, "[õ|ò|ó|ô]", "o");
            Text = Regex.Replace(Text, "[Õ|Ó|Ò|Ô]", "O");
            Text = Regex.Replace(Text, "[ù|ú|û|µ]", "u");
            Text = Regex.Replace(Text, "[Ú|Ù|Û]", "U");
            Text = Regex.Replace(Text, "[ý|ÿ]", "y");
            Text = Regex.Replace(Text, "[Ý]", "Y");
            Text = Regex.Replace(Text, "[ç|č]", "c");
            Text = Regex.Replace(Text, "[Ç|Č]", "C");
            Text = Regex.Replace(Text, "[Ð]", "D");
            Text = Regex.Replace(Text, "[ñ]", "n");
            Text = Regex.Replace(Text, "[Ñ]", "N");
            Text = Regex.Replace(Text, "[š]", "s");
            Text = Regex.Replace(Text, "[Š]", "S");

            return Text;

        }

        public string FindAndRenamePicture()
        {
            return "";
        }

        public string GenerateIdFromBarcode(string pBarcode)
        {
            pBarcode = Regex.Replace(pBarcode, "[^0-9]", string.Empty);
            return pBarcode.Substring(0, Math.Min(6, pBarcode.Length)).ToString();  // Längenbegrenzung auf 6 Zeichen
        }

        public string GenerateMail(string pMail)
        {
            if (IsValidEmail(pMail))
            {
                return pMail;
            }
            return FBereinigen(Name) + "." + FBereinigen(Firstname) + "." + Id + "@bkb.krbor.de";
        }

        bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        public string generateKurzname()
        {
            return FBereinigen(Name) + FBereinigen(Firstname) + Id;
        }
    }
}
