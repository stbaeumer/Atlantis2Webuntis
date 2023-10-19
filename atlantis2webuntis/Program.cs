using Ionic.Zip;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace atlantis2webuntis
{
    class Program
    {
        public const string ConnectionStringAtlantis = @"Dsn=Atlantis17u;uid=DBA";
        public const string ConnectionStringUntis = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source=M:\\Data\\gpUntis.mdb;";
        public static string PfadZuAtlantisFotos = @"\\fs01\SoftwarehausHeider\Atlantis\Dokumente\jpg";
        public static string ImportdateiFürWebuntis = "Importdatei_für_Webuntis_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".csv";
        public static string MoodlekurseAusKlassenAnlegenPfad = "MoodlekurseAusKlassenAnlegen_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".csv";
        public static string SchülerInKlassenMoodlekurseEinschreiben = "SchülerInKlassenMoodlekurseEinschreiben_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".csv";

        static void Main(string[] args)
        {
            System.Net.ServicePointManager.ServerCertificateValidationCallback = ((sender, certificate, chain, sslPolicyErrors) => true);

            string aktuellerDatensatz = "";
            List<string> datensätze = new List<string>();

            try
            {
                Console.WriteLine("Atlantis2Webuntis (Version 20231019)");
                Console.WriteLine("====================================");
                Console.WriteLine("");
                
                int sj = (DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1);
                string aktSjUntis = sj.ToString() + (sj + 1);
                string aktSjAtlantis = sj.ToString() + "/" + (sj + 1 - 2000);

                Adressen adressen = new Adressen(ConnectionStringAtlantis, aktSjAtlantis);

                // MassMailSenden(adressen);
                
                Betriebe betriebe = new Betriebe(ConnectionStringAtlantis, aktSjAtlantis);
                Bemerkungen bemerkungen = new Bemerkungen(ConnectionStringAtlantis, aktSjAtlantis);
                Periodes periodes = new Periodes();
                var periode = (from p in periodes where p.Bis >= DateTime.Now.Date where DateTime.Now.Date >= p.Von select p.IdUntis).FirstOrDefault();
                Klasses klasses = new Klasses(periode);
                Schuelers schuelers = new Schuelers(ConnectionStringAtlantis, betriebe, adressen, klasses, bemerkungen);
                Schuelers untisSchuelers = new Schuelers();

                Console.WriteLine("");
                Console.WriteLine("Folgende Schülerinnen und Schüler existieren in Untis, aber nicht in Atlantis. Sie bekommen in Untis");
                Console.WriteLine("automatisch ein Austrittsdatum gesetzt:");

                Console.WriteLine("");

                int i = 1;

                foreach (var schueler in untisSchuelers)
                {
                    if (!(from s in schuelers where s.Id == schueler.Id select s).Any())
                    {
                        // Wenn bereits ein Austrittsdatum in der Vergangenheit gesetzt ist, wird das nicht geändert

                        if (!(schueler.Austrittsdatum < DateTime.Now))
                        {
                            if ((from k in klasses where k.NameUntis == schueler.Grade select k).Any())
                            {
                                Console.WriteLine(i.ToString().PadLeft(5) + ". " + schueler.Kurzname.PadRight(20) + " Austrittsdatum: " + DateTime.Now.ToShortDateString().PadRight(10) + " bisherige Klasse:" + schueler.Grade);
                                schueler.Austrittsdatum = DateTime.Now;
                                schueler.Grade = "";
                                schuelers.Add(schueler);
                            }
                        }
                    }
                }

                Console.WriteLine("");
                schuelers.GeneriereImportdateiFürWebuntis(ImportdateiFürWebuntis);
                //schuelers.MoodlekurseAusKlassenAnlegen(MoodlekurseAusKlassenAnlegenPfad);
                //schuelers.SchülerInKlassenMoodlekurseEinschreiben(SchülerInKlassenMoodlekurseEinschreiben);
                schuelers.ZippeBilder(PfadZuAtlantisFotos);
                
                try
                {
                    Process.Start(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                try
                {
                    System.Diagnostics.Process.Start(ImportdateiFürWebuntis);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }


                Console.WriteLine("");
                Console.WriteLine("!!!!  Das Stichtagsdatum für die Klassenzugehörigkeit muss auf das Tagesdatum gesetzt werden !!!!");
                Console.WriteLine("==================================================================================================");
                Console.WriteLine("");
                Console.WriteLine("");
                Console.WriteLine("");
                Console.WriteLine("Vorgehen:");
                Console.WriteLine("=========");
                Console.WriteLine("1. Stammdaten > Schüler*innen > Import");
                Console.WriteLine("2. Datei " + ImportdateiFürWebuntis + " auswählen. Dann 'Weiter' klicken.");
                Console.WriteLine("3. Profil 'Schuelerimport' auswählen.");
                Console.WriteLine("4. Administration > Benutzer > Benutzerverwaltung > 'Benutzer für Schüler*innen anlegen'. Dann 'Benutzer erstellen'");
                Console.WriteLine("5. Administration > Benutzer > Benutzerverwaltung > Import");
                Console.WriteLine("6. Datei " + ImportdateiFürWebuntis + " auswählen. Dann 'Weiter' klicken. ");
                Console.WriteLine("7. O365-Profil wählen. 'Import' klicken. Jetzt sollte die fehlende O365-Identität hinzugefügt worden sein.");


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                MailSenden("Fehler in atlantis2webuntis", ex.ToString() + "Fehler beim aktuellen Datensatz: " + aktuellerDatensatz);
            }
            finally
            {
                Console.WriteLine("ANYKEY klicken, um zu beenden.");
                Console.ReadKey();
            }
        }

        private static void MassMailSenden(Adressen adressen)
        {
            ExchangeService exchangeService = new ExchangeService();

            exchangeService.UseDefaultCredentials = true;
            exchangeService.TraceEnabled = false;
            exchangeService.TraceFlags = TraceFlags.All;
            exchangeService.Url = new Uri("https://ex01.bkb.local/EWS/Exchange.asmx");

            int ss = 1;
            foreach (var item in (from f in adressen where f.Email != null where f.Email != "" where f.Art == "Schüler" select f).ToList())
            {
                if (ss > 414)
                {
                    try
                    {
                        EmailMessage message = new EmailMessage(exchangeService);

                        //message.ToRecipients.Add("stefan.baeumer@berufskolleg-borken.de");
                        message.ToRecipients.Add(item.Email);
                        message.Subject = "Aufgrund der Sturmwarnung bleibt das Berufskolleg Borken am 10.02.20 geschlossen";

                        message.Body = @"Guten Tag " + item.Vorname + " " + item.Nachname + ", " +
                            "<br /><br />aufgrund der Sturmwarnung fällt für alle Schülerinnen und Schüler des Berufskollegs Borken der Unterricht am 10.02.20 aus. Schülerinnen und Schüler in einer dualen Ausbildung stehen an diesem Tag ihrem Ausbildungsbetrieb zur Verfügung." +
                            "<br /><br />Bitte schreiben Sie das auch in Ihre Klassengruppe, damit wirklich alle die Nachricht bekommen." +
                            "<br />" +
                            "<br />" +
                            "Aktuelle Informationen finden Sie auf der Homepage." +
                            "<br />" +
                            "<br />" +
                            "<br />Freundliche Grüße" +
                            "<br />" +
                            "<br />" +
                            "<br />i. V. Stefan Bäumer" +
                            "<br />Stellvertretender Schulleiter" +
                            "<br />Berufskolleg Borken" +
                            "<br />Josefstraße 10" +
                            "<br />46325 Borken" +
                            "<br />fon + 49(0) 2861 90990-0" +
                            "<br />fax + 49(0) 2861 90990-55" +
                            "<br />e-mail stefan.baeumer@berufskolleg-borken.de";

                        message.SendAndSaveCopy();
                        Console.WriteLine(ss.ToString().PadLeft(5) + " " + message.Body.ToString().Substring(0, 80) + " ... per Mail gesendet.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ss.ToString().PadLeft(5) + ex.ToString());
                    }
                    finally
                    {
                        Thread.Sleep(1000);
                    }    
                }                
                
                ss++;
            }            
        }

        internal static void MailSenden(string subject, string body)
        {
            ExchangeService exchangeService = new ExchangeService();

            exchangeService.UseDefaultCredentials = true;
            exchangeService.TraceEnabled = false;
            exchangeService.TraceFlags = TraceFlags.All;
            exchangeService.Url = new Uri("https://ex01.bkb.local/EWS/Exchange.asmx");

            EmailMessage message = new EmailMessage(exchangeService);

            message.ToRecipients.Add("stefan.baeumer@berufskolleg-borken.de");

            message.Subject = subject;

            message.Body = body;

            message.SendAndSaveCopy();
            Console.WriteLine(subject + " ... per Mail gesendet.");
        }
    }
}