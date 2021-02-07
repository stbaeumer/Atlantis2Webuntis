using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;

namespace atlantis2webuntis
{
    public class Betriebe : List<Betrieb>
    {
        private string connectionStringAtlantis;

        public Betriebe(string connectionStringAtlantis, string aktSjAtlantis)
        {
            try
            {
                using (OdbcConnection connection = new OdbcConnection(connectionStringAtlantis))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"SELECT 
DBA.schue_sj.pu_id AS idAtlantis,
DBA.schue_sj.vorgang_schuljahr,
DBA.adresse.ad_id,
DBA.adresse.name_1 AS Name,
DBA.adresse.strasse AS Strasse,
DBA.adresse.lkz AS LandKürzel,
DBA.adresse.plz AS PLZ,
DBA.adresse.ort AS Ort,
DBA.adresse.tel_1 AS Telefon,
DBA.adresse.www AS WWW,
DBA.adresse.email AS Mail,
DBA.adresse.plz_postfach AS PLZPostfach,
DBA.adresse.postfach AS Postfach,
DBA.pj_bt.dat_von AS AusbildungsVertragsBeginn,
DBA.pj_bt.dat_bis AS AusbildungsVertragsEnde
FROM(DBA.adresse JOIN DBA.pj_bt ON DBA.adresse.bt_id = DBA.pj_bt.bt_id) JOIN DBA.schue_sj ON DBA.pj_bt.pj_id = DBA.schue_sj.pj_id
WHERE vorgang_schuljahr = '" + aktSjAtlantis + "';", connection);

                    Console.Write("Betriebe ".PadRight(75,'.'));

                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.klasse");

                    foreach (DataRow theRow in dataSet.Tables["DBA.klasse"].Rows)
                    {
                        Betrieb betrieb = new Betrieb()
                        {
                            SchuelerIdAtlantis = theRow["idAtlantis"] == null ? -99 : Convert.ToInt32(theRow["idAtlantis"]),
                            Adressart = "Betrieb",
                            Name = theRow["Name"] == null ? "" : theRow["Name"].ToString(),
                            Straße = theRow["Strasse"] == null ? "" : theRow["Strasse"].ToString(),
                            PLZ = theRow["PLZ"] == null ? "" : theRow["PLZ"].ToString(),
                            Ort = theRow["Ort"] == null ? "" : theRow["Ort"].ToString(),
                            Telefon = theRow["Telefon"] == null ? "" : theRow["Telefon"].ToString(),
                            EMail = theRow["Mail"] == null ? "" : theRow["Mail"].ToString(),
                            WWW = theRow["WWW"] == null ? "" : theRow["WWW"].ToString(),
                            Postfach = theRow["Postfach"] == null ? "" : theRow["Postfach"].ToString(),
                            LandKürzel = theRow["LandKürzel"] == null ? "" : theRow["LandKürzel"].ToString(),
                            PostfachPlz = theRow["PLZPostfach"] == null ? "" : theRow["PLZPostfach"].ToString(),
                            AusbildungsVertragsBeginn = theRow["AusbildungsVertragsBeginn"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["AusbildungsVertragsBeginn"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture),
                            AusbildungsVertragsEnde = theRow["AusbildungsVertragsEnde"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["AusbildungsVertragsEnde"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture),
                        };

                        // Nur wenn jetzt innerhalb der Ausbildung liegt, wird der Betrieb hinzugefügt.
                        if (betrieb.AusbildungsVertragsBeginn < DateTime.Now && DateTime.Now < betrieb.AusbildungsVertragsEnde)
                        {
                            this.Add(betrieb);
                        }

                    }
                    connection.Close();
                    Console.WriteLine((" " + this.Count.ToString()).PadLeft(30,'.'));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}