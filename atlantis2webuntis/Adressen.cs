using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;

namespace atlantis2webuntis
{
    public class Adressen : List<Adresse>
    {     
        public Adressen(string connectionStringAtlantis, string aktSjAtlantis)
        {
            try
            {
                Console.Write("Adressen ".PadRight(75,'.'));
                using (OdbcConnection connection = new OdbcConnection(connectionStringAtlantis))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"SELECT 
DBA.schue_sj.vorgang_schuljahr, 
DBA.schue_sj.pu_id AS IdAtlantis,
DBA.adresse.s_typ_adr AS Art, 
DBA.adresse.plz AS PLZ, 
DBA.adresse.strasse AS Strasse, 
DBA.adresse.ort AS Ort, 
DBA.adresse.email AS email, 
DBA.adresse.tel_1, 
DBA.adresse.s_anrede AS Anrede, 
DBA.adresse.sorge_berechtigt_jn AS SorgeberechtigtJn, 
DBA.adresse.tel_2, 
DBA.adresse.name_1 AS Nachname, 
DBA.adresse.name_2 AS Vorname
FROM DBA.schue_sj JOIN DBA.adresse ON DBA.schue_sj.pu_id = DBA.adresse.pu_id 
WHERE vorgang_schuljahr = '" + aktSjAtlantis + "';", connection);

                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.klasse");

                    foreach (DataRow theRow in dataSet.Tables["DBA.klasse"].Rows)
                    {
                        var adresse = new Adresse();
                        adresse.IdAtlantis = theRow["IdAtlantis"] == null ? -99 : Convert.ToInt32(theRow["IdAtlantis"]);
                        adresse.Art = theRow["Art"] == null ? "" : theRow["Art"].ToString();
                        adresse.Nachname = theRow["Nachname"] == null || theRow["Nachname"].ToString() == "" ? "NN" : theRow["Nachname"].ToString();
                        adresse.SorgeberechtigtJn = theRow["SorgeberechtigtJn"] == null ? "" : theRow["SorgeberechtigtJn"].ToString();
                        adresse.Anrede = theRow["Anrede"] == null ? "" : theRow["Anrede"].ToString();
                        adresse.Vorname = theRow["Vorname"] == null ? "" : theRow["Vorname"].ToString();
                        if (theRow["tel_1"] != null)
                            if (theRow["tel_1"].ToString() != "")
                                adresse.Telefons.Add(theRow["tel_1"].ToString());
                        if (theRow["tel_2"] != null)
                            if (theRow["tel_2"].ToString() != "")
                                adresse.Telefons.Add(theRow["tel_2"].ToString());
                        adresse.Plz = theRow["plz"] == null ? "" : theRow["plz"].ToString();
                        adresse.Ort = theRow["ort"] == null ? "" : theRow["ort"].ToString();
                        adresse.Strasse = theRow["strasse"] == null ? "" : theRow["strasse"].ToString();
                        adresse.Email = theRow["email"] == null ? "" : theRow["email"].ToString();

                       
                        this.Add(adresse);
                       
                    }
                    connection.Close();
                    Console.WriteLine((" " + this.Count.ToString()).PadLeft(30, '.'));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}