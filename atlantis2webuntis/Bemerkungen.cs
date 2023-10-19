using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;

namespace atlantis2webuntis
{
    public class Bemerkungen : List<Bemerkung>
    {
        private string connectionStringAtlantis;

        public Bemerkungen(string connectionStringAtlantis, string aktSjAtlantis)
        {
            try
            {
                using (OdbcConnection connection = new OdbcConnection(connectionStringAtlantis))
                {
                    DataSet dataSet = new DataSet();
                    OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"SELECT DBA.schueler_info.pu_id AS AtlantisSchuelerId,
DBA.schueler_info.s_typ_puin AS Name,
DBA.schueler_info.s_typ_puin_2,
DBA.schueler_info.datum AS Von,
DBA.schueler_info.datum_2 AS Bis
FROM DBA.schueler_info
WHERE s_typ_puin = 'B';", connection);

                    Console.Write("Bemerkungen ".PadRight(75, '.'));

                    connection.Open();
                    schuelerAdapter.Fill(dataSet, "DBA.klasse");

                    foreach (DataRow theRow in dataSet.Tables["DBA.klasse"].Rows)
                    {
                        Bemerkung bemerkung = new Bemerkung()
                        {
                            AtlantisSchuelerId = theRow["AtlantisSchuelerId"] == null ? -99 : Convert.ToInt32(theRow["AtlantisSchuelerId"]),
                            Kürzel = theRow["Name"] == null ? "" : theRow["Name"].ToString(),
                            Von = theRow["Von"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["Von"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture),
                            Bis = theRow["Bis"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["Bis"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture),
                        };

                        if (bemerkung.AtlantisSchuelerId == 154312)
                        {
                            string a = "";
                        }
                        // Nur wenn jetzt innerhalb der Ausbildung liegt, wird der Betrieb hinzugefügt.
                        if (bemerkung.Von < DateTime.Now && DateTime.Now < bemerkung.Bis)
                        {
                            this.Add(bemerkung);
                        }
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