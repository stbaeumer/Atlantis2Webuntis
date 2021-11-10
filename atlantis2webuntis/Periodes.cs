using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data.SqlClient;
using System.IO;
using System.Text;

namespace atlantis2webuntis
{
    public class Periodes : List<Periode>
    {
        public Periodes()
        {
            SqlConnection sqlConnection;
            sqlConnection = new SqlConnection(@"Data Source=SQL01\UNTIS;Initial Catalog=master;Integrated Security=True");

            try
            {
                sqlConnection.Open();

                string queryString = @"SELECT DISTINCT
Terms.TERM_ID, 
Terms.Name, 
Terms.Longname, 
Terms.DateFrom, 
Terms.DateTo
FROM Terms
WHERE (((Terms.SCHOOLYEAR_ID)= " + (DateTime.Now.Month >= 7 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString() + (DateTime.Now.Month >= 7 ? DateTime.Now.Year + 1 : DateTime.Now.Year).ToString() + ")  AND ((Terms.SCHOOL_ID)=177659)) ORDER BY Terms.TERM_ID;";

                using (SqlCommand sqlCommand = new SqlCommand(queryString, sqlConnection))
                {
                    SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();

                    while (sqlDataReader.Read())
                    {
                        Periode periode = new Periode()
                        {
                            IdUntis = sqlDataReader.GetInt32(0),
                            Name = Global.SafeGetString(sqlDataReader, 1),
                            Langname = Global.SafeGetString(sqlDataReader, 2),
                            Von = DateTime.ParseExact((sqlDataReader.GetInt32(3)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                            Bis = DateTime.ParseExact((sqlDataReader.GetInt32(4)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
                        };

                        if (DateTime.Now > periode.Von && DateTime.Now < periode.Bis)
                            this.AktuellePeriode = periode.IdUntis;

                        this.Add(periode);
                    };

                    // Korrektur des Periodenendes

                    for (int i = 0; i < this.Count - 1; i++)
                    {
                        this[i].Bis = this[i + 1].Von.AddDays(-1);
                    }

                    sqlDataReader.Close();
                }

                if (this.AktuellePeriode == 0)
                {
                    Console.WriteLine("Es kann keine aktuelle Periode ermittelt werden. Das ist z. B. während der Sommerferien der Fall.");
                    Console.WriteLine("Es wird die Periode " + this.Count + " als aktuelle Periode angenommen.");
                    this.AktuellePeriode = this.Count;
                }
                else
                {
                    
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                sqlConnection.Close();
            }
        }

        public int AktuellePeriode { get; private set; }
    }
}