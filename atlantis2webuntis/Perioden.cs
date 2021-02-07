using System;
using System.Collections.Generic;
using System.Data.OleDb;

namespace atlantis2webuntis
{
    public class Perioden : List<Periode>
    {
        public Perioden(string connectionStringUntis, string aktSjUntis)
        {

            Console.Write("Perioden in Untis ".PadRight(30, '.'));

            using (OleDbConnection oleDbConnection = new OleDbConnection(connectionStringUntis))
            {
                try
                {
                    string queryString = @"SELECT       Terms.TERM_ID, 
                                                        Terms.Name, 
                                                        Terms.Longname, 
                                                        Terms.DateFrom, 
                                                        Terms.DateTo
                                                        FROM Terms
                                                        WHERE (((Terms.SCHOOLYEAR_ID)= " + aktSjUntis + ")AND ((Terms.SCHOOL_ID)=177659)) ORDER BY Terms.TERM_ID;";

                    OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                    oleDbConnection.Open();
                    OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                    while (oleDbDataReader.Read())
                    {
                        Periode periode = new Periode()
                        {
                            IdUntis = oleDbDataReader.GetInt32(0),
                            Name = Global.SafeGetString(oleDbDataReader, 1),
                            Langname = Global.SafeGetString(oleDbDataReader, 2),
                            Von = DateTime.ParseExact((oleDbDataReader.GetInt32(3)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                            Bis = DateTime.ParseExact((oleDbDataReader.GetInt32(4)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
                        };
                        this.Add(periode);
                    };

                    oleDbDataReader.Close();

                    Console.WriteLine((" " + this.Count.ToString()).PadLeft(30, '.'));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                finally
                {
                    oleDbConnection.Close();
                }
            }
        }

        public int GetAktuellePerionde()
        {
            foreach (var periode in this)
            {
                if (periode.Von < DateTime.Now && DateTime.Now <= periode.Bis)
                {
                    return periode.IdUntis;
                }
            }
            return 0;
        }
    }
}