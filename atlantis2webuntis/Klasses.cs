using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace atlantis2webuntis
{
    public class Klasses : List<Klasse>
    {
        public Klasses(int periode)
        {
            using (SqlConnection odbcConnection = new SqlConnection(Global.ConnectionStringUntis))
            {
                try
                {
                    string queryString = @"SELECT DISTINCT 
Class.Class_ID, 
Class.Name,
Class.TeacherIds,
Class.Longname, 
Teacher.Name, 
Class.ClassLevel,
Class.PERIODS_TABLE_ID
FROM Class LEFT JOIN Teacher ON Class.TEACHER_ID = Teacher.TEACHER_ID WHERE (((Class.SCHOOLYEAR_ID)=" + Global.AktSj[0] + Global.AktSj[1] + ") AND (((Class.TERM_ID)=" + periode + ")) AND ((Teacher.SCHOOLYEAR_ID)=" + Global.AktSj[0] + Global.AktSj[1] + ") AND ((Teacher.TERM_ID)=" + periode + ")) OR (((Class.SCHOOLYEAR_ID)=" + Global.AktSj[0] + Global.AktSj[1] + ") AND ((Class.TERM_ID)=" + periode + ") AND ((Class.SCHOOL_ID)=177659) AND ((Teacher.SCHOOLYEAR_ID) Is Null) AND ((Teacher.TERM_ID) Is Null)) ORDER BY Class.Name ASC;";

                    SqlCommand odbcCommand = new SqlCommand(queryString, odbcConnection);
                    odbcConnection.Open();
                    SqlDataReader sqlDataReader = odbcCommand.ExecuteReader();

                    while (sqlDataReader.Read())
                    {                        
                        bool istVollzeit = istVollzeitKlasse(Global.SafeGetString(sqlDataReader, 1));

                        Klasse klasse = new Klasse()
                        {
                            IdUntis = sqlDataReader.GetInt32(0),
                            NameUntis = Global.SafeGetString(sqlDataReader, 1),
                            IstVollzeit = istVollzeit
                        };

                        this.Add(klasse);
                    };

                    sqlDataReader.Close();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    throw new Exception(ex.ToString());
                }
                finally
                {
                    odbcConnection.Close();
                }
            }
        }

        private bool istVollzeitKlasse(string klassenname)
        {
            var vollzeitBeginn = new List<string>() { "BS", "BW", "BT", "FM", "FS", "G", "HB" };

            foreach (var item in vollzeitBeginn)
            {
                if (klassenname.StartsWith(item))
                {
                    return true;
                }
            }
            return false;
        }
    }
}