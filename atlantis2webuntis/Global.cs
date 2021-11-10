using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace atlantis2webuntis
{
    internal class Global
    {
        public const string ConnectionStringUntis = @"Data Source=SQL01\UNTIS;Initial Catalog=master;Integrated Security=True";

        public static List<string> AktSj = new List<string>() {
            (DateTime.Now.Month >= 7 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
            (DateTime.Now.Month >= 7 ? DateTime.Now.Year + 1 : DateTime.Now.Year).ToString()
        };

        public static string SafeGetString(SqlDataReader reader, int colIndex)
        {
            if (!reader.IsDBNull(colIndex))
                return reader.GetString(colIndex);
            return string.Empty;
        }
    }
}