using System.Data.OleDb;

namespace atlantis2webuntis
{
    public class Global
    {
        public static string SafeGetString(OleDbDataReader reader, int colIndex)
        {
            if (!reader.IsDBNull(colIndex))
                return reader.GetString(colIndex);
            return string.Empty;
        }
    }
}