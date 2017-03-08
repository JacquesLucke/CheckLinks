using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;

namespace CheckLinks
{
    class AccessInterface
    {
        public static OleDbConnection GetConnection(String path)
        {
            string accessString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path;
            return new OleDbConnection(accessString);
        }

        public static OleDbDataAdapter ExecuteSelectCommand(OleDbConnection connection, string sql)
        {
            OleDbCommand accessCommand = new OleDbCommand(sql, connection);
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(accessCommand);
            return dataAdapter;
        }
        
        public static void ExecuteUpdateCommand(OleDbConnection connection, string sql)
        {
            OleDbCommand command = new OleDbCommand(sql, connection);
            command.ExecuteNonQuery();
        }
        

        public static string ExtractAddress(string link)
        {
            if (link.Length < 2) return String.Empty;
            if (link[0] == '#' && link[1] != '#')
            {
                return link.Substring(1, link.IndexOf('#', 1) - 1).Trim();
            }
            if (link[0] != '#')
            {
                int firstIndex = link.IndexOf('#');
                int secondIndex = link.IndexOf('#', firstIndex + 1);
                return link.Substring(firstIndex + 1, secondIndex - firstIndex - 1).Trim();
            }
            return String.Empty;
        }
    }
}
