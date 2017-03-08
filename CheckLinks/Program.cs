using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CheckLinks
{
    class Program
    {
        static void Main(string[] args)
        {
            /*
            InputParameters input = new InputParameters(
                @"C:\Users\lu\Desktop\Jacques\TestLinkCheck.accdb",
                "Dokument_Browser",
                "Link",
                "LinkOK");
            */
            InputParameters input = InputParameters.FromArray(args);

            Execute(input);
            Console.WriteLine("\nDone.");
            Console.ReadKey();
        }

        private static void Execute(InputParameters input)
        {
            OleDbConnection connection = AccessInterface.GetConnection(input.FileName);
            connection.Open();
            Execute(connection, input);
            //Clean(connection, input, true);
            connection.Close();
        }

        private static void Execute(OleDbConnection connection, InputParameters input)
        {
            SetNullsNotOk(connection, input);
            foreach (DataRow row in GetNonNullDataRows(connection, input))
            {
                string content = Convert.ToString(row[input.LinkField]);
                string address = AccessInterface.ExtractAddress(content);

                bool exists = File.Exists(address);

                if (exists) Console.WriteLine("Yes: " + address);
                else Console.WriteLine("No: " + address);

                string sql = $"UPDATE [{input.TableName}] SET [{input.OkField}] = {exists} WHERE [{input.LinkField}] = \"{content}\"";
                AccessInterface.ExecuteUpdateCommand(connection, sql);
            }
        }

        private static DataRowCollection GetNonNullDataRows(OleDbConnection connection, InputParameters input)
        {
            String sql = $"SELECT [{input.LinkField}] FROM [{input.TableName}] WHERE [{input.LinkField}] is not null";
            OleDbDataAdapter dataAdapter = AccessInterface.ExecuteSelectCommand(connection, sql);
            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet, input.TableName);
            DataTable table = dataSet.Tables[input.TableName];
            return table.Rows;
        }

        private static void SetNullsNotOk(OleDbConnection connection, InputParameters input)
        {
            string sql = $"UPDATE [{input.TableName}] SET [{input.OkField}] = False WHERE [{input.LinkField}] is null";
            AccessInterface.ExecuteUpdateCommand(connection, sql);
        }
        
        private static void Clean(OleDbConnection connection, InputParameters input, bool state)
        {
            string sql = $"UPDATE [{input.TableName}] SET [{input.OkField}] = {state}";
            AccessInterface.ExecuteUpdateCommand(connection, sql);
        }
    }
}
