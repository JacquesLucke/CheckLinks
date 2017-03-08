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
            Console.WriteLine("CheckLinks started.");
            /*
            InputParameters input = new InputParameters(
                @"C:\Users\lu\Desktop\Jacques\TestLinkCheck.accdb",
                "Dokument_Browser",
                "Link",
                "LinkOK");
            */
            InputParameters input = null;
            try { input = InputParameters.FromArray(args); }
            catch
            {
                Console.WriteLine("4 arguments required: <path> <table_name> <link_field> <ok_field>");
            }

            if (input != null)
            {
                Execute(input);
                Console.WriteLine("\nDone.");
            }
            Console.ReadKey();
        }

        private static void Execute(InputParameters input)
        {
            Console.WriteLine("Trying to connect...");
            OleDbConnection connection = AccessInterface.GetConnection(input.FileName);
            connection.Open();
            Console.WriteLine("Connected.");
            Execute(connection, input);
            //Clean(connection, input, false);
            connection.Close();
        }

        private static void Execute(OleDbConnection connection, InputParameters input)
        {
            SetNullsNotOk(connection, input);
            HashSet<string> validLinks = new HashSet<string>();
            foreach (DataRow row in GetNonNullDataRows(connection, input))
            {
                string content = Convert.ToString(row[input.LinkField]);
                string address = AccessInterface.ExtractAddress(content);
                
                if (File.Exists(address))
                {
                    Console.WriteLine("Yes: " + address);
                    validLinks.Add(content);
                }
                else
                {
                    Console.WriteLine("No: " + address);
                }
            }

            Console.WriteLine("\nExecuting query...");
            string condition = getCondition(input, validLinks);
            string sqlTrue = $"UPDATE [{input.TableName}] SET [{input.OkField}] = True WHERE {condition}";
            AccessInterface.ExecuteUpdateCommand(connection, sqlTrue);
            string sqlFalse = $"UPDATE [{input.TableName}] SET [{input.OkField}] = False WHERE NOT {condition}";
            AccessInterface.ExecuteUpdateCommand(connection, sqlFalse);
            Console.WriteLine("Query executed");
        }

        private static string getCondition(InputParameters input, HashSet<string> links)
        { 
            List<string> elements = new List<string>();
            foreach(string s in links)
            {
                elements.Add($"\"{s}\"");
            }
            return $"([{input.LinkField}] IN ({String.Join(",", elements)}))";
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
