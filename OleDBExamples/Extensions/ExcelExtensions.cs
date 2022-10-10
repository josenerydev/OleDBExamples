using Newtonsoft.Json;
using System.Data;
using System.Data.OleDb;

namespace OleDBExamples.Extensions
{
    public class ExcelExtensions
    {
        public static List<T> XlsxToObject<T>(string path, string commandText)
        {
            using (OleDbConnection connection = new OleDbConnection())
            {
                DataTable dataTable = new DataTable();
                connection.ConnectionString = $"Provider=Microsoft.ACE.OLEDB.12.0; Data Source={path};Extended Properties=Excel 12.0;";
                using (OleDbCommand command = connection.CreateCommand())
                {
                    command.CommandText = commandText;
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter())
                    {
                        adapter.SelectCommand = command;
                        adapter.Fill(dataTable);
                    }
                }

                return JsonConvert.DeserializeObject<List<T>>(JsonConvert.SerializeObject(dataTable));
            }
        }
    }
}
