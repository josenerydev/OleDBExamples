using OleDBExamples.Extensions;
using OleDBExamples.Models;

namespace OleDBExamples
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<Product> products = new List<Product>();

            const string commandText = "SELECT CInt(Id) AS Id, CStr(Name) AS Name FROM [Sheet1$]";
            var fileName = "C:\\Temp\\products.xlsx";
            try
            {
                products = ExcelExtensions.XlsxToObject<Product>(fileName, commandText);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}