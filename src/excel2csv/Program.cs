using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace EXCEL2CSV
{
    class Program
    {
        static void Main(string[] args)
        {
            string fullName;
            Console.WriteLine("Input Excel file name: ");
            fullName = Console.ReadLine();
            if (fullName.Length == 0)
            {
                fullName = @"C:/123/test.xlsx";
            }
            var csvFiles = ExcelUtils.ConverseExcelToCSV(fullName);

            for (int i = 0; i < csvFiles.Count; ++i)
            {
                var csv = csvFiles[i];
                Console.Write(csv);
                FileStream fs = new FileStream(string.Format("{0}.{1}.csv", fullName, i), FileMode.Create);
                byte[] data = new UTF8Encoding().GetBytes(csv);
                fs.Write(data, 0, data.Length);
                fs.Flush();
                fs.Close();
            }
        }
    }
}
