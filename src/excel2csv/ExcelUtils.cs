using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;

namespace EXCEL2CSV
{

    public class ExcelUtils
    {
        public static List<string> ConverseExcelToCSV(string fileName)
        {
            List<string> strList = new List<string>();
            FileStream stream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            var util = TransferDataFactory.GetUtil(fileName);
            var tables = util.GetTables(stream);

            foreach (var table in tables)
            {
                var csv = TransferDataFactory.GetUtil(DataFileType.CSV);
                //var mStream = util.GetStream(data);
                var mStream = csv.GetStream(table);

                if (mStream.CanRead)
                {
                    StreamReader reader = new StreamReader(mStream, Encoding.UTF8);
                    strList.Add(reader.ReadToEnd());
                }
            }
            return strList;
        }
    }
}
