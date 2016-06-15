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
            FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);

            try
            {
                var util = TransferDataFactory.GetUtil(fileName);
                var tables = util.GetTables(fs);

                foreach (var table in tables)
                {
                    var csv = TransferDataFactory.GetUtil(DataFileType.CSV);
                    var data = csv.GetStream(table);

                    if (data.CanRead)
                    {
                        StreamReader reader = new StreamReader(data, Encoding.UTF8);
                        strList.Add(reader.ReadToEnd());
                    }
                }
            }
            catch (System.Exception ex)
            {
            	
            }
            finally
            {
                fs.Close();
            }

            return strList;
        }
    }
}
