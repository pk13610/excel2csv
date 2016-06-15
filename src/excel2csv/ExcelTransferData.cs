//////////////////////////////////////////////////////////////////////////
// Ref: http://www.cnblogs.com/qisheng/p/3441902.html
// Modify: vavava
// INFO: add sheets support by GetTables
// TODO: scv.GetTables()
// DATE: 20160615
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;


namespace EXCEL2CSV
{
    public interface ITransferData
    {
        Stream GetStream(DataTable table);
        DataTable GetData(Stream stream);
        List<DataTable> GetTables(Stream stream);
    }

    public abstract class ExcelTransferData : ITransferData
    {
        protected IWorkbook _workBook;

        public virtual Stream GetStream(DataTable table)
        {
            var sheet = _workBook.CreateSheet();
            if (table != null)
            {
                var rowCount = table.Rows.Count;
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    var row = sheet.CreateRow(i);
                    for (int j = 0; j < table.Columns.Count; j++)
                    {
                        var cell = row.CreateCell(j);
                        if (table.Rows[i][j] != null)
                        {
                            var str = table.Rows[i][j].ToString();
                            cell.SetCellValue(str);
                        }
                    }
                }
            }
            MemoryStream ms = new MemoryStream();
            _workBook.Write(ms);
            return ms;
        }

        public virtual DataTable GetData(Stream stream)
        {
            using (stream)
            {
                var sheet = _workBook.GetSheetAt(0);
                if (sheet != null)
                {
                    var headerRow = sheet.GetRow(0);
                    DataTable dt = new DataTable();
                    int columnCount = headerRow.Cells.Count;
                    for (int i = 0; i < columnCount; i++)
                    {
                        dt.Columns.Add("col_" + i.ToString());
                    }
                    var row = sheet.GetRowEnumerator();
                    while (row.MoveNext())
                    {
                        var dtRow = dt.NewRow();
                        var excelRow = row.Current as NPOI.SS.UserModel.IRow;
                        for (int i = 0; i < columnCount; i++)
                        {
                            var cell = excelRow.GetCell(i);

                            if (cell != null)
                            {
                                dtRow[i] = GetValue(cell);
                            }
                        }
                        dt.Rows.Add(dtRow);
                    }
                    return dt;
                }
            }

            return null;
        }

        public virtual List<DataTable> GetTables(Stream stream)
        {
            List<DataTable> tables = new List<DataTable>();
            using (stream)
            {

                for (int index = 0; ; ++index)
                {
                    var sheet = _workBook.GetSheetAt(index);
                    if (sheet == null)
                    {
                        break;
                    }
                    var headerRow = sheet.GetRow(0);
                    if (headerRow == null)
                    {
                        break;
                    }
                    DataTable dt = new DataTable();
                    int columnCount = headerRow.Cells.Count;
                    for (int i = 0; i < columnCount; i++)
                    {
                        dt.Columns.Add("col_" + i.ToString());
                    }
                    var row = sheet.GetRowEnumerator();
                    while (row.MoveNext())
                    {
                        var dtRow = dt.NewRow();
                        var excelRow = row.Current as NPOI.SS.UserModel.IRow;
                        for (int i = 0; i < columnCount; i++)
                        {
                            var cell = excelRow.GetCell(i);

                            if (cell != null)
                            {
                                dtRow[i] = GetValue(cell);
                            }
                        }
                        dt.Rows.Add(dtRow);
                    }
                    tables.Add(dt);
                }
            }

            return tables;
        }

        private object GetValue(ICell cell)
        {
            object value = null;
            switch (cell.CellType)
            {
                case CellType.Blank:
                    break;
                case CellType.Boolean:
                    value = cell.BooleanCellValue ? "1" : "0"; break;
                case CellType.Error:
                    value = cell.ErrorCellValue; break;
                case CellType.Formula:
                    value = "=" + cell.CellFormula; break;
                case CellType.Numeric:
                    value = cell.NumericCellValue.ToString(); break;
                case CellType.String:
                    value = cell.StringCellValue; break;
                case CellType.Unknown:
                    break;
            }
            return value;
        }

    }


    /// <summary>
    /// 2003
    /// </summary>
    public class XlsTransferData : ExcelTransferData
    {
        public override Stream GetStream(DataTable table)
        {
            base._workBook = new HSSFWorkbook();
            return base.GetStream(table);
        }

        public override DataTable GetData(Stream stream)
        {
            base._workBook = new HSSFWorkbook(stream);
            return base.GetData(stream);
        }

        public override List<DataTable> GetTables(Stream stream)
        {
            base._workBook = new HSSFWorkbook(stream);
            return base.GetTables(stream);
        }
    }

    /// <summary>
    /// 2007
    /// </summary>
    public class XlsxTransferData : ExcelTransferData
    {

        public override Stream GetStream(DataTable table)
        {
            base._workBook = new NPOI.XSSF.UserModel.XSSFWorkbook();
            return base.GetStream(table);
        }

        public override DataTable GetData(Stream stream)
        {
            base._workBook = new NPOI.XSSF.UserModel.XSSFWorkbook(stream);
            return base.GetData(stream);
        }

        public override List<DataTable> GetTables(Stream stream)
        {
            base._workBook = new NPOI.XSSF.UserModel.XSSFWorkbook(stream);
            return base.GetTables(stream);
        }
    }

    public enum DataFileType
    {
        CSV,
        XLS,
        XLSX
    }

    public class TransferDataFactory
    {
        public static ITransferData GetUtil(string fileName)
        {
            var array = fileName.Split('.');
            var dataType = (DataFileType)Enum.Parse(typeof(DataFileType), array[array.Length - 1], true);
            return GetUtil(dataType);
        }

        public static ITransferData GetUtil(DataFileType dataType)
        {
            switch (dataType)
            {
                case DataFileType.CSV: return new CsvTransferData();
                case DataFileType.XLS: return new XlsTransferData();
                case DataFileType.XLSX: return new XlsxTransferData();
                default: return new CsvTransferData();
            }
        }

    }
}
