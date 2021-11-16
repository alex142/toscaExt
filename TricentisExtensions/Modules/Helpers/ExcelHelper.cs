using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;


namespace TricentisExtensions.Modules.Helpers
{
    public static class ExcelHelper
    {
        public static void UpdateExcel(Excel._Worksheet sheet, int row, int col, string data)
        {
            try
            {
                sheet.Cells[row, col] = data;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public static void WriteColumn(Excel._Worksheet sheet, string columnName, IEnumerable<string> data)
        {
            var lst = data.ToList();
            Excel.Range xlRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[lst.Count, 10]];

            for (int i = 1; i < xlRange.Columns.Count; i++)
            {
                if (!String.IsNullOrEmpty(xlRange.Cells[1, i].Value) && (xlRange.Cells[1, i].Value.ToUpper() == columnName.ToUpper()))
                {
                    for (int j = 2; j < xlRange.Rows.Count + 2; j++)
                    {
                        xlRange.Cells[j, i].Value = lst[j - 2];
                    }
                    break;
                }
            }
        }

        public static void CleanColumn(Excel._Worksheet sheet, string columnName)
        {
            Excel.Range colRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 10]];
            for (int i = 1; i < colRange.Columns.Count; i++)
            {
                if (!string.IsNullOrEmpty(colRange.Cells[1, i].Value) && (colRange.Cells[1, i].Value.ToUpper() == columnName.ToUpper()))
                {
                    Excel.Range rowRange = sheet.Range[sheet.Cells[2, i], sheet.Cells[200, i]];
                    rowRange.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                    break;
                }
            }
        }

        public static void CleanSheet(Excel._Worksheet sheet, int startRow = 1, int startCol = 1, int endRow = 1000, int endCol = 100)
        {
            Excel.Range xlRange = sheet.Range[sheet.Cells[startRow, startCol], sheet.Cells[endRow, endCol]];

            xlRange.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
        }

        //Connection string to excel workbook
        private static string ConnectionToExcel(string filePath)
        {
            return $"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {filePath}; Extended Properties = 'Excel 12.0 XML;HDR=YES;'; ";
            //return $"Provider = Microsoft.ACE.OLEDB.15.0; Data Source = {filePath}; Extended Properties = 'Excel 12.0 XML;HDR=YES;'; ";
        }

        /* Desc: Converts excel sheet to DataTable type
         * Returns: New DataTable object
         */
        public static DataTable ConvertExcelToDataTable(string filePath, string sheetName)
        {
            if (sheetName.Contains('.'))
            {
                sheetName = sheetName.Replace('.', '#');
            }

            DataTable data = new DataTable();

            using (OleDbConnection conn = new OleDbConnection(ConnectionToExcel(filePath)))
            {
                OleDbCommand comm = new OleDbCommand($"Select * from [{sheetName}$]", conn);
                conn.Open();
                using (OleDbDataAdapter sda = new OleDbDataAdapter(comm))
                {
                    sda.Fill(data);
                }
            }

            return data;
        }

        /* Desc: Removes empty rows from DataTable objects and returns copy of initial without empty rows
         * Returns: New DataTable object
         */
        public static DataTable RemoveEmptyRows(DataTable table)
        {
            return table.Rows.Cast<DataRow>().Where(row => !row.ItemArray.All(field => field is DBNull || string.IsNullOrWhiteSpace(field as string))).CopyToDataTable();
        }

        /* Desc: Compares cell by cell of two DataTables
         * Returns: True if all cells matched, otherwise false
         */
        public static bool AreTablesEqual(DataTable target, DataTable source)
        {
            bool isPresent = false;
            foreach (DataRow row in target.Rows)
            {
                foreach (var targetItem in row.ItemArray)
                {
                    isPresent = false;
                    foreach (DataRow sourceItem in source.Rows)
                    {
                        if (sourceItem.ItemArray.Contains(targetItem))
                        {
                            isPresent = true;
                        }
                    }
                }
            }
            return isPresent;
        }

        /* Desc: Compares two DataTables
         * Returns: IEnumerable<DataRow> collection with all rows that present in source file but not in Target
         */
        public static IEnumerable<DataRow> GetTablesDifference(DataTable target, DataTable source)
        {
            return source.AsEnumerable().Except(target.AsEnumerable(), DataRowComparer.Default);
        }

        private static DataTable GetColumn(DataTable dt, string columnName)
        {
            int index = dt.Columns.IndexOf(columnName);

            if (index != -1)
            {
                DataTable resDt = new DataTable();
                resDt.Columns.Add(columnName);

                foreach (DataRow row in dt.Rows)
                {
                    resDt.Rows.Add(row[index]);
                }

                return resDt;
            }
            else
            {
                throw new IndexOutOfRangeException("Column with given name is not found");
            }

        }

        public static DataTable GetColumn(DataTable dt, int index)
        {
            if (index > -1)
            {
                DataTable resDt = new DataTable();
                resDt.Columns.Add($"Column {index}");

                foreach (DataRow row in dt.Rows)
                {
                    resDt.Rows.Add(row[index]);
                }

                return resDt;
            }
            else
            {
                throw new IndexOutOfRangeException("Column with given name is not found");
            }

        }


        public static void DeleteDataTableRows(DataTable dt, int startAt, int finishAt)
        {
            for (int i = startAt; i < finishAt; i++)
            {
                dt.Rows[i].Delete();
            }

            dt.AcceptChanges();
        }

        public static void DeleteDataTableColumns(DataTable dt, int startAt, int finishAt = 0)
        {
            if (finishAt == 0)
            {
                finishAt = dt.Columns.Count;
            }

            for (int i = startAt; i < finishAt; i++)
            {
                dt.Columns.RemoveAt(startAt);
            }
            dt.AcceptChanges();
        }

        public static void ExportToExcel(DataTable tbl, string excelFilePath)
        {
            try
            {
                if (tbl == null || tbl.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                // load excel, and create a new workbook
                var xlApp = new Excel.Application();
                Excel._Workbook xlBook;
                bool existingFile = !string.IsNullOrEmpty(excelFilePath) && File.Exists(excelFilePath);

                if (existingFile)
                {
                    xlBook = xlApp.Workbooks.Open(excelFilePath);
                }
                else
                {
                    xlBook = xlApp.Workbooks.Add();
                }
                // single worksheet
                Excel._Worksheet xlSheet = xlApp.ActiveSheet;

                CleanSheet(xlSheet, 1, 1, xlSheet.Rows.Count, xlSheet.Columns.Count);

                // column headings
                for (var i = 0; i < tbl.Columns.Count; i++)
                {
                    xlSheet.Cells[1, i + 1] = tbl.Columns[i].ColumnName;
                }

                // rows
                for (var i = 0; i < tbl.Rows.Count; i++)
                {
                    // to do: format datetime values before printing
                    for (var j = 0; j < tbl.Columns.Count; j++)
                    {
                        xlSheet.Cells[i + 2, j + 1] = tbl.Rows[i][j];
                    }
                }

                // check file path

                try
                {
                    if (existingFile)
                    {
                        xlBook.Save();
                    }
                    else
                    {
                        xlBook.SaveAs(excelFilePath);
                    }

                    xlBook.Close();
                    xlApp.Quit();
                }
                catch (Exception ex)
                {
                    throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                                        + ex.Message);
                }

            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }


        public static IEnumerable<DataRow> GetColumnsDifference(DataTable target, DataTable source, string columnName)
        {
            var sourceCol = GetColumn(source, columnName);
            var targetCol = GetColumn(target, columnName);           
            

            var dif = GetTablesDifference(targetCol, sourceCol);
            return dif;
        }

        public static bool AreColumnsEqual(DataTable target, DataTable source, string columnName)
        {
            var sourceCol = GetColumn(source, columnName);
            var targetCol = GetColumn(target, columnName);
            sourceCol = RemoveEmptyRows(sourceCol);
            targetCol = RemoveEmptyRows(targetCol);

            if (sourceCol.Rows.Count != targetCol.Rows.Count)
            {
                return false;
            }

            var dif1 = GetTablesDifference(targetCol, sourceCol).Count();
            var dif2 = GetTablesDifference(sourceCol, targetCol).Count();

            return true ? dif1 == dif2 && dif1 == 0 : false;
        }
    }
}
