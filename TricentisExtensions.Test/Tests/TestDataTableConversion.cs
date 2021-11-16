using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;

namespace TricentisExtensions.Test
{
    [TestClass]
    public class TestDataTableConversion
    {
        [TestMethod]
        public void TestDataTable()
        {
            DataSet ds = new DataSet();
            //var dt3 = new DataTable();
            var rows = new List<string>();
            var dt1 = ConvertExcelToDataTable(@"C:\Users\Oleksii.Kukharenko\Desktop\Refference.xlsx", "GL Details Report");
            var dt2 = ConvertExcelToDataTable(@"C:\Users\Oleksii.Kukharenko\Desktop\Refference1.xlsx", "GL Details Report");
            dt1.TableName = "Source";
            dt2.TableName = "Target";
            ds.Tables.Add(dt1);
            ds.Tables.Add(dt2);
            System.IO.File.WriteAllText(@"C:\Users\Oleksii.Kukharenko\Desktop\WriteLines.txt", $"\n");

            System.IO.File.AppendAllText(@"C:\Users\Oleksii.Kukharenko\Desktop\WriteLines.txt", $"Before cleaning [Rows X Columns] :{dt1.Rows.Count} X {dt1.Columns.Count}\n");

            //dt1 = dt1.Rows.Cast<DataRow>().
            //    Where(row => !row.ItemArray.All(field => field is DBNull || string.IsNullOrWhiteSpace(field as string))).
            //    CopyToDataTable();

            //DataRow rowToRemove = dt1.Rows.Cast<DataRow>().Select(row => row.ItemArray.Contains("Report Run Date/Time"));
            //int i = 0;
            //foreach (DataColumn col in dt1.Columns)
            //{
            //    i++;
            //    System.IO.File.AppendAllText(@"C:\Users\Oleksii.Kukharenko\Desktop\WriteLines.txt", $"Index : {i}; Name : {col.ColumnName}\n");

            //}
            //System.IO.File.AppendAllText(@"C:\Users\Oleksii.Kukharenko\Desktop\WriteLines.txt", $"After cleaning [Rows X Columns] :{dt1.Rows.Count} X {dt1.Columns.Count}\n");

            var result = dt1.AsEnumerable().Except(dt2.AsEnumerable(), DataRowComparer.Default);

            foreach (DataRow row in result)
            {
                foreach (var item in row.ItemArray)
                {
                    System.IO.File.AppendAllText(@"C:\Users\Oleksii.Kukharenko\Desktop\WriteLines.txt", $"\t{item.ToString()}");
                }
            }
        }

        private string ConnectionToExcel(string filePath)
        {
            return $"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {filePath}; Extended Properties = 'Excel 12.0 XML;HDR=YES;'; ";
        }

        public DataTable ConvertExcelToDataTable(string filePath, string sheetName)
        {
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

        public static DataTable CompareTwoDataTable(DataTable dt1, DataTable dt2)
        {
            dt1.AcceptChanges();
            dt1.Merge(dt2);
            dt1.AcceptChanges();
            DataTable d3 = dt1.DefaultView.ToTable(true);
            return d3;
        }
    }
}
