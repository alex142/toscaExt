using System.Windows;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System;
using System.Collections.Generic;
using System.Linq;
using TricentisExtensions.Modules;
using AutoIt;
using TricentisExtensions.Test.Utils;
using TricentisExtensions.Modules.Helpers;
using System.Data;
using System.IO;
using TricentisExtensions.Modules.CustomEngine;

namespace TricentisExtensions.Test.Tests
{
    enum ColumnType
    {
        undefined,
        Layout,
        Dimension,
        VarList,
        DefaultPrompts,
        ActiveFilters
    }

    [TestClass]
    public class UnitTest1
    {

        [TestMethod]
        public void TestMethod1()
        {
            Excel.Application xlApp;

            xlApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

            Excel.Workbooks books = xlApp.Workbooks;

            File.WriteAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", $"{books.Count}\n");
            Array tempFiltersArray = (Array)((object)xlApp.Run("SAPListOfDynamicFilters", "DS_1", "Key"));


            foreach (var item in tempFiltersArray)
            {
                File.AppendAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", $"{item.ToString()}\n");

            }

            foreach (Excel.Workbook book in books)
            {
                File.AppendAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", $"{book.Name}\n");
            }

            //System.IO.File.WriteAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", $"{ xlApp.Application.Run("SapGetProperty", "LastError", "Text")}\n");


            if (tempFiltersArray.Rank > 1)
            {
                var tempFiltersArray1 = (object[,])tempFiltersArray;
                for (int i = 1; i <= tempFiltersArray1.GetLength(0); i++)
                {
                    File.AppendAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", $"\n");

                    for (int j = 1; j <= tempFiltersArray1.GetLength(1); j++)
                    {
                        if (tempFiltersArray1[i, j].ToString().ToUpper() == "MEASURES")
                        {
                            File.AppendAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", $"{tempFiltersArray1[i, j]}\t");
                            File.AppendAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", $"{tempFiltersArray1[i, j + 1]}\t");
                        }
                    }
                }
            }
            //System.IO.File.AppendAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", $"\n{tempFiltersArray.Rank}\n");
            //xlApp.Run("SAPExecuteCommand", "Refresh", "DS_1");
            //var name = xlApp.Run("SAPGetVariable", "DS_1", "Ledger", "TECHNICALNAME");
            //xlApp.Run("SAPSetVariable", name, "AK", "INPUT_STRING", "DS_1");

            //System.IO.File.AppendAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", $"{name}\n");
        }

        [TestMethod]
        public void TestDataComparison()
        {
            WriteToFile.Reset();
            var refFile = @"C:\Users\okukharenko\Desktop\TestRefference.xlsx";
            var refSheet = "Outstanding Invoice LI Data";

            var xl1 = @"C:\Tosca_Projects\Tosca_Workspaces\RandA\ExcelTmp\dt1.xlsx";
            var xl2 = @"C:\Tosca_Projects\Tosca_Workspaces\RandA\ExcelTmp\dt2.xlsx";
            //var xl1 = @"dt1.xlsx";
            //var xl2 = @"dt2.xlsx";

            var dt1 = ExcelHelper.ConvertExcelToDataTable(refFile, refSheet);


            var xlapp = new FileHelper();

            xlapp.TryGetInstance("Eng Detail Acc");

            var path = xlapp.XlBook.FullName;

            var sheet = xlapp.XlSheet.Name;

            if (sheet.Contains('.'))
            {
                sheet = sheet.Replace('.', '#');
            }

            WriteToFile.Write($"Opening: {path}\nSheet: {sheet}\n");

            var dt2 = ExcelHelper.ConvertExcelToDataTable(path, sheet);

            ExcelHelper.DeleteDataTableRows(dt2, 0, 9);
            ExcelHelper.DeleteDataTableColumns(dt2,22);

            var dt11 = dt1.Copy();
            var dt12 = dt2.Copy();


            dt2 = ExcelHelper.RemoveEmptyRows(dt2);

            ExcelHelper.ExportToExcel(dt1, xl1);
            ExcelHelper.ExportToExcel(dt2, xl2);

            var xldt1 = ExcelHelper.ConvertExcelToDataTable(xl1, "Sheet1");
            var xldt2 = ExcelHelper.ConvertExcelToDataTable(xl2, "Sheet1");

            WriteToFile.Write($"DT1 Cols: {xldt1.Columns.Count}\nRows: {xldt1.Rows.Count}\n");            

            WriteToFile.Write($"DT2 Cols: {xldt2.Columns.Count}\nRows: {xldt2.Rows.Count}\n");

            var result2 = ExcelHelper.GetTablesDifference(xldt1, xldt2);

            foreach (var item in result2)
            {
                foreach (var value in item.ItemArray)
                {
                    WriteToFile.Write($"{value.ToString()} -- ");
                }
                WriteToFile.Write($"\n");
            }

            WriteToFile.Write($"Diff 1: {result2.Count()}\n");

            //result2 = ExcelHelper.GetTablesDifference(dt11, dt12);

            //WriteToFile.Write($"Diff 2: {result2.Count()}\n"); 

        }

        [TestMethod]
        public void TestMethod3()
        {
            WriteToFile.Reset();
            WriteToFile.Write($"{Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Tosca Temp Files", "xl1.xlsx")}\n");
            WriteToFile.Write($"{Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)}\n");
            WriteToFile.Write($"{Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)}\n");
            WriteToFile.Write($"{Environment.GetFolderPath(Environment.SpecialFolder.Templates)}\n");

        }

        [TestMethod]
        public void CompareTesting()
        {
            WriteToFile.Reset();

            var xl1 = @"C:\Users\okukharenko\Desktop\dt1.xlsx";
            var xl2 = @"C:\Users\okukharenko\Desktop\dt2.xlsx";


            var xldt1 = ExcelHelper.ConvertExcelToDataTable(xl1, "Sheet1");
            var xldt2 = ExcelHelper.ConvertExcelToDataTable(xl2, "Sheet1");

            var col1 = ExcelHelper.GetColumn(xldt1, 0);
            var col2 = ExcelHelper.GetColumn(xldt2, 0);
            col1 = ExcelHelper.RemoveEmptyRows(col1);
            col2 = ExcelHelper.RemoveEmptyRows(col2);


            //foreach (DataRow item in col2.Rows)
            //{
            //    foreach (var value in item.ItemArray)
            //    {
            //        WriteToFile.Write($"{value.ToString()}\n");
            //    }
            //}

            foreach (DataColumn col in col1.Columns)
            {
                WriteToFile.Write($"Col Name: {col.ColumnName}\n");
            }

            WriteToFile.Write($"Col 1 Rows: {col1.Rows.Count}\n");

            WriteToFile.Write($"Col 2 Rows: {col2.Rows.Count}\n");

            var result1 = ExcelHelper.AreColumnsEqual(col1, col2, "Column 0");
            var result2 = ExcelHelper.AreColumnsEqual(col2, col1, "Column 0");


            WriteToFile.Write($"Diff 1: {result1}\n");
            WriteToFile.Write($"Diff 1: {result2}\n");

            //foreach (var item in result1)
            //{
            //    foreach (var value in item.ItemArray)
            //    {
            //        WriteToFile.Write($"{value.ToString()}\n");
            //    }
            //}

            Assert.AreEqual(true, result1);

        }

        [TestMethod]
        public void TestMethod4()
        {

            WriteToFile.Reset();
            var aoApp = new AOOperations();
            //Excel.Application xlApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            //var result  = xlApp.Run("SAPListOfMembers", "DS_1", "FILTER");
            //var result = xlApp.Run("SAPGetVariable", "DS_1", "EDNEE5BDF1XLAKLFPS8RE67FK", "DESCRIPTION");

            var rows = aoApp.GetRows();

            foreach (var item in rows)
            {
                WriteToFile.Write($"{item}\n");
            }

            //var measures = aoApp.GetMeasures();
            var measures = aoApp.GetMeasures();
            
            foreach (var item in measures)
            {
                
                WriteToFile.Write($"{aoApp.GetDimensionTechName(item)}\n");
            }


            //var result = xlApp.Run("SAPCallMemberSelector", "DS_1", "FILTER", "EDNEE5BDF1XLAKLFPS8RE67FK");

            //object[,] dimensionArray = xlApp.Run("SAPListOfDimensions", "DS_1", "Description");

            //for (int i = 1; i <= dimensionArray.GetLength(0); i++)
            //{
            //    if ((((string)dimensionArray[i, 2]).ToUpper()).Equals("SUMMARY"))
            //    {
            //        WriteToFile.Write($"{(string)dimensionArray[i, 1]}\n");
            //    }
            //}
            //WriteToFile.Write($"{rows.Count}  {measures.Count}\n");


            //var filt = ao.SetFilter("MEASURES", "EDNEE5BDF1XLAKLFPS8LC5NA8");
            //WriteToFile.Write($"{filt}\n");

        }

        [TestMethod]
        public void TestMethod5()
        {

            WriteToFile.Reset();

            var ao = new AOOperations();

            ao.SetFilter("Summary", "!4WIF9NHC2SDFE0I72W05JXUZX");

            var result = ao.GetActiveFilters();
            
            var measures = result.Values.Where(key => !key.ToUpper().Equals("MEASURES"));



            //foreach (var item in measures)
            //{
            //    WriteToFile.Write($"{item}\n");

            //}

            foreach (var pair in result)
            {
                if (!pair.Key.ToUpper().Equals("MEASURES"))
                {
                    measures = pair.Value.Split(';');
                }
                              
            }

            foreach (var item in measures)
            {
                WriteToFile.Write($"----{item.Trim()}\n");
            }


            //foreach (var item in result)
            //{
            //    WriteToFile.Write($"{item}\n");
            //}



            //var filt = ao.SetFilter("MEASURES", "EDNEE5BDF1XLAKLFPS8LC5NA8");
            //WriteToFile.Write($"{filt}\n");

        }


    }


}
