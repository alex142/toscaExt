using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using AutoIt;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TricentisExtensions.Modules;
using TricentisExtensions.Modules.CustomEngine;
using TricentisExtensions.Modules.Helpers;
using TricentisExtensions.Test.Utils;
using Excel = Microsoft.Office.Interop.Excel;


namespace TricentisExtensions.Test.Tests
{
    [TestClass]
    public class ExcelTests
    {
        [TestMethod]
        public void TestExcelWriter()
        {
            var resFile = @"C:\Users\okukharenko\Desktop\Result.xlsx";
            //var refFile = @"C:\Users\okukharenko\Desktop\Refference.xlsx";
            var resSheet = "Test";
            //var dimSheet = "GL Details Dimensions";

            var xlApp = new AOOperations();
            var res = new FileHelper(resFile, resSheet);

            var layout = xlApp.GetRows();
            layout.AddRange(xlApp.GetMeasures());
            var dimensions = xlApp.GetDimensions();
            var prompts = xlApp.GetVariables("PROMPTS_FILLED");

            res.FillColumn("Dimensions", dimensions);
            res.FillColumn("Default Layout", layout);
            res.FillColumn("Variable Name", prompts.Keys);
            res.FillColumn("Variable Value", prompts.Values);


            //var result = ResultVerification.VerifyMetadata(resFile, refFile, defSheet, defSheet);

            //Assert.IsTrue(result);
        }

        [TestMethod]
        public void TestCleanSheet()
        {
            var resFile = @"C:\Users\okukharenko\Desktop\Result.xlsx";
            var resSheet = "Test";

            var res = new FileHelper(resFile, resSheet);

            ExcelHelper.CleanColumn(res.XlSheet, "Dimensions");
            res.XlBook.Save();
            res.XlBook.Close();

            //var result = ResultVerification.VerifyMetadata(resFile, refFile, defSheet, defSheet);

            //Assert.IsTrue(result);
        }

        [TestMethod]
        public void TestVars()
        {
            var resFile = @"C:\Users\okukharenko\Desktop\Result.xlsx";
            var defSheet = "Variables";

            var xlApp = new AOOperations();
            var res = new FileHelper(resFile, defSheet);

            var vars = xlApp.GetVariables("PROMPTS_FILLED");
        }

        [TestMethod]
        public void TestWriteColumn()
        {
            var resFile = @"C:\Users\okukharenko\Desktop\dt1.xlsx";
            var defSheet = "Sheet 1";
            var col1 = "Veriable Name";
            var col2 = "Variable value";


            using (var res = new FileHelper(resFile, defSheet))
            {

                var aoApp = new AOOperations();

                var vars = aoApp.GetVariables("PROMPTS_FILLED");

                //var vars = new Dictionary<string, string>() {
                //{"key 1", "value 1" },
                //{"key 2", "value 2"},
                //{"key 3", "value 3" },
                //{"key 4", "value 4"}};

                res.FillColumn(col1, vars.Keys);
                res.FillColumn(col2, vars.Values);
            }            
        }

        [TestMethod]
        public void TestAddin()
        {
            var xlApp = new FileHelper();

            var app = xlApp.GetActiveInstance();

            foreach (Excel.Workbook book in app.Workbooks)
            {
                book.Save();
                book.Close();
            }

            xlApp.Dispose();
        }

        [TestMethod]
        public void TestThreading()
        {
            var xlApp = new FileHelper();

            xlApp.OpenExcel(@"C:\Users\okukharenko\Desktop\dev2.xlsx");

            xlApp.EnableComAddin("","","");

            //(new Thread(() =>
            //{
            //    AutoItX.WinWaitActive("Logon to SAP BusinessObjects BI Platform", "", 12);
            //    if (AutoItX.WinExists("Logon to SAP BusinessObjects BI Platform", "") == 1)
            //    {
            //        AutoItX.WinClose("Logon to SAP BusinessObjects BI Platform", "");
            //    }
            //    //AutoItX.WinActivate("Sign in");
            //})).Start();

            //var logon = xlApp.GetXlApp().Run("SAPLogon", "DS_1", "010", "T-RNA-GLB", "Testing@123");

            //var restart = xlApp.GetXlApp().Run("SAPExecuteCommand", "Refresh", "DS_1");

            //System.IO.File.WriteAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", $"Error - {xlApp.GetXlApp().Run("SapGetProperty", "LastError", "Text")}\n");
            //System.IO.File.AppendAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", $"Logon : {logon}\nRestart : {restart}");
            //var lst = xlApp.GetXlApp().Run("SAPGetProperty", "IsDataSourceActive ", "DS_1");
            //var lst1 = xlApp.GetXlApp().Run("SAPGetProperty", "IsConnected", "DS_1");
            //System.IO.File.AppendAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", $"DS is Active : {lst}\nDS is connected: {lst1}\n");

            //xlApp.GetXlApp().Run("SAPExecuteCommand", "Refresh", "DS_1");
            //var aoApp = new AOOperations();
            //aoApp.SetVariable("Company Code", "1001");
            //aoApp.SetFilter("Fiscal Year", @"4025/2018");
            //aoApp.SetFilter("GL Account", "14600001");




            //xlApp.Dispose();
        }

        [TestMethod]
        public void TestSaveAndClose()
        {

            var xlApp = new FileHelper();

            xlApp.TryGetInstance("EP");

            if (xlApp.GetXlApp() != null)
            {
                xlApp.GetXlApp().ActiveWorkbook.Close(true);
                xlApp.XlBook = null;
            }

            xlApp.Dispose();
        }

        [TestMethod]
        public void TestWriteColumnFromReport()
        {
            var srcFile = @"C:\Tosca_Projects\Tosca_Workspaces\RandA\Downloads\APF Report  - 02-14-20_08_36.xlsx";
            var srcSheet = "APF by Time Dimension";

            var dt = ExcelHelper.ConvertExcelToDataTable(srcFile, srcSheet);
            ExcelHelper.DeleteDataTableColumns(dt, 1);
            ExcelHelper.DeleteDataTableRows(dt, 0, 6);

            WriteToFile.Reset();
            //WriteToFile.Write($"{dt.Rows.Count} X {dt.Columns.Count}");

            foreach (DataRow item in dt.Rows)
            {
                foreach (var value in item.ItemArray)
                {
                    WriteToFile.Write($"{value.ToString()}\n");
                }
            }

            //foreach (DataColumn col in dt.Columns)
            //{
            //    WriteToFile.Write($"Col Name: {col.ColumnName}\n");
            //}
        }
    }
}
