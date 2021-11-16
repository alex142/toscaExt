using System.Collections.Generic;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System;
using System.Threading;
using TricentisExtensions.Modules.WindowForms;
using AutoIt;

namespace TricentisExtensions.Modules.Helpers

{
    public class FileHelper : IDisposable
    {
        private Excel.Application XlApp { get; set;}
        public Excel._Workbook XlBook { get; set; }
        public Excel._Worksheet XlSheet { get; set; }

       

        public FileHelper()
        {
        }

        public FileHelper(string resFile, string resSheet)
        {
            XlApp = new Excel.Application();
            XlBook = XlApp.Workbooks.Open(resFile);
            XlSheet = (Excel._Worksheet)XlBook.Worksheets[resSheet];
        }

        public Excel.Application GetXlApp()
        {
            return XlApp;
        }

        public void SetXlApp(Excel.Application xlApp)
        {
            XlApp = xlApp;
        }

        public Excel.Application GetActiveInstance()
        {
            try
            {
                XlApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                return XlApp;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public void TryGetInstance(string workbookName)
        {
            try
            {
                XlApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                foreach (Excel.Workbook excel in XlApp.Workbooks)
                {
                    //var name = excel.Name.Substring(0, excel.Name.IndexOf(".xlsx"));
                    if (excel.Name.Contains(workbookName)) XlBook = excel;
                }
                XlSheet = XlBook.ActiveSheet;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public void CloseAll()
        {
            try
            {
                XlApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                foreach (Excel.Workbook excel in XlApp.Workbooks)
                {
                    excel.Close();
                }
                XlApp.Quit();
            }
            catch (Exception e)
            {
                throw e;
            }
        }        

        public void OpenExcel(string filePath)
        {
            if (XlApp != null)
            {
                throw new Exception("Application is already opened");
            }
            XlApp = new Excel.Application
            {
                Visible = true
            };
            XlApp.WindowState = Excel.XlWindowState.xlMaximized;
            XlBook = XlApp.Workbooks.Open(filePath);
            XlSheet = (Excel._Worksheet)XlBook.ActiveSheet;
        }

        public void EnableComAddin(string client, string user, string password, string progId = "SapExcelAddIn")
        {
            foreach (Microsoft.Office.Core.COMAddIn addin in XlApp.COMAddIns)
            {
                if (addin.ProgId.ToUpper() == progId.ToUpper())
                {
                    if (!addin.Connect)
                    {
                        new Thread(() =>
                        {
                            var win = new Window("Logon to SAP BusinessObjects BI Platform", "", 25);
                            win.Close();

                            var logonWin = new LoginWindow(25);
                            logonWin.LogIn(client, user, password);

                            var prompts = new Window("Prompts","", 25);
                            prompts.Close();

                        }).Start();

                        addin.Connect = true;                        
                    }
                    else
                    {
                        addin.Connect = false;

                        new Thread(() =>
                        {
                            var win = new Window("Logon to SAP BusinessObjects BI Platform", "", 25);
                            win.Close();

                            var logonWin = new LoginWindow(25);
                            logonWin.LogIn(client, user, password);

                            var prompts = new Window("Prompts", "", 25);
                            prompts.Close();

                        }).Start();

                        addin.Connect = true;

                    }
                }
            }
        }

        public void EnableComAddin2(string progId = "SapExcelAddIn")
        {
            foreach (Microsoft.Office.Core.COMAddIn addin in XlApp.COMAddIns)
            {
                if (addin.ProgId.ToUpper() == progId.ToUpper())
                {
                    if (!addin.Connect)
                    {                      
                        addin.Connect = true;
                    }
                    else
                    {
                        addin.Connect = false;

                        addin.Connect = true;
                    }
                }
            }
        }


        public void FillColumn(string colName, IEnumerable<string> data)
        {
            ExcelHelper.CleanColumn(XlSheet, colName);

            ExcelHelper.WriteColumn(XlSheet, colName, data);

            XlBook.Save();
        }

        public void CleanSheet()
        {
            ExcelHelper.CleanSheet(XlSheet, 1, 1);

            XlBook.Save();
        }

        public static bool VerifyMetadata(string resFile, string refFile, string resSheet, string refSheet)
        {
            var source = ExcelHelper.ConvertExcelToDataTable(refFile, refSheet);
            var target = ExcelHelper.ConvertExcelToDataTable(resFile, resSheet);

            IEnumerable<DataRow> result = ExcelHelper.GetTablesDifference(target, source);

            return true ? result.Count() == 0 : false;
        }


        bool disposed = false;

        // Public implementation of Dispose pattern callable by consumers.
        public void Dispose()
        {
            Close(true);
            GC.SuppressFinalize(this);
        }

        // Protected implementation of Dispose pattern.
        protected virtual void Close(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                // Free any other managed objects here.
                //
            }

            if (XlBook != null)
            {
                XlBook.Save();
                XlBook.Close();
                XlBook = null;
            }
            if (XlApp != null)
            {
                XlApp.Quit();
                XlApp = null;
            }

            disposed = true;
        }

        //~FileHelper()
        //{
        //    Close(false);
        //}
    }
}
