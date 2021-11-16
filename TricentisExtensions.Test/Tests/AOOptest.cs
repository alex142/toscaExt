using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using TricentisExtensions.Modules;
using TricentisExtensions.Test.Utils;
using TricentisExtensions.Modules.CustomEngine;



namespace TricentisExtensions.Test.Tests
{
    [TestClass]
    public class AOOptest
    {
        AOOperations xlApp;
        [TestInitialize]
        public void SetUp()
        {
            WriteToFile.Reset();
            xlApp = new AOOperations("DS_2");
        }

        [TestCleanup]
        public void CleanUp()
        {
            xlApp = null;
        }

        [TestMethod]
        public void TestMethod1()
        {
            var vars = xlApp.GetVariables();

            var allDims = xlApp.GetDimensions();

            var filters = xlApp.GetActiveFilters();

            var measures = xlApp.GetMeasures();

            var allDimsTech = xlApp.GetDimensionTechNames();

            var defRows = xlApp.GetRows();

            var defCols = xlApp.GetColumns();

            var filterNames = xlApp.GetActiveFilterNames();

            WriteToFile.Write($"Variables : {vars.Count}\n");
            foreach (var pair in vars)
            {
                WriteToFile.Write($"{pair.Key} - {pair.Value}\n");
            }

            WriteToFile.Write($"Filters : {filters.Count}\n");
            foreach (var pair in filters)
            {
                WriteToFile.Write($"{pair.Key} - {pair.Value}\n");
            }

            WriteToFile.Write($"Filter Names : {filterNames.Count}\n");
            foreach (var name in filterNames)
            {
                WriteToFile.Write($"Name : {name}\n");
            }

            WriteToFile.Write($"\nMeasures\n");
            foreach (var measure in measures)
            {
                WriteToFile.Write($"{measure}\n");
            }
            WriteToFile.Write($"\nCount of all dimensions : {allDims.Length}\nCount of all dim Tech Names : {allDimsTech.Length}\nCount of rows in layout : {defRows.Count}\nCount of columns : {defCols.Count}");

            WriteToFile.Write($"\nName - Tech Name\n");
            for (int i = 0; i < allDims.Length; i++)
            {
                WriteToFile.Write($"{allDims[i]} - {allDimsTech[i]}\n");
            }

            WriteToFile.Write($"\nRows\n");
            foreach (var row in defRows)
            {
                WriteToFile.Write($"{row}\n");
            }

            WriteToFile.Write($"\nColumns\n");
            foreach (var col in defCols)
            {
                WriteToFile.Write($"{col}\n");
            }
        }

        [Ignore]
        [TestMethod]
        public void TestMethod0()
        {
            var vars = xlApp.GetVariables();
            var prompts = vars.Keys;
            var techPrompts = new List<string>();


            foreach (var prompt in prompts)
            {
                techPrompts.Add(xlApp.GetVariableTechnicalName(prompt));
                WriteToFile.Write($"{xlApp.GetVariableTechnicalName(prompt)}\n");
            }

            var techs = techPrompts.Where(x => !string.IsNullOrEmpty(x));

            WriteToFile.Write($"\nInitial : {prompts.Count}; Technical : {techPrompts.Count}; techs : {techs.Count()} \n");
        }


        [TestMethod]
        public void TestMethod2()
        {
            var res = xlApp.SetFilter("Vendor", "ALLMEMBERS");
            res = xlApp.SetFilter("GL Account", "ALLMEMBERS");

            var dict = new Dictionary<string, string> {
                {"Fiscal Year", @"4025/2018"},
                {"GL Account","14600001"}
            };
            res = xlApp.SetFilters(dict);
            Assert.AreEqual(1, res);
            res = xlApp.ResetAllFilters();
            Assert.AreEqual(1, res);
        }


        [TestMethod]
        public void TestMethod3()
        {
            var message = xlApp.GetLastError();
            WriteToFile.Write($"{message}\n");
            var measures = xlApp.GetMeasures();
            var activeFilters = xlApp.GetActiveFilterNames();

            foreach (var measure in measures)
            {
                WriteToFile.Write($"{measure}\n");
            }

            foreach (var filter in activeFilters)
            {
                WriteToFile.Write($"{filter}\n");
            }
        }

        [TestMethod]
        public void TestMethod4()
        {
            //xlApp.ActivateVariableSubmit();
            //var res = xlApp.SetVariable1("!CDS_F_2CIFILEDGER", "0L");
            //Assert.AreEqual(1, res);
            //WriteToFile.Write($"Initial : {res}");

            var res = xlApp.AddToColumns("Fiscal Year");
            Assert.AreEqual(1, res);

            //var dict = new Dictionary<string, string> {
            //    {"Company Code", "1001"},
            //    {"Fiscal Period", "10"}
            //};
            //xlApp.PauseRefresh();
            //foreach (var pair in dict)
            //{
            //    WriteToFile.Write($"Initial : {pair.Key} ; Technical: {xlApp.GetVariableTechnicalName(pair.Key)}");
            //}
            //xlApp.ActivateRefresh();

            ////res = xlApp.SetVariables(dict);
            //Assert.AreEqual(1, res);


        }
    }
}
