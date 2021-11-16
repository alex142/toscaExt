using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using AutoIt;
using TricentisExtensions.Modules.WindowForms;


namespace TricentisExtensions.Test.Tests
{
    /// <summary>
    /// Summary description for AutoItTest
    /// </summary>
    [TestClass]
    public class AutoItTest
    {
        public AutoItTest()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod]
        public void TestMethod1()
        {
            var window = new CancelDialog(5);

            var result = window.WaitWindowDisappear(20);

            System.IO.File.WriteAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", $"{result}\n");

            //var win = new Window("Prompts");
            //var clientCtr = new Control(win, "mClientTextBox");
            //var uNameCtr = new Control(win, "mUserTextBox");
            //var passCtr = new Control(win, "mPasswordTextBox");
            //var okBtn = new Control(win, "mOkButton");

            //clientCtr.SetText("030");
            //uNameCtr.SetText("T-RNA-GLB");
            //passCtr.SetText(@"Testing@123");
            //win.Activate();
            //okBtn.Click();

            //win.Close();

            //AutoItX.ControlSend();
        }

        [TestMethod]
        public void TestWaitForPrompt()
        {
            try
            {
                _ = new PromptWindow(5);
            }
            catch (NullReferenceException e)
            {
                System.IO.File.WriteAllText(@"C:\Users\okukharenko\Desktop\WriteLines.txt", $"{e.Message}\n");
                Assert.Fail();
            }

        }
    }
}
