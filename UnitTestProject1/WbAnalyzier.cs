using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Converter;
using ExcelRLibrary;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTestProject1
{
    /// <summary>
    /// Summary description for WbAnalyzier
    /// </summary>
    [TestClass]
    public class WbAnalyzier
    {
        private WorkbooksAnalyzier analyzier;
        public WbAnalyzier()
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
        public void CreateInstanceFromCustomWb()
        {
            const string testWbPath = @"D:\Google Drive\Rway\Visual Studio 2013\Projects\ConvertDB\UnitTestProject1\bin\Release\Atests.xlsx";
            analyzier = new WorkbooksAnalyzier(testWbPath);
            analyzier.CheckWorkbook(testWbPath);
            analyzier.ComparedColumns.ForEach(pair =>
            {
                Assert.AreEqual(pair.Key.CodeName,pair.Value.First());
            });
        }
    }
}
