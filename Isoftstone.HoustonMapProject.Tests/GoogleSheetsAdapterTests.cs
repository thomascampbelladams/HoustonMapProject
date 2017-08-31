// --------------------------------------------------------------------------------------------------------------------
// <copyright file="GoogleSheetsAdapterTests.cs" company="iSoftStone">
//   Copyright 2017 iSoftStone
// </copyright>
// <summary>
//   Defines the GoogleSheetsAdapterTests type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Isoftstone.HoustonMapProject.Tests
{
    using System;
    using System.Configuration;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The google sheets adapter tests.
    /// </summary>
    [TestClass]
    public class GoogleSheetsAdapterTests
    {
        /// <summary>
        /// The google sheets adapter.
        /// </summary>
        private readonly GoogleSheetsAdapter googleSheetsAdapter = new GoogleSheetsAdapter();

        /// <summary>
        /// The test with null spreadsheet id and name.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void TestWithNullSpreadsheetIdAndName()
        {
            this.googleSheetsAdapter.GetRows(null, null);
        }

        /// <summary>
        /// The test with null spreadsheet id.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void TestWithNullSpreadsheetId()
        {
            this.googleSheetsAdapter.GetRows(null, ConfigurationManager.AppSettings["Spreadsheet Name"]);
        }

        /// <summary>
        /// The test with null spreadsheet name.
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void TestWithNullSpreadsheetName()
        {
            this.googleSheetsAdapter.GetRows(ConfigurationManager.AppSettings["Spreadsheet Id"], null);
        }
    }
}
