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
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;

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
        /// The number of rows in test spreadsheet.
        /// </summary>
        private readonly int numberOfRowsInTestSpreadsheet = 4680;

        /// <summary>
        /// The number of rows with addresses that haven't been added to the map.
        /// </summary>
        private readonly int numberOfRowsWithUnaddedAddresses = 495;

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

        /// <summary>
        /// The test with invalid sheet name.
        /// </summary>
        [TestMethod]
        public void TestWithInvalidSheetName()
        {
            this.googleSheetsAdapter.GetRows("invalid", "invalid");
        }

        /// <summary>
        /// The gets rows test.
        /// </summary>
        [TestMethod]
        public void GetsAtLeastOneRowTest()
        {
            IEnumerable<SheetsRow> ret =
                this.googleSheetsAdapter.GetRows(
                    ConfigurationManager.AppSettings["Spreadsheet Id"],
                    ConfigurationManager.AppSettings["Spreadsheet Name"]);

            Assert.AreNotEqual(ret.Count(), 0);
        }

        /// <summary>
        /// The gets right number of rows.
        /// </summary>
        [TestMethod]
        public void GetsRightNumberOfRowsTest()
        {
            IEnumerable<SheetsRow> ret =
                from r in
                this.googleSheetsAdapter.GetRows(
                    ConfigurationManager.AppSettings["Spreadsheet Id"],
                    ConfigurationManager.AppSettings["Spreadsheet Name"],
                    Convert.ToInt32(ConfigurationManager.AppSettings["Rows to skip"]))
                where
                !string.IsNullOrEmpty(r.ExactAddress)
                && !(string.IsNullOrEmpty(r.City) && string.IsNullOrEmpty(r.ZipCode))
                select r;

            Assert.AreEqual(this.numberOfRowsInTestSpreadsheet, ret.Count());
        }

        /// <summary>
        /// The gets right number of rows that aren't added to the map test.
        /// </summary>
        [TestMethod]
        public void HoustonNumbersTest()
        {
            IEnumerable<SheetsRow> ret =
                from r in
                this.googleSheetsAdapter.GetRows(
                    ConfigurationManager.AppSettings["Spreadsheet Id"],
                    ConfigurationManager.AppSettings["Spreadsheet Name"],
                    Convert.ToInt32(ConfigurationManager.AppSettings["Rows to skip"]))
                where
                r.City != null
                && (r.City.ToLower().Equals("houston") || r.City.ToLower().Equals("hoiston")
                    || r.City.ToLower().Equals("hoston") || r.City.ToLower().Equals("houston tx")
                    || r.City.ToLower().Equals("houston, tx") || r.City.ToLower().Equals("houston."))
                select r;

            Assert.AreEqual(ret.Count(), this.numberOfRowsWithUnaddedAddresses);
        }
    }
}
