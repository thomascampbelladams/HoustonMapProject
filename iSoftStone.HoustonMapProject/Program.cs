// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Program.cs" company="iSoftStone">
//   Copyright 2017 iSoftStone
// </copyright>
// <summary>
//   Defines the Program type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Isoftstone.HoustonMapProject
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;

    using Microsoft.Azure.WebJobs;

    // To learn more about Microsoft Azure WebJobs SDK, please see https://go.microsoft.com/fwlink/?LinkID=320976

    /// <summary>
    /// The program.
    /// </summary>
    public class Program
    {
        /// <summary>
        /// The application name.
        /// </summary>
        private static string applicationName = "iSoftStone Houston Map Project";

        // Please set the following connection strings in app.config for this WebJob to run:
        // AzureWebJobsDashboard and AzureWebJobsStorage

        /// <summary>
        /// The main.
        /// </summary>
        public static void Main()
        {
            var config = new JobHostConfiguration();
            GoogleSheetsAdapter googleSheetsAdapter = new GoogleSheetsAdapter();
            IEnumerable<SheetsRow> returnedRows =
                from r in googleSheetsAdapter.GetRows(
                    ConfigurationManager.AppSettings["Spreadsheet Id"],
                    ConfigurationManager.AppSettings["Spreadsheet Name"],
                    Convert.ToInt32(ConfigurationManager.AppSettings["Rows to skip"]))
                where r.IsAlreadyInMaps
                select r;

            if (config.IsDevelopment)
            {
                config.UseDevelopmentSettings();
            }

            foreach (SheetsRow returnedRow in returnedRows)
            {
                Console.WriteLine($"Address: {returnedRow.ExactAddress}");
                Console.WriteLine($"Is it on the map?: {returnedRow.IsAlreadyInMaps}");
            }

            var host = new JobHost();

            // The following code ensures that the WebJob will be running continuously
            host.RunAndBlock();
        }
    }
}
