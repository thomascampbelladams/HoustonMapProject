// --------------------------------------------------------------------------------------------------------------------
// <copyright file="GoogleSheetsAdapter.cs" company="iSoftStone">
//   Copyright 2017 iSoftStone
// </copyright>
// <summary>
//   Defines the GoogleSheetsAdapter type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Isoftstone.HoustonMapProject
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.IO;
    using System.Linq;
    using System.Threading;

    using Google.Apis.Auth.OAuth2;
    using Google.Apis.Services;
    using Google.Apis.Sheets.v4;
    using Google.Apis.Sheets.v4.Data;
    using Google.Apis.Util.Store;

    /// <summary>
    /// The google sheets adapter.
    /// </summary>
    public class GoogleSheetsAdapter
    {
        /// <summary>
        /// The scopes.
        /// </summary>
        private static readonly string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };

        /// <summary>
        /// The service.
        /// </summary>
        private readonly SheetsService service;

        /// <summary>
        /// Initializes a new instance of the <see cref="GoogleSheetsAdapter"/> class.
        /// </summary>
        public GoogleSheetsAdapter()
        {
            UserCredential credential;

            using (
                FileStream stream = new FileStream(
                    ConfigurationManager.AppSettings["Client Secret File Name"],
                    FileMode.Open,
                    FileAccess.Read))
            {
                string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");
                credential =
                    GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
            }

            this.service = new SheetsService(new BaseClientService.Initializer
                                                 {
                                                     HttpClientInitializer = credential,
                                                     ApplicationName = ConfigurationManager.AppSettings["Application Name"]
                                                 });
        }

        /// <summary>
        /// The get rows.
        /// </summary>
        /// <param name="sheetId">
        /// The sheet id.
        /// </param>
        /// <param name="sheetName">
        /// The sheet name.
        /// </param>
        /// <param name="rowsToSkip">
        /// The rows To Skip.
        /// </param>
        /// <returns>
        /// The <see cref="List"/>.
        /// </returns>
        public List<SheetsRow> GetRows(string sheetId, string sheetName, int rowsToSkip = 0)
        {
            if (string.IsNullOrEmpty(sheetId))
            {
                throw new ArgumentNullException(nameof(sheetId));
            }

            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ArgumentNullException(nameof(sheetName));
            }

            List<SheetsRow> ret = new List<SheetsRow>();
            SpreadsheetsResource.GetRequest request = this.service.Spreadsheets.Get(sheetId);
            request.IncludeGridData = true;
            Spreadsheet responseSpreadsheet = request.Execute();
            Sheet spreadSheet = (from s in responseSpreadsheet.Sheets
                                 where s.Properties.Title.Equals(sheetName)
                                 select s).FirstOrDefault();
            int count = 0;

            if (spreadSheet != null)
            {
                foreach (RowData rowData in spreadSheet.Data[0].RowData)
                {
                    if (rowsToSkip > 0)
                    {
                        count++;

                        if (count <= rowsToSkip)
                        {
                            continue;
                        }
                    }

                    ret.Add(new SheetsRow(rowData));
                }
            }

            return ret;
        }
    }
}
