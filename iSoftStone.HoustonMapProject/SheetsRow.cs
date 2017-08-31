// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SheetsRow.cs" company="iSoftStone">
//   Copyright 2017 iSoftStone
// </copyright>
// <summary>
//   Defines the SheetsRow type.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Isoftstone.HoustonMapProject
{
    using System;
    using System.Globalization;

    using Google.Apis.Sheets.v4.Data;

    /// <summary>
    /// The sheets row.
    /// </summary>
    public class SheetsRow
    {
        /// <summary>
        /// The _row data.
        /// </summary>
        private RowData rowData;

        /// <summary>
        /// Initializes a new instance of the <see cref="SheetsRow"/> class.
        /// </summary>
        /// <param name="rowData">
        /// The row data.
        /// </param>
        public SheetsRow(RowData rowData)
        {
            this.rowData = rowData;
        }

        /// <summary>
        /// Gets the notes.
        /// </summary>
        public string Notes => this.rowData.Values[0].FormattedValue;

        /// <summary>
        /// Gets or sets the last contacted.
        /// </summary>
        public DateTime LastContacted => DateTime.Parse(this.rowData.Values[1].FormattedValue, new CultureInfo("en-us", true));

        /// <summary>
        /// Gets or sets the time stamp.
        /// </summary>
        public DateTime TimeStamp => DateTime.Parse(this.rowData.Values[2].FormattedValue, new CultureInfo("en-us", true));

        /// <summary>
        /// Gets or sets the status.
        /// </summary>
        public string Status => this.rowData.Values[3].FormattedValue;

        /// <summary>
        /// Gets or sets the full name.
        /// </summary>
        public string FullName => this.rowData.Values[4].FormattedValue;

        /// <summary>
        /// Gets or sets the exact address.
        /// </summary>
        public string ExactAddress => this.rowData.Values[5].FormattedValue;

        /// <summary>
        /// Gets or sets the city.
        /// </summary>
        public string City => this.rowData.Values[6].FormattedValue;

        /// <summary>
        /// Gets or sets the zip code.
        /// </summary>
        public string ZipCode => this.rowData.Values[7].FormattedValue;

        /// <summary>
        /// Gets or sets the phone number.
        /// </summary>
        public ulong PhoneNumber => Convert.ToUInt64(this.rowData.Values[8].FormattedValue);

        /// <summary>
        /// Gets or sets the number of adults.
        /// </summary>
        public uint NumberOfAdults => Convert.ToUInt32(this.rowData.Values[9].FormattedValue);

        /// <summary>
        /// Gets or sets the number of children.
        /// </summary>
        public uint NumberOfChildren => Convert.ToUInt32(this.rowData.Values[10].FormattedValue);

        /// <summary>
        /// Gets or sets the number of pets.
        /// </summary>
        public uint NumberOfPets => Convert.ToUInt32(this.rowData.Values[11].FormattedValue);

        /// <summary>
        /// Gets or sets the special considerations.
        /// </summary>
        public string SpecialConsiderations => this.rowData.Values[12].FormattedValue;

        /// <summary>
        /// Gets or sets the date of birth.
        /// </summary>
        public DateTime DateOfBirth => DateTime.Parse(this.rowData.Values[13].FormattedValue, new CultureInfo("en-us", true));

        /// <summary>
        /// Gets or sets the physical descriptions.
        /// </summary>
        public string PhysicalDescriptions => this.rowData.Values[14].FormattedValue;

        /// <summary>
        /// Gets or sets a value indicating whether agreed under oath.
        /// </summary>
        public bool AgreedUnderOath => this.rowData.Values[18].FormattedValue.Equals("I Agree");

        /// <summary>
        /// Gets or sets the emergency status.
        /// </summary>
        public string EmergencyStatus => this.rowData.Values[19].FormattedValue;

        /// <summary>
        /// Gets or sets the height of water in feet.
        /// </summary>
        public decimal HeightOfWaterInFeet => Convert.ToDecimal(this.rowData.Values[20].FormattedValue);

        /// <summary>
        /// Gets or sets a value indicating whether is already in maps.
        /// </summary>
        public bool IsAlreadyInMaps
            =>
                this.rowData.Values[5].EffectiveFormat?.TextFormat?.ForegroundColor != null
                && this.rowData.Values[5].EffectiveFormat.TextFormat.ForegroundColor.Blue > 0;

        /// <summary>
        /// The reporting person.
        /// </summary>
        public ReportingPerson ReportingP => new ReportingPerson
                                                 {
                                                     Name = this.rowData.Values[15].FormattedValue,
                                                     Email = this.rowData.Values[16].FormattedValue,
                                                     PhoneNumber = this.rowData.Values[17].FormattedValue
                                                 };

        /// <summary>
        /// The reporting person.
        /// </summary>
        public class ReportingPerson
        {
            /// <summary>
            /// Gets or sets the name.
            /// </summary>
            public string Name { get; set; }

            /// <summary>
            /// Gets or sets the phone number.
            /// </summary>
            public string PhoneNumber { get; set; }

            /// <summary>
            /// Gets or sets the email.
            /// </summary>
            public string Email { get; set; }
        }
    }
}
