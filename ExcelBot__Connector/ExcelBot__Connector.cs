//  ExcelBot__Connector.cs
//
//  Author:
//       Victor L. Senior (VLS) <betselection(&)gmail.com>
//
//  Web: 
//       http://betselection.cc/betsoftware/
//
//  Sources:
//       http://github.com/betselection/
//
//  Copyright (c) 2014 Victor L. Senior
//
//  This program is free software: you can redistribute it and/or modify
//  it under the terms of the GNU General Public License as published by
//  the Free Software Foundation, either version 3 of the License, or
//  (at your option) any later version.
//
//  This program is distributed in the hope that it will be useful,
//  but WITHOUT ANY WARRANTY; without even the implied warranty of
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//  GNU General Public License for more details.
//
//  You should have received a copy of the GNU General Public License
//  along with this program.  If not, see <http://www.gnu.org/licenses/>.

/// <summary>
/// ExcelBot connector.
/// </summary>
namespace ExcelBot__Connector
{
    // Directives
    using System;
    using System.Drawing;
    using System.IO;
    using System.Reflection;
    using System.Windows.Forms;
    using NetOffice;
    using NetOffice.OfficeApi.Enums;
    using Excel = NetOffice.ExcelApi;
    using Office = NetOffice.OfficeApi;

    /// <summary>
    /// ExcelBot connector class.
    /// </summary>
    public class ExcelBot__Connector : Form
    {
        /// <summary>
        /// The marshal object.
        /// </summary>
        private object marshal = null;

        /// <summary>
        /// The last row.
        /// </summary>
        private int lastRow = 5;

        /// <summary>
        /// The excel application.
        /// </summary>
        private Excel.Application excelApplication = null;

        /// <summary>
        /// Inits the instance.
        /// </summary>
        /// <param name="passedMarshal">Passed marshal.</param>
        public void Init(object passedMarshal)
        {
            // Set marshal
            this.marshal = passedMarshal;

            // Set icon
            this.Icon = (Icon)this.marshal.GetType().GetProperty("Icon").GetValue(this.marshal, null);

            // Set open excel file dialog
            OpenFileDialog openExcelSheetFile = new OpenFileDialog();

            // Set initial directory 
            openExcelSheetFile.InitialDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            // Set title
            openExcelSheetFile.Title = "Open ExcelBot File";

            // No multiselect
            openExcelSheetFile.Multiselect = false;

            // Validate file name
            openExcelSheetFile.ValidateNames = true;

            // Validate existence
            openExcelSheetFile.CheckFileExists = true;
            openExcelSheetFile.CheckPathExists = true;

            // Set filter to excel format
            openExcelSheetFile.Filter = "Excel files|*.xls*";

            // Get dialog result
            DialogResult dialogResult = openExcelSheetFile.ShowDialog();

            // Check if there's a valid file 
            if (dialogResult == DialogResult.OK)
            {
                // Set excel application
                this.excelApplication = new Excel.Application();

                // Open excel sheet
                this.excelApplication.Workbooks.Open(openExcelSheetFile.FileName);

                // Show Excel
                this.excelApplication.Visible = true;
                
                // Change event
                this.excelApplication.SheetChangeEvent += new Excel.Application_SheetChangeEventHandler(this.SheetChangeEventHandler);
            }
        }

        /// <summary>
        /// Processes input.
        /// </summary>
        public void Input()
        {
            // Set last
            string last = (string)this.marshal.GetType().GetProperty("Last").GetValue(this.marshal, null);

            // Check for undo
            if (last == "-U")
            {
                // Remove number
            }
            else
            {
                // Add passed number
                this.excelApplication.Cells[this.lastRow, 1].Value2 = last;

                // Rise last row
                this.lastRow++;
            }
        }

        /// <summary>
        /// Sheet change event handler.
        /// </summary>
        /// <param name="sh">Sheet object.</param>
        /// <param name="range">Passed range.</param>
        private void SheetChangeEventHandler(COMObject sh, Excel.Range range)
        {
            // Get location
            string betLocation = (this.excelApplication.Cells[this.lastRow, 2] as Excel.Range).Value2.ToString();

            // Get amount
            string betAmount = (this.excelApplication.Cells[this.lastRow, 5] as Excel.Range).Value2.ToString();

            // Trim amount comma
            betAmount = betAmount.Trim(new char[] { ',' });

            // Check there's something
            if (betAmount.Length > 0)
            {
                // Bet amount array
                string[] betAmountArray = betAmount.Split(',');

                // Trim location comma
                betLocation = betLocation.Trim(new char[] { ',' });

                // Bet location array
                string[] betLocationArray = betLocation.Split(',');

                // Iterate bets
                for (int i = 0; i < betAmountArray.Length; i++)
                {
                    // Add current bet
                    this.marshal.GetType().InvokeMember("AddBet", BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public, null, this.marshal, new object[] { this.ExcelBotToBetSoftware(betLocationArray[i], betAmountArray[i]) });
                }
            }
        }

        /// <summary>
        /// Changes ExcelBot bet format to BetSoftware's format.
        /// </summary>
        /// <returns>The bet in BetSoftware format.</returns>
        /// <param name="betLocation">Bet location.</param>
        /// <param name="betAmount">Bet amount.</param>
        private string ExcelBotToBetSoftware(string betLocation, string betAmount)
        {
            // BetSoftware's translated location
            string bstLocation = string.Empty;

            // Switch location
            switch (betLocation)
            {
            /* Dozens */

                case "DZ1":
                    bstLocation = "D1";
                    break;

                case "DZ2":
                    bstLocation = "D2";
                    break;

                case "DZ3":
                    bstLocation = "D3";
                    break;

            /* Columns */

                case "CL1":
                    bstLocation = "C1";
                    break;

                case "CL2":
                    bstLocation = "C2";
                    break;

                case "CL3":
                    bstLocation = "C3";
                    break;

            /* Double Streets */

                case "DS1":
                    bstLocation = "1-6";
                    break;

                case "DS2":
                    bstLocation = "7-12";
                    break;

                case "DS3":
                    bstLocation = "13-18";
                    break;

                case "DS4":
                    bstLocation = "19-24";
                    break;

                case "DS5":
                    bstLocation = "25-30";
                    break;

                case "DS6":
                    bstLocation = "31-36";
                    break;

            /* Corners */

                case "Q1":
                    bstLocation = "1-5";
                    break;

                case "Q2":
                    bstLocation = "2-6";
                    break;

                case "Q3":
                    bstLocation = "4-8";
                    break;

                case "Q4":
                    bstLocation = "5-9";
                    break;

                case "Q5":
                    bstLocation = "7-11";
                    break;

                case "Q6":
                    bstLocation = "8-12";
                    break;

                case "Q7":
                    bstLocation = "10-14";
                    break;

                case "Q8":
                    bstLocation = "11-15";
                    break;

                case "Q9":
                    bstLocation = "13-17";
                    break;

                case "Q10":
                    bstLocation = "14-18";
                    break;

                case "Q11":
                    bstLocation = "16-20";
                    break;

                case "Q12":
                    bstLocation = "17-21";
                    break;

                case "Q13":
                    bstLocation = "19-23";
                    break;

                case "Q14":
                    bstLocation = "20-24";
                    break;

                case "Q15":
                    bstLocation = "22-26";
                    break;

                case "Q16":
                    bstLocation = "23-27";
                    break;

                case "Q17":
                    bstLocation = "25-29";
                    break;

                case "Q18":
                    bstLocation = "26-30";
                    break;

                case "Q19":
                    bstLocation = "28-32";
                    break;

                case "Q20":
                    bstLocation = "29-33";
                    break;

                case "Q21":
                    bstLocation = "31-35";
                    break;

                case "Q22":
                    bstLocation = "32-36";
                    break;

            /* Streets */

                case "S1":
                    bstLocation = "1-3";
                    break;

                case "S2":
                    bstLocation = "4-6";
                    break;

                case "S3":
                    bstLocation = "7-9";
                    break;

                case "S4":
                    bstLocation = "10-12";
                    break;

                case "S5":
                    bstLocation = "13-15";
                    break;

                case "S6":
                    bstLocation = "16-18";
                    break;

                case "S7":
                    bstLocation = "19-21";
                    break;

                case "S8":
                    bstLocation = "22-24";
                    break;

                case "S9":
                    bstLocation = "25-27";
                    break;

                case "S10":
                    bstLocation = "28-30";
                    break;

                case "S11":
                    bstLocation = "31-33";
                    break;

                case "S12":
                    bstLocation = "34-36";
                    break;

            /* Splits */

                case "SPL1/4":
                    bstLocation = "1-4";
                    break;

                case "SPL2/5":
                    bstLocation = "2-5";
                    break;

                case "SPL3/6":
                    bstLocation = "3-6";
                    break;

                case "SPL7/10":
                    bstLocation = "7-10";
                    break;

                case "SPL8/11":
                    bstLocation = "8-11";
                    break;

                case "SPL9/12":
                    bstLocation = "9-12";
                    break;

                case "SPL13/16":
                    bstLocation = "13-16";
                    break;

                case "SPL14/17":
                    bstLocation = "14-17";
                    break;

                case "SPL15/18":
                    bstLocation = "15-18";
                    break;

                case "SPL19/22":
                    bstLocation = "19-22";
                    break;

                case "SPL20/23":
                    bstLocation = "20-23";
                    break;

                case "SPL21/24":
                    bstLocation = "21-24";
                    break;

                case "SPL25/28":
                    bstLocation = "25-28";
                    break;

                case "SPL26/29":
                    bstLocation = "26-29";
                    break;

                case "SPL27/30":
                    bstLocation = "27-30";
                    break;

                case "SPL31/34":
                    bstLocation = "31-34";
                    break;

                case "SPL32/35":
                    bstLocation = "32-35";
                    break;

                case "SPL33/36":
                    bstLocation = "33-36";
                    break;

            /* Even chances, straight */
                default:

                    bstLocation = betLocation;
                    break;
            }

            // Return translated location
            return betAmount + "@" + bstLocation;
        }
    }
}