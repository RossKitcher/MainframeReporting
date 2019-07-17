using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace start
{
    public partial class Main_Form : Form
    {
        public Main_Form()
        {
            InitializeComponent();
        }

        private void runReportButton_Click(object sender, EventArgs e)
        {

            
            if (reportBgWorker.IsBusy != true)
            {
                reportStatusProgress.Maximum = 100;
                reportStatusProgress.Step = 1;
                reportStatusProgress.Value = 0;
                reportStatusLabel.Text = "Running";
                // Start the asynchronous operation.
                reportBgWorker.RunWorkerAsync();
            }
        }

        private void reportBgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var backgroundWorker = sender as BackgroundWorker;
            backgroundWorker.ReportProgress(0);
            
            backgroundWorker.ReportProgress(10);
            
            string fileName = textBox1.Text;

            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;

            try
            {
                oXL = new Excel.Application();

                //Get existing workbook
                oWB = (Excel._Workbook)(oXL.Workbooks.Open(fileName));
                var xlSheets = oWB.Sheets as Excel.Sheets;

                //Create APP and vlookup for Warton4 data
                oSheet = ChangeSheet("Warton4 Data", oWB);
                InsertAppColumn(oSheet); //Call function to insert column

                //Create pivot for Warton4 data on new sheet
                var PivotSheet = (Excel.Worksheet)xlSheets.Add(Type.Missing, xlSheets[4], Type.Missing, Type.Missing);
                PivotSheet.Name = "Warton4 Pivot"; //Name new sheet
                CreatePivotSheet(oWB, oSheet, PivotSheet, "Warton4 Pivot", false); //Call function to create pivot

                backgroundWorker.ReportProgress(25);
                

                //Do the same for Brought data
                oSheet = ChangeSheet("Brought Data", oWB);
                InsertAppColumn(oSheet);
                //Create Pivot
                PivotSheet = (Excel.Worksheet)xlSheets.Add(Type.Missing, xlSheets[4], Type.Missing, Type.Missing);
                PivotSheet.Name = "Brought Pivot";
                CreatePivotSheet(oWB, oSheet, PivotSheet, "Brought Pivot", false);

                backgroundWorker.ReportProgress(35);

                //Now CHAD data
                oSheet = ChangeSheet("CHAD Data", oWB);
                InsertAppColumn(oSheet);
                //CHAD Pivot
                PivotSheet = (Excel.Worksheet)xlSheets.Add(Type.Missing, xlSheets[4], Type.Missing, Type.Missing);
                PivotSheet.Name = "CHAD Pivot";
                CreatePivotSheet(oWB, oSheet, PivotSheet, "CHAD Pivot", false);

                backgroundWorker.ReportProgress(50);

                //Warton2 data
                oSheet = ChangeSheet("Warton2 Data", oWB);
                InsertAppColumn(oSheet);
                //Warton2 Pivot
                PivotSheet = (Excel.Worksheet)xlSheets.Add(Type.Missing, xlSheets[4], Type.Missing, Type.Missing);
                PivotSheet.Name = "Warton2 Pivot";
                CreatePivotSheet(oWB, oSheet, PivotSheet, "Warton2 Pivot", false);

                backgroundWorker.ReportProgress(60);

                //Once all pivots are created, create a sheet unifying all data from created pivots,
                //this is done to prevent excel from reaching the row limit.
                CreateConsolidatedSheet(oWB, xlSheets, backgroundWorker);

                backgroundWorker.ReportProgress(100);



                //Give user control of excel spreadsheet once all processing is finished
                oXL.Visible = true;
                oXL.UserControl = true;
            }
            //Catch any errors that may arise
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                errorMessage = String.Concat(errorMessage, ". Ensure correct excel file has been used.");

                MessageBox.Show(errorMessage, "Error");

                backgroundWorker.ReportProgress(0);
            }

        }

        //Change worksheet
        private Excel._Worksheet ChangeSheet(string sheetName, Excel._Workbook oWB)
        {
            
            if (sheetName == "Warton4 Data")
                return (Excel.Worksheet)oWB.Worksheets["Warton4 Data"];
            else if (sheetName == "Brought Data")
                return (Excel.Worksheet)oWB.Worksheets["Brought Data"];
            else if (sheetName == "CHAD Data")
                return (Excel.Worksheet)oWB.Worksheets["CHAD Data"];
            else if (sheetName == "Warton2 Data")
                return (Excel.Worksheet)oWB.Worksheets["Warton2 Data"];
            else
                return null;
        }


        //Creates pivot table
        private void CreatePivotSheet(Excel._Workbook workbook, Excel._Worksheet dataSheet, Excel._Worksheet pivotSheet, string tableName, bool unified)
        {
            //If consolidated sheet, range selects up until column C, if normal pivot, selects up until K
            string col;
            if (unified == true)
            {
                col = "C";
            }
            else
            {
                col = "K";
            }
            //Get last used row
            var lastUsedRow = getLastUsedRow(dataSheet); //Select all data from starting cell to last column + row
            var dataRange = dataSheet.get_Range("A1", col + lastUsedRow);
            var pivotRange = pivotSheet.Cells[1, 1]; //Select target location
            var oPivotCache = (Excel.PivotCache)workbook.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, dataRange); //Create cache specifying data is coming from a table
            var oPivotTable = (Excel.PivotTable)pivotSheet.PivotTables().Add(PivotCache: oPivotCache, TableDestination: pivotRange, TableName: tableName);//Create table

            if (unified == true)//If consolidated sheet
            {
                //Set Row field to 'APP'
                var RowPivotField = (Excel.PivotField)oPivotTable.PivotFields("APP");
                RowPivotField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                //Set Values field to 'Total'
                var SumPivotField = (Excel.PivotField)oPivotTable.PivotFields("Total");
                SumPivotField.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                SumPivotField.Function = Excel.XlConsolidationFunction.xlSum;
                SumPivotField.Name = "CPU Time";

                //Set Column field to 'LPAR'
                var ColPivotField = (Excel.PivotField)oPivotTable.PivotFields("LPAR");
                ColPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;

            }
            else //If normal sheet
            {
                //Set Row field to 'APP'
                Excel.PivotField RowPivotField = (Excel.PivotField)oPivotTable.PivotFields("APP");
                RowPivotField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                //Set Values field to 'CPUTIME'
                Excel.PivotField SumPivotField = (Excel.PivotField)oPivotTable.PivotFields("CPUTIME");
                SumPivotField.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                SumPivotField.Function = Excel.XlConsolidationFunction.xlSum;
                SumPivotField.Name = "CPU Time";
            }
        }

        //Insert column 'APP' for the given worksheet
        private void InsertAppColumn(Excel._Worksheet oWS)
        {

            Excel.Range oRng;


            oRng = oWS.Columns["D"]; // Select column D

            //Insert column and shift existing columns to the right
            oRng.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight,
                Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            oWS.Cells[1, 4] = "APP"; //Set header of the column

            //Get the last row in the spreadsheet containing data
            var lastUsedRow = getLastUsedRow(oWS);

            oRng = oWS.get_Range("D2", "D" + lastUsedRow); //Select all rows in column D below the title
            oRng.Formula = "=VLOOKUP(C2,Lookup!A:B,2,TRUE)"; //Set each cell to perform a vlookup
        }

        //Create new column for each pivot containing the corresponding LPAR
        private void CreateLPAR(Excel._Worksheet oSheet, string sheetName)
        {
            //Get the last row in the spreadsheet containing data
            var lastUsedRow = getLastUsedRow(oSheet);

            var oRng = oSheet.get_Range("C3", "C" + (lastUsedRow - 1)); //Get range of data excluding titles & grand total
            oRng.Value2 = sheetName; //Change contents to the LPAR
            oSheet.Cells[2, 3] = "LPAR";//Set title of column
        }

        //Creates unifying sheet
        private void CreateConsolidatedSheet(Excel._Workbook workbook, Excel.Sheets xlSheets, BackgroundWorker backgroundWorker)
        {
            //Create empty sheet
            var ConsolidatedSheet = (Excel.Worksheet)xlSheets.Add(Type.Missing, xlSheets[8], Type.Missing, Type.Missing);
            ConsolidatedSheet.Name = "Consolidated Data";//Rename it

            //Create empty pivot sheet
            var ConsolidatedPivot = (Excel.Worksheet)xlSheets.Add(Type.Missing, xlSheets[9], Type.Missing, Type.Missing);
            ConsolidatedPivot.Name = "Consolidated Pivot";

            backgroundWorker.ReportProgress(70);

            //Add LPAR column to all pivot's by calling CreateLPAR()
            var Warton4Pivot = (Excel.Worksheet)workbook.Sheets[8];
            CreateLPAR(Warton4Pivot, "Warton4");

            var BroughtPivot = (Excel.Worksheet)workbook.Sheets[7];
            CreateLPAR(BroughtPivot, "Brought");

            var CHADPivot = (Excel.Worksheet)workbook.Sheets[6];
            CreateLPAR(CHADPivot, "CHAD");

            var Warton2Pivot = (Excel.Worksheet)workbook.Sheets[5];
            CreateLPAR(Warton2Pivot, "Warton2");

            //Get the last row in the spreadsheet containing data
            var lastUsedRow = getLastUsedRow(Warton4Pivot);

            var Warton4Range = Warton4Pivot.get_Range("A2", "C" + (lastUsedRow - 1));//Get range of data including title, excluding grandtotal
            var ConsRange = ConsolidatedSheet.Range["A1", "C" + (lastUsedRow - 2)];//Get range of cells to copy data too
            Warton4Range.Copy(ConsRange);//Copy the data

            var startOfNextData = lastUsedRow - 1;//Set where the next set of data needs to be appended to

            //Find new last used row
            lastUsedRow = getLastUsedRow(BroughtPivot);
            var BroughtRange = BroughtPivot.get_Range("A3", "C" + (lastUsedRow - 1));//Get data from Brough
            ConsRange = ConsolidatedSheet.Range["A" + startOfNextData, "C" + (startOfNextData + lastUsedRow - 2)];//Copy target location
            BroughtRange.Copy(ConsRange);//Move Brought's data

            startOfNextData += lastUsedRow - 3;
            lastUsedRow = getLastUsedRow(CHADPivot);

            var CHADRange = CHADPivot.get_Range("A3", "C" + (lastUsedRow - 1));
            ConsRange = ConsolidatedSheet.Range["A" + startOfNextData, "C" + (startOfNextData + lastUsedRow - 2)];
            CHADRange.Copy(ConsRange);
            backgroundWorker.ReportProgress(80);

            startOfNextData += lastUsedRow - 3;
            lastUsedRow = getLastUsedRow(Warton2Pivot);

            var Warton2Range = Warton2Pivot.get_Range("A3", "C" + (lastUsedRow - 1));
            ConsRange = ConsolidatedSheet.Range["A" + startOfNextData, "C" + (startOfNextData + lastUsedRow - 2)];
            Warton2Range.Copy(ConsRange);

            backgroundWorker.ReportProgress(90);

            //Now create Pivot for the unified data using the unified sheet
            CreatePivotSheet(workbook, ConsolidatedSheet, ConsolidatedPivot, "Consolidated Pivot", true);
            FormatConsolidatedSheet(workbook, ConsolidatedPivot);
            CreateChart(ConsolidatedPivot, backgroundWorker);
        }

        private void FormatConsolidatedSheet(Excel._Workbook workbook, Excel._Worksheet consolidatedPivot)
        {
            Excel.Range oRng = consolidatedPivot.get_Range("A1", "F2");
            oRng.Interior.Color = ColorTranslator.ToOle(Color.LightBlue);
            oRng.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;

            oRng = consolidatedPivot.get_Range("A3", "F" + getLastUsedRow(consolidatedPivot));
            oRng.Interior.Color = ColorTranslator.ToOle(Color.White);
        }

        private void CreateChart(Excel._Worksheet oSheet, BackgroundWorker backgroundWorker)
        {
            //Connect to workbook
            Excel._Workbook oWB = (Excel._Workbook)oSheet.Parent;

            //Get range of data in which the chart will reference to
            Excel.Range oRng = oSheet.get_Range("A1", "C" + getLastUsedRow(oSheet));

            //Create chart using ChartWizard
            Excel._Chart oChart = (Excel._Chart)oWB.Charts.Add(Missing.Value, Missing.Value,
                Missing.Value, Missing.Value);
            oChart.ChartWizard(oRng, Excel.XlChartType.xlColumnStacked, Missing.Value,
                Excel.XlRowCol.xlRows, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            oChart.Location(Excel.XlChartLocation.xlLocationAsObject, oSheet.Name);

            //Set location, height and width of the chart
            oSheet.Shapes.Item(1).Top = 75;
            oSheet.Shapes.Item(1).Left = 600;
            oSheet.Shapes.Item(1).Width = 1304;
            oSheet.Shapes.Item(1).Height = 419;

            backgroundWorker.ReportProgress(95);
        }

        //Returns last used row
        private int getLastUsedRow(Excel._Worksheet oSheet)
        {
            return oSheet.Cells.Find("*", Missing.Value,
                               Missing.Value, Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, Missing.Value, Missing.Value).Row;
        }

        private void reportBgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            reportStatusProgress.Value = e.ProgressPercentage;

        }
        // This event handler deals with the results of the background operation.
        private void reportBgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                reportStatusLabel.Text = "Canceled!";
            }
            else if (e.Error != null)
            {
                reportStatusLabel.Text = "Error: " + e.Error.Message;
            }
            else
            {
                reportStatusLabel.Text = "Finished";
            }
        }
        private void browseFilesButton_Click(object sender, EventArgs e)
        {
            // Show the dialog and get result.
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName; //Save filepath to variable
                try
                {
                    //Set textbox to the filepath
                    textBox1.Text = file;
                }
                catch (IOException)
                {
                }
            }
        }

        private void dataFormatButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Rules for data:\n4 Data sheets required, each with names:\n\'Brought Data\', \'CHAD Data\', \'Warton2 Data\', \'Warton4 Data\'.\nLookup table must be included in the same excel workbook with name,\n\'Lookup\'", "Data Format");
        }

        private void userGuideButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Instructions:\n1.Browse files and upload Mainframe Data.\n2.Run Report. ~2 mins\n3.Once finished, edit any charts as needed.\n4.Save & upload to sharepoint under filename:\n\t-\'Top 25 Batch Utilisation April 2019.\'\n\t(current month)", "Instructions");
        }


    }
}
