using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Diagnostics;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;

namespace MyRevitCommands
{
    internal class GenericToevoegen
    {
        public Result GenericExecute(ExternalCommandData commandData, ref string message, ElementSet elements,
            string MapPath, ArrayList revitSchedules, string fileName)
        {
            string xlSheetName;
            Directory.CreateDirectory(MapPath);
            var r = Result.Succeeded;
            // Is pas nodig vanaf EPPlus v5. Hier gebruiken we v4 nog die ondere een andere license valt maar depreciated is 
            // ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var emptyFirstCellDocument = "";
            var TotalDocument = "";

            // Get UIDocument
            var uidoc = commandData.Application.ActiveUIDocument;
            // Get Document
            var doc = uidoc.Document;

            var col = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule));

            // Export Options on how the .txt file will look. Text "" is gone
            // FieldDelimiter is TAB replaced with ,
            var opt = new ViewScheduleExportOptions
            {
                TextQualifier = ExportTextQualifier.None,
                FieldDelimiter = ","
            };

            // Formating for writing to xlsx
            var format = new ExcelTextFormat
            {
                Culture = CultureInfo.InvariantCulture,
                Delimiter = ','

                // Other properties
                // EOL, DataTypes, Encoding, SkipLinesBeginning/End
            };

            // Creates new excelpackage 
            using (var excelEngine = new ExcelPackage())
            {
                using (var xlPackage = new ExcelPackage())
                {
                    var wbUitzondering = xlPackage.Workbook.Worksheets.Add("Uitzondering");
                    var wbAllesIngevuld = xlPackage.Workbook.Worksheets.Add("AllesIngevuld");
                    foreach (ViewSchedule vs in col)
                        // Searches for schedules containing AE E60 M52 en M57 ventilatierooster
                        // dit zijn de schedules waarbij het met aantallen is.

                        if (CheckValues(vs.Name, revitSchedules))
                        {
                            // Checks if vs.name is not longer than 30 characters because Excels crashes.
                            // Keep in mind for the vs.names that they have to be different under 30 characters
                            // Excel does not like the same names
                            if (vs.Name.Length > 30)
                                xlSheetName = vs.Name.Substring(0, 30);
                            else
                                xlSheetName = vs.Name;
                            //create a WorkSheet
                            var ws1 = excelEngine.Workbook.Worksheets.Add(xlSheetName);
                            // Export c:\\temp --> Will be save as
                            var filename = Environment.UserName + vs.Name;
                            // Exports .csv file to mappath
                            vs.Export(MapPath, filename + ".csv", opt);

                            // Document settings, start empty, sets pathfile, reads all lines, delimiter, start with 1 
                            var normalDocument = "";
                            var StringPathFile = MapPath + filename + ".csv";
                            var lines = File.ReadAllLines(StringPathFile);
                            char[] delimitChars = {','};
                            var i = 1;
                            var strLineCompare = ""; 


                            foreach (var line in lines)
                            {
                                // Gets first 2 row of each Schedule, name and properties.
                                // This is done for visibility and making the excel easier to read
                                // Means, header and first row, count, family, type, ... 
                                if (i < 3)
                                {
                                    // Every line in the doc gets added as a newLine.
                                    emptyFirstCellDocument += line + Environment.NewLine;
                                    i++;
                                }

                                // Looks for first value if this is null or blank ""
                                // Geo-IT number has to be first for this line
                                // Checks if it is blank or not
                                if (line.Split(delimitChars)[0] == "")
                                    // Empty first cell .csv line
                                    emptyFirstCellDocument += line + Environment.NewLine;
                                else
                                    // If the first cell is filled, continue on this .csv line
                                    // Only checks for blanks or not!
                                    normalDocument += line + Environment.NewLine;

                                //if (line.Split(delimitChars)[0] == strLineCompare.Split(delimitChars)[0])
                                //{
                                //    line.Split(delimitChars)[1] + strLineCompare.Split(delimitChars)[1]

                                //    countDoc += 
                                //}


                                strLineCompare = line;
                            }
                            

                            // Gets spacing for each schedule.
                            // Increasing visibility for reading the excel.
                            emptyFirstCellDocument += Environment.NewLine;
                            TotalDocument += normalDocument + Environment.NewLine;

                            File.WriteAllText(MapPath + filename + ".csv", normalDocument);
                            var file = new FileInfo(MapPath + filename + ".csv");
                            // Adds Worksheet as first in the row because excel keeps adding in the back
                            ws1.Workbook.Worksheets.MoveToStart(xlSheetName);
                            // EPPlus method for reading a .csv to excel in the format as written above here line +-40
                            ws1.Cells["A1"].LoadFromText(file, format);
                            // the path of the file
                            var filePath = MapPath + fileName + ".xlsx";
                            // Write the file to the disk
                            var fi = new FileInfo(filePath);
                            excelEngine.SaveAs(fi);

                            // Exporting .csv file when first cell is empty
                            File.WriteAllText(MapPath + "Uitzonderingen.csv", emptyFirstCellDocument);
                            var fileUitzondering = new FileInfo(MapPath + "Uitzonderingen.csv");
                            // Load the empty.csv to the worksheet
                            wbUitzondering.Cells["A1"].LoadFromText(fileUitzondering, format);
                            var stringPath = MapPath + fileName + "Overzicht.xlsx";

                            // When first cell is not blank, write to filled.csv
                            File.WriteAllText(MapPath + "TotaalIngevuld.csv", TotalDocument);
                            var fileTotaalIngevuld = new FileInfo(MapPath + "TotaalIngevuld.csv");
                            wbAllesIngevuld.Cells["A1"].LoadFromText(fileTotaalIngevuld, format);
                            wbAllesIngevuld.Cells["A:I"].Sort(0, false);


                            // Still testing. q
                            for (var rowNum = 1; rowNum <= wbAllesIngevuld.Dimension.End.Row; rowNum++)
                            {
                                var row = wbAllesIngevuld.Cells[string.Format("{0}:{0}", rowNum)];
                               // Debug.WriteLine(row.Value.ToString());
                            }


                            //for (int currentRow = wbAllesIngevuld.Dimension.Start.Row; currentRow <= wbAllesIngevuld.Dimension.End.Row; currentRow++)
                            //{
                            //    string firstCellValue = wbAllesIngevuld.Cells[currentRow, 1].Value == null || wbAllesIngevuld.Cells[currentRow, 1].Value.ToString() == "" ? null : wbAllesIngevuld.Cells[currentRow, 1].Value.ToString();
                            //    string secondCellValue = wbAllesIngevuld.Cells[currentRow, 2].Value == null || wbAllesIngevuld.Cells[currentRow, 2].Value.ToString() == "" ? null : wbAllesIngevuld.Cells[currentRow, 2].Value.ToString();

                            //    if (firstCellValue == secondCellValue)
                            //    {
                            //    wbAllesIngevuld.Cells[currentRow, 1, currentRow, 2].Merge = true;
                            //    }


                            // Write the file to the disk
                            var fileInfoUitzondering = new FileInfo(stringPath);
                            xlPackage.SaveAs(fileInfoUitzondering);

                            // For each loop this deletes the .csv files to keep the folder clean
                            File.Delete(MapPath + filename + ".csv");
                            File.Delete(MapPath + "Uitzonderingen.csv");
                            File.Delete(MapPath + "TotaalIngevuld.csv");




                        }
                        // Closes both excel packages 
                        xlPackage.Dispose();
                }

                excelEngine.Dispose();
            }

            return r;
        }

        // Checks values with have been added through the buttons
        private bool CheckValues(string name, ArrayList toCheck)
        {
            var current = false;
            foreach (string item in toCheck)
                if (name.Contains(item))
                    current = true;
            return current;
        }
    }
}