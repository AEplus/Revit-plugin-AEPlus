using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyRevitCommands
{
    class GenericToevoegen
    {
        /*
            ...
            MapPath example value:  @"c:\\temp\\E_60\" 
         */
        public Result GenericExecute(ExternalCommandData commandData, ref string message, ElementSet elements, string MapPath, ArrayList revitSchedules)
        {
            string xlSheetName;
            System.IO.Directory.CreateDirectory(MapPath);
            Result r = Result.Succeeded;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string emptyFirstCellDocument = "";
            string TotalDocument = "";

            // Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            // Get Document
            Document doc = uidoc.Document;

            FilteredElementCollector col = new FilteredElementCollector(doc)
            .OfClass(typeof(ViewSchedule));

            // Export Options on how the .txt file will look. Text "" is gone
            // FieldDelimiter is TAB replaced with ,
            ViewScheduleExportOptions opt = new ViewScheduleExportOptions()
            {
                TextQualifier = ExportTextQualifier.None,
                FieldDelimiter = ",",
            };

            // Formating for writing to xlsx
            var format = new ExcelTextFormat()
            {
                Culture = CultureInfo.InvariantCulture,
                // Escape character for values containing the Delimiter
                // ex: "A,Name",1 --> two cells, not three
                TextQualifier = '"'
                // Other properties
                // EOL, DataTypes, Encoding, SkipLinesBeginning/End
            };
            // Creates new excelpackage this
            using (ExcelPackage excelEngine = new ExcelPackage())
            {
                using (ExcelPackage xlPackage = new ExcelPackage())
                {
                    ExcelWorksheet wbUitzondering = xlPackage.Workbook.Worksheets.Add("Uitzondering");
                    ExcelWorksheet wbAllesIngevuld = xlPackage.Workbook.Worksheets.Add("AllesIngevuld");
                    foreach (ViewSchedule vs in col)
                    {
                        // Searches for schedules containing AE E60 M52 en M57 ventilatierooster
                        // dit zijn de schedules waarbij het met aantallen is.

                        if(checkValues(vs.Name, revitSchedules))
                        {
                            // Checks if vs.name is not longer than 30 characters because Excels crashes.
                            // Keep in mind for the vs.names that they have to be different under 30 characters
                            // Excel does not like the same names
                            if (vs.Name.Length > 30)
                            {
                                xlSheetName = vs.Name.Substring(0, 30);
                            }
                            else
                            {
                                xlSheetName = vs.Name;
                            }
                            //create a WorkSheet
                            ExcelWorksheet ws1 = excelEngine.Workbook.Worksheets.Add(xlSheetName);
                            // Export c:\\temp --> Will be save as
                            string filename = Environment.UserName + vs.Name;
                            vs.Export(MapPath, filename + ".csv", opt);

                            string normalDocument = "";
                            string StringPathFile = MapPath + filename + ".csv";
                            string[] lines = File.ReadAllLines(StringPathFile);
                            char[] delimitChars = { ',' };
                            int i = 1;

                            foreach (string line in lines)
                            {
                                // Gets first 2 row of each Schedule, name and properties.
                                // This is done for visibility and making the excel easier to read
                                if (i < 3)
                                {
                                    emptyFirstCellDocument += line + Environment.NewLine;
                                    i++;
                                }

                                // Looks for first value if this is null or blank ""
                                // Geo-IT number has to be first for this line
                                // Checks if it is blank or not
                                if (line.Split(delimitChars)[0] == ""
                                    || line.Split(delimitChars)[0] == null)
                                {
                                    // Empty first cell .csv line
                                    emptyFirstCellDocument += line + Environment.NewLine;
                                }
                                else
                                {
                                    // If the first cell is filled, continue on this .csv line
                                    // Only checks for blanks or not!
                                    normalDocument += line + Environment.NewLine;
                                }
                            }
                            // Gets spacing for each schedule.
                            // Increasing visibility for reading the excel.
                            emptyFirstCellDocument += Environment.NewLine + Environment.NewLine;
                            TotalDocument += normalDocument + Environment.NewLine;

                            File.WriteAllText(MapPath + filename + ".csv", normalDocument.ToString());
                            FileInfo file = new FileInfo(MapPath + filename + ".csv");
                            // Adds Worksheet as first in the row because excel keeps adding in the back
                            ws1.Workbook.Worksheets.MoveToStart(xlSheetName);
                            // EPPlus method for reading a .csv to excel in the format as written above here line +-40
                            ws1.Cells["A1"].LoadFromText(file, format);
                            // the path of the file
                            string filePath = MapPath + "Excel_E_60.xlsx";
                            // Write the file to the disk
                            FileInfo fi = new FileInfo(filePath);
                            excelEngine.SaveAs(fi);

                            // Exporting .csv file when first cell is empty
                            File.WriteAllText(MapPath + "Uitzonderingen.csv", emptyFirstCellDocument);
                            FileInfo fileUitzondering = new FileInfo(MapPath + "Uitzonderingen.csv");
                            // Load the empty.csv to the worksheet
                            wbUitzondering.Cells["A1"].LoadFromText(fileUitzondering, format);
                            string stringPath = MapPath + "Overzicht.xlsx";

                            // When first cell is not blank, write to filled.csv
                            File.WriteAllText(MapPath + "TotaalIngevuld.csv", TotalDocument);
                            FileInfo fileTotaalIngevuld = new FileInfo(MapPath + "TotaalIngevuld.csv");
                            wbAllesIngevuld.Cells["A1"].LoadFromText(fileTotaalIngevuld, format);

                            // Write the file to the disk
                            FileInfo fileInfoUitzondering = new FileInfo(stringPath);
                            xlPackage.SaveAs(fileInfoUitzondering);

                            // For each loop this deletes the .csv files to keep the folder clean
                            File.Delete(MapPath + filename + ".csv");
                            File.Delete(MapPath + "Uitzonderingen.csv");
                        }
                        //FileInfo fileIngevuld = new FileInfo(stringpath);
                        //xlPackage.SaveAs(fileIngevuld);
                    }
                    xlPackage.Dispose();
                }
                excelEngine.Dispose();
            }
            return r;
        }

        private Boolean checkValues(string name, ArrayList toCheck)
        {
            Boolean current = false;
            foreach (string item in toCheck)
            {
                if (name.Contains(item))
                {
                    current = true;
                }
            }
            return current;
        }

    }
}
