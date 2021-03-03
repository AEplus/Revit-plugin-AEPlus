//using Autodesk.Revit.Attributes;
//using Autodesk.Revit.DB;
//using Autodesk.Revit.UI;
//using Microsoft.VisualBasic.FileIO;
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;

//namespace MyRevitCommands
//{
//    [TransactionAttribute(TransactionMode.ReadOnly)]
//    public class ExportSchedule : IExternalCommand
//    {
//        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
//        {
//            // Get UIDocument
//            UIDocument uidoc = commandData.Application.ActiveUIDocument;
//            // Get Document
//            Document doc = uidoc.Document;
//            FilteredElementCollector col = new FilteredElementCollector(doc)
//                                                .OfClass(typeof(ViewSchedule));
//            /*
//            Export Options on how the .txt file will look. Text "" is gone
//            FieldDelimiter is TAB replaced with ,  
//             */
//            ViewScheduleExportOptions opt = new ViewScheduleExportOptions
//            {
//                TextQualifier = ExportTextQualifier.None,
//                FieldDelimiter = ","
//            };

//            // MisValue for excel files, unknown values
//            object misValue = System.Reflection.Missing.Value;

//            // Launch Excel application and check if the application is available 
//            Excel.Application xlApp = new
//            Microsoft.Office.Interop.Excel.Application();

//            if (xlApp == null)
//            {
//                MessageBox.Show("Excel is not properly installed!");
//            }

//            foreach (ViewSchedule vs in col)
//            {
//                Directory.CreateDirectory(@"c:\\temp");
//                vs.Export(@"c:\\temp", Environment.UserName + DateTime.Today + vs.Name + ".csv", opt);

//                // search temp folder
//                var ListOfcsvFilenames = Directory.GetFiles("c:\\temp", Environment.UserName + "*.csv");

//                // Copy files and move to temp folder 

//                // read newest csv files to array
//                List<CustomSchedule> customSchedules = new List<CustomSchedule>();
//                for (csvFileNames in ListOfcsvFilenames)
//                {
//                    customSchedules.Add(new CustomSchedule(File.ReadAllLines(csvFileNames.ToString())));
//                }

//                // making excel


//                // create excel file
//                Excel.Application oXL = new Excel.Application();
//                oXL.Visible = false;
//                Excel.Workbook oWB = oXL.Workbooks.Add(misValue);

//                // Excel inladen
//                foreach (CustomSchedule cs in customSchedules)
//                {
//                    // create excel sheet
//                    Excel.Worksheet oSheet = oWB.Sheets.Add(misValue,
//                                                            misValue,
//                                                            1,
//                                                            misValue)
//                                                            as Excel.Worksheet;

//                    // add excel sheet name
//                    oSheet.Name = cs.getTitle();
//                    oSheet.Cells[1, 1] = cs.getTitle();

//                    // add excel sheet headers
//                    int i = 1;
//                    foreach (string header in cs.getHeaders())
//                    {
//                        oSheet.Cells[2, i] = header;
//                        i++;
//                    }

//                    // looping over an excel row that will be filled in with data from 
//                    // CustomSchedule object 
//                    int y = 3;
//                    foreach (List<string> row in cs.getData())
//                    {
//                        // this is an excel row
//                        int x = 1;
//                        foreach (string item in row)
//                        {
//                            // this is an excel row column
//                            oSheet.Cells[x, y] = item;
//                            x++;
//                        }
//                        y++;
//                    }
//                }

//                // Save project to path, close and save changes, quit excel 
//                oWB.SaveAs("C:\\test1\test1.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
//                oWB.Close(true, misValue, misValue);
//                oXL.UserControl = true;
//                oXL.Quit();
//            }
//            return Result.Succeeded;
//        }
//    }
//    public class CustomSchedule
//    {
//        private string title;
//        private List<string> headers;
//        private List<List<string>> data;
//        public CustomSchedule(string[] data)
//        {
//            this.data = new List<List<string>>();
//            for (int i = 0; i < data.Length; i++)
//            {
//                Console.Out.WriteLine(data[i]);
//                if (i == 0)
//                {
//                    // read title from file
//                    this.title = data[i];
//                }
//                else if (i == 1)
//                {
//                    // read headers
//                    this.headers = new List<string>(data[i].Split(','));
//                }
//                else
//                {
//                    // read rows 
//                    this.data.Add(
//                            new List<string>(data[i].Split(',')));
//                }
//            }
//        }
//        // Excel title, headers and the rest of the data 
//        public CustomSchedule(string title, List<string> headers, List<List<string>> data)
//        {
//            this.title = title;
//            this.headers = headers;
//            this.data = data;
//        }
//        public string getTitle()
//        {
//            return this.title;
//        }
//        public List<string> getHeaders()
//        {
//            return this.headers;
//        }
//        public List<List<string>> getData()
//        {
//            return this.data;
//        }
//    }
//}
