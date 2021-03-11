using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using OfficeOpenXml;
using System.IO;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Collections.Generic;

namespace E_60Toevoegen
{
	[TransactionAttribute(TransactionMode.ReadOnly)]
	public class CSV_E60Toevoegen : IExternalCommand
	{
		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			System.IO.Directory.CreateDirectory(@"c:\\temp\\E_60");
			Result r = Result.Succeeded;
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			// MisValue for excel files, unknown values
			object misValue = System.Reflection.Missing.Value;

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
				TextQualifier = ExportTextQualifier.DoubleQuote,
				FieldDelimiter = ","
			};

			// Creates new excelpackage this 
			using (ExcelPackage excelEngine = new ExcelPackage())
			{
				foreach (ViewSchedule vs in col)
				{
					// Searches for schedules containing AE E60 M52 en M57 ventilatierooster
					// dit zijn de schedules waarbij het met aantallen is.
					if (vs.Name.Contains("AE_E60")
						|| vs.Name.Contains("AE_M52")
						|| vs.Name.Contains("AE_M57_ Ventilatieroosters")
						|| vs.Name.Contains("AE_M57_Toestellen VENT")
						|| vs.Name.Contains("AE_M50_Toestellen HVAC coll"))
					{
						//create a WorkSheet
						ExcelWorksheet ws1 = excelEngine.Workbook.Worksheets.Add(vs.Name);
						// Export c:\\temp --> Will be save as
						vs.Export(@"c:\\temp\\E_60\", Environment.UserName + vs.Name + ".csv", opt);
						FileInfo file = new FileInfo(@"c:\\temp\\E_60\" + Environment.UserName + vs.Name + ".csv");

						// Adds Worksheet as first in the row 
						ws1.Workbook.Worksheets.MoveToStart(vs.Name);

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
						
						ws1.Cells["A1"].LoadFromText(file, format);

						// excelEngine.Workbook.Worksheets.MoveBefore(i);
						// the path of the file
						string filePath = "C:\\temp\\E_60\\Excel_E_60.xlsx";

						// Write the file to the disk
						FileInfo fi = new FileInfo(filePath);
						excelEngine.SaveAs(fi);

						// Excel terug uitlezen en kijken naar welke cell er leeg in de eerste kolom
						// om vervolgens de rij te kopieren. 
						


					}
				}
			}
			// result succeeded 
			return r;
		}
			// Test voor het zoeken van lege eerste cel in een rij. Deze dan vervolgens te verwijderen
			// Beter is als deze verplaats worden naar een ander workbook
    }

}



