using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OmnicellBlueprintingTool.Visio;
using Excel = Microsoft.Office.Interop.Excel;


///
/// helper URL http://csharp.net-informations.com/excel/csharp-format-excel.htm
/// 

namespace OmnicellBlueprintingTool.ExcelHelpers
{
	public class CreateExcelDataFile
	{
		private Excel.Application _xlApp = null;
		private Excel.Workbook _xlWorkbook = null;
		private Excel.Worksheet _xlWorksheet = null;
		//Excel.Sheets _worksheets = null;		// _xlWorkbook.Worksheets; 

		//object misValue = System.Reflection.Missing.Value;


		public CreateExcelDataFile()
		{
		}

		public bool PopulateExcelDataFile(Dictionary<int, ShapeInformation> shapesMap, string namePath)
		{
			// if file already exists display a message box asking the user
			// if the file can be overwritten or needs to be saved off
			// or just backup the file and move on
			if (openFile(namePath))
			{
				// error
				return true;
			}

			if (_xlWorksheet != null)
			{
				if (writeExcelDataSheet(_xlWorksheet, shapesMap))
				{
					closeExcel();
					return true;
				}

				// save and close the excel file
				saveFile(namePath);
			}
			return false;
		}


		private bool openFile(string fileNamePath)
		{
			// check if existing file name exists
			// if so lets overwright it (give warning)
			// open the file for wright
			// declare the application object
			_xlApp = new Excel.Application();
			if (_xlApp == null)
			{
				MessageBox.Show("Excel is not properly installed!!");
				return true;   // error
			}

			// open new excel file
			_xlWorkbook = _xlApp.Workbooks.Add(Type.Missing);
			// _xlWorksheet = (Excel.Worksheet)_xlWorkbook.Worksheets.get_Item(1);
			_xlWorksheet = (Excel.Worksheet)_xlWorkbook.ActiveSheet; 

			// open existing excel file
			//_xlWorkbook = _xlApp.Workbooks.Open(fileNamePath);

			return false;
		}

		private Excel.Workbook createNewWorkbook(string sWorkbookName)
		{
			return _xlApp.Workbooks.Add(sWorkbookName);
		}

		private bool saveFile(string fileNamePath)
		{
			if (_xlWorkbook != null)
			{
				//_xlWorksheet.SaveAs("your-file-name.xls");
				_xlWorkbook.SaveAs(fileNamePath, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
										Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
			}
			closeExcel();

			return false;
		}

		private void closeExcel()
		{
			if (_xlWorkbook != null)
			{
				_xlWorkbook.Close(true, Type.Missing, Type.Missing);
			}
			if (_xlApp != null)
			{
				_xlApp.Quit();
			}

			Marshal.ReleaseComObject(_xlWorksheet);
			Marshal.ReleaseComObject(_xlWorkbook);
			Marshal.ReleaseComObject(_xlApp);

			_xlWorksheet = null;
			_xlWorkbook = null;
			_xlApp = null;
		}

		private void addNexWorksheet(string sheetName)
		{
			_xlApp.DisplayAlerts = false;

			//var xlNewSheet = (Excel.Worksheet)Worksheets.Add(Worksheets[1], Type.Missing, Type.Missing, Type.Missing);
			//xlNewSheet.Name = sheetName;
			//xlNewSheet.Cells[1, 1] = "New sheet content";
		}

		private void selectWorkSheet(string sheetName)
		{
			int nIdx = 1;  // should be the first sheet

			// get the sheet index for the given name to make this correct

			selectworkSheet(nIdx);
		}

		private void selectworkSheet(int nIdx)
		{
			// check to ensure the nIdx value is withing range

			_xlWorksheet = (Excel.Worksheet)_xlWorkbook.Worksheets.get_Item(nIdx);
			_xlWorksheet.Select();

		}

		/// <summary>
		/// writeVisioDataSheet
		/// Write the data to the Excel file
		/// Write to the VisioData tab (index = 1)
		/// </summary>
		/// <returns></returns>
		private bool writeExcelDataSheet(Excel.Worksheet workSheet, Dictionary<int, ShapeInformation> shapesMap)
		{
			//Create COM Objects. Create a COM object for everything that is referenced
			//Excel.Application xlApp = new Excel.Application();
			//Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
			//Excel._Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

			//Excel.Range xlRange = xlWorksheet.UsedRange;

			int rowCount = 2; // skip the header
			//int colCount = Enum.GetNames(typeof(ExcelVariables.CellIndex)).Length;

			//System.Array myArray = (System.Array)xlRange.Cells.Value2;
			try
			{
				// write the header
				// write data to the excel file
				if (writeHeader(_xlWorksheet, ExcelVariables.GetExcelHeaderNames()))
				{
					return true;	// error
				}

				foreach (var shape in shapesMap)
				{
					// break apart the object and update the excel row based on the column value from the shapesMap
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.VisioPage]).Value = shape.Value.VisioPage;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ShapeType]).Value = "Shape";
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.UniqueKey]).Value = shape.Value.UniqueKey;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.StencilImage]).Value = shape.Value.StencilImage;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.StencilLabel]).Value = shape.Value.StencilLabel;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.StencilLabelFontSize]).Value = shape.Value.StencilLabelFontSize;

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Mach_name]).Value = string.Empty;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Mach_id]).Value = string.Empty;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Site_id]).Value = string.Empty;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Site_name]).Value = string.Empty;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Site_address]).Value = string.Empty;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Omnis_name]).Value = string.Empty;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Omnis_id]).Value = string.Empty;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.SiteIdOmniId]).Value = string.Empty;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.IP]).Value = string.Empty;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Ports]).Value = string.Empty;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.DevicesCount]).Value = string.Empty;
					
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.PosX]).Value = shape.Value.Pos_x;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.PosY]).Value = shape.Value.Pos_y;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Width]).Value = shape.Value.Width;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Height]).Value = shape.Value.Height;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FillColor]).Value = shape.Value.FillColor;
					
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ConnectFrom]).Value = shape.Value.ConnectFrom;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromLineLabel]).Value = shape.Value.FromLineLabel;

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromLinePattern]).Value = shape.Value.FromLinePattern;
					if (shape.Value.FromLinePattern <= 1)
					{
						((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromLinePattern]).Value = "";
					}
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromArrowType]).Value = shape.Value.FromArrowType;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromLineColor]).Value = shape.Value.FromLineColor;

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ConnectTo]).Value = shape.Value.ConnectTo;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToLineLabel]).Value = shape.Value.ToLineLabel;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToLinePattern]).Value = shape.Value.ToLinePattern;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToArrowType]).Value = shape.Value.ToArrowType;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToLineColor]).Value = shape.Value.ToLineColor;
					rowCount++;
				}
			}
			catch(Exception ex)
			{
				MessageBox.Show(string.Format("Exception::writeExcelDataSheet - {0}"), ex.Message);
			}
			return false;
		}

		/// <summary>
		/// wruteSystemInfoSheet
		/// Write to the SystemInfo sheet
		/// </summary>
		/// <returns></returns>
		private bool writeSystemInfoSheet()
		{
			return false;
		}


		private bool writeInterfacesSheet()
		{
			return false;
		}

		private bool writeTableSheet()
		{
			return false;
		}

		private bool writeHeader(Excel.Worksheet workSheet, Dictionary<int, string> headerNames)
		{
			// check to ensure file is open
			// if not open it.
			// start at the first row and column
			// write the header using the ExcelHeaderNames enum
			string headerName = string.Empty;
			int row = 1;
			// header names map starts with 0 index
			for (int col = 0; col < headerNames.Count; col++)
			{
				// we only need to get the first few columns to determine what to do
				if (!headerNames.TryGetValue(col, out headerName))
				{
					MessageBox.Show(string.Format("writeHeader::Error writing header.  column:{0}-Name:{1}", col, headerName));
					return true;
				}
				((Excel.Range)workSheet.Cells[row, col+1]).Value = headerName;
			}
			return false;
		}
	}
}
