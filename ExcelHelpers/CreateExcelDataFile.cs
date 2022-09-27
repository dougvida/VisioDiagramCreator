using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using VisioDiagramCreator.Visio;
using Excel = Microsoft.Office.Interop.Excel;

///
/// helper URL http://csharp.net-informations.com/excel/csharp-format-excel.htm
/// 

namespace VisioDiagramCreator.ExcelHelpers
{
	public class CreateExcelDataFile
	{
		private Excel.Application _xlApp = null;
		private Excel.Workbook _xlWorkbook = null;
		private Excel.Worksheet _xlWorksheet = null;
		//Excel.Sheets _worksheets = null;		// _xlWorkbook.Worksheets; 
		
		object misValue = System.Reflection.Missing.Value;

		Dictionary<int, string> excelHeaderNames = new Dictionary<int, string>{

			// Excel data file header.  Must be in this sequence
			{ 0, "Visio Page"},		      // Page indicator to place this shape
			{ 1, "Shape Type"},           // key
			{ 2, "Stencil Key"},          // device unique Key used for connecting visio shapes
			{ 3, "Stencil Image"},        // device visio image name
			{ 4, "Stencil Label"},        // device label
			{ 5, "Stencil Font Size"},		// default is what the stencil font size is   (use 12:B for 12 pt. Bold or 12 for 12 pt)
			{ 6, "Mach_name"},				// device machine Name
			{ 7, "Mach_id"},					// device machine Id
			{ 8, "Site_id"},					// device site Id
			{ 9, "Site_name"},				// deivce site name
			{ 10, "Site_address"},			// device site address
			{ 11, "Omnis_name"},				// device name
			{ 12, "Omnis_id"},				// device Id
			{ 13, "SiteIdOmniId"},			// site_id+omni_id
			{ 14, "IP"},						// device IP address
			{ 15, "Ports"},					// device Ports
			{ 16, "DevicesCount"},			// number of Devices for this type (part of a group)

			{ 17, "PosX"},						// Shape position X
			{ 18, "PosY"},						// shape position Y
			{ 19, "Width"},					// shape width
			{ 20, "Height"},					// shape height
			{ 21, "Fill Color"},          // color to fill stincel

			{ 22, "Connect From"},        // used to link this visio shape to another visio shape
			{ 23, "From LineLabel"},      // Arrow Text
			{ 24, "From LinePattern"},    // Line pattern solid = 1
			{ 25, "From ArrowType"},      // Can contain one of these [None, Start, End, Both]
			{ 26, "From LineColor"},      // Arrow Color

			{ 27, "Connect To"},          // used to link this visio shape to another visio shape
			{ 28, "To LineLabel"},        // Arrow Text
			{ 29, "To LinePattern"},      // Line pattern solid = 1
			{ 30, "To ArrowType"},        // Can contain one of these [None, Start, End, Both]
			{ 31, "To LineColor"}			// Arrow Color
		};

		public CreateExcelDataFile()
		{
		}

		public bool PopulateExcelDataFile(Dictionary<int, ShapeInformation>shapesMap, string namePath )
		{
			// if file already exists display a message box asking the user
			// if the file can be overwritten or needs to be saved off
			// or just backup the file and move on
			if (openFile( namePath ))
			{
				// error
				return true;
			}

			if (_xlWorksheet != null)
			{
				// write data to the excel file


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
				return true;	// error
			}

			// open new excel file
			_xlWorkbook = _xlApp.Workbooks.Add(misValue);
			_xlWorksheet = (Excel.Worksheet)_xlWorkbook.Worksheets.get_Item(1);

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
				_xlWorkbook.SaveAs(fileNamePath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, 
															Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
			}
			closeExcel();

			return false;
		}

		private void closeExcel()
		{
			if (_xlWorkbook != null)
			{
				_xlWorkbook.Close(true, misValue, misValue);
			}
			if (_xlApp != null)
			{
				_xlApp.Quit();
			}

			Marshal.ReleaseComObject(_xlWorksheet);
			Marshal.ReleaseComObject(_xlWorkbook);
			Marshal.ReleaseComObject(_xlApp);
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
			int nIdx = 1;	// should be the first sheet

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
		private bool writeVisioDataSheet()
		{

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

		private bool writeHeader(Dictionary<int, string> headerNames)
		{
			// check to ensure file is open
			// if not open it.
			// start at the first row and column
			// write the header using the excelHeaderNames enum

			return false;
		}
	}
}
