using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace VisioDiagramCreator.ExcelHelpers
{
	public class CreateExcelDataFile
	{
		private Excel.Application xlApp = null;
		private Excel.Workbook xlWorkbook = null;

		Dictionary<int, string> excelHeaderNames = new Dictionary<int, string>{

			// Excel data file header.  Must be in this sequence
			{ 0, "Visio Page"},		      // Page indicator to place this shape
			{ 1, "Shape Type"},           // key
			{ 2, "Stencil Key"},          // device unique Key used for connecting visio shapes
			{ 3, "Stencil Image"},        // device visio image name
			{ 4, "Stencil Label"},        // device label
			{ 5, "Stencil Label Font Size"},// default is what the stencil font size is   (use 12:B for 12 pt. Bold or 12 for 12 pt)
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

		private bool openFile(string fileNamePath)
		{
			// check if existing file name exists
			// if so lets overwright it (give warning)
			// open the file for wright
			// declare the application object
			xlApp = new Excel.Application();

			// open a file
			xlWorkbook = xlApp.Workbooks.Open(fileNamePath);

			return false;
		}

		private void CloseExcelFile()
		{
			Marshal.ReleaseComObject(this.xlWorkbook);
			Marshal.ReleaseComObject(this.xlApp);
		}

		private bool writeHeader()
		{
			// check to ensure file is open
			// if not open it.
			// start at the first row and column
			// write the header using the excelHeaderNames enum

			return false;
		}
	}
}
