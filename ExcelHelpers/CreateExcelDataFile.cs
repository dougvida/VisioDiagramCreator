using System;
using System.Drawing;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Visio;
using OmnicellBlueprintingTool.Models;
using OmnicellBlueprintingTool.Visio;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;




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

		public CreateExcelDataFile()
		{
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

		private int writeHeader(Excel.Worksheet workSheet, Dictionary<int, string> headerNames, int row = 1)
		{
			string headerName = string.Empty;

			// header names map starts with 0 index
			for (int col = 0; col < headerNames.Count; col++)
			{
				// we only need to get the first few columns to determine what to do
				if (!headerNames.TryGetValue(col, out headerName))
				{
					MessageBox.Show(string.Format("writeHeader::Error writing header.  column:{0}-Name:{1}", col, headerName));
					return -1;
				}
				((Excel.Range)workSheet.Cells[row, col + 1]).Value = headerName;
				switch(col+1)
				{
					case 1:	// Visio Page
					case 2:	// Shape Type
					case 3:	// Unique Key
					case 4:	// Stencil Image
					case 19:	// PosX
					case 20:	// PosY
						workSheet.Cells[row, col + 1].Interior.Color = Excel.XlRgbColor.rgbRed;
						break;

					default:
						workSheet.Cells[row, col + 1].Interior.Color = Excel.XlRgbColor.rgbBeige;
						break;
				}
				((Excel.Range)workSheet.Cells[row, col + 1]).Borders.LineStyle = XlLineStyle.xlContinuous;
			}
			return row;
		}

		private int writeConfiguration(Excel.Worksheet workSheet, DiagramData diagramData, int cellIndex, int nRow)
		{
			ShapeInformation shpObj = null;

			// Write comment section named "Configuration"
			shpObj = new ShapeInformation();
			shpObj.VisioPage = 0;
			shpObj.ShapeType = "Configuration";
			shpObj.UniqueKey = String.Empty;
			shpObj.StencilLabel = String.Empty;
			if (writeData(workSheet, shpObj, nRow, true))
			{
				MessageBox.Show("CreateExeclDataFile:: writeConfiguration Error.  Failed to write Comment data");
				return -1;
			}

			// Write Template section
			shpObj = new ShapeInformation();
			nRow++;
			shpObj.VisioPage = 0;
			shpObj.ShapeType = "Template";
			shpObj.UniqueKey = string.Format(@"{0}",diagramData.VisioTemplateFilePath + VisioVariables.DefaultBlueprintingTemplateFile);
			shpObj.StencilLabel = string.Format("Use the Blueprinting Visio Template.  Already contains the {0}",VisioVariables.DefaultBlueprintingTemplateFile);
			if (writeData(workSheet, shpObj, nRow, true))
			{
				MessageBox.Show("CreateExeclDataFile:: writeConfiguration Error.  Failed to write Template data");
				return -2;
			}

			// Write the Stencil data
			shpObj = new ShapeInformation();
			nRow++;
			shpObj.VisioPage = 0;
			shpObj.ShapeType = "Stencil";
			shpObj.UniqueKey = string.Format(@"{0}", diagramData.VisioStencilFilePaths[0] +VisioVariables.DefaultBlueprintingStencilFile);
			shpObj.StencilLabel = string.Format("Use the Blueprinting Visio Stencil.  Already contains the Stencil file:{0}",VisioVariables.DefaultBlueprintingStencilFile);
			if (writeData(workSheet, shpObj, nRow, false))
			{
				MessageBox.Show("CreateExeclDataFile:: writeConfiguration Error.  Failed to write Stincel data");
				return -3;
			}

			// Write Page setup Section
			shpObj = new ShapeInformation();
			nRow++;
			shpObj.VisioPage = 0;
			shpObj.ShapeType = "Page Setup";
			shpObj.UniqueKey = VisioVariables.VisioPageOrientation.Portrait + ":" + VisioVariables.VisioPageSize.Legal;
			shpObj.StencilLabel = "• Orientation: Landscape or Portrait (default)\r\n• Size: Letter (default), Tabloid, Ledger, Legal, A3, A4";
			if (writeData(workSheet, shpObj, nRow, true))
			{
				MessageBox.Show("CreateExeclDataFile:: writeConfiguration Error.  Failed to Setup Page data");
				return -4;
			}

			shpObj = new ShapeInformation();
			nRow++;
			shpObj.VisioPage = 0;
			shpObj.ShapeType = "Page Setup";
			shpObj.UniqueKey = "Autosize:true";
			shpObj.StencilLabel = "• true - Autosize all pages\r\n• false - (default) don't Autosize the pages";
			if (writeData(workSheet, shpObj, nRow, true))
			{
				MessageBox.Show("CreateExeclDataFile:: writeConfiguration Error.  Failed to Setup Page data");
				return -5;
			}

			// Write comment section named "Visio Shapes"
			shpObj = new ShapeInformation();
			nRow++;
			shpObj.VisioPage = 0;
			shpObj.ShapeType = "Visio Section";
			shpObj.UniqueKey = string.Empty;
			shpObj.StencilLabel = String.Empty;
			if (writeData(workSheet, shpObj, nRow, true))
			{
				MessageBox.Show("CreateExeclDataFile:: writeConfiguration Error.  Failed to write Comment2 data");
				return -6;
			}
			return nRow;
		}


		public bool PopulateExcelDataFile(DiagramData diagramData, Dictionary<int, ShapeInformation> shapesMap, string namePath)
		{
			int nRow = 1;

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
				try
				{
					// write the header
					// write data to the excel file
					nRow = writeHeader(_xlWorksheet, ExcelVariables.GetExcelHeaderNames(), nRow);
					if (nRow < 0)
					{
						MessageBox.Show("Error::PopulateExcelDataFile\n Writing the header");
						closeExcel();
						return true;
					}

					nRow = writeConfiguration(_xlWorksheet, diagramData, ExcelVariables.GetHeaderCount(), ++nRow);
					if (nRow < 0)
					{
						MessageBox.Show("Error::PopulateExcelDataFile\n Writing the configuration section");
						closeExcel();
						return true;
					}

					// write the Stencil data
					nRow = writeAllData(_xlWorksheet, shapesMap, ++nRow);
					if (nRow < 0)
					{
						closeExcel();
						return true;
					}

					FormatWorkSheet(_xlWorksheet);

					// this should stop the check Compatibility diaglog from poping up
					_xlWorkbook.DoNotPromptForConvert = true;             
					
						// save and close the excel file
					saveFile(namePath);
				}
				catch(Exception ex)
				{
					MessageBox.Show(string.Format("Exception::PopulateExcelDataFile\n{0}", ex.Message));
					return true;
				}
			}
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
				//Here saving the file in xlsx
				_xlWorkbook.SaveAs(fileNamePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing,
				Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

				//_xlWorkbook.SaveAs(fileNamePath, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
				//						Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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
		/// <param name="workSheet"></param>
		/// <param name="shapesMap"></param>
		/// <param name="row"></param>
		/// <returns>int<text>
		/// <text>if > 0 total number of rows in the excel data file</text>
		/// <text>if <= 0 an error has occured</text>
		/// </returns>
		private int writeAllData(Excel.Worksheet workSheet, Dictionary<int, ShapeInformation> shapesMap, int rowCount)
		{
			try
			{
				foreach (KeyValuePair<int, ShapeInformation> keyValue in shapesMap)
				{	
					if (string.IsNullOrEmpty(keyValue.Value.ShapeType))
					{
						keyValue.Value.ShapeType = "Shape";
					}
					writeData(workSheet, keyValue.Value, rowCount++, false);
					//rowCount++;
				}
			}
			catch(Exception ex)
			{
				MessageBox.Show(string.Format("Exception::writeExcelDataSheet - {0}"), ex.Message);
				return -1;	// error
			}
			return rowCount;
		}

		private bool writeData(Excel.Worksheet workSheet, ShapeInformation shape, int rowCount, bool IsComment)
		{
			try
			{
				string sTmp = string.Empty;

				// break apart the object and update the excel row based on the column value from the shapesMap
				sTmp = shape.VisioPage.ToString();
				if (IsComment)
				{
					sTmp = ";";
				}
				((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.VisioPage]).Value = sTmp;
				((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ShapeType]).Value = shape.ShapeType;
				((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.UniqueKey]).Value = shape.UniqueKey;
				((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.StencilImage]).Value = shape.StencilImage;
				((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.StencilLabel]).Value = shape.StencilLabel;

				// if not a comment and above the header and configurations rows fill these cells
				if (!IsComment && rowCount > 7)
				{
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.StencilLabelPosition]).Value = shape.StencilLabelPosition;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.StencilLabelFontSize]).Value = shape.StencilLabelFontSize;

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

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.PosX]).Value = shape.Pos_x;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.PosX]).NumberFormat = "#0.000";

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.PosY]).Value = shape.Pos_y;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.PosY]).NumberFormat = "#0.000";

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Width]).Value = shape.Width;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Width]).NumberFormat = "#0.000";

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Height]).Value = shape.Height;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Height]).NumberFormat = "#0.000";

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FillColor]).Value = shape.FillColor;

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ConnectFrom]).Value = shape.ConnectFrom;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromLineLabel]).Value = shape.FromLineLabel;
					
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromLinePattern]).Value = shape.FromLinePattern;
					if (shape.FromLinePattern <= 1)
					{
						((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromLinePattern]).Value = "";
					}
					
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromArrowType]).Value = shape.FromArrowType;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromLineColor]).Value = shape.FromLineColor;

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ConnectTo]).Value = shape.ConnectTo;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToLineLabel]).Value = shape.ToLineLabel;
					
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToLinePattern]).Value = shape.ToLinePattern;
					if (shape.ToLinePattern <= 1)
					{
						((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToLinePattern]).Value = "";
					}
					
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToArrowType]).Value = shape.ToArrowType;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToLineColor]).Value = shape.ToLineColor;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Exception::writeExcelDataSheet - {0}"), ex.Message);
				return true;
			}
			if (IsComment)
			{
				Excel.Range range = workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, ExcelVariables.GetHeaderCount()]];
				range.Interior.Color = Excel.XlRgbColor.rgbYellow;
			}
			return false;
		}

		public void FormatWorkSheet(Excel.Worksheet workSheet)
		{
			Excel.Range xlRange = workSheet.UsedRange;
			int rowCount = xlRange.Rows.Count;
			int colCount = xlRange.Columns.Count;

			// format each cell to be center justified and Left aligned in the row
			workSheet.Cells.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
			workSheet.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

			//Set Text-Wrap for all rows true//
			workSheet.Rows.WrapText = true;

			for (int nCol = 1; nCol < colCount; nCol++)
			{
				var data = (Excel.Range)workSheet.Cells[1, nCol];
				if (data != null)
				{
					string sTmp = data.Text.ToString().ToUpper();
					switch (sTmp)
					{
						case "UNIQUE KEY":
						case "STENCIL LABEL":
							((Excel.Range)workSheet.Cells[1, nCol]).ColumnWidth = 37.00;
							break;

						case "VISIO PAGE":
						case "SHAPE TYPE":
							((Excel.Range)workSheet.Cells[1, nCol]).ColumnWidth = 16.00;
							break;

						case "POSX":
						case "POSY":
						case "WIDTH":
						case "HEIGHT":
							((Excel.Range)workSheet.Cells[1, nCol]).ColumnWidth = 8.00;
							((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.PosX]).NumberFormat = "#0.000";
							break;
						case "MACH_NAME":
						case "MACH_ID":
						case "SITE_ID":
						case "SITE_NAME":
						case "SITE_ADDRESS":
						case "OMNIS_NAME":
						case "OMNIS_ID":
						case "SITEID_OMNIID":
						case "FILL COLOR":
							((Excel.Range)workSheet.Cells[1, nCol]).ColumnWidth = 8.00;
							break;

						default:
							((Excel.Range)workSheet.Cells[1, nCol]).ColumnWidth = 12.00;
							break;
					}
				}
			}
			// lets set borders around each cell
			Excel.Range range = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[rowCount, ExcelVariables.GetHeaderCount()]];
			range.Borders.LineStyle = XlLineStyle.xlContinuous;
			range.Rows.AutoFit();		// auto aize the rows
			//range.Columns.AutoFit();
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

	}
}
