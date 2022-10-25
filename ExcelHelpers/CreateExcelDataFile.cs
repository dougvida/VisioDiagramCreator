﻿using System;
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
				MessageBox.Show("ERROR::CreateExcelDataFile\n\nExcel is not properly installed!!", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return true;   // error
			}

			// open new excel file
			_xlWorkbook = _xlApp.Workbooks.Add(Type.Missing);

			addNewWorksheet("VisioData");
			addNewWorksheet("SystemInfo");
			addNewWorksheet("Interfaces");
			addNewWorksheet("Tables");

			// select "VisioData" sheet
			_xlWorksheet = (Excel.Worksheet)_xlWorkbook.Worksheets.get_Item(1);
			
			// lets freeze the top row
			_xlWorksheet.Activate();
			_xlWorksheet.Application.ActiveWindow.SplitRow = 1; 
			_xlWorksheet.Application.ActiveWindow.FreezePanes = true;

			//_xlWorksheet = (Excel.Worksheet)_xlWorkbook.ActiveSheet;
			//_xlWorksheet.Name = "VisioData";

			deleteWorkSheet("Sheet1");	// this must be called after _xlWorksheet has been initialized

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
					string sTmp = string.Format("writeHeader::Error writing header.  column:{0}-Name:{1}", col, headerName);
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
			string sTmp = string.Empty;
			try
			{
				// Write comment section named "Configuration"
				shpObj = new ShapeInformation();
				shpObj.VisioPage = 0;
				shpObj.ShapeType = "Configuration";
				shpObj.UniqueKey = String.Empty;
				shpObj.StencilLabel = String.Empty;
				if (writeData(workSheet, shpObj, nRow, true))
				{
					 sTmp = "CreateExeclDataFile::writeConfiguration \n\nFailed to write Comment data";
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);

					return -1;
				}

				// Write Template section
				shpObj = new ShapeInformation();
				nRow++;
				shpObj.VisioPage = 0;
				shpObj.ShapeType = "Template";
				shpObj.UniqueKey = string.Format(@"{0}", diagramData.VisioTemplateFilePath + VisioVariables.DefaultBlueprintingTemplateFile);
				shpObj.StencilLabel = string.Format("Use the Blueprinting Visio Template.  Already contains the {0}", VisioVariables.DefaultBlueprintingTemplateFile);
				if (writeData(workSheet, shpObj, nRow, true))
				{
					sTmp = "CreateExeclDataFile::writeConfiguration Error\n\nFailed to write Template data";
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return -2;
				}

				// Write the Stencil data
				shpObj = new ShapeInformation();
				nRow++;
				shpObj.VisioPage = 0;
				shpObj.ShapeType = "Stencil";
				shpObj.UniqueKey = string.Format(@"{0}", diagramData.VisioStencilFilePaths[0] + VisioVariables.DefaultBlueprintingStencilFile);
				shpObj.StencilLabel = string.Format("Use the Blueprinting Visio Stencil.  Already contains the Stencil file:{0}", VisioVariables.DefaultBlueprintingStencilFile);
				if (writeData(workSheet, shpObj, nRow, false))
				{
					sTmp = "CreateExeclDataFile::writeConfiguration Error\n\nFailed to write Stincel data";
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
					sTmp = "CreateExeclDataFile::writeConfiguration Error\n\nFailed to Setup Page data";
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
					sTmp = "CreateExeclDataFile::writeConfiguration Error\n\nFailed to Setup Page data";
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
					sTmp = "CreateExeclDataFile::writeConfiguration Error\n\nFailed to write Visio Section Comment data";
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return -6;
				}
			}
			catch(Exception ex)
			{
				sTmp = string.Format("CreateExeclDataFile::writeConfiguration Exception\n\n{0}-{1}", ex.Message, ex.StackTrace);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			return nRow;
		}


		public bool PopulateExcelDataFile(DiagramData diagramData, Dictionary<int, ShapeInformation> shapesMap, string namePath)
		{
			int nRow = 1;
			string sTmp = string.Empty;

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
						sTmp = "CreateExcelDataFile::PopulateExcelDataFile\n\nWriting the header";
						MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
						closeExcel();
						return true;
					}

					nRow = writeConfiguration(_xlWorksheet, diagramData, ExcelVariables.GetHeaderCount(), ++nRow);
					if (nRow < 0)
					{
						sTmp = "CreateExcelDataFile::PopulateExcelDataFile\n\nWriting the configuration section";
						MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
						closeExcel();
						return true;
					}

					// write the Stencil data
					nRow = writeAllData(_xlWorksheet, shapesMap, ++nRow);
					if (nRow < 0)
					{
						sTmp = string.Format("CreateExcelDataFile::PopulateExcelDataFile\n\nWriting All Data:{0}",nRow);
						MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
						closeExcel();
						return true;
					}

					// format the VisioData sheet
					formatVisioDataSheet(_xlWorksheet);

					// populate the Tables sheet
					writeTableSheet(diagramData);

					// some column use a dropdown list so we need to setup it up
					setColumnsDropdownList(diagramData);

					// this should stop the check Compatibility diaglog from poping up
					_xlWorkbook.DoNotPromptForConvert = true;             
					
					// save and close the excel file
					saveFile(namePath);
				}
				catch(Exception ex)
				{
					sTmp = string.Format("Exception::PopulateExcelDataFile\n{0}", ex.Message);
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return true;
				}
			}
			return false;
		}

		/// <summary>
		/// addNewWorksheet
		/// add a new worksheet to the workbook and give it a name
		/// the new sheet will be added after the last sheet
		/// </summary>
		/// <param name="sheetName"></param>
		private void addNewWorksheet(string sheetName)
		{
			Excel.Sheets xlSheets = _xlWorkbook.Worksheets;
			var xlNewSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
			xlNewSheet.Name = sheetName;

			int totalSheets = _xlApp.Application.ActiveWorkbook.Sheets.Count;
			((Excel.Worksheet)_xlApp.Application.ActiveSheet).Move(
				 _xlApp.Application.Worksheets[totalSheets]);
		}

		private Excel.Worksheet selectWorkSheet(string sheetName)
		{
			Excel.Worksheet workSheet = _xlWorkbook.Sheets[sheetName];
			workSheet.Activate();
			return workSheet;
		}

		private void selectWorkSheet(int nIdx)
		{
			// check to ensure the nIdx value is withing range

			_xlWorksheet = (Excel.Worksheet)_xlWorkbook.Worksheets.get_Item(nIdx);
			_xlWorksheet.Select();

		}

		/// <summary>
		/// deleteWorkSheet
		/// delete the worksheet by name from the workbook
		/// </summary>
		/// <param name="name"></param>
		private void deleteWorkSheet(string name)
		{
			for (int xx = 1; xx <= _xlWorkbook.Worksheets.Count; xx++)
			{
				Excel.Worksheet workSheet = (Excel.Worksheet)_xlWorkbook.Worksheets.get_Item(xx);
				if (workSheet.Name == name)
				{
					// delete;
					((Excel.Worksheet)_xlApp.Application.ActiveWorkbook.Sheets[xx]).Delete();
					return;
				}
			}
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
				string sTmp = string.Format("Exception::writeExcelDataSheet - Exception\n\n{0}\n{1}",ex.Message, ex.StackTrace);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
				string sTmp = string.Format("Exception::writeExcelDataSheet - Exception\n\n{0}\n{1}", ex.Message, ex.StackTrace);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return true;
			}
			if (IsComment)
			{
				Excel.Range range = workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, ExcelVariables.GetHeaderCount()]];
				range.Interior.Color = Excel.XlRgbColor.rgbYellow;
			}
			return false;
		}

		private void formatVisioDataSheet(Excel.Worksheet workSheet)
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
						case "STENCIL IMAGE":
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

						case "CONNECT FROM":
						case "CONNECT TO":
						case "FROM LINE LABEL":
						case "TO LINE LABEL":
							((Excel.Range)workSheet.Cells[1, nCol]).ColumnWidth = 20.00;
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
		/// writeSystemInfoSheet
		/// Not yet implemented
		/// </summary>
		/// <returns></returns>
		private bool writeSystemInfoSheet()
		{
			return false;
		}

		/// <summary>
		/// writeInterfacesSheet
		/// Not yet implemented
		/// </summary>
		/// <returns></returns>
		private bool writeInterfacesSheet()
		{
			return false;
		}

		/// <summary>
		/// writeTableSheet
		/// write the data to the table sheet
		/// the user can use this sheet as list to the excel file
		/// </summary>
		/// <returns></returns>
		private bool writeTableSheet(DiagramData diagramData)
		{
			Excel.Worksheet xlNewSheet = selectWorkSheet("Tables");

			// column A is Colors
			((Excel.Range)xlNewSheet.Cells[1, 1]).Value = "Color";
			((Excel.Range)xlNewSheet.Cells[1, 1]).ColumnWidth = 20.00;
			xlNewSheet.Cells[1, 1].Interior.Color = Excel.XlRgbColor.rgbDarkSeaGreen;

			((Excel.Range)xlNewSheet.Cells[2, 1]).Value = "";
			((Excel.Range)xlNewSheet.Cells[3, 1]).Value = "Black";
			((Excel.Range)xlNewSheet.Cells[4, 1]).Value = "Blue";
			((Excel.Range)xlNewSheet.Cells[5, 1]).Value = "Cyan";
			((Excel.Range)xlNewSheet.Cells[6, 1]).Value = "Gray";
			((Excel.Range)xlNewSheet.Cells[7, 1]).Value = "Green";
			((Excel.Range)xlNewSheet.Cells[8, 1]).Value = "Light Blue";
			((Excel.Range)xlNewSheet.Cells[9, 1]).Value = "Light Green";
			((Excel.Range)xlNewSheet.Cells[10, 1]).Value = "Orange";
			((Excel.Range)xlNewSheet.Cells[11, 1]).Value = "Red";
			((Excel.Range)xlNewSheet.Cells[12, 1]).Value = "Yellow";

			Excel.Range range = xlNewSheet.Range[xlNewSheet.Cells[1, 1], xlNewSheet.Cells[12, 1]];
			range.Borders.LineStyle = XlLineStyle.xlContinuous;
			range.Rows.AutoFit();      // auto aize the rows

			// Column C is Arrows
			((Excel.Range)xlNewSheet.Cells[1, 3]).Value = "Arrows";
			((Excel.Range)xlNewSheet.Cells[1, 3]).ColumnWidth = 20.00;
			xlNewSheet.Cells[1, 3].Interior.Color = Excel.XlRgbColor.rgbDarkSeaGreen;

			((Excel.Range)xlNewSheet.Cells[2, 3]).Value = "";
			((Excel.Range)xlNewSheet.Cells[3, 3]).Value = "None";
			((Excel.Range)xlNewSheet.Cells[4, 3]).Value = "Start";
			((Excel.Range)xlNewSheet.Cells[5, 3]).Value = "End";
			((Excel.Range)xlNewSheet.Cells[6, 3]).Value = "Both";

			range = xlNewSheet.Range[xlNewSheet.Cells[1, 3], xlNewSheet.Cells[6, 3]];
			range.Borders.LineStyle = XlLineStyle.xlContinuous;
			range.Rows.AutoFit();      // auto aize the rows

			// Column E is Stencil Label Font size
			((Excel.Range)xlNewSheet.Cells[1, 5]).Value = "Stencil Label Font Size";
			((Excel.Range)xlNewSheet.Cells[1, 5]).ColumnWidth = 20.00;
			xlNewSheet.Cells[1, 5].Interior.Color = Excel.XlRgbColor.rgbDarkSeaGreen;

			((Excel.Range)xlNewSheet.Cells[2, 5]).Value = "";
			((Excel.Range)xlNewSheet.Cells[3, 5]).Value = "6";
			((Excel.Range)xlNewSheet.Cells[4, 5]).Value = "6:B";
			((Excel.Range)xlNewSheet.Cells[5, 5]).Value = "8";
			((Excel.Range)xlNewSheet.Cells[6, 5]).Value = "8:B";
			((Excel.Range)xlNewSheet.Cells[7, 5]).Value = "9";
			((Excel.Range)xlNewSheet.Cells[8, 5]).Value = "9:B";
			((Excel.Range)xlNewSheet.Cells[9, 5]).Value = "10";
			((Excel.Range)xlNewSheet.Cells[10, 5]).Value = "10:B";
			((Excel.Range)xlNewSheet.Cells[11, 5]).Value = "11";
			((Excel.Range)xlNewSheet.Cells[12, 5]).Value = "11:B";
			((Excel.Range)xlNewSheet.Cells[13, 5]).Value = "12";
			((Excel.Range)xlNewSheet.Cells[14, 5]).Value = "12:B";
			((Excel.Range)xlNewSheet.Cells[15, 5]).Value = "14";
			((Excel.Range)xlNewSheet.Cells[16, 5]).Value = "14:B";

			range = xlNewSheet.Range[xlNewSheet.Cells[1, 5], xlNewSheet.Cells[16, 5]];
			range.Borders.LineStyle = XlLineStyle.xlContinuous;
			range.Rows.AutoFit();      // auto aize the rows

			// Column G is Line Pattern
			((Excel.Range)xlNewSheet.Cells[1, 7]).Value = "Line Pattern";
			((Excel.Range)xlNewSheet.Cells[1, 7]).ColumnWidth = 20.00;
			xlNewSheet.Cells[1, 7].Interior.Color = Excel.XlRgbColor.rgbDarkSeaGreen;

			((Excel.Range)xlNewSheet.Cells[2, 7]).Value = "";
			((Excel.Range)xlNewSheet.Cells[3, 7]).Value = "Solid";
			((Excel.Range)xlNewSheet.Cells[4, 7]).Value = "Dashed";
			((Excel.Range)xlNewSheet.Cells[5, 7]).Value = "Dotted";
			((Excel.Range)xlNewSheet.Cells[6, 7]).Value = "Dash_Dot";

			range = xlNewSheet.Range[xlNewSheet.Cells[1, 7], xlNewSheet.Cells[6, 7]];
			range.Borders.LineStyle = XlLineStyle.xlContinuous;
			range.Rows.AutoFit();      // auto aize the rows

			// Column I is Stencil Label Position
			((Excel.Range)xlNewSheet.Cells[1, 9]).Value = "Stencil Label Position";
			((Excel.Range)xlNewSheet.Cells[1, 9]).ColumnWidth = 20.00;
			xlNewSheet.Cells[1, 9].Interior.Color = Excel.XlRgbColor.rgbDarkSeaGreen;

			((Excel.Range)xlNewSheet.Cells[2, 9]).Value = "";
			((Excel.Range)xlNewSheet.Cells[3, 9]).Value = "Top";
			((Excel.Range)xlNewSheet.Cells[4, 9]).Value = "Bottom";

			range = xlNewSheet.Range[xlNewSheet.Cells[1, 9], xlNewSheet.Cells[4, 9]];
			range.Borders.LineStyle = XlLineStyle.xlContinuous;
			range.Rows.AutoFit();      // auto aize the rows

			// Column K Shape Type
			((Excel.Range)xlNewSheet.Cells[1, 11]).Value = "Shape Type";
			((Excel.Range)xlNewSheet.Cells[1, 11]).ColumnWidth = 20.00;
			xlNewSheet.Cells[1, 11].Interior.Color = Excel.XlRgbColor.rgbDarkSeaGreen;
			
			((Excel.Range)xlNewSheet.Cells[2, 11]).Value = "";
			((Excel.Range)xlNewSheet.Cells[3, 11]).Value = "Template";
			((Excel.Range)xlNewSheet.Cells[4, 11]).Value = "Stencil";
			((Excel.Range)xlNewSheet.Cells[5, 11]).Value = "Page Setup";
			((Excel.Range)xlNewSheet.Cells[6, 11]).Value = "Shape";

			range = xlNewSheet.Range[xlNewSheet.Cells[1, 11], xlNewSheet.Cells[6, 11]];
			range.Borders.LineStyle = XlLineStyle.xlContinuous;
			range.Rows.AutoFit();      // auto aize the rows

			// Column N is OC_Blueprinting stencil names  (may get this from a list to make dymanic)
			((Excel.Range)xlNewSheet.Cells[1, 14]).Value = "Default Stencil Names";
			((Excel.Range)xlNewSheet.Cells[1, 14]).ColumnWidth = 20.00;
			xlNewSheet.Cells[1, 11].Interior.Color = Excel.XlRgbColor.rgbDarkSeaGreen;

			((Excel.Range)xlNewSheet.Cells[2, 14]).Value = "";

			range = xlNewSheet.Range[xlNewSheet.Cells[1, 14], xlNewSheet.Cells[2, 14]];
			range.Borders.LineStyle = XlLineStyle.xlContinuous;
			range.Rows.AutoFit();      // auto aize the rows

			// format each cell to be center justified and Left aligned in the row
			xlNewSheet.Cells.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
			xlNewSheet.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

			return false;
		}

		private void setColumnsDropdownList(DiagramData diagramData)
		{
			Excel.Range xlRange = _xlWorksheet.UsedRange;
			int startingRow = 2;
			int rowCount = xlRange.Rows.Count;
			int colCount = xlRange.Columns.Count;

			// the count will be dynamic based on the json data in the OmnicellBlueprintingTool.json.json file
			string tablesColorColumn = String.Format("=Tables!$A${0}:$A${1}",startingRow, diagramData.AppConfig.Colors.Count);
			string tablesArrowsColumn = String.Format("=Tables!$C${0}:$C${1}", startingRow, diagramData.AppConfig.Arrows.Count);
			string tablesLabelFontSizeColumn = String.Format("=Tables!$E${0}:$E${1}", startingRow, diagramData.AppConfig.LabelFontSizes.Count);
			string tablesLinePatternColumn = String.Format("=Tables!$G${0}:$G{1}", startingRow, diagramData.AppConfig.LinePatterns.Count);
			string tablesLabelPositionColumn = String.Format("=Tables!$I${0}:$I{1}", startingRow, diagramData.AppConfig.StencilLabelPosition.Count);
			string tablesShapeTypeColumn = String.Format("=Tables!$K${0}:$K${1}", startingRow, diagramData.AppConfig.ShapeTypes.Count);

			// Shape Type column
			Excel.Range xlRange1 = _xlWorksheet.get_Range(string.Format("B{0}2:B{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList,Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesShapeTypeColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// Stencil Label position
			xlRange1 = _xlWorksheet.get_Range(string.Format("F{0}:F{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesLabelPositionColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// Stencil Label Font Size
			xlRange1 = _xlWorksheet.get_Range(string.Format("G{0}:G{1}", startingRow,rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesLabelFontSizeColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// Fill Color
			xlRange1 = _xlWorksheet.get_Range(string.Format("W{0}:W{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesColorColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// From Line Pattern
			xlRange1 = _xlWorksheet.get_Range(string.Format("Z{0}:Z{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesLinePatternColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// From Arrow
			xlRange1 = _xlWorksheet.get_Range(string.Format("AA{0}:AA{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesArrowsColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// From Line Color
			xlRange1 = _xlWorksheet.get_Range(string.Format("AB{0}:AB{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesColorColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// To Line Pattern
			xlRange1 = _xlWorksheet.get_Range(string.Format("AE{0}:AE{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesLinePatternColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// To Arrow
			xlRange1 = _xlWorksheet.get_Range(string.Format("AF{0}:AF{1}",	startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesArrowsColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// To Line Color
			xlRange1 = _xlWorksheet.get_Range(string.Format("AG{0}:AG{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesColorColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;
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

	}
}
