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
using System.Text.RegularExpressions;
using System.Linq;
using System.Diagnostics;


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
		Process[] _excelProcsOld = null;

		public CreateExcelDataFile()
		{
		}

		/// <summary>
		/// openFile
		/// Open the Excel Data file for writing
		/// </summary>
		/// <param name="fileNamePath"></param>
		/// <returns>bool false:success</returns>
		private bool openFile(string fileNamePath)
		{
			_excelProcsOld = Process.GetProcessesByName("EXCEL");

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

		/// <summary>
		/// wroteHeader
		/// Write the Header information to the Excel file
		/// </summary>
		/// <param name="workSheet"></param>
		/// <param name="headerNames"></param>
		/// <param name="row"></param>
		/// <returns>int : Positive value = row number | negative number = error</returns>
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
				switch(headerName)
				{
					// these are the Excel data file column's
					// if you change these positions you must change it here also
					case "Visio Page":	// 1 Visio Page
					case "Shape Type":	// 2 Shape Type
					case "Unique Key":	// 3 Unique Key
					case "Stencil Image":	// 4 Stencil Image
					case "PosX":	// 19 PosX
					case "PosY":	// PosY
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

		/// <summary>
		/// writeConfiguration
		/// Write the configuration section to the Excel Data file
		/// </summary>
		/// <param name="workSheet"></param>
		/// <param name="diagramData"></param>
		/// <param name="visioHelper"></param>
		/// <param name="cellIndex"></param>
		/// <param name="nRow"></param>
		/// <returns>int : Positive value = row number | negative number = error</returns>
		private int writeConfiguration(Excel.Worksheet workSheet, DiagramData diagramData, VisioHelper visioHelper,int cellIndex, int nRow)
		{
			ShapeInformation shpObj = null;
			string sTmp = string.Empty;
			try
			{
				// Write comment section named "Configuration"
				shpObj = new ShapeInformation();
				shpObj.VisioPage = "";
				shpObj.ShapeType = "Disabled";
				shpObj.UniqueKey = "Visio Configuration";
				shpObj.StencilLabel = String.Empty;
				if (_writeData(workSheet, visioHelper, shpObj, nRow, true))
				{
					 sTmp = "CreateExeclDataFile::writeConfiguration \n\nFailed to write Comment data";
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);

					return -1;
				}

				// Write Template section
				shpObj = new ShapeInformation();
				nRow++;
				shpObj.VisioPage = "";
				shpObj.ShapeType = "Template";
				shpObj.UniqueKey = string.Format(@"{0}", diagramData.VisioTemplateFilePath + VisioVariables.DefaultBlueprintingTemplateFile);
				shpObj.StencilLabel = string.Format("• Use the default Application Template or use your own custom Template.\nAdd your custom stencils to the template\n\nDefault template name:'{0}'", VisioVariables.DefaultBlueprintingTemplateFile);
				if (_writeData(workSheet, visioHelper, shpObj, nRow, true))
				{
					sTmp = "CreateExeclDataFile::writeConfiguration Error\n\nFailed to write Template data";
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return -2;
				}

				// Write the Stencil data
				shpObj = new ShapeInformation();
				nRow++;
				shpObj.VisioPage = "";
				shpObj.ShapeType = "Stencil";
				shpObj.UniqueKey = string.Format(@"{0}", diagramData.VisioStencilFilePaths[0] + VisioVariables.DefaultBlueprintingStencilFile);
				shpObj.StencilLabel = string.Format("• Omnicell Blueprinting tool Stencil \"{0}\"\r\n• File Location should be where the application is installed, in the subfolder \"Data\\Stencils\"", VisioVariables.DefaultBlueprintingStencilFile);
				if (_writeData(workSheet, visioHelper, shpObj, nRow, false))
				{
					sTmp = "CreateExeclDataFile::writeConfiguration Error\n\nFailed to write Stincel data";
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return -4;
				}

				// Write the Custom Stencil data
				shpObj = new ShapeInformation();
				nRow++;
				shpObj.VisioPage = "";
				shpObj.ShapeType = "Stencil";
				shpObj.UniqueKey = string.Format(@"{0}CS_CustomStencils.vssx", diagramData.VisioStencilFilePaths[0]);
				shpObj.StencilLabel = string.Format("• Custom Stencil specific to an account\r\n• Enter the full path and file name in the Unique Key where custom stencil is located");
				if (_writeData(workSheet, visioHelper, shpObj, nRow, false))
				{
					sTmp = "CreateExeclDataFile::writeConfiguration Error\n\nFailed to write Custom Stincel data";
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return -5;
				}

				// Write Page setup Section
				shpObj = new ShapeInformation();
				nRow++;
				shpObj.VisioPage = "";
				shpObj.ShapeType = "Page Setup";
				shpObj.UniqueKey = VisioVariables.VisioPageOrientation.Portrait + ":" + VisioVariables.VisioPageSize.Legal;
				shpObj.StencilLabel = "• Orientation: Landscape or Portrait (default)\r\n• Size: Letter (default), Tabloid, Ledger, Legal, A3, A4";
				if (_writeData(workSheet, visioHelper, shpObj, nRow, true))
				{
					sTmp = "CreateExeclDataFile::writeConfiguration Error\n\nFailed to Setup Page data";
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return -6;
				}

				// write Page setup Section
				shpObj = new ShapeInformation();
				nRow++;
				shpObj.VisioPage = "";
				shpObj.ShapeType = "Page Setup";
				shpObj.UniqueKey = "Autosize:true";
				shpObj.StencilLabel = "• true - Autosize all pages\r\n• false - (default) don't Autosize the pages.\n\nIf you are using a CustomConfig.json file this should be false";
				if (_writeData(workSheet, visioHelper, shpObj, nRow, false))
				{
					sTmp = "CreateExeclDataFile::writeConfiguration Error\n\nFailed to Setup Page data";
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return -7;
				}

				// Write comment section named "Visio Section for shapes"
				shpObj = new ShapeInformation();
				nRow++;
				shpObj.VisioPage = "";
				shpObj.ShapeType = "Disabled";
				shpObj.UniqueKey = "Visio Section";
				shpObj.StencilLabel = String.Empty;
				if (_writeData(workSheet, visioHelper, shpObj, nRow, true))
				{
					sTmp = "CreateExeclDataFile::writeConfiguration Error\n\nFailed to write Visio Section Comment data";
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return -8;
				}
			}
			catch(Exception ex)
			{
				sTmp = string.Format("CreateExeclDataFile::writeConfiguration Exception\n\n{0}-{1}", ex.Message, ex.StackTrace);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			return nRow;
		}

		/// <summary>
		/// PopulateExcelDataFile
		/// higher level function that will populate the Excel Data File
		/// </summary>
		/// <param name="diagramData"></param>
		/// <param name="visioHelper"></param>
		/// <param name="shapesMap"></param>
		/// <param name="namePath"></param>
		/// <returns>bool false:success</returns>
		public bool PopulateExcelDataFile(DiagramData diagramData, VisioHelper visioHelper, Dictionary<string, ShapeInformation> shapesMap, string namePath)
		{
			bool bSave = false;
			bool bErr = false;	// successful

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
						closeExcel(false);
						return true;
					}

					nRow = writeConfiguration(_xlWorksheet, diagramData, visioHelper, ExcelVariables.GetHeaderCount(), ++nRow);
					if (nRow < 0)
					{
						sTmp = "CreateExcelDataFile::PopulateExcelDataFile\n\nWriting the configuration section";
						MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
						closeExcel(false);
						return true;
					}

					// write the Stencil data
					nRow = writeAllData(_xlWorksheet, visioHelper, shapesMap, ++nRow);
					if (nRow < 0)
					{
						sTmp = string.Format("CreateExcelDataFile::PopulateExcelDataFile\n\nWriting All Data:{0}",nRow);
						MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
						closeExcel(false);
						return true;
					}

					// format the VisioData sheet
					formatVisioDataSheet(_xlWorksheet);

					// populate the Tables sheet
					writeTableSheet(diagramData, visioHelper);

					// some column use a dropdown list so we need to setup it up
					setColumnsDropdownList(diagramData, visioHelper);

					// this should stop the check Compatibility diaglog from poping up
					_xlWorkbook.DoNotPromptForConvert = true;

					bSave = true;		// prompt user to save Excel file
				}
				catch(Exception ex)
				{
					sTmp = string.Format("Exception::PopulateExcelDataFile\n\n{0}\n{1}", ex.Message,ex.StackTrace);
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
					closeExcel(false);
					return true;
				}
				finally
				{
					if (bSave)
					{
						// save and close the excel file
						if (saveFile(namePath))
						{
							sTmp = string.Format("MainForm::The Excel data file has been created successfully\n\n{0}", namePath);
							MessageBox.Show(sTmp, "INFORMATION", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
						else
						{
							bErr = true;
						}
					}
					else
					{
						closeExcel(false);   // don't save file just close Excel 
					}
				}
			}
			return bErr;
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

		/// <summary>
		/// selectWorkSheet
		/// set focus to the Excel worksheet specified by the argument
		/// </summary>
		/// <param name="sheetName">WorkSheet Name</param>
		/// <returns>Excel.Worksheet</returns>
		private Excel.Worksheet selectWorkSheet(string sheetName)
		{
			Excel.Worksheet workSheet = _xlWorkbook.Sheets[sheetName];
			workSheet.Activate();
			return workSheet;
		}

		/// <summary>
		/// selectWorkSheet
		/// set foces to the Excel worksheet specified by the argument
		/// </summary>
		/// <param name="nIdx">index</param>
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
		private int writeAllData(Excel.Worksheet workSheet, VisioHelper visioHelper, Dictionary<string, ShapeInformation> shapesMap, int rowCount)
		{
			string visioPageName = string.Empty;
			try
			{
				int nCnt = 0;
				foreach (KeyValuePair<string, ShapeInformation> keyValue in shapesMap)
				{
					if (string.IsNullOrEmpty(keyValue.Value.ShapeType))
					{
						keyValue.Value.ShapeType = "Shape";
					}

					// we want to skip the first time
					if (!keyValue.Value.VisioPage.Equals(visioPageName, StringComparison.OrdinalIgnoreCase) && nCnt > 0)
					{
						// we have a new page lets make a comment break
						visioPageName = keyValue.Value.VisioPage; // update the new page
						ShapeInformation shpInfo = new ShapeInformation();
						shpInfo.VisioPage = visioPageName;
						shpInfo.ShapeType = "Disabled";
						shpInfo.UniqueKey = "New Page";
						_writeData(workSheet, visioHelper, shpInfo, rowCount++, true);

						// remember to write the real row after the comment row
						nCnt++;
						_writeData(workSheet, visioHelper, keyValue.Value, rowCount++, false);
					}
					else
					{
						nCnt++;
						_writeData(workSheet, visioHelper, keyValue.Value, rowCount++, false);
						visioPageName = keyValue.Value.VisioPage;
					}
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

		/// <summary>
		/// _WriteData
		/// write the shape information to the Excel row/cell
		/// </summary>
		/// <param name="workSheet">Excel worksheet</param>
		/// <param name="visioHelper">Helper functions</param>
		/// <param name="shape">Shape information to write</param>
		/// <param name="rowCount">Excel Row</param>
		/// <param name="IsComment">true:is comment</param>
		/// <returns>bool false:success</returns>
		private bool _writeData(Excel.Worksheet workSheet, VisioHelper visioHelper, ShapeInformation shape, int rowCount, bool IsComment)
		{
			try
			{
				string sTmp = string.Empty;
				string sTmp2 = string.Empty;

				// break apart the object and update the excel row based on the column value from the shapesMap
				sTmp = shape.VisioPage.Trim();;
				sTmp2 = shape.ShapeType.Trim();
				if (IsComment)
				{
					sTmp = ";";
					if (shape.ShapeType.Equals("New Page", StringComparison.CurrentCultureIgnoreCase) || rowCount > 8)
					{ 
						sTmp2 = "Disabled";
					}
				}

				((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.VisioPage]).Value = sTmp;
				((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ShapeType]).Value = sTmp2;
				((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.UniqueKey]).Value = shape.UniqueKey;
				((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.StencilImage]).Value = shape.StencilImage;

				((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.StencilLabel]).Value = shape.StencilLabel;
				((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.IP]).Value = string.Empty;
				((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Ports]).Value = string.Empty;
				
				// we need to check if an IP address in in the label name
				// if so cut it out and place into the IP excel cell (xxx.xxx.xxx.xxx)
				// also check if there is a PORT and place into the Port excel cell PORT: xxxxxx
				string sLabel = shape.StencilLabel;
				if (!string.IsNullOrEmpty(shape.StencilImage))
				{
					if (shape.StencilImage.IndexOf("Server", StringComparison.CurrentCultureIgnoreCase) >= 0)
					{
						const string sIP_Port = @"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}:\d{1,5}\b";
						Regex ip = new Regex(sIP_Port, RegexOptions.IgnoreCase);
						MatchCollection result = ip.Matches(sLabel);
						if (result.Count > 0)
						{
							// we have something to work on
							string sIP = result[0].Value.ToString().Trim();
							((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.IP]).Value = sIP;

							string[] saTmp = sIP.Split(':');
							if (saTmp.Length > 0)
							{
								((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.IP]).Value = saTmp[0];
								((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Ports]).Value = saTmp[1];

								// we need to strip out the IP:Port information from the label
								string sLbl = sLabel;
								int foundIdx = sLabel.IndexOf(sIP);
								int fountLen = sIP.Length; // should be the length of the IP:Port string

								sTmp = sLabel.Substring(0, foundIdx - 1); // get the first part of the original string minus the IP:Port
																						// now we need to remove the IP:Port and append the rest of the original string is there is anything
								if ((foundIdx + sIP.Length) < sLabel.Length)
								{
									sTmp2 = sLabel.Substring((foundIdx + sIP.Length));
									sTmp += string.Format(" {0}", sTmp2);
								}
								shape.StencilLabel = sTmp;
							}
						}
						else
						{
							const string sIP_Only = @"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b";
							ip = new Regex(sIP_Only, RegexOptions.IgnoreCase);
							result = ip.Matches(sLabel);
							if (result.Count > 0)
							{
								// we have something to work on
								string sIP = result[0].Value.ToString().Trim();
								((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.IP]).Value = sIP;

								// we need to strip out the IP:Port information from the label
								string sLbl = sLabel;
								int foundIdx = sLabel.IndexOf(sIP);
								sTmp = sLabel.Substring(0, foundIdx - 1); // first part
								
								// now we need to remove the IP:Port and append the rest of the original string is there is anything
								if ((foundIdx + sIP.Length) < sLabel.Length)
								{
									sTmp2 = sLabel.Substring((foundIdx + sIP.Length));
									sTmp += string.Format(" {0}", sTmp2);
								}
								shape.StencilLabel = sTmp;
							}
						}
					}
				}

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
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.DevicesCount]).Value = string.Empty;

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.PosX]).Value = shape.Pos_x;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.PosX]).NumberFormat = "#0.000";

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.PosY]).Value = shape.Pos_y;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.PosY]).NumberFormat = "#0.000";

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Width]).Value = shape.Width;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Width]).NumberFormat = "#0.000";

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Height]).Value = shape.Height;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.Height]).NumberFormat = "#0.000";

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FillColor]).Value = shape.FillColor; // should be color name
					if (string.IsNullOrEmpty(shape.FillColor))
					{
						// we don't want to set the fill color for these shapes
						if (shape.StencilImage.IndexOf("OC_Logo", StringComparison.OrdinalIgnoreCase) < 0 && 
							 shape.StencilImage.IndexOf("OC_Title", StringComparison.OrdinalIgnoreCase) < 0 &&
							 shape.StencilImage.IndexOf("OC_Footer", StringComparison.OrdinalIgnoreCase) < 0)
						{
							((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.rgbFillColor]).Value = shape.RGBFillColor;
						}
					}

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ConnectFrom]).Value = shape.ConnectFrom;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromLineLabel]).Value = shape.FromLineLabel;

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromLinePattern]).Value = "";  // default to solid
					// we only want to populate this field if we are connected to another shape
					if (!string.IsNullOrEmpty(shape.ConnectFrom))
					{
						sTmp = visioHelper.GetConnectorLinePattern(shape.FromLinePattern);
						if (!string.IsNullOrEmpty(sTmp))
						{
							((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromLinePattern]).Value = sTmp;
						}
					}
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromArrowType]).Value = shape.FromArrowType;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromLineColor]).Value = shape.FromLineColor;

					// get the To Line Weight value.  We only need to set it if it's not default 1 pt
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromLineWeight]).Value = ""; // set default
					sTmp = shape.FromLineWeight;
					if (!string.IsNullOrEmpty(sTmp))
					{
						if (visioHelper.IsConnectorLineWeight(sTmp))
						{
							((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.FromLineWeight]).Value = sTmp;
						}
					}

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ConnectTo]).Value = shape.ConnectTo;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToLineLabel]).Value = shape.ToLineLabel;

					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToLinePattern]).Value = ""; // use default
					// we only want to populate this field if we are connected to another shape
					if (!string.IsNullOrEmpty(shape.ConnectTo))
					{
						sTmp = visioHelper.GetConnectorLinePattern(shape.ToLinePattern);
						if (!string.IsNullOrEmpty(sTmp))
						{
							((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToLinePattern]).Value = sTmp;
						}
					}
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToArrowType]).Value = shape.ToArrowType;
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToLineColor]).Value = shape.ToLineColor;

					// get the To Line Weight value
					((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToLineWeight]).Value = "";	// set default
					sTmp = shape.ToLineWeight;
					if (!string.IsNullOrEmpty(sTmp))
					{
						if (visioHelper.IsConnectorLineWeight(sTmp))
						{
							// not "1 pt" so lets check if valid entry and if so persist it
							((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.ToLineWeight]).Value = sTmp;
						}
					}
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

		/// <summary>
		/// formatVisioDataSheet
		/// this function will format the VisioData sheet
		/// Header, columns widths, Colors
		/// </summary>
		/// <param name="workSheet"></param>
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
					string sTmp = data.Text.ToString().Trim();
					switch (sTmp)
					{
						case "Unique Key":
						case "Stencil Label":
							((Excel.Range)workSheet.Cells[1, nCol]).ColumnWidth = 37.00;
							break;

						case "Visio Page":
						case "Shape Type":
						case "Stencil Image":
							((Excel.Range)workSheet.Cells[1, nCol]).ColumnWidth = 16.00;
							break;

						case "PosX":
						case "PosY":
						case "Width":
						case "Height":
							((Excel.Range)workSheet.Cells[1, nCol]).ColumnWidth = 8.00;
							((Excel.Range)workSheet.Cells[rowCount, ExcelVariables.CellIndex.PosX]).NumberFormat = "#0.000";
							break;

						case "Mach_name":
						case "Mach_id":
						case "Site_id":
						case "Site_name":
						case "Site_address":
						case "Omnis_name":
						case "Omnis_id":
						case "SiteId_OmnisId":
						case "Fill Color":
							((Excel.Range)workSheet.Cells[1, nCol]).ColumnWidth = 8.00;
							break;

						case "Connect From":
						case "Connect To":
						case "From Line Label":
						case "To Line Label":
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
		private bool writeTableSheet(DiagramData diagramData, VisioHelper visioHelper)
		{
			int startingRow = 1;
			Excel.Worksheet xlNewSheet = selectWorkSheet("Tables");

			try
			{
				// column A is Colors
				((Excel.Range)xlNewSheet.Cells[1, 1]).Value = "Color";
				((Excel.Range)xlNewSheet.Cells[1, 1]).ColumnWidth = 20.00;
				xlNewSheet.Cells[startingRow++, 1].Interior.Color = Excel.XlRgbColor.rgbDarkSeaGreen;
				List<string> lTmp = visioHelper.GetAllColorNames();
				if (lTmp != null && lTmp.Count > 0)
				{
					foreach (var item in lTmp)
					{
						((Excel.Range)xlNewSheet.Cells[startingRow++, 1]).Value = item;
					}
				}
				Excel.Range range = xlNewSheet.Range[xlNewSheet.Cells[2, 1], xlNewSheet.Cells[startingRow - 1, 1]];
				range.Borders.LineStyle = XlLineStyle.xlContinuous;
				range.Rows.AutoFit();      // auto aize the rows

				// Column C is Arrows
				((Excel.Range)xlNewSheet.Cells[1, 3]).Value = "Arrows";
				((Excel.Range)xlNewSheet.Cells[1, 3]).ColumnWidth = 20.00;
				xlNewSheet.Cells[1, 3].Interior.Color = Excel.XlRgbColor.rgbDarkSeaGreen;
				List<string> strArray = visioHelper.GetConnectorArrowTypes();
				if (strArray.Count > 0)
				{
					int nRow = 2;
					foreach (string str in strArray)
					{
						((Excel.Range)xlNewSheet.Cells[nRow++, 3]).Value = str;
					}
				}
				range = xlNewSheet.Range[xlNewSheet.Cells[2, 3], xlNewSheet.Cells[strArray.Count + 1, 3]];
				range.Borders.LineStyle = XlLineStyle.xlContinuous;
				range.Rows.AutoFit();      // auto aize the rows

				// Column E is Stencil Label Font size
				((Excel.Range)xlNewSheet.Cells[1, 5]).Value = "Stencil Label Font Size";
				((Excel.Range)xlNewSheet.Cells[1, 5]).ColumnWidth = 20.00;
				xlNewSheet.Cells[1, 5].Interior.Color = Excel.XlRgbColor.rgbDarkSeaGreen;
				strArray = visioHelper.GetStencilLabelFontSize();
				if (strArray.Count > 0)
				{
					int nRow = 2;
					foreach (string str in strArray)
					{
						((Excel.Range)xlNewSheet.Cells[nRow++, 5]).Value = str;
					}
				}
				range = xlNewSheet.Range[xlNewSheet.Cells[2, 5], xlNewSheet.Cells[strArray.Count + 1, 5]];
				range.Borders.LineStyle = XlLineStyle.xlContinuous;
				range.Rows.AutoFit();      // auto aize the rows

				// Column G is Line Pattern
				((Excel.Range)xlNewSheet.Cells[1, 7]).Value = "Line Pattern";
				((Excel.Range)xlNewSheet.Cells[1, 7]).ColumnWidth = 20.00;
				xlNewSheet.Cells[1, 7].Interior.Color = Excel.XlRgbColor.rgbDarkSeaGreen;
				strArray = visioHelper.GetConnectorLinePatterns();
				if (strArray.Count > 0)
				{
					int nRow = 2;
					foreach(string str in strArray)
					{
						((Excel.Range)xlNewSheet.Cells[nRow++, 7]).Value = str;
					}
				}
				range = xlNewSheet.Range[xlNewSheet.Cells[2, 7], xlNewSheet.Cells[strArray.Count + 1, 7]];
				range.Borders.LineStyle = XlLineStyle.xlContinuous;
				range.Rows.AutoFit();      // auto aize the rows

				// Column I is Stencil Label Position
				((Excel.Range)xlNewSheet.Cells[1, 9]).Value = "Stencil Label Position";
				((Excel.Range)xlNewSheet.Cells[1, 9]).ColumnWidth = 20.00;
				xlNewSheet.Cells[1, 9].Interior.Color = Excel.XlRgbColor.rgbDarkSeaGreen;
				strArray = visioHelper.GetStencilLabelPositions();
				if (strArray.Count > 0)
				{
					int nRow = 2;
					foreach (string str in strArray)
					{
						((Excel.Range)xlNewSheet.Cells[nRow++, 9]).Value = str;
					}
				}
				range = xlNewSheet.Range[xlNewSheet.Cells[2, 9], xlNewSheet.Cells[strArray.Count + 1, 9]];
				range.Borders.LineStyle = XlLineStyle.xlContinuous;
				range.Rows.AutoFit();      // auto aize the rows

				// Column K Shape Type
				((Excel.Range)xlNewSheet.Cells[1, 11]).Value = "Shape Type";
				((Excel.Range)xlNewSheet.Cells[1, 11]).ColumnWidth = 20.00;
				xlNewSheet.Cells[1, 11].Interior.Color = Excel.XlRgbColor.rgbDarkSeaGreen;
				 strArray = visioHelper.GetShapeTypes();
				if (strArray.Count > 0)
				{
					int nRow = 2;
					foreach (string str in strArray)
					{
						((Excel.Range)xlNewSheet.Cells[nRow++, 11]).Value = str;
					}
				}
				range = xlNewSheet.Range[xlNewSheet.Cells[2, 11], xlNewSheet.Cells[strArray.Count + 1, 11]];
				range.Borders.LineStyle = XlLineStyle.xlContinuous;
				range.Rows.AutoFit();      // auto aize the rows

				// Column M Connector Line Weight
				((Excel.Range)xlNewSheet.Cells[1, 13]).Value = "Line Weight";
				((Excel.Range)xlNewSheet.Cells[1, 13]).ColumnWidth = 20.00;
				xlNewSheet.Cells[1, 13].Interior.Color = Excel.XlRgbColor.rgbDarkSeaGreen;
				strArray = visioHelper.GetConnectorLineWeights();
				if (strArray.Count > 0)
				{
					int nRow = 2;
					foreach (string str in strArray)
					{
						((Excel.Range)xlNewSheet.Cells[nRow++, 13]).Value = str;
					}
				}
				range = xlNewSheet.Range[xlNewSheet.Cells[2, 13], xlNewSheet.Cells[strArray.Count + 1, 13]];
				range.Borders.LineStyle = XlLineStyle.xlContinuous;
				range.Rows.AutoFit();      // auto aize the rows

				// Column O is OC_Blueprinting stencil names  (may get this from a list to make dymanic)
				((Excel.Range)xlNewSheet.Cells[1, 15]).Value = "Default Stencil Names";
				((Excel.Range)xlNewSheet.Cells[1, 15]).ColumnWidth = 20.00;
				xlNewSheet.Cells[1, 15].Interior.Color = Excel.XlRgbColor.rgbDarkSeaGreen;
				strArray = visioHelper.GetDefaultStencilNames();
				if (strArray.Count > 0)
				{
					int nRow = 2;
					foreach (string str in strArray)
					{
						((Excel.Range)xlNewSheet.Cells[nRow++, 15]).Value = str;
					}
				}
				range = xlNewSheet.Range[xlNewSheet.Cells[2, 15], xlNewSheet.Cells[strArray.Count + 1, 15]];
				range.Borders.LineStyle = XlLineStyle.xlContinuous;
				range.Rows.AutoFit();      // auto aize the rows

				// format each cell to be center justified and Left aligned in the row
				xlNewSheet.Cells.Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
				xlNewSheet.Cells.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
			}
			catch(IndexOutOfRangeException ex)
			{
				ConsoleOut.writeLine(String.Format("CreateExcelDataFile::writeTableSheet - Exception\n\n{0}\n{1}",ex.Message, ex.StackTrace));
			}
			return false;
		}

		/// <summary>
		/// setColumnsDropdownList
		/// this is the connection between the Tables sheet entries and the VisioData sheet columns
		/// I need to make this more dynamic like look at the VisioData sheet for column names to get the column identifier to use here
		/// </summary>
		/// <param name="diagramData"></param>
		private void setColumnsDropdownList(DiagramData diagramData, VisioHelper visioHelper)
		{
			Excel.Range xlRange = _xlWorksheet.UsedRange;
			int startingRow = 2;
			int rowCount = xlRange.Rows.Count;
			int colCount = xlRange.Columns.Count;

			// the count will be dynamic based on the json data in the OmnicellBlueprintingTool.json.json file
			string tablesColorColumn = String.Format("=Tables!$A${0}:$A${1}", startingRow, visioHelper.GetAllColorNames().Count + 1);
			string tablesArrowsColumn = String.Format("=Tables!$C${0}:$C${1}", startingRow, visioHelper.GetConnectorArrowTypes().Count + 1);
			string tablesLabelFontSizeColumn = String.Format("=Tables!$E${0}:$E${1}", startingRow, visioHelper.GetStencilLabelFontSize().Count + 1);
			string tablesLinePatternColumn = String.Format("=Tables!$G${0}:$G${1}", startingRow, visioHelper.GetConnectorLinePatterns().Count + 1);
			string tablesLabelPositionColumn = String.Format("=Tables!$I${0}:$I${1}", startingRow, visioHelper.GetStencilLabelPositions().Count + 1);
			string tablesShapeTypeColumn = String.Format("=Tables!$K${0}:$K${1}", startingRow, visioHelper.GetShapeTypes().Count + 1);
			string tablesLineWeightColumn = String.Format("=Tables!$M${0}:$M${1}", startingRow, visioHelper.GetConnectorLineWeights().Count + 1);
			string tablesDefaultStencilNamesColumn = String.Format("=Tables!$O${0}:$O${1}", startingRow, visioHelper.GetDefaultStencilNames().Count + 1);

			// now lets link the data list to the excel columns on the VisioData sheet
			// Shape Type column
			Excel.Range xlRange1 = _xlWorksheet.get_Range(string.Format("B{0}:B{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList,Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesShapeTypeColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// Stencil Images column
			xlRange1 = _xlWorksheet.get_Range(string.Format("D{0}:D{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesDefaultStencilNamesColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// Stencil Label position column
			xlRange1 = _xlWorksheet.get_Range(string.Format("F{0}:F{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesLabelPositionColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// Stencil Label Font Size column
			xlRange1 = _xlWorksheet.get_Range(string.Format("G{0}:G{1}", startingRow,rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesLabelFontSizeColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// Fill Color column
			xlRange1 = _xlWorksheet.get_Range(string.Format("W{0}:W{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesColorColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// From Line Pattern column
			xlRange1 = _xlWorksheet.get_Range(string.Format("AA{0}:AA{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesLinePatternColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// From Arrow column
			xlRange1 = _xlWorksheet.get_Range(string.Format("AB{0}:AB{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesArrowsColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// From Line Color column
			xlRange1 = _xlWorksheet.get_Range(string.Format("AC{0}:AC{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesColorColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// From Line Weight column
			xlRange1 = _xlWorksheet.get_Range(string.Format("AD{0}:AD{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesLineWeightColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// To Line Pattern column
			xlRange1 = _xlWorksheet.get_Range(string.Format("AG{0}:AG{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesLinePatternColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// To Arrow column
			xlRange1 = _xlWorksheet.get_Range(string.Format("AH{0}:AH{1}",	startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesArrowsColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// To Line Color column
			xlRange1 = _xlWorksheet.get_Range(string.Format("AI{0}:AI{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesColorColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;

			// To Line Weight column
			xlRange1 = _xlWorksheet.get_Range(string.Format("AJ{0}:AJ{1}", startingRow, rowCount)).EntireColumn;
			xlRange1.Validation.Delete();
			xlRange1.Validation.Add(XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop,
						Excel.XlFormatConditionOperator.xlBetween, tablesLineWeightColumn, Type.Missing);
			xlRange1.Validation.InCellDropdown = true;
			xlRange1.Validation.IgnoreBlank = false;
		}
		
		/// <summary>
		/// saveFile
		/// Prompt the user if they want to save the file
		/// if No the file will not be saved
		/// if Yes the file will be saved into the ExcelData folder or a folder that was specificed
		/// All excel objects will be closed
		/// </summary>
		/// <param name="fileNamePath"></param>
		/// <returns>bool true:success</returns>
		private bool saveFile(string fileNamePath)
		{
			bool bSave = true;
			try
			{
				if (_xlWorkbook != null)
				{
					//Here saving the file in xlsx
					_xlWorkbook.SaveAs(fileNamePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing,
					Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

					//_xlWorkbook.SaveAs(fileNamePath, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					//						Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				}
			}
			catch (Exception)
			{
				bSave = false;
			}
			finally
			{
				closeExcel(bSave);
			}
			return bSave;
		}

		/// <summary>
		/// closeExcel
		/// 
		/// </summary>
		/// <param name="bSave">
		/// <option>true display save prompt</option>
		/// <option>false do not show save prompt</option>
		/// </param>
		private void closeExcel(bool bSave)
		{
			//quit and release
			_xlWorkbook.Close(bSave, Type.Missing, Type.Missing);
			_xlApp.Quit();

			//release com objects to fully kill excel process from running in the background
			releaseObject(_xlApp);
			releaseObject(_xlWorkbook);
			releaseObject(_xlWorksheet);

			killExcelProcesses(_excelProcsOld);
		}

		private void releaseObject(object obj)
		{
			try
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
				obj = null;
			}
			catch (Exception ex)
			{
				obj = null;
				string sTmp = string.Format("ProcessExcelDataFile::releaseObject - Exception\n\nUnable to release the Object:{0}\n{1}", ex.Message, ex.StackTrace);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				GC.Collect();
			}
		}

		/// <summary>
		/// killExcelProcesses
		/// kill all the excell processes to ensure everything is released
		/// Keep any existing Excel processes already opened before this app has started
		/// </summary>
		/// <param name="excelProcsOld"></param>
		private void killExcelProcesses(Process[] excelProcsOld)
		{
			//Compare the EXCEL ID and Kill it 
			Process[] excelProcsNew = Process.GetProcessesByName("EXCEL");
			foreach (Process procNew in excelProcsNew)
			{
				int exist = 0;
				if (excelProcsOld != null)
				{
					foreach (Process procOld in excelProcsOld)
					{
						if (procNew.Id == procOld.Id)
						{
							exist++;
						}
					}
				}
				if (exist == 0)
				{
					procNew.Kill();
				}
			}
		}

	}
}
