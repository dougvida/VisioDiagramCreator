using OmnicellBlueprintingTool.Models;
using OmnicellBlueprintingTool.Visio;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab
using OmnicellBlueprintingTool.ExcelHelpers;
using static OmnicellBlueprintingTool.Visio.VisioVariables;
using OmnicellBlueprintingTool.Configuration;
using System.Text.RegularExpressions;

namespace OmnicellBlueprintingTool.ExcelHelpers
{
	public class ProcessExcelDataFile
	{

		/// <summary>
		/// parseExcelFile
		/// Parse the Excel data into a DiagramData class.
		/// this class will hold all the excel data that will be used to transfer into Visio diagram data
		/// </summary>
		/// <param name="file">Visio File to load</param>
		/// <param name="diagData">DiagramData</param>
		/// <returns>DiagramData</returns>
		/// <exception cref="ArgumentNullException"></exception>
		/// <exception cref="Exception"></exception>
		public DiagramData parseExcelFile(string file, DiagramData diagData, ref VisioHelper visioHelper)
		{
			string sTmp = string.Empty;

			if (string.IsNullOrEmpty(file))
			{
				// Error file is empty
				sTmp = string.Format("ProcessExcelDataFile::parseExcelFile - Exception\n\n(File is missing: {0})", nameof(file));
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return null;
			}

			List<Device> devices = new List<Device>();
			Device device = null;
			diagData.VisioStencilFilePaths = new List<string>();

			Process[] excelProcsOld = Process.GetProcessesByName("EXCEL");       
			
			//Create COM Objects. Create a COM object for everything that is referenced
			Excel.Application xlApp = new Excel.Application();
			Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
			Excel._Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

			Excel.Range xlRange = xlWorksheet.UsedRange;

			int rowCount = xlRange.Rows.Count;
			int colCount = xlRange.Columns.Count;

			System.Array myArray = (System.Array)xlRange.Cells.Value2;
			try
			{
				for (int row = 2; row <= xlRange.Rows.Count; row++)
				{
					// Visio Page column can contain
					// 1. first character ';' this is a comment and this line will be ignored
					// 2. the Page name text.  if a single numeric value is use the page will be "Page-#"
					// 3. if Blank it could be part of configuration section  Viso Page number are not needed for non Stencil types
					var visPage = myArray.GetValue(row, (int)ExcelVariables.CellIndex.VisioPage);
					if (visPage == null) // value is null skip this column
					{
						// because first column will contain ';' or a numeric or blank value
						// this is a cluster don't have time to fix correctly
						// so look at the 2nd column because this is normally filled. if VisioPage is blank
						// it's a good row to process
						if (myArray.GetValue(row, (int)ExcelVariables.CellIndex.ShapeType) == null)
						{
							continue;   // both VisioPage and Shape Type are blank so lets skip this row
						}
					}
					else
					{
						// check if this is a comment we need to skip
						if (visPage.ToString().StartsWith(";"))
						{
							// this is a comment skip it
							continue;
						}
					}

					var shapeType = myArray.GetValue(row, (int)ExcelVariables.CellIndex.ShapeType);
					if (shapeType != null)
					{
						if (shapeType.ToString().Trim().Equals("Disabled", StringComparison.OrdinalIgnoreCase))
						{
							continue;	// skip this entry it like a comment
						}
						if (visPage == null && shapeType.ToString().Trim().Equals("Shape",StringComparison.OrdinalIgnoreCase ))
						{
							// error condition.  Excel data is formatted incorrect.  Must have a valid Visio Page value
							sTmp = string.Format("parseExcelFile - Bad formatting found in the Excel file.  Column 'Visio Page' at Row:{0}\nThis value cannot be null if ShapeType is set to 'Stencil'.", row);
							ConsoleOut.writeLine(sTmp);
							MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return null;	// signal error
						}
					}

					// If we got here than this is something we need to process
					var data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.ShapeType);
					if (data != null)
					{
						switch ((string)data.ToString().Trim().ToUpper())
						{
							case "DISABLED":
								// Should not get here should be handled above
								// ShapeType is Disabled so we will ignore this entry
								ConsoleOut.writeLine(string.Format("parseExcelFile - Ignoring this row.  ShapeType:'{0}', Row:{1}", data, row));
								break;

							case "PAGE SETUP":           // Visio Page setup/Size
								data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.UniqueKey);
								if (data != null)
								{
									// need to split up the string  orientation:size
									data = (string)data.ToString().Trim();
									if (data != null)
									{
										string[] saTmp = data.ToString().Split(':');
										switch (saTmp[0].ToUpper())
										{
											case "AUTOSIZE":
												diagData.AutoSizeVisioPages = false;
												if (!string.IsNullOrEmpty(saTmp[1]))
												{
													// the string must be true anything else is default false
													if (saTmp[1].ToUpper().Equals("TRUE"))
													{
														// user wants to Autosize the pages
														diagData.AutoSizeVisioPages = true;
													}
												}
												ConsoleOut.writeLine(string.Format("parseExcelFile - Processing Page Setup - Autosize:{0} found at Row:{1}", saTmp[1], row));
												break;

											case "PORTRAIT":
												diagData.VisioPageOrientation = VisioVariables.VisioPageOrientation.Portrait;
												diagData.VisioPageSize = VisioVariables.VisioPageSize.Letter;
												if (!string.IsNullOrEmpty(saTmp[1]))
												{
													// change the page size if needed
													diagData.VisioPageSize = VisioVariables.GetVisioPageSize(saTmp[1].Trim());
												}
												ConsoleOut.writeLine(string.Format("parseExcelFile - Processing Page Setup - Portrait:{0} found at Row:{1}", saTmp[1], row));
												break;

											case "LANDSCAPE":
												diagData.VisioPageOrientation = VisioVariables.VisioPageOrientation.Landscape;
												diagData.VisioPageSize = VisioVariables.VisioPageSize.Letter;
												if (!string.IsNullOrEmpty(saTmp[1]))
												{
													// change the page size if needed
													diagData.VisioPageSize = VisioVariables.GetVisioPageSize(saTmp[1].Trim());
												}
												ConsoleOut.writeLine(string.Format("parseExcelFile - Processing Page Setup - Landscape:{0} found at Row:{1}", saTmp[1], row));
												break;
											
											default:
												diagData.AutoSizeVisioPages = false;
												diagData.VisioPageOrientation = VisioVariables.VisioPageOrientation.Portrait;
												diagData.VisioPageSize = VisioVariables.VisioPageSize.Letter;
												ConsoleOut.writeLine(string.Format("parseExcelFile - Processing Page Setup - Using Default  AutoSizeVisioPage=false, Portraid:Letter found at Row:{0}", row));
												break;
										}
									}
								}
								break;

							case "TEMPLATE":
								// Open a template.  This may be used with existing stencilsList already in the document
								data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.UniqueKey);
								if (data != null)
								{
									diagData.VisioTemplateFilePath = (string)data.ToString().Trim();
									ConsoleOut.writeLine(string.Format("parseExcelFile - Adding Template file:'{0}' found at Row:{1}", diagData.VisioTemplateFilePath, row));
								}
								break;

							case "STENCIL":            // stencilsList to add
								data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.UniqueKey);
								string stencilFile = data.ToString();// but we only want the first part of the key
								if (!string.IsNullOrEmpty(stencilFile))
								{
									diagData.VisioStencilFilePaths.Add(stencilFile);
								}
								ConsoleOut.writeLine(string.Format("parseExcelFile - Adding Stencil file:'{0}' found at Row:{1}", stencilFile, row));
								break;

							case "SHAPE":  
								// device will contain all the Visio shapes to be drawn on the Visio Diagram
								// myArray contains all the excel rows to process
								// Excel Document column "Visio Page" can be numeric or Alphanumeric text   This will be the page name on the Visio Document
								//     you can have multiple pages within the Excel Data file if you want to seperate drawings on different pages
								// I.E. if the value of the column "Visio Page" is "1" this will use a name 'Page-1'  Visio will not allow for single digit tabs
								// I.E. if the value of the column "Visio Page" is "This is first Page", this will be the name of the tab
								// page tab order will depend on the Excel Data file shapes order
								// please group all the shapes by Page when possible
								device = _parseExcelData(ref visioHelper, myArray, row);
								if (device != null)
								{
									devices.Add(device);
									if (diagData.AllShapesMap.ContainsKey(device.ShapeInfo.UniqueKey))
									{
										sTmp = string.Format("\"parseExcelFile - Duplicate key '{0}' found at Row:{1}.\n\nRemember uniqueKey values can't be same accross pages.", device.ShapeInfo.UniqueKey, row);
										ConsoleOut.writeLine(sTmp);
										MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
									}
									else
									{
										diagData.AllShapesMap.Add(device.ShapeInfo.UniqueKey, device);
										ConsoleOut.writeLine(string.Format("parseExcelFile - Adding Shape '{0}' found at Row:{1}", device.ShapeInfo.UniqueKey, row));
									}
								}
								break;

							default:
								// the finally will be called to clean / close up everything
								sTmp = String.Format("ProcessExcelDataFile::parseExcelFile\n\nInvalid 'ShapeType' found:({0}) at Row:{1}\n\nPlease resolve this issue in the Excel Data file", data, row);
								ConsoleOut.writeLine(sTmp);
								MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
								return null;
						}
					}
				}
			}
			catch (Exception ex)
			{
				ConsoleOut.writeLine(ex.Message + " - " + ex.StackTrace);
				sTmp = string.Format("ProcessExcelDataFile::parseExcelFile - Exception\n\nDuplicate key:({0}) found.\nPlease resolve this issue in the Excel Data file\n{1}\n{2}", device.ShapeInfo.UniqueKey, ex.Message, ex.StackTrace);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return null;
			}
			finally
			{
				if (diagData != null)
				{
					diagData.Devices = devices;
				}

				//quit and release
				xlWorkbook.Close(true, Type.Missing, Type.Missing);
				xlApp.Quit();

				//release com objects to fully kill excel process from running in the background
				releaseObject(xlApp);
				releaseObject(xlWorkbook);
				releaseObject(xlWorksheet);
				releaseObject(xlRange);

				killExcelProcesses(excelProcsOld);
			}

			// set the MaxVisioPages variable.  Only used when creating the Visio Document.  Use the pageObj[#] for document count and names
			diagData.MaxVisioPages = visioHelper.GetVisioPageNames().Count;
			return diagData;
		}

		/// <summary>
		/// _parseExcelData
		/// this will parse the data from the excel file
		/// 
		/// </summary>
		/// <param name="myArray">excel data</param>
		/// <param name="row">array row to index on</param>
		/// <returns>Device</returns>
		private Device _parseExcelData(ref VisioHelper visioHelper, System.Array myArray, int row)
		{
			//string sColor = string.Empty;
			Device device = new Device();
			ShapeInformation visioInfo = new ShapeInformation();
			try
			{
				var data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.VisioPage);
				if (data != null)
				{
					int value = -1;
					if (Regex.Matches(data.ToString().Trim(), @"[a-zA-Z]").Count <= 0)
					{
						// No characters found convert to Int.    should be 1-XX for just a page number
						if (Int32.TryParse(data.ToString().Trim(), out value))
						{
							visioInfo.VisioPage = string.Format("Page-{0}", value);
						}
					}
					else
					{
						visioInfo.VisioPage = data.ToString().Trim();
					}
					visioHelper.AddVisioPageName(visioInfo.VisioPage);
				}

				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.ShapeType);
				if (data != null)
				{
					visioInfo.ShapeType = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.UniqueKey);
				if (data != null)
				{
					visioInfo.UniqueKey = data.ToString().Trim();   // unique key for this shape
				}

				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.StencilImage);
				if (data != null)
				{
					visioInfo.StencilImage = data.ToString().Trim(); // must match exactly the name in the visio stencil
				}

				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.StencilLabel);
				if (data != null)
				{
					visioInfo.StencilLabel = data.ToString().Trim(); // text to add to the stencil image
				}

				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.StencilLabelPosition);
				if (data != null)
				{
					if (data.ToString().Trim().ToUpper().Equals(VisioVariables.STINCEL_LABEL_POSITION_BOTTOM.ToUpper()))
					{
						visioInfo.StencilLabelPosition = VisioVariables.STINCEL_LABEL_POSITION_BOTTOM; // text to add to the stencil image
					}
				}

				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.StencilLabelFontSize);
				if (data != null)
				{
					visioInfo.StencilLabelFontSize = data.ToString().Trim(); // stencil fontsize to use.  If blank use stencil text size
				}
				// decode font size if needed
				if (!string.IsNullOrEmpty(visioInfo.StencilLabelFontSize))
				{
					// get the value can be like   "12:B" or just a number
					// check if size if over 14 default to stencil size
					// if < 6 default to stencil size
					// if the letter 'B' is found we need to make bold
					// use regex to separate
					string[] saTmp = visioInfo.StencilLabelFontSize.Split(':');
					visioInfo.StencilLabelFontSize = saTmp[0].Trim();
					if (Int32.Parse(visioInfo.StencilLabelFontSize) > 14 || Int32.Parse(visioInfo.StencilLabelFontSize) < 6)
					{
						visioInfo.StencilLabelFontSize = String.Empty;  // too small or too large so default to stencil size
						visioInfo.isStencilLabelFontBold = false;       // also change to narmal weight
					}
					else
					{
						if (saTmp.Length > 1)
						{
							if (saTmp[1].ToUpper() == "B")
							{
								visioInfo.isStencilLabelFontBold = true;
								//visioInfo.LineWeight = VisioVariables.sLINE_WEIGHT_2;
							}
						}
					}
				}

				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.Mach_name);
				if (data != null)
				{
					device.MachineName = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.Mach_id);
				if (data != null)
				{
					device.MachineId = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.Site_id);
				if (data != null)
				{
					device.SiteId = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.Site_name);
				if (data != null)
				{
					device.SiteName = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.Site_address);
				if (data != null)
				{
					device.SiteAddress = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.Omnis_name);
				if (data != null)
				{
					device.OmniName = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.Omnis_id);
				if (data != null)
				{
					device.OmniId = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.SiteIdOmniId);
				if (data != null)
				{
					device.SiteId_OmniId = data.ToString().Trim();
				}

				// shape IP address value
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.IP);
				if (data != null)
				{
					device.OmniIP = data.ToString().Trim();
				}

				// shape Port value
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.Ports);
				if (data != null)
				{
					device.OmniPorts = data.ToString().Trim();
				}

				// shape position X
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.PosX);
				if (data != null)
				{
					visioInfo.Pos_x = Convert.ToDouble(data);
				}

				// shape position Y
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.PosY);
				if (data != null)
				{
					visioInfo.Pos_y = Convert.ToDouble(data);
				}

				// shape width
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.Width);
				if (data != null)
				{
					visioInfo.Width = Convert.ToDouble(data);
				}

				// shape height
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.Height);
				if (data != null)
				{
					visioInfo.Height = Convert.ToDouble(data);
				}

				// shape label
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.DevicesCount);
				if (data != null)
				{
					visioInfo.StencilLabel += " / " + data.ToString().Trim();
				}

				// shape fill color.  Should be the text color if being read from the excel data file not rgb
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.FillColor);
				if (data != null)
				{
					visioInfo.FillColor = data.ToString().Trim();
				}

				// shape rgb fill color.  This is normally populated by processing a visio diagram into an excel data file
				// however, if the color you want is not available you may put the RGB value here "RGB(128,128,128)"
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.rgbFillColor);
				if (data != null)
				{
					visioInfo.rgbFillColor = data.ToString().Trim();
				}

				// connector from shape
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.ConnectFrom);
				if (data != null)
				{
					visioInfo.ConnectFrom = data.ToString().Trim();
				}

				// connector label
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.FromLineLabel);
				if (data != null)
				{
					visioInfo.FromLineLabel = data.ToString().Trim();
				}

				// connector Line pattern
				string sTmp = string.Empty;
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.FromLinePattern);
				if (data != null)
				{
					sTmp = data.ToString().Trim();
				}
				switch (sTmp)
				{
					case VisioVariables.sLINE_PATTERN_DASHED:
						visioInfo.FromLinePattern = VisioVariables.LINE_PATTERN_DASH;
						break;

					case VisioVariables.sLINE_PATTERN_DOTTED:
						//visioInfo.LineWeight = VisioVariables.sLINE_WEIGHT_2;
						visioInfo.FromLinePattern = VisioVariables.LINE_PATTERN_DOTTED;
						break;

					case VisioVariables.sLINE_PATTERN_DASHDOT:
						visioInfo.FromLinePattern = VisioVariables.LINE_PATTERN_DASHDOT;
						break;

					default:
						case VisioVariables.sLINE_PATTERN_SOLID:	
						visioInfo.FromLinePattern = VisioVariables.LINE_PATTERN_SOLID;
						break;
				}

				// connector Arrow type
				sTmp = string.Empty;
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.FromArrowType);
				if (data != null)
				{
					sTmp = data.ToString().Trim();
				}
				switch (sTmp)
				{
					case VisioVariables.sARROW_START:
						visioInfo.FromArrowType = VisioVariables.sARROW_START;
						break;
					case VisioVariables.sARROW_END:
						visioInfo.FromArrowType = VisioVariables.sARROW_END;
						break;
					case VisioVariables.sARROW_BOTH:
						visioInfo.FromArrowType = VisioVariables.sARROW_BOTH;
						break;
					default:
						visioInfo.FromArrowType = VisioVariables.sARROW_NONE;
						break;
				}

				// connector line color
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.FromLineColor);
				visioInfo.FromLineColor = "";
				if (data != null)
				{
					if (!string.IsNullOrEmpty(visioHelper.GetRGBColor(data.ToString())))
					{
						// value was found as a color
						visioInfo.FromLineColor = data.ToString();
					}
				}

				// From connector Line weight
				sTmp = string.Empty;
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.FromLineWeight);
				if (data != null)
				{
					sTmp = data.ToString().Trim();
					sTmp = visioHelper.FindConnectorLineWeight(sTmp);
					if (!string.IsNullOrEmpty(sTmp))
					{
						visioInfo.FromLineWeight = sTmp;
					}
				}


				//
				// Get the To section
				//

				// connect to shape
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.ConnectTo);
				if (data != null)
				{
					if (!string.IsNullOrEmpty(data.ToString().Trim()))    // unique key for the To shape identifier - will match another shape object field 2 or empty for no connection
					{
						visioInfo.ConnectTo = data.ToString().Trim();
					}
				}

				// connector label
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.ToLineLabel);
				if (data != null)
				{
					visioInfo.ToLineLabel = data.ToString().Trim();
				}

				// connector line pattern
				sTmp = string.Empty;
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.ToLinePattern);
				if (data != null)
				{
					sTmp = data.ToString().Trim();
				}
				switch (sTmp)
				{
					case VisioVariables.sLINE_PATTERN_DASHED:
						visioInfo.ToLinePattern = VisioVariables.LINE_PATTERN_DASH;
						break;

					case VisioVariables.sLINE_PATTERN_DOTTED:
						//visioInfo.LineWeight = VisioVariables.sLINE_WEIGHT_2;
						visioInfo.ToLinePattern = VisioVariables.LINE_PATTERN_DOTTED;
						break;

					case VisioVariables.sLINE_PATTERN_DASHDOT:
						visioInfo.ToLinePattern = VisioVariables.LINE_PATTERN_DASHDOT;
						break;

					default:
					case VisioVariables.sLINE_PATTERN_SOLID:
						visioInfo.ToLinePattern = VisioVariables.LINE_PATTERN_SOLID;
						break;
				}

				// connector Arrow type
				sTmp = string.Empty;
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.ToArrowType);
				if (data != null)
				{
					sTmp = data.ToString().Trim();
				}
				switch (sTmp)
				{
					case VisioVariables.sARROW_START:
						visioInfo.ToArrowType = VisioVariables.sARROW_START;
						break;
					case VisioVariables.sARROW_END:
						visioInfo.ToArrowType = VisioVariables.sARROW_END;
						break;
					case VisioVariables.sARROW_BOTH:
						visioInfo.ToArrowType = VisioVariables.sARROW_BOTH;
						break;
					default:
						visioInfo.ToArrowType = VisioVariables.sARROW_NONE;
						break;
				}

				// connector line color
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.ToLineColor);
				visioInfo.ToLineColor = "";
				if (data != null)
				{
					if (!string.IsNullOrEmpty(visioHelper.GetRGBColor(data.ToString())))
					{
						// data color was found
						visioInfo.ToLineColor = data.ToString().Trim();
					}
				}

				// To connector Line weight
				sTmp = string.Empty;
				data = myArray.GetValue(row, (int)ExcelVariables.CellIndex.ToLineWeight);
				if (data != null)
				{
					sTmp = data.ToString().Trim();
					sTmp = visioHelper.FindConnectorLineWeight(sTmp);
					if (!string.IsNullOrEmpty(sTmp))
					{
						visioInfo.ToLineWeight = sTmp;
					}
				}

				// save it
				device.ShapeInfo = visioInfo;
			}
			catch (Exception exp)
			{
				ConsoleOut.writeLine(exp.Message + " - " + exp.StackTrace);
				return null;
			}
			//ConsoleOut.writeLine("adding stencil:{0}",visioInfo.UniqueKey);
			return device;
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
				string sTmp = string.Format("ProcessExcelDataFile::releaseObject - Exception\n\nUnable to release the Object:{0}\n{1}",ex.Message, ex.StackTrace);
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
				foreach (Process procOld in excelProcsOld)
				{
					if (procNew.Id == procOld.Id)
					{
						exist++;
					}
				}
				if (exist == 0)
				{
					procNew.Kill();
				}
			}
		}

		private bool openExcelFile(string file)
		{
			Excel.Application excelApp = null;
			Excel.Workbooks wkbks = null;
			Excel.Workbook wkbk = null;

			bool wasFoundRunning = false;

			Excel.Application tApp = null;
			//Checks to see if excel is opened
			try
			{
				tApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
				if (tApp.Caption.Contains("Architect"))
				{
					wasFoundRunning = true;
				}
			}
			catch (Exception)//Excel not open
			{
				wasFoundRunning = false;
			}
			finally
			{
				if (true == wasFoundRunning)
				{
					excelApp = tApp;
					wkbk = excelApp.Workbooks.Add(Type.Missing);
				}
				else
				{
					excelApp = new Excel.Application();
					wkbks = excelApp.Workbooks;
					wkbk = wkbks.Add(Type.Missing);
				}
				//Release the temp if in use
				if (null != tApp)
				{
					Marshal.FinalReleaseComObject(tApp);
				}
				tApp = null;
			}
			//Initialize the sheets in the new workbook
			return wasFoundRunning;
		}

	}
}
