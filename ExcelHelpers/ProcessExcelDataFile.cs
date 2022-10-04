using OmnicellBlueprintingTool.Models;
using OmnicellBlueprintingTool.Visio;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab


namespace OmnicellBlueprintingTool.ExcelHelpers
{
	public class ProcessExcelDataFile
	{
		private enum _cellIndex
		{
			// NOTE ****
			// the order of this enum must match the column order in the Excel file
			VisioPage = 1,       // Page indicator to place this shape
			ShapeType,           // key
			StencilKey,          // device unique Key used for connecting visio shapes
			StencilImage,        // device visio image name
			StencilLabel,        // device label
			StencilFontSize,     // default is what the stencil font size is   (use 12:B for 12 pt. Bold or 12 for 12 pt)
			Mach_name,           // device machine Name
			Mach_id,             // device machine Id
			Site_id,             // device site Id
			Site_name,           // deivce site name
			Site_address,        // device site address
			Omnis_name,          // device name
			Omnis_id,            // device Id
			SiteIdOmniId,        // site_id+omni_id
			IP,                  // device IP address
			Ports,               // device Ports
			DevicesCount,        // number of Devices for this type (part of a group)

			PosX,                // Shape position X
			PosY,                // shape position Y
			Width,               // shape width
			Height,              // shape height
			FillColor,           // color to fill stincel
			ConnectFrom,         // used to link this visio shape to another visio shape
			FromLineLabel,       // Arrow Text
			FromLinePattern,     // Line pattern solid = 1
			FromArrowType,       // Can contain one of these [None, Start, End, Both]
			FromLineColor,       // Arrow Color
			ConnectTo,           // used to link this visio shape to another visio shape
			ToLineLabel,         // Arrow Text
			ToLinePattern,       // Line pattern solid = 1
			ToArrowType,         // Can contain one of these [None, Start, End, Both]
			ToLineColor,         // Arrow Color
		}

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
		public DiagramData parseExcelFile(string file, DiagramData diagData)
		{
			object misValue = System.Reflection.Missing.Value; 
			
			if (string.IsNullOrEmpty(file))
			{
				// Error file is empty
				MessageBox.Show(string.Format("Exception:parseExcelFile(File is missing: {0})", nameof(file)));
				return null;
			}

			List<Device> devices = new List<Device>();
			Device device = null;
			diagData.visioStencilFilePaths = new List<string>();

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
					// we only need to get the first few columns to determine what to do
					var data = myArray.GetValue(row, (int)_cellIndex.VisioPage);
					if (data == null) // value is null skip this column
					{
						continue;   // should not happen
					}
					if (data.ToString().StartsWith(";"))   // first row is a header so skip
					{
						continue;
					}
					if ((Convert.ToInt32(data) > diagData.MaxVisioPages))
					{
						diagData.MaxVisioPages = Convert.ToInt32(data);
					}

					data = myArray.GetValue(row, (int)_cellIndex.ShapeType);
					if (data != null)
					{
						switch ((string)data.ToString().Trim().ToUpper())
						{
							case "TEMPLATE":           // Open a template.  This may be used with existing stencils already in the document
								data = myArray.GetValue(row, (int)_cellIndex.StencilKey);
								if (data != null)
								{
									diagData.visioTemplateFilePath = (string)data.ToString().Trim();
								}
								break;

							case "BLANK DOCUMENT":     // create a new blank Visio document.  No existing stencils attached.  Not using a Template
								diagData.visioTemplateFilePath = "";
								break;

							case "STENCIL":            // stencils to add
								data = myArray.GetValue(row, (int)_cellIndex.StencilKey);
								string stencilFile = data.ToString();// but we only want the first part of the key
								if (!string.IsNullOrEmpty(stencilFile))
								{
									diagData.visioStencilFilePaths.Add(stencilFile);
								}
								break;

							case "SHAPE":             // stencils to create on the document.  pass myArray object, row # and column count
								device = _parseExcelData(myArray, row);
								if (device != null)
								{
									devices.Add(device);
									diagData.AllShapesMap.Add(device.ShapeInfo.UniqueKey, device);
								}
								break;

							default:
								// the finally will be called to clean / close up everything
								MessageBox.Show(String.Format("ERROR::parseExcelFile\n\nInvalid value for the field 'ShapeType'\nFound:({0}) at Row:{1}\n\nPlease resolve this issue in the Excel Data file", data, row));
								return null;
						}
					}
					//ConsoleOut.writeLine(string.Format("parseExcelFile - ShapeType:{0}, Row{1}", cells[(int)_cellIndex.StencilImage].ToString().Trim(), row));
				}
			}
			catch (Exception ex)
			{
				ConsoleOut.writeLine(ex.Message + " - " + ex.StackTrace);
				MessageBox.Show(String.Format("Exception::parseExcelFile - Duplicate key:({0}) found.\nPlease resolve this issue in the Excel Data file\n{1}", device.ShapeInfo.UniqueKey, ex.Message)); //, ex.StackTrace.ToString);
				return null;
			}
			finally
			{
				if (diagData != null)
				{
					diagData.Devices = devices;
				}

				//quit and release
				xlWorkbook.Close(true, misValue, misValue);
				xlApp.Quit();

				//release com objects to fully kill excel process from running in the background
				releaseObject(xlApp);
				releaseObject(xlWorkbook);
				releaseObject(xlWorksheet);
				releaseObject(xlRange);

				killExcelProcesses(excelProcsOld);
			}
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
		private Device _parseExcelData(System.Array myArray, int row)
		{
			Device device = new Device();
			ShapeInformation visioInfo = new ShapeInformation();
			try
			{
				var data = myArray.GetValue(row, (int)_cellIndex.VisioPage);
				if (data != null)
				{
					visioInfo.VisioPage = Convert.ToInt32(data);
				}

				data = myArray.GetValue(row, (int)_cellIndex.StencilKey);
				if (data != null)
				{
					visioInfo.UniqueKey = data.ToString().Trim();   // unique key for this shape
				}

				data = myArray.GetValue(row, (int)_cellIndex.StencilImage);
				if (data != null)
				{
					visioInfo.StencilImage = data.ToString().Trim(); // must match exactly the name in the visio stencil
				}

				data = myArray.GetValue(row, (int)_cellIndex.StencilLabel);
				if (data != null)
				{
					visioInfo.StencilLabel = data.ToString().Trim(); // text to add to the stencil image
				}

				data = myArray.GetValue(row, (int)_cellIndex.StencilFontSize);
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
								visioInfo.LineWeight = VisioVariables.LINE_WEIGHT_2;
							}
						}
					}
				}

				data = myArray.GetValue(row, (int)_cellIndex.Mach_name);
				if (data != null)
				{
					device.MachineName = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)_cellIndex.Mach_id);
				if (data != null)
				{
					device.MachineId = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)_cellIndex.Site_id);
				if (data != null)
				{
					device.SiteId = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)_cellIndex.Site_name);
				if (data != null)
				{
					device.SiteName = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)_cellIndex.Site_address);
				if (data != null)
				{
					device.SiteAddress = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)_cellIndex.Omnis_name);
				if (data != null)
				{
					device.OmniName = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)_cellIndex.Omnis_id);
				if (data != null)
				{
					device.OmniId = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)_cellIndex.SiteIdOmniId);
				if (data != null)
				{
					device.SiteId_OmniId = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)_cellIndex.IP);
				if (data != null)
				{
					device.OmniIP = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)_cellIndex.Ports);
				if (data != null)
				{
					device.OmniPorts = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)_cellIndex.PosX);
				if (data != null)
				{
					visioInfo.Pos_x = Convert.ToDouble(data);
				}

				data = myArray.GetValue(row, (int)_cellIndex.PosY);
				if (data != null)
				{
					visioInfo.Pos_y = Convert.ToDouble(data);
				}

				data = myArray.GetValue(row, (int)_cellIndex.Width);
				if (data != null)
				{
					visioInfo.Width = Convert.ToDouble(data);
				}

				data = myArray.GetValue(row, (int)_cellIndex.Height);
				if (data != null)
				{
					visioInfo.Height = Convert.ToDouble(data);
				}

				data = myArray.GetValue(row, (int)_cellIndex.DevicesCount);
				if (data != null)
				{
					visioInfo.StencilLabel += " / " + data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)_cellIndex.FillColor);
				if (data != null)
				{
					// should be a string like
					visioInfo.FillColor = data.ToString().Trim();
				}

				// Get the ShpFromObj section
				data = myArray.GetValue(row, (int)_cellIndex.ConnectFrom);
				if (data != null)
				{
					visioInfo.ConnectFrom = data.ToString().Trim();
				}

				data = myArray.GetValue(row, (int)_cellIndex.FromLineLabel);
				if (data != null)
				{
					visioInfo.FromLineLabel = data.ToString().Trim();
				}

				// Arrow type to use if enabled
				string sTmp = string.Empty;
				data = myArray.GetValue(row, (int)_cellIndex.FromLinePattern);
				if (data != null)
				{
					sTmp = data.ToString().Trim().ToUpper();
				}
				switch (sTmp)
				{
					case "SOLID":
						visioInfo.FromLinePattern = (double)VisioVariables.LINE_PATTERN_SOLID;
						break;

					case "DASH":
						visioInfo.FromLinePattern = VisioVariables.LINE_PATTERN_DASH;
						break;

					case "DOTTED":
						visioInfo.FromLinePattern = VisioVariables.LINE_PATTERN_DOTTED;
						break;

					case "DASH_DOT":
						visioInfo.FromLinePattern = VisioVariables.LINE_PATTERN_DASHDOT;
						break;

					default:
						visioInfo.FromLinePattern = VisioVariables.LINE_PATTERN_SOLID;
						break;
				}

				// set the ShpFromObj ArrowType
				sTmp = string.Empty;
				data = myArray.GetValue(row, (int)_cellIndex.FromArrowType);
				if (data != null)
				{
					sTmp = data.ToString().Trim().ToUpper();
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

				data = myArray.GetValue(row, (int)_cellIndex.FromLineColor);
				if (data != null)
				{
					visioInfo.FromLineColor = data.ToString().Trim();
				}
				if (string.IsNullOrEmpty(visioInfo.FromLineColor))
				{
					visioInfo.FromLineColor = VisioVariables.COLOR_BLACK;
				}

				// Get the To section
				data = myArray.GetValue(row, (int)_cellIndex.ConnectTo);
				if (data != null)
				{
					if (!string.IsNullOrEmpty(data.ToString().Trim()))    // unique key for the To shape identifier - will match another shape object field 2 or empty for no connection
					{
						visioInfo.ConnectTo = data.ToString().Trim();
					}
				}

				data = myArray.GetValue(row, (int)_cellIndex.ToLineLabel);
				if (data != null)
				{
					visioInfo.ToLineLabel = data.ToString().Trim();
				}

				// Arrow type to use if enabled
				sTmp = string.Empty;
				data = myArray.GetValue(row, (int)_cellIndex.ToLinePattern);
				if (data != null)
				{
					sTmp = data.ToString().Trim().ToUpper();
				}
				switch (sTmp)
				{
					case "SOLID":
						visioInfo.ToLinePattern = (double)VisioVariables.LINE_PATTERN_SOLID;
						break;

					case "DASH":
						visioInfo.ToLinePattern = VisioVariables.LINE_PATTERN_DASH;
						break;

					case "DOTTED":
						visioInfo.ToLinePattern = VisioVariables.LINE_PATTERN_DOTTED;
						break;

					case "DASH_DOT":
						visioInfo.ToLinePattern = VisioVariables.LINE_PATTERN_DASHDOT;
						break;

					default:
						visioInfo.ToLinePattern = VisioVariables.LINE_PATTERN_SOLID;
						break;
				}

				// do we want to have a start arrow
				sTmp = string.Empty;
				data = myArray.GetValue(row, (int)_cellIndex.ToArrowType);
				if (data != null)
				{
					sTmp = data.ToString().Trim().ToUpper();
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

				data = myArray.GetValue(row, (int)_cellIndex.ToLineColor);
				if (data != null)
				{
					visioInfo.ToLineColor = data.ToString().Trim();
				}
				if (string.IsNullOrEmpty(visioInfo.ToLineColor))
				{
					visioInfo.ToLineColor = VisioVariables.COLOR_BLACK;
				}
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
				MessageBox.Show("Unable to release the Object " + ex.ToString());
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
