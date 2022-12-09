using Microsoft.Office.Interop.Visio;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using OmnicellBlueprintingTool.Models;
using Visio1 = Microsoft.Office.Interop.Visio;
using System.Xml.Linq;
using static OmnicellBlueprintingTool.Visio.VisioVariables;
using OmnicellBlueprintingTool.Properties;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrayNotify;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Color = System.Drawing.Color;
using System.Drawing.Imaging;
using System.Linq;

namespace OmnicellBlueprintingTool.Visio
{
	public class VisioHelper
	{
		public Visio1.Application appVisio = null;
		public Visio1.Documents vDocuments = null;
		public Visio1.Document vDocument = null;

		List<Visio1.Document> stencilsList = new List<Visio1.Document>();

		//** *********************************************************************************************************** **//
		// this section from json file for excel table entries for Visio
		private static StringComparer comparer = StringComparer.OrdinalIgnoreCase;
		private static List<string> _shapeTypesList = null;
		private static List<string> _connectorArrowsList = null;
		private static List<string> _connectorLinePatternsList = null;
		private static List<string> _stencilLabelPositionsList = null;
		private static List<string> _stencilLabelFontSizesList = null;
		private static List<string> _connectorLineWeightsList = null;
		private static List<string> _defaultStencilNames = null;
		private static List<string> _visioPageNamesList = null;

		private static Dictionary<string, string> _visioColorsMap = null; // new Dictionary<string, string>(comparer); 

		Dictionary<string, Color> _appColorsMap = null;

		//** *********************************************************************************************************** **//

		public VisioHelper()
		{
		}

		/// <summary>
		/// setAutoSizeDiagram
		/// set the AutoSize parameter for each page in the Document
		/// </summary>
		/// <param name="bMode"><option>true - Autosize diagram</option><option>false - Dont autosize diagram (default)</option></param>
		private void setAutoSizeDiagram(bool bMode = false)
		{
			if (bMode)
			{
				Visio1.Pages pagesObj = appVisio.ActiveDocument.Pages;
				foreach (Visio1.Page page in pagesObj)
				{
					// set the AutoSize and AutoSizeDrawing for each page in the Document
					page.AutoSize = true;
					page.AutoSizeDrawing();
				}
			}
		}

		/// <summary>
		/// SetupPage
		/// Set the Visio diagram page Orientation and page size
		/// </summary>
		/// <param name="currentPage">Current visio page</param>
		/// <param name="orientation"><options>"Portrait" or "Landscape"</options></param>
		/// <param name="size"><options>"Letter", "Tabloid", "Ledger", "Legal", "A3", "A4"</options></param>
		/// <return>bool<options>true error or false success</options></return>
		private bool setupDiagramPage(Visio1.Page currentPage, VisioPageOrientation orientation, VisioPageSize size)
		{
			Visio1.Shape sheet = currentPage.PageSheet;
			string width = string.Empty;
			string height = string.Empty;

			if (currentPage == null)
			{
				string sTmp = string.Format("VisioHelper::setupDiagramPage - Error\n\nOne of the following is null or empty: Page{0}, Orientation:{1}, Size:{3}", currentPage, orientation, size);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return true;
			}

			switch (size)
			{
				case VisioVariables.VisioPageSize.Tabloid:
					width = "11 in";
					height = "17 in";
					break;
				case VisioVariables.VisioPageSize.Ledger:
					width = "17 in";
					height = "11 in";
					break;
				case VisioVariables.VisioPageSize.Legal:
					width = "8.5 in";
					height = "14 in";
					break;
				case VisioVariables.VisioPageSize.A3:
					width = "11.69 in";
					height = "16.54 in";
					break;
				case VisioVariables.VisioPageSize.A4:
					width = "8.27 in";
					height = "11.60 in";
					break;
				case VisioVariables.VisioPageSize.Letter:
				default:
					width = "8.5 in";
					height = "11 in";
					break;
			}

			switch (orientation)
			{
				case VisioVariables.VisioPageOrientation.Landscape:
					currentPage.PageSheet.Cells["PageWidth"].FormulaU = height;
					currentPage.PageSheet.Cells["PageHeight"].FormulaU = width;
					currentPage.PageSheet.Cells["PrintPageOrientation"].FormulaU = "2";

					//currentPage.PageSheet.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionObject, (short)Visio1.VisRowIndices.visRowPrintProperties, (short)Visio1.VisCellIndices.visPageWidth).FormulaU = height;
					//currentPage.PageSheet.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionObject, (short)Visio1.VisRowIndices.visRowPrintProperties, (short)Visio1.VisCellIndices.visPageHeight).FormulaU = width;
					//currentPage.PageSheet.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionObject, (short)Visio1.VisRowIndices.visRowPrintProperties, (short)Visio1.VisCellIndices.visPageDrawSizeType).FormulaU = "3";
					break;

				case VisioVariables.VisioPageOrientation.Portrait:
				default:
					currentPage.PageSheet.Cells["PageWidth"].FormulaU = width;
					currentPage.PageSheet.Cells["PageHeight"].FormulaU = height;
					currentPage.PageSheet.Cells["PrintPageOrientation"].FormulaU = "1";

					//sheet.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionObject, (short)Visio1.VisRowIndices.visRowPrintProperties, (short)Visio1.VisCellIndices.visPageWidth).FormulaU = height;
					//sheet.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionObject, (short)Visio1.VisRowIndices.visRowPrintProperties, (short)Visio1.VisCellIndices.visPageHeight).FormulaU = width;
					//sheet.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionObject, (short)Visio1.VisRowIndices.visRowPrintProperties, (short)Visio1.VisCellIndices.visPageDrawSizeType).FormulaU = "1";
					break;
			}
			return false;  // successful
		}

		/// <summary>
		/// setupVisioDiagram
		/// this will setup the visio document page size and orientation
		/// also based on tghe dspMode the document is noShow or Show
		/// </summary>
		/// <param name="allData">DiagramData</param>
		/// <param name="dspMode">enum VisioVariables.ShowDiagram</param>
		/// <returns>Visio.Pages</returns>
		private Visio1.Pages setupVisioDiagram(DiagramData diagramData, VisioVariables.ShowDiagram dspMode)
		{
			string sErr = string.Empty;

			// Start Visio
			this.appVisio = new Visio1.Application();
			this.ShowVisioDiagram(appVisio, dspMode);             // don't show the diagram

			this.vDocuments = appVisio.Documents;
			try
			{
				if (!string.IsNullOrEmpty(diagramData.VisioTemplateFilePath))
				{
					// we need to check if the file is a template file or not
					// this will open a template file
					// Create a new document. but you will need to add a master stencil
					this.vDocument = appVisio.Documents.Add(diagramData.VisioTemplateFilePath);
				}
				else
				{
					// create a new blank Visio document without a Template file
					this.vDocument = appVisio.Documents.Add("");
				}
			}
			catch (Exception ex1)
			{
				string sTmp = string.Format("VisioHelper::setupVisioDiagram - Exception\n\nwith the Template file\n{0}\n{1}",ex1.Message, ex1.StackTrace);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return null;
			}
			string sStencil = string.Empty;
			try
			{
				// this gives a count of all the stencilsList on the status bar
				int countStencils = vDocument.Masters.Count;

				// get the current draw page
				Visio1.Page currentPage = vDocument.Pages[1];

				// lets add stencilsList to the template if they don't alredy exist using the Excel Data File
				foreach (var stencil in diagramData.VisioStencilFilePaths)
				{
					if (this.vDocuments != null)  // do we have any stencilsList attached to this template?
					{
						var vPage = vDocument.Application.ActivePage;

						// Load the stencil we want
						sStencil = stencil.ToString();
						Visio1.Document nStencil = vDocuments.OpenEx(stencil, (short)Visio1.VisOpenSaveArgs.visOpenDocked);
						stencilsList.Add(nStencil);
					}
					else
					{
						sErr = "Error with vDocument being null.  This should not happen";
						throw new Exception(string.Format("ERROR::setupVisioDiagram - {0}", sErr));
					}
				}
			}
			catch (Exception ex2)
			{
				string sTmp = string.Format("VisioHelper::setupVisioDiagram - Exception\n\nloading Stencil file:({0})\nMost likely the wrong stencil file name or path location\n{1}",sStencil, ex2.Message);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return null;
			}

			Visio1.Pages pagesObj = appVisio.ActiveDocument.Pages;

			// The new document will have one page, get the a reference to it.
			Visio1.Page page1 = vDocument.Pages[1];

			// check if we have pages in the configuration
			List<string> visioPageNames = this.GetVisioPageNames();
			if (visioPageNames.Count > 0)
			{
				diagramData.MaxVisioPages = visioPageNames.Count;
				page1.Name= visioPageNames[0];	// Rename the first page
			}

			//Assuming 'No theme' is set for the page, no arrow will be shown so change theme to see connector arrow
			page1.SetTheme("Office Theme");

			// Page 1 is Standard
			if (!setupDiagramPage(page1, diagramData.VisioPageOrientation, diagramData.VisioPageSize))
			{
				double xPosition = page1.PageSheet.get_CellsU("PageWidth").ResultIU;
				double yPosition = page1.PageSheet.get_CellsU("PageHeight").ResultIU;
				var pageOrientation = page1.PageSheet.get_CellsU("PrintPageOrientation").ResultIU;
				ConsoleOut.writeLine(string.Format("Adding Visio page:'{0}', Height:{1}, Width:{2} and Orientation:'{3}'", page1.Name, yPosition, xPosition, diagramData.VisioPageOrientation));
			}

			int cnt = this.vDocuments.Count;

			// this section is if we want to add more than one page to the document
			// At this point a page has already been created above so you need to adjust the count by 1
			Visio1.Page page = null;        
			for (int i = 0; i < diagramData.MaxVisioPages - 1; i++)
			{
				page = vDocument.Pages.Add();
				if (visioPageNames.Count > 0)
				{
					// use the names from the list
					page.Name = visioPageNames[i + 1];
				}

				//Assuming 'No theme' is set for the page, no arrow will be shown so change theme to see connector arrow
				page.SetTheme("Office Theme");

				// Page 1 is Standard
				if (!setupDiagramPage(page, diagramData.VisioPageOrientation, diagramData.VisioPageSize))
				{
					double xPosition = page.PageSheet.get_CellsU("PageWidth").ResultIU;
					double yPosition = page.PageSheet.get_CellsU("PageHeight").ResultIU;
					var pageOrientation = page.PageSheet.get_CellsU("PrintPageOrientation").ResultIU;
					ConsoleOut.writeLine(string.Format("Adding Visio page:'{0}', Height:{1}, Width:{2} and Orientation:'{3}'", page.Name, yPosition, xPosition, diagramData.VisioPageOrientation));
				}
			}

			// set the active page to the first page
			SetActivePage(1);

			return pagesObj;
		}

		/// <summary>
 		/// SetPage1Active
		/// set the active page
		/// pageNumber may not be less than 1 and can't be greater than the max number of pages
		/// </summary>
		/// <param name="pageNumber">Range is 1-max page</param>
		public void SetActivePage(int pageNumber)
		{
			Visio1.Pages pagesObj = appVisio.ActiveDocument.Pages;
			if (pageNumber < 1 || pageNumber > pagesObj.Count)
			{
				// default to first page
				this.appVisio.ActiveWindow.Page = pagesObj[1];
				ConsoleOut.writeLine(string.Format("Setting the Active Page to Default page:'{0}'", pagesObj[1].Name));
			}
			else 
			{
				// the visio document should contain all the pages at this point
				// set the active page
				this.appVisio.ActiveWindow.Page = pagesObj[pageNumber];
				ConsoleOut.writeLine(string.Format("Setting the Active Page to:'{0}'", pagesObj[pageNumber].Name));
			}
		}

		/// <summary>
		/// SetPage1Active
		/// set the active page
		/// pageNumber may not be less than 1 and can't be greater than the max number of pages
		/// </summary>
		/// <param name="pageName">Should be the name of the page or "1"</param>
		public void SetActivePage(string pageName)
		{
			Visio1.Pages pagesObj = appVisio.ActiveDocument.Pages;
			if (string.IsNullOrEmpty(pageName))
			{
				SetActivePage(1); // default to first page
				return;
			}
			else
			{
				// remember index to Visio pages does not use zero index must start with 1  (Visual Basic thing i guess)
				for (int nIdx = 1;  nIdx <= pagesObj.Count; nIdx++)
				{
					var name = pagesObj[nIdx].Name.Trim();
					if (name.ToString().ToUpper().Trim().Equals(pageName.ToUpper().Trim()))
					{
						SetActivePage(nIdx); // set the active page
						return;
					}
				}
				SetActivePage(1); // if all others failed set to the first page
			}
		}


		/// <summary>
		/// GetNumberOfVisioPages
		/// return the number of pages in the document
		/// </summary>
		/// <returns></returns>
		public int GetNumberOfVisioPages()
		{
			return appVisio.ActiveDocument.Pages.Count;
		}

		/// <summary>
		/// _drawShape
		/// draw the visio stencil on the visio document
		/// update the dictionaries in the DiagramData object for connection info
		/// </summary>
		/// <param name="device">Device</param>
		/// <param name="vPages">Visio.Pages</param>
		/// <returns>Visio.Shape</returns>
		/// <exception cref="Exception"></exception>		
		private Visio1.Shape _drawShape(Device device, Visio1.Pages vPages)
		{
			Visio1.Shape shpObj = null;
			Visio1.Master stnObj = null;

			// lets look if the stencil is part of the Document Stencil master
			// this is the issue.  the count will go up with each stincel added to the diagram so this logic is messed up

			try
			{
				if (vDocument.Masters.get_ItemU(device.ShapeInfo.StencilImage) != null)
				{
					// it part of the Document Stencil
					stnObj = vDocument.Masters[device.ShapeInfo.StencilImage];
				}
			}
			catch (System.Runtime.InteropServices.COMException com)
			{
				// if we get this exception the item was not found
				// stencil not found here so lets try looking if any other stencil files have been added
				// Fall through and continue
			}

			try
			{
				if (stnObj == null)
				{
					// else look to see if the Stencil is part of the added stincel files
					if (stencilsList.Count > 0)
					{
						foreach (Visio1.Document stencil in stencilsList)
						{
							try
							{
								stnObj = stencil.Masters[device.ShapeInfo.StencilImage];
							}
							catch (System.Runtime.InteropServices.COMException com)
							{
								// if we get this exception the item was not found so lets try searching the next stencil
								//Console.WriteLine(string.Format("failed to locate this Stencil:{0} for this stencil file:{1}", device.ShapeInfo.StencilImage, stencil.Template));
								continue;
							}
							if (stnObj != null)
							{
								break;   // found get out of the loop
							}
						}
					}
				}
				if (stnObj == null)
				{
					string sTmp = string.Format("VisioHelper::_drawShape  Error\n\nCan't find Stencil:{0}", device.ShapeInfo.StencilImage);
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine(sTmp);
					return null;
				}

				Visio1.Pages pagesObj = this.appVisio.ActiveDocument.Pages;

				SetActivePage(device.ShapeInfo.VisioPage);

				// draw the shape on the document
				shpObj = this.appVisio.ActivePage.Drop(stnObj, device.ShapeInfo.Pos_x, device.ShapeInfo.Pos_y);
				shpObj.NameU = device.ShapeInfo.UniqueKey;
				shpObj.Name = device.ShapeInfo.StencilImage;

				#region keep
				// keep this section because it provides an offset when resizing shapes
				// normal stencilsList are normal (east-width and south-height)
				double WidthAdjustment = 0.0; // Math.Truncate(shpObj.get_Cells("Width").ResultIU * 1000) / 1000;
				double HeightAdjustment = 0.0; // Math.Truncate(shpObj.get_Cells("Height").ResultIU * 1000) / 1000;
				if (device.ShapeInfo.StencilImage.IndexOf("Group", StringComparison.OrdinalIgnoreCase) >= 0 ||
					 device.ShapeInfo.StencilImage.IndexOf("Dash", StringComparison.OrdinalIgnoreCase) >= 0)
				{
					WidthAdjustment = 0.750;   // shape stencil size
					HeightAdjustment = 0.268;  // shape stencil size
				}
				#endregion
				if (device.ShapeInfo.Width > 0.0)
				{
					// we need to stretch East
					shpObj.Resize(VisResizeDirection.visResizeDirE, device.ShapeInfo.Width - WidthAdjustment, VisUnitCodes.visDrawingUnits);
				}
				if (device.ShapeInfo.Height > 0.0)
				{
					// we need to stretch south
					shpObj.Resize(VisResizeDirection.visResizeDirS, device.ShapeInfo.Height - HeightAdjustment, VisUnitCodes.visDrawingUnits);
				}

				//var linePatternCell = shpConn.get_CellsU("LinePattern");

				// Look at the Fill color name to use.  if empty check the rgbFillColor field
				string rgb = GetRGBColor(device.ShapeInfo.FillColor);
				if (string.IsNullOrEmpty(rgb))
				{
					// FillColor is empty now check if we should use the rgbFillColor
					if (!string.IsNullOrEmpty(device.ShapeInfo.rgbFillColor))
					{
						// yes use the rgbFillColor
						rgb = device.ShapeInfo.rgbFillColor;

						// we don't want to fill these objects
						// (Logo, Footer, Title)
						if (device.ShapeInfo.StencilImage.IndexOf("Logo") >= 0 ||
							device.ShapeInfo.StencilImage.IndexOf("Title") >= 0 ||
							device.ShapeInfo.StencilImage.IndexOf("Footer") >= 0 )
						{
							rgb = "";
						}
					}
				}
				if (!string.IsNullOrEmpty(rgb))
				{
					if (device.ShapeInfo.StencilImage.Trim().IndexOf("OC_DashOutline") >= 0)
					{
						// handle this shape differently.
						// Only the line color should be set.  No fill
						shpObj.get_Cells("LineColor").FormulaU = rgb;
					}
					else
					{
						// visFillForegnd is used with fill object is solid fill pattern 1
						shpObj.get_CellsSRC(
						(short)VisSectionIndices.visSectionObject,
						(short)VisRowIndices.visRowFill,
						(short)VisCellIndices.visFillForegnd).FormulaU = rgb;

						shpObj.get_CellsSRC(
							 (short)Visio1.VisSectionIndices.visSectionObject,
							 (short)Visio1.VisRowIndices.visRowFill,
							 (short)Visio1.VisCellIndices.visFillBkgnd).FormulaU = rgb;

						// for an shape to be filled this needs to be set
						shpObj.get_CellsSRC(
							 (short)Visio1.VisSectionIndices.visSectionObject,
							 (short)Visio1.VisRowIndices.visRowFill,
							 (short)Visio1.VisCellIndices.visFillPattern).FormulaU = "1";

						//shpObj.get_Cells("LineColor").FormulaForceU = rgb; // set the stencil outline color same as stencil fill color
						//shpObj.get_Cells("LineColor").FormulaU = rgb;		// set the stencil outline color same as stencil fill color

						// we normally want an outline on most of the shapes so set this to BLACK
						shpObj.get_Cells("LineColor").FormulaU = GetRGBColor(VisioVariables.sCOLOR_BLACK);
					}
				}

				// we want to keep the shape outline color Black for this Stencil
				if (device.ShapeInfo.UniqueKey.Trim().StartsWith("OC_TableCell"))
				{
					shpObj.get_Cells("LineColor").FormulaU = GetRGBColor(VisioVariables.sCOLOR_BLACK);
				}

				if (!string.IsNullOrEmpty(device.ShapeInfo.StencilLabel))
				{
					shpObj.Text = device.ShapeInfo.StencilLabel;
					if (!string.IsNullOrEmpty(device.ShapeInfo.StencilLabelFontSize))
					{
						if (device.ShapeInfo.isStencilLabelFontBold)
						{
							// bold font
							shpObj.get_CellsSRC(
								 (short)Visio1.VisSectionIndices.visSectionCharacter,
								 (short)0,
								 (short)Visio1.VisCellIndices.visCharacterStyle).FormulaU = "1";
							//shpObj.TextStyleKeepFmt = "Bold";    // Using this code would not allow the font size to be changed
						}
						//shpObj.get_Cells("Char.Size").Formula = "=" + device.ShapeInfo.StencilLabelFontSize + " pt";	// was
						shpObj.get_Cells("Char.Size").FormulaU = "=" + device.ShapeInfo.StencilLabelFontSize + " pt";   // changed to
																																						//shpObj.Cells("Char.Size").FormulaU = device.ShapeInfo.StencilLabelFontSize + " pt";
																																						//string fontSize = shpObj.get_Cells("Char.Size").Formula;
					}

					// check if we have an IP that needs to be displayed
					if (!string.IsNullOrEmpty(device.OmniIP))
					{
						shpObj.Text += "\n" + device.OmniIP;
					}
					if (!string.IsNullOrEmpty(device.OmniPorts))
					{
						shpObj.Text += ":" + device.OmniPorts;
					}
					int textLen = shpObj.Text.Length;

					// check if we need to move the text box to the bottom of the stencil
					if ((!string.IsNullOrEmpty(device.ShapeInfo.StencilLabelPosition) || (device.ShapeInfo.StencilLabelPosition.IndexOf(VisioVariables.STINCEL_LABEL_POSITION_BOTTOM)>= 0)) && textLen > 0)
					{
						short exists = shpObj.RowExists[(short)Visio1.VisSectionIndices.visSectionObject,
												 (short)Visio1.VisRowIndices.visRowTextXForm,
												 (short)Visio1.VisExistsFlags.visExistsAnywhere];
						if (exists == 0)
						{
							shpObj.AddRow((short)Visio1.VisSectionIndices.visSectionObject, (short)Visio1.VisRowIndices.visRowTextXForm, (short)Visio1.VisRowTags.visTagDefault);
						}

						//	Set the text transform formulas:
						shpObj.get_CellsU("TxtHeight").FormulaForceU = "Height*0";
						shpObj.get_CellsU("TxtPinY").FormulaForceU = "Height*0";

						//	Set the paragraph alignment formula:
						shpObj.get_CellsU("VerticalAlign").FormulaForceU = "0";
					}

					// dont resize the text for the Title and Footer stencilsList or if the Stencil label font size is set to a value
					if (string.IsNullOrEmpty(device.ShapeInfo.StencilLabelFontSize))
					{
						if (device.ShapeInfo.StencilImage.IndexOf("Title", StringComparison.OrdinalIgnoreCase) <= 0 &&
							device.ShapeInfo.StencilImage.IndexOf("Footer", StringComparison.OrdinalIgnoreCase) <= 0)
						{
							int nSize = 0;
							string[] saTmp = device.ShapeInfo.StencilLabel.Split('\n');
							foreach (string saTmpStr in saTmp)
							{
								if (saTmpStr.Length > nSize)
								{
									nSize = saTmpStr.Length;
								}
							}
							//	Set the text transform formulas:
							//var lHeight = Math.Truncate(shpObj.get_CellsU("TxtHeight").ResultIU * 1000) / 1000;
							//shpObj.get_CellsU("TxtPinY").FormulaForceU = "Height*0";
							double lWidth = Math.Truncate(shpObj.get_CellsU("TxtWidth").ResultIU * 1000) / 1000;
							double fSize = Math.Truncate(shpObj.get_CellsU("Char.Size").ResultIU * 1000) / 1000;
							double xx = Math.Truncate(((fSize * nSize) - lWidth) * 1000) / 1000;
							if (xx > lWidth)
							{
								//	Set the paragraph alignment formula:
								// shpObj.get_CellsU("VerticalAlign").FormulaForceU = "0";
								//scale = shpObj.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionUser, (short)Visio1.VisRowIndices.visRowUser, (short)Visio1.VisCellIndices.visUserValue).ResultIU;
								double scale = 0.5;
								// Then set the font, and the TextMargins(for any that are non - zero) with the following(assuming the normal font size is 12 and the left margin is 4pt.:
								shpObj.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionCharacter, 0, (short)Visio1.VisCellIndices.visCharacterSize).FormulaU = (scale * 12).ToString() + "pt";
								shpObj.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionObject, (short)Visio1.VisRowIndices.visRowText, (short)Visio1.VisCellIndices.visTxtBlkLeftMargin).FormulaU = (scale * 2).ToString() + "pt";
							}
						}
					}
				}
			}
			catch(Exception ep)
			{
				string sTmp = string.Format("VisioHelper::_drawShape.  Drawing:{0}-{1}\n\n{3}", device.ShapeInfo.StencilImage, device.ShapeInfo.UniqueKey, ep.Message);
				MessageBox.Show(sTmp, "EXCEPTION", MessageBoxButtons.OK, MessageBoxIcon.Error);
				ConsoleOut.writeLine(sTmp);
				return null;
			}
			ConsoleOut.writeLine(String.Format("VisioHelper::_drawShape.  Drawing:{0} - {1}",device.ShapeInfo.StencilImage, device.ShapeInfo.UniqueKey));
			return shpObj;
		}

		/// <summary>
		/// ShowVisioDiagram
		/// This method can control if the Visio document is visiable or not.
		/// One thing I have noticed is if the visio document is hidden you will need to
		/// open Visio to open the hidden file to save it or discard it.
		/// </summary>
		/// <param name="appV">Visio1.Application</param>
		/// <param name="show"><options>VisioVariables.ShowDiagram.NoShow or VisioVariables.ShowDiagram.Show</options></param>
		public void ShowVisioDiagram(Visio1.Application appV, VisioVariables.ShowDiagram show)
		{
			if (show == VisioVariables.ShowDiagram.Show)
			{
				appV.Visible = true; // show the diagram
			}
			else
			{
				appV.Visible = false;   // dont show the diagram
			}
		}

		/// <summary>
		/// ClearStencilList
		/// this will clear the stencilsList list for reuse
		/// </summary>
		public void ClearStencilList()
		{
			if (stencilsList != null)
			{
				// must clear this list otherwise an Exception will occur dealing with RPS miss leading error when app is ran again without closing
				stencilsList.Clear();
			}
		}

		/// <summary>
		/// SaveDiagram
		/// Save the Visio Diagram using the name passed from the fileNamePath argument
		/// if argument is null or empty Visio app will prompt the user for the file name and path
		/// if argument is valid use it as the path and file name to saveAs
		/// return true if error
		/// return false if successful
		/// </summary>
		/// <param name="fileNamePath"></param>
		/// <param name="bSave">
		/// false - User does not want the document to be saved.  However, we must set the saved flag to true
		///         Don't want to be prompted to save when closing Visio
		/// true - save document and set saved flag to true
		/// </param>
		/// <returns>true - error</returns>
		public bool SaveDocument(string fileNamePath, bool bSave)
		{
			try
			{
				if (string.IsNullOrEmpty(fileNamePath))
				{
					return true;
				}

				if (bSave)	// user requested this document to be saved
				{
					if (this.vDocuments != null)
					{
						if (!string.IsNullOrEmpty(fileNamePath))
						{
							this.appVisio.ActiveDocument.SaveAs(fileNamePath);
						}
					}
				}
				// set document has been saved
				// this should stop the application from asking the user to save the document again
				if (this.vDocument != null)
				{
					this.vDocument.Saved = true;
				}
			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				string sTmp = string.Format("VisioHelper::SaveDocument - Exception\n\nSaving Visio File:'{0}'\n\n{1}", fileNamePath, ex.Message);
				//MessageBox.Show(sTmp, "EXCEPTION", MessageBoxButtons.OK, MessageBoxIcon.Error);
				ConsoleOut.writeLine(sTmp);
			}
			return false;
		}

		/// <summary>
		/// VisioForceCloseAll
		/// Close all the Visio documents
		/// if no file name is present Visio will give the option to saveAs
		/// if a file name is present Visio will save using the path and file name
		/// User has the ability to not save, Cancel or Save
		/// </summary>
		/// <param bSave>bool</param>
		/// <values>Default is false (save) if true dont save</values>
		public void VisioForceCloseAll(bool bSave = false)
		{
			try
			{
				ClearStencilList();
				ClearVisioPageNamesList();

				if (this.vDocuments != null)
				{
					while (this.vDocuments.Count > 0)
					{
						// set document has been saved
						if (bSave)
						{
							this.vDocument.Saved = true;
						}
						
						this.vDocument.Close();
						//this.vDocuments.Application.ActiveDocument.Close();
					}
					this.vDocuments = null;
				}
				if (this.appVisio != null)
				{
					this.appVisio.Quit();
					this.appVisio = null;
				}

				GC.Collect();
				GC.WaitForPendingFinalizers();
				GC.Collect();
				GC.WaitForPendingFinalizers();
			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				string sTmp = string.Format("VisioHelper::VisioForceCloseAll - Exception\n\nUser closed the Visio document before exiting the application\n\n{0}", ex.Message);
				//MessageBox.Show(sTmp, "EXCEPTION", MessageBoxButtons.OK, MessageBoxIcon.Error);
				ConsoleOut.writeLine(sTmp);
			}
		}

		/// <summary>
		/// DrawAllShapes
		/// Draw all the visio stencilsList obtained from the data file
		/// also based on tghe dspMode the document is noShow or Show
		/// </summary>
		/// <param name="diagramData">DiagramData</param>
		/// <param name="dspMode">enum VisioVariables.ShowDiagram</param>
		/// <exception cref="Exception"></exception>
		public bool DrawAllShapes(DiagramData diagramData, VisioVariables.ShowDiagram dspMode)
		{
			Visio1.Pages vPages = setupVisioDiagram(diagramData, dspMode);
			if (vPages == null)
			{
				// error
				return true;
			}

			int count = vDocument.Masters.Count;

			Visio1.Shape shpObj = null;
			foreach (Device device in diagramData.Devices)
			{
				try
				{
					// draw other shapes
					// add list of shaps to ignore
					shpObj = _drawShape(device, vPages);
					if (shpObj == null)
					{
						// there was an error so lets abort
						return true;
					}
					device.ShapeInfo.ShpObj = shpObj;   // save this stencil object
				}
				catch (Exception ex)
				{
					string sTmp = string.Format("VisioHelper::DrawAllShapes - Exception\n\nStencil Image:{0} not found.\nShape Key:{1}\n{2}", device.ShapeInfo.StencilImage, device.ShapeInfo.UniqueKey, ex.Message);
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine(string.Format("Exception::DrawAllShapes - Stencil Image:{0} not found.  Shape Key:{1}\n{2}", device.ShapeInfo.StencilImage, device.ShapeInfo.UniqueKey, ex.Message));
					return true;
				}
			}

			// use before saving AutoSizeDrawing
			appVisio.AutoLayout = true;

			// using true will auto size the document.
			// sometimes this is not needed/wanted
			// can use AutoSize:false "Page Setup" in excel data file
			setAutoSizeDiagram(diagramData.AutoSizeVisioPages);

			return false;
		}

		/// <summary>
		/// ConnectShapes
		/// this function will connect all the shapes that are required.
		/// it will use the All shaps map to lookup the shape object for the connectFrom and connectTo values
		/// special handling is if the Shape object for the ShpFromObj and To are null don't draw a connection
		/// </summary>
		/// <param name="diagData">DiagramData</param>
		/// <returns>bool<values>false success</values></returns>
		public bool ConnectShapes(DiagramData diagData)
		{
			ShapeConnection lookupShapeConnection = null;
			Visio1.Shape shpConn = null;
			int nCnt = 0;
			try
			{
				// iterate over the ShapeConnectionsMap to determine if a connection shape is needed
				for (nCnt = 0; nCnt < diagData.ShapeConnectionsMap.Count; nCnt++)
				{
					// nCnt is the key
					if (diagData.ShapeConnectionsMap.TryGetValue(nCnt, out lookupShapeConnection))
					{
						// Drop the built-in connector object on the lower left corner of the page:
						// need to drop on another page
						Visio1.Pages pagesObj = appVisio.ActiveDocument.Pages;

						SetActivePage(lookupShapeConnection.device.ShapeInfo.VisioPage);

						// draw the object on the Visio diagram
						shpConn = appVisio.ActivePage.Drop(pagesObj.Application.ConnectorToolDataObject, 0.0, 0.0);

						// Set the connector properties, using different arrows, colors, and patterns for many-to-many relationships.
						shpConn.get_CellsU("ShdwPattern").ResultIU = VisioVariables.SHDW_PATTERN;
						shpConn.get_CellsU("BeginArrow").ResultIU = VisioVariables.ARROW_NONE;
						shpConn.get_CellsU("EndArrow").ResultIU = VisioVariables.ARROW_NONE;
						shpConn.get_CellsU("LineColor").FormulaU = GetRGBColor(VisioVariables.sCOLOR_BLACK);
						shpConn.get_CellsU("Rounding").ResultIU = VisioVariables.ROUNDING;
						shpConn.get_CellsU("LinePattern").ResultIU = VisioVariables.LINE_PATTERN_SOLID;
						shpConn.get_CellsU("LineWeight").FormulaU = VisioVariables.sLINE_WEIGHT_1;

						//if (lookupShapeConnection.device.ShapeInfo.LineWeight != VisioVariables.sLINE_WEIGHT_1)
						//{
						//	shpConn.get_CellsU("LineWeight").FormulaU = lookupShapeConnection.device.ShapeInfo.LineWeight;
						//}
						if (lookupShapeConnection.LineWeight != VisioVariables.sLINE_WEIGHT_1)
						{
							shpConn.get_CellsU("LineWeight").FormulaU = lookupShapeConnection.LineWeight;
						}
						if (lookupShapeConnection.LinePattern > 0)
						{
							shpConn.get_CellsU("LinePattern").ResultIU = lookupShapeConnection.LinePattern;
						}

						switch ((string)lookupShapeConnection.ArrowType.Trim())
						{
							case VisioVariables.sARROW_START:
								shpConn.get_CellsU("BeginArrow").ResultIU = VisioVariables.BEGIN_ARROW;
								break;
							case VisioVariables.sARROW_END:
								shpConn.get_CellsU("EndArrow").ResultIU = VisioVariables.END_ARROW;
								break;
							case VisioVariables.sARROW_BOTH:
								shpConn.get_CellsU("BeginArrow").ResultIU = VisioVariables.BEGIN_ARROW;
								shpConn.get_CellsU("EndArrow").ResultIU = VisioVariables.END_ARROW;
								break;
							default:
								shpConn.get_CellsU("BeginArrow").ResultIU = VisioVariables.ARROW_NONE;
								shpConn.get_CellsU("EndArrow").ResultIU = VisioVariables.ARROW_NONE;
								break;
						}

						// set connection text
						if (!string.IsNullOrEmpty(lookupShapeConnection.LineLabel))
						{
							shpConn.Text = lookupShapeConnection.LineLabel;
						}

						//var linePatternCell = shpConn.get_CellsU("LinePattern");
						string rgbColor = string.Empty;
						rgbColor = GetRGBColor(lookupShapeConnection.LineColor.Trim().ToUpper());
						if (string.IsNullOrEmpty(rgbColor))
						{
							rgbColor = GetRGBColor(VisioVariables.sCOLOR_BLACK);
						}
						//shpConn.get_CellsU("LineColor").Formula = "="+rgbColor;	// was
						shpConn.get_CellsU("LineColor").FormulaU = "=" + rgbColor;	// changed to

						// default the connection Text color Black
						shpConn.get_CellsU("Char.Color").FormulaU = "=" + GetRGBColor(VisioVariables.sCOLOR_BLACK);

						//set the shape back color
						shpConn.get_CellsSRC((short)VisSectionIndices.visSectionObject,
					 							(short)VisRowIndices.visRowFill,
												(short)VisCellIndices.visFillForegnd).FormulaU = "="+rgbColor;

						
						//shpConn.get_CellsSRC((short)VisSectionIndices.visSectionObject,
					 	//						(short)VisRowIndices.visRowFill,
						//						(short)VisCellIndices.visFillBkgnd).FormulaU = "=" + rgbColor;
						//
						//shpConn.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionObject,
						//							(short)Visio1.VisRowIndices.visRowCharacter,
						//							(short)Visio1.VisCellIndices.visTxtBlkBkgndTrans).FormulaU = "1";
						////(short)Visio1.VisCellIndices.visTxtBlkBkgnd).FormulaU = "0";

						// now connect the connector to the objects
						if (lookupShapeConnection.ShpFromObj != null && lookupShapeConnection.ShpToObj != null)
						{
							//shpConn.AutoConnect(lookupShapeConnection.ShpFromObj, Visio1.VisAutoConnectDir.visAutoConnectDirNone);
							//shpConn.AutoConnect(lookupShapeConnection.ShpToObj, Visio1.VisAutoConnectDir.visAutoConnectDirNone);

							// Connect its Begin to the 'ShpFromObj' shape:
							shpConn.get_CellsU("BeginX").GlueTo(lookupShapeConnection.ShpFromObj.get_CellsU("PinX"));

							// Connect its End to the 'To' shape:
							shpConn.get_CellsU("EndX").GlueTo(lookupShapeConnection.ShpToObj.get_CellsU("PinX"));
						}
						else
						{
							ConsoleOut.writeLine(string.Format("SKIP drawing this connection from:{0} To:{1}", lookupShapeConnection.device.ShapeInfo.ConnectFrom, lookupShapeConnection.device.ShapeInfo.ConnectTo));
						}
					}
				}
			}
			catch (Exception ex)
			{
				if (ex.Message.ToString().IndexOf("Inappropriate target", StringComparison.OrdinalIgnoreCase) != -1)
				{
					string sTmp = "Exception occured\n\nPlease check you Excel data file column 'Visio Page' values.\nYou may have a page value in the wrong location\n\nYou can't have page number Intermingled";
					MessageBox.Show(sTmp, "EXCEPTION", MessageBoxButtons.OK, MessageBoxIcon.Error);
					ConsoleOut.writeLine(sTmp);
					return true;		// error
				}
				//throw new Exception(string.Format("VisioHelper::ConnectShapes - Exception\n\n{0}\n{1}", ex.Message, ex.StackTrace));
			}
			return false;	// success
		}

		/// <summary>
		/// ListStencils
		/// List all the stencilsList in the master stencil document
		/// </summary>
		/// <param name="diagData">DiagramData</param>
		/// <param name="dspMode">enum VisioVariables.ShowDiagram</param>
		/// <returns>true = error, fase = success</returns>
		public bool ListDocumentStencils(DiagramData diagramData, VisioVariables.ShowDiagram dspMode)
		{
			Visio1.Pages vPages = setupVisioDiagram(diagramData, dspMode);
			if (vPages == null)
			{
				// error
				return true;
			}

			// using true will auto size the document.
			// sometimes this is not needed/wanted
			// can use AutoSize:false "Page Setup" in excel data file
			//setAutoSizeDiagram(diagramData.AutoSizeVisioPages);

			Console.WriteLine("\n\n");

			ArrayList masterArray_0 = new ArrayList();
			Visio1.Document doc_0 = vDocuments[1];    // Document stencil figures
			Visio1.Document doc_1 = vDocuments[2];    // Stencil figures
			Visio1.Masters masters_0 = doc_0.Masters;
			Visio1.Masters masters_1 = doc_1.Masters;
			int nCnt = 1;
			foreach (Visio1.Master master in masters_0)
			{
				// Document stencil figures
				masterArray_0.Add(master.NameU);   // THIS WILL CONTAIN DIAGRAM FIGURES
				Console.WriteLine(string.Format("ListDocumentStencils - Master0 - {0} : ID:{1} Name:{2} NameU:{3}", nCnt++, master.ID, master.Name, master.NameU));
			}
			Console.WriteLine("\n\n");
			nCnt = 1;
			foreach (Visio1.Master master in masters_1)
			{
				// Document stencil figures
				masterArray_0.Add(master.NameU);   // THIS WILL CONTAIN DIAGRAM FIGURES
				Console.WriteLine(string.Format("ListDocumentStencils - Master1 - {0} : ID:{1} Name:{2} NameU:{3}", nCnt++, master.ID, master.Name, master.NameU));
			}
			Console.WriteLine("\n\n");
			if (this.vDocument != null)
			{
				this.vDocument.Saved = true;
			}
			this.VisioForceCloseAll();

			return false;
		}

		/// <summary>
		/// setShapeTextBottom
		/// Adjusts the text block of selected shapes so that
		/// the text is at the bottom of the shape. This matches
		/// the default text position for inserted images.
		/// </summary>
		private void setShapeTextBottom()
		{
			short exists = -1;
			Visio1.Selection sel = this.appVisio.ActiveWindow.Selection;
			foreach (Visio1.Shape shp in sel)
			{
				// '// 'Add' the Text Transfomrm section, if it's not there:
				exists = shp.RowExists[(short)Visio1.VisSectionIndices.visSectionObject,
										 (short)Visio1.VisRowIndices.visRowTextXForm,
										 (short)Visio1.VisExistsFlags.visExistsAnywhere];
				if (exists == 0)
				{
					shp.AddRow((short)Visio1.VisSectionIndices.visSectionObject,
						        (short)Visio1.VisRowIndices.visRowTextXForm,
								  (short)Visio1.VisRowTags.visTagDefault);
				}

				//	Set the text transform formulas:
				shp.get_CellsU("TxtHeight").FormulaForceU = "Height*0";
				shp.get_CellsU("TxtPinY").FormulaForceU = "Height*0";

				//	Set the paragraph alignment formula:
				shp.get_CellsU("VerticalAlign").FormulaForceU = "0";
			}
		}


		/** ************************************************************************************** **/

		public bool SetShapeTypesList(List<string> values)
		{
			if (values == null || values.Count <= 0)
			{
				return true;   // error
			}
			if (_shapeTypesList == null)
			{
				_shapeTypesList = new List<string>();
			}
			foreach (string value in values)
			{
				_shapeTypesList.Add(value);
			}
			return false;  // success
		}

		public List<string> GetShapeTypes()
		{
			if (_shapeTypesList == null)
			{
				return null;
			}
			return _shapeTypesList;
		}

		public string FindShapeType(string value)
		{
			if (_shapeTypesList == null)
			{
				return "";  // Use default value
			}
			if (string.IsNullOrEmpty(value))
			{
				return ""; // Use default value
			}
			foreach (string item in _shapeTypesList)
			{
				if (item.Equals(value.Trim(), StringComparison.OrdinalIgnoreCase))
				{
					return item;
				}
			}
			return "";  // Use default value
		}

		/** ************************************************************************************** **/

		public bool SetConnectorArrowsList(List<string> values)
		{
			if (values == null || values.Count <= 0)
			{
				return true;   // error
			}
			if (_connectorArrowsList == null)
			{
				_connectorArrowsList = new List<string>();
			}
			foreach (string value in values)
			{
				_connectorArrowsList.Add(value);
			}
			return false;  // success
		}

		public List<string> GetConnectorArrows()
		{
			if (_connectorArrowsList == null)
			{
				return null;
			}
			return _connectorArrowsList;
		}

		public string FindConnectorArrow(string value)
		{
			if (_connectorArrowsList == null)
			{
				return "";  // Use default value
			}
			if (string.IsNullOrEmpty(value))
			{
				return ""; // Use default value
			}
			foreach (string item in _connectorArrowsList)
			{
				if (item.Equals(value.Trim(), StringComparison.OrdinalIgnoreCase))
				{
					return item;
				}
			}
			return "";  // Use default value
		}

		/** ************************************************************************************** **/

		public bool SetConnectorLinePatternsList(List<string> values)
		{
			if (values == null || values.Count <= 0)
			{
				return true;   // error
			}

			if (_connectorLinePatternsList == null)
			{
				_connectorLinePatternsList = new List<string>();
			}
			foreach (string value in values)
			{
				_connectorLinePatternsList.Add(value);
			}
			return false;  // success
		}

		public List<string> GetConnectorLinePatterns()
		{
			if (_connectorLinePatternsList == null)
			{
				return null;
			}
			return _connectorLinePatternsList;
		}

		/// <summary>
		/// GetConnectorLinePatternText
		/// Because Visio uses a numeric value for line patterns we need
		/// to convert to text when writing to Excel
		/// </summary>
		/// <param name="value"></param>
		/// <returns></returns>
		public string GetConnectorLinePatternText(double value)
		{
			if (_connectorLinePatternsList == null)
			{
				return "";  // Use default value
			}
			if (value < 1)
			{
				return ""; // Use default value
			}

			// convert value to int (use as an index)
			switch(value)
			{
				case 2:
					return VisioVariables.sLINE_PATTERN_DASHED;
				case 3:
					return VisioVariables.sLINE_PATTERN_DOTTED;
				case 4:
					return VisioVariables.sLINE_PATTERN_DASHDOT;
				case 1:  // Solid
				default:	// solid
					return VisioVariables.sLINE_PATTERN_SOLID;
			}
		}

		public string FindConnectorLinePattern(string value)
		{
			if (_connectorLinePatternsList == null)
			{
				return "";  // Use default value
			}
			if (string.IsNullOrEmpty(value))
			{
				return ""; // Use default value
			}
			foreach (string item in _connectorLinePatternsList)
			{
				if (item.Equals(value.Trim(), StringComparison.OrdinalIgnoreCase))
				{
					return item;
				}
			}
			return "";  // Use default value
		}

		/** ************************************************************************************** **/

		public bool SetStencilLabelPositionsList(List<string> values)
		{
			if (values == null || values.Count <= 0)
			{
				return true;   // error
			}
			if (_stencilLabelPositionsList == null)
			{
				_stencilLabelPositionsList = new List<string>();
			}
			foreach (string value in values)
			{
				_stencilLabelPositionsList.Add(value);
			}

			return false;  // success
		}

		public List<string> GetStencilLabelPositions()
		{
			if (_stencilLabelPositionsList == null)
			{
				return null;
			}
			return _stencilLabelPositionsList;
		}

		public string FindStencilLabelPosition(string value)
		{
			if (_stencilLabelPositionsList == null)
			{
				return "";  // Use default value
			}
			if (string.IsNullOrEmpty(value))
			{
				return ""; // Use default value
			}
			foreach (string item in _stencilLabelPositionsList)
			{
				if (item.Equals(value.Trim(), StringComparison.OrdinalIgnoreCase))
				{
					return item;
				}
			}
			return "";  // Use default value
		}

		/** ************************************************************************************** **/

		public bool SetStencilLabelFontSizeList(List<string> values)
		{
			if (values == null || values.Count <= 0)
			{
				return true;   // error
			}
			if (_stencilLabelFontSizesList == null)
			{
				_stencilLabelFontSizesList = new List<string>();
			}
			foreach (string value in values)
			{
				_stencilLabelFontSizesList.Add(value);
			}
			return false;  // success
		}

		public List<string> GetStencilLabelFontSize()
		{
			if (_stencilLabelFontSizesList == null)
			{
				return null;
			}
			return _stencilLabelFontSizesList;
		}

		public string FindStencilLabelFontSize(string value)
		{
			if (_stencilLabelFontSizesList == null)
			{
				return "";  // Use default value
			}
			if (string.IsNullOrEmpty(value))
			{
				return ""; // Use default value
			}
			foreach (string item in _connectorLineWeightsList)
			{
				if (item.Equals(value.Trim(), StringComparison.OrdinalIgnoreCase))
				{
					return item;
				}
			}
			return "";  // Use default value
		}

		/** ************************************************************************************** **/

		public bool SetConnectorLineWeightsList(List<string> values)
		{
			if (values == null || values.Count <= 0)
			{
				return true;   // error
			}
			if (_connectorLineWeightsList == null)
			{
				_connectorLineWeightsList = new List<string>();
			}
			foreach (string value in values)
			{
				_connectorLineWeightsList.Add(value);
			}
			return false;  // success
		}

		public List<string> GetConnectorLineWeights()
		{
			if (_connectorLineWeightsList == null)
			{
				return null;
			}
			return _connectorLineWeightsList;
		}

		/// <summary>
		/// FindConnectorLineWeight
		/// search the list for the paramater
		/// if found use that value as the To or From Line Weight value as a string
		/// if not found null will be returned so use the default value
		/// ignore case
		/// </summary>
		/// <param name="value">lookup</param>
		/// <returns>Found value or null</returns>
		public string FindConnectorLineWeight(string value)
		{
			if (_connectorLineWeightsList == null)
			{
				return "";  // use default value
			}
			if (string.IsNullOrEmpty(value))
			{
				return "";  // Use default value
			}
			foreach (string item in _connectorLineWeightsList)
			{
				if (item.Equals(value.Trim(), StringComparison.OrdinalIgnoreCase))
				{
					return item;
				}
			}
			return "";  // use default value
		}

		/** ************************************************************************************** **/

		public bool SetDefaultStencilNamesList(List<string> names)
		{
			if (names == null || names.Count <= 0)
			{
				return true;   // error
			}
			if (_defaultStencilNames == null)
			{
				_defaultStencilNames = new List<string>();
			}
			foreach (string name in names)
			{
				_defaultStencilNames.Add(name);
			}
			return false;  // success
		}

		public List<string> GetDefaultStencilNames()
		{
			if (_defaultStencilNames == null)
			{
				return null;
			}
			return _defaultStencilNames;
		}


		/// <summary>
		/// FindDefaultStencilName
		/// Search if stencil map for the name value argument
		/// ignore case
		/// </summary>
		/// <param name="name">search name</param>
		/// <returns>null - if not found else the stencil name</returns>
		public string FindDefaultStencilName(string name)
		{
			if (_defaultStencilNames == null)
			{
				return "";  // use default value
			}
			if (string.IsNullOrEmpty(name))
			{
				return "";  // use default value
			}
			foreach (string item in _defaultStencilNames)
			{
				if (item.Equals(name.Trim(), StringComparison.OrdinalIgnoreCase))
				{
					return item;
				}
			}
			return "";  // use default value
		}

		/** ************************************************************************************** **/

		public void ClearVisioPageNamesList()
		{
			if (_visioPageNamesList != null)
			{
				_visioPageNamesList.Clear();
			}
		}

		/// <summary>
		/// AddVisioPageName
		/// if list does not exists create it
		/// add a new page name to the list
		/// Note this is only necessary during the processing of reading the Excel data file.
		/// Once the Visio diagram has been created use pagObj[x] to obtain the tabs
		/// </summary>
		/// <param name="name">add this value</param>
		/// <returns>bool true - failure; false - success</returns>
		public bool AddVisioPageName(string name)
		{
			if (string.IsNullOrEmpty(name))
			{
				return true;   // error
			}

			if (_visioPageNamesList == null)
			{
				_visioPageNamesList = new List<string>();
			}

			if (!_visioPageNamesList.Contains(name.Trim()))
			{
				_visioPageNamesList.Add(name.Trim());
			}
			return false;      // success
		}

		/// <summary>
		/// GetVisioPageNames
		/// return a list of strings containing page names
		/// this is really only used during the processing of reading the Excel data file
		/// after the Visio diagram has been created use pagesObj[x] to obtain the lames of tabs
		/// </summary>
		/// <returns>List<string></returns>
		public List<string> GetVisioPageNames()
		{
			if (_visioPageNamesList == null)
			{
				_visioPageNamesList = new List<string>();
			}
			return _visioPageNamesList;
		}

		/// <summary>
		/// GetVisioPageNumberByName
		/// find the Visio page number based on the Visio page name
		/// </summary>
		/// <param string>lookup by page name</param>
		/// <returns>int - page number</returns>
		//public int GetVisioPageNumberByName(string name)
		//{
		//	string value = string.Empty;
		//	if (_visioPageNamesList == null)
		//	{
		//		_visioPageNamesList = new List<string>();
		//	}
		//	if (string.IsNullOrEmpty(name))
		//	{
		//		return 1;
		//	}

		//	// search for the page
		//	for (int nIdx = 0; nIdx <= _visioPageNamesList.Count; nIdx++)
		//	{
		//		if (_visioPageNamesList[nIdx] == name)
		//		{
		//			return nIdx+1;	// zero base index.   can't return 0 out of index for Visio everything valuid is > 0
		//		}
		//	}
		//	return 1;  // use default value
		//}

		/// <summary>
		/// GetVisioPageNameByNumber
		/// get the page name for the given page number
		/// </summary>
		/// <param int>lookup vakye by page number</param>
		/// <returns string>page name or "1" if name not found</returns>
		//public string GetVisioPageNameByNumber(int value)
		//{
		//	if ( value <= 0 || _visioPageNamesList == null)
		//	{
		//		return "1";
		//	}
		//	return _visioPageNamesList[value];
		//}


		/** ************************************************************************************** **/

		public bool SetColorsMap(Dictionary<string, string> colorsMap)
		{
			if (colorsMap == null || colorsMap.Count <= 0)
			{
				return true;   // error
			}
			if (_visioColorsMap == null)
			{
				_visioColorsMap = new Dictionary<string,string>();
			}
			foreach (KeyValuePair<string,string> item in colorsMap)
			{
				_visioColorsMap.Add(item.Key, item.Value);
			}

			setColorNameColorMap();
			
			return false;  // success
		}

		public Dictionary<string,string> GetColorsMap()
		{
			if (_visioColorsMap == null)
			{
				_visioColorsMap = new Dictionary<string, string>();
			}
			return _visioColorsMap;
		}

		/// <summary>
		/// GetRGBColor
		/// return the RGB color value based on the color string argument
		/// color argument "Black" will return "RGB(0,0,0)"
		/// </summary>
		/// <param name="color">search value</param>
		/// <returns>"RGB(???,???,???)"</returns>
		public string GetRGBColor(string color)
		{
			string value = string.Empty;
			if (string.IsNullOrEmpty(color) || _visioColorsMap == null)
			{
				return "";
			}

			foreach (KeyValuePair<string, string> kvp in _visioColorsMap)
			{
				if (string.Equals(kvp.Key, color, StringComparison.OrdinalIgnoreCase))
				{
					return kvp.Value.Trim().ToString();
				}
			}
			return "";
		}

		/// <summary>
		/// FindColorValueFromRGB
		/// return the color string based on the rgb value argument
		/// search "RGB(0,0,0)" will return "Black"
		/// </summary>
		/// <param name="rgb"></param>
		/// <returns>string</returns>
		/// <text>color name</text>
		public string GetColorValueFromRGB(string rgb)
		{
			if (string.IsNullOrEmpty(rgb) || _visioColorsMap == null)
			{
				return "";
			}
			foreach (KeyValuePair<string, string> item in _visioColorsMap)
			{
				if (item.Value.Equals(rgb.Trim()))
				{
					return item.Key;
				}
			}
			return "";
		}

		/// <summary>
		/// GetAllColorNames
		/// return a list of color names
		/// </summary>
		/// <returns>List<string></returns>
		public List<string> GetAllColorNames()
		{
			List<string> saTmp2 = new List<string>();	
			foreach (KeyValuePair<string, string> keyValue in _visioColorsMap)
			{
				// adjust the index to be minus 1 bacause we added a row outside the array
				saTmp2.Add(keyValue.Key.Trim());
			}
			return saTmp2;
		}

		/// <summary>
		/// GetAllRGBValues
		/// Return a list of RGB values
		/// </summary>
		/// <returns>List<string></returns>
		public List<string> GetAllRGBValues()
		{
			List<string> saTmp2 = new List<string>();
			foreach (KeyValuePair<string, string> keyValue in _visioColorsMap)
			{
				// adjust the index to be minus 1 bacause we added a row outside the array
				saTmp2.Add(keyValue.Value.Trim());
			}
			return saTmp2;
		}

		public string FindColorbyName(string name)
		{
			string value = string.Empty;
			if (string.IsNullOrEmpty(name) || _visioColorsMap == null)
			{
				return "";
			}

			foreach (KeyValuePair<string, string> kvp in _visioColorsMap)
			{
				if (string.Equals(kvp.Key, name, StringComparison.OrdinalIgnoreCase))
				{
					return kvp.Key.Trim().ToString();
				}
			}
			return "";
		}


		/// <summary>
		/// SetColorNameColorMap
		/// populate the map after the the base color map has been created
		/// </summary>
		private void setColorNameColorMap()
		{
			_appColorsMap = new Dictionary<string, Color>();

			Regex regex = new Regex(@"^RGB\((?<r>\d{1,3}),(?<g>\d{1,3}),(?<b>\d{1,3})\)");
			foreach (KeyValuePair<string, string> item in GetColorsMap())
			{
				string nStr = item.Value.Replace(", ", ",");    // remove any spaces after the ',' character
				Match match = regex.Match(nStr.ToUpper());
				int r = int.Parse(match.Groups["r"].Value);
				int g = int.Parse(match.Groups["g"].Value);
				int b = int.Parse(match.Groups["b"].Value);
				_appColorsMap.Add(item.Key, Color.FromArgb(255, r, g, b));
			}
		}

		/// <summary>
		/// GetColorNameColor
		/// return a Dictionary<string,Color>
		/// this is used to help determine a valid color based on an rgb color
		/// </summary>
		/// <returns>Dictionary<string,Color></returns>
		public Dictionary<string, Color> GetColorNameColorsMap()
		{
			if (_appColorsMap == null)
			{
				setColorNameColorMap();
			}
			return _appColorsMap;
		}


		/** ************************************************************************************** **/

	}
}



// fill a shap with color
// PPS Solved! To remove the dependency on those sub-shapes, you need to change their Fillstyle to Normal.
// Just add new line of code ssh.FillStyle = 'Normal'.
//
//
//import win32com.client as w32
//visio = w32.Dispatch("visio.Application")
//visio.Visible = 1
//# create document based on Detailed Network Diagram template (use full path)
//doc = visio.Documents.Add("C:\Program Files\Microsoft Office\root\Office16\visio content\1033\dtlnet_m.vstx")
//# use one of docked stencilsList 
//stn2 = visio.Documents("PERIPH_M.vssx")
//#define 'Server' master-shape
//server = stn2.Masters("Server")
//#define page
//page = doc.Pages.Item(1)
//# rename page
//page.name = "My drawing"
//# drop master-shape on page, define 'Server' instance
//serv = page.Drop(server, 0, 0)
//# iterate sub-shapes (side edges) 
//for i in range (2,6):
//    #define one od side edges from 'Server'
//    ssh = serv.shapes(i)
//    # Change Fill Style to 'Normal'
//    ssh.FillStyle = 'Normal'
//    # fix FillForegnd cell for side edge
//    ssh.Cells('Fillforegnd').FormulaForceU = 'Guard(Sheet.' + str(serv.id) + '!FillForegnd)'
//    # fix FillBkgnd cell for side edge
//    ssh.Cells('FillBkgnd').FormulaForceU = 'Guard(Sheet.' + str(serv.id) + '!FillBkgnd)'
//    # instead formula 'Guard(x)' rewrite formula 'Guard(1)'    
//    ssh.Cells('FillPattern').FormulaForceU = 'Guard(1)'
//
//# fill main shape in 'Server' master-shape
//serv.Cells("FillForegnd").FormulaForceU = '5'
//