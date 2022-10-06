using Microsoft.Office.Interop.Visio;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using OmnicellBlueprintingTool.Models;
using Visio1 = Microsoft.Office.Interop.Visio;
using System.Xml.Linq;

namespace OmnicellBlueprintingTool.Visio
{
	public class VisioHelper
	{
		public Visio1.Application appVisio = null;
		public Visio1.Documents vDocuments = null;
		public Visio1.Document vDocument = null;

		List<Visio1.Document> stencils = new List<Visio1.Document>();

		public VisioHelper()
		{
		}

		/// <summary>
		/// SetupPage
		/// Set the Visio diagram page Orientation and page size
		/// </summary>
		/// <param name="currentPage">Current visio page</param>
		/// <param name="orientation"><options>"Portrait" or "Landscape"</options></param>
		/// <param name="size"><options>"Letter", "Tabloid", "Ledger", "Legal", "A3", "A4"</options></param>
		/// <return>bool<options>true error or false success</options></return>
		private bool SetupDiagramPage(Visio1.Page currentPage, string orientation, string size)
		{
			Visio1.Shape sheet = currentPage.PageSheet;
			string width = string.Empty;
			string height = string.Empty;

			if (currentPage == null || string.IsNullOrEmpty(orientation) || string.IsNullOrEmpty(size))
			{
				MessageBox.Show(string.Format("Error one of the following is null or empty: Page{0}, Orientation:{1}, Size:{3}", currentPage, orientation, size));
				return true;
			}

			switch (size.ToUpper())
			{
				case "TABLOID":
					width = "8.5 in";
					height = "11 in";
					break;
				case "LEDGER":
					width = "8.5 in";
					height = "11 in";
					break;
				case "LEGAL":
					width = "8.5 in";
					height = "11 in";
					break;
				case "A3":
					width = "8.5 in";
					height = "11 in";
					break;
				case "A4":
					width = "8.5 in";
					height = "11 in";
					break;
				case "LETTER":
				default:
					width = "8.5 in";
					height = "11 in";
					break;
			}

			switch (orientation.ToUpper())
			{
				case "LANDSCAPE":
					currentPage.PageSheet.Cells["PageWidth"].FormulaU = height;
					currentPage.PageSheet.Cells["PageHeight"].FormulaU = width;
					currentPage.PageSheet.Cells["PrintPageOrientation"].FormulaU = "2";

					//currentPage.PageSheet.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionObject, (short)Visio1.VisRowIndices.visRowPrintProperties, (short)Visio1.VisCellIndices.visPageWidth).FormulaU = height;
					//currentPage.PageSheet.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionObject, (short)Visio1.VisRowIndices.visRowPrintProperties, (short)Visio1.VisCellIndices.visPageHeight).FormulaU = width;
					//currentPage.PageSheet.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionObject, (short)Visio1.VisRowIndices.visRowPrintProperties, (short)Visio1.VisCellIndices.visPageDrawSizeType).FormulaU = "3";
					break;
				case "PORTRAIT":
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
				if (!string.IsNullOrEmpty(diagramData.visioTemplateFilePath))
				{
					// we need to check if the file is a template file or not
					// this will open a template file
					// Create a new document. but you will need to add a master stencil
					this.vDocument = appVisio.Documents.Add(diagramData.visioTemplateFilePath);
				}
				else
				{
					// create a new blank document
					this.vDocument = appVisio.Documents.Add("");
				}
			}
			catch (Exception ex1)
			{
				sErr = "Error with the Template file";
				MessageBox.Show(string.Format("Exception::setupVisioDiagram - {0}\n{1}", sErr, ex1));
				return null;
			}
			try
			{
				// this gives a count of all the stencils on the status bar
				int countStencils = vDocument.Masters.Count;

				// get the current draw page
				Visio1.Page currentPage = vDocument.Pages[1];

				// lets add stencils to the template if they don't alredy exist using the Excel Data File
				foreach (var stencil in diagramData.visioStencilFilePaths)
				{
					if (this.vDocuments != null)  // do we have any stencils attached to this template?
					{
						var vPage = vDocument.Application.ActivePage;

						// Load the stencil we want
						Visio1.Document nStencil = vDocuments.OpenEx(stencil, (short)Visio1.VisOpenSaveArgs.visOpenDocked);
						stencils.Add(nStencil);

						// show the stencil window
						//Visio1.Window stencilWindow = currentPage.Document.OpenStencilWindow();
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
				sErr = "Error with the stencil file.  Possible issue is the stencil file name changed\nDoes not match what the Template file is using";
				MessageBox.Show(string.Format("Exception::setupVisioDiagram - {0}\n{1}", sErr, ex2));
				return null;
			}

			Visio1.Pages pagesObj = appVisio.ActiveDocument.Pages;
			//appVisio.Visible = true;

			// The new document will have one page, get the a reference to it.
			Visio1.Page page1 = vDocument.Pages[1];
			page1.Name = "Page-1";
			page1.AutoSize = false;
			//page1.AutoSizeDrawing();

			//Assuming 'No theme' is set for the page, no arrow will be shown so change theme to see connector arrow
			page1.SetTheme("Office Theme");

			// Page 1 is Standard
			if (!SetupDiagramPage(page1, diagramData.VisioPageOrientation, diagramData.VisioPageSize))
			{
				double xPosition = page1.PageSheet.get_CellsU("PageWidth").ResultIU;
				double yPosition = page1.PageSheet.get_CellsU("PageHeight").ResultIU;
				var pageOrientation = page1.PageSheet.get_CellsU("PrintPageOrientation").ResultIU;
				ConsoleOut.writeLine(string.Format("page:{0}, Height:{1}, Width:{2} and Orientation:{3}", page1.Name, yPosition, xPosition, diagramData.VisioPageOrientation));
			}

			int cnt = this.vDocuments.Count;
			// this section is if we want to add more than one page to the document
			// At this point page-1 has already been created
			for (int i = 0; i < diagramData.MaxVisioPages - 1; i++)
			{
				Visio1.Page page = vDocument.Pages.Add();
				// Name the pages. This is what is shown in the page tabs.
				page.Name = "Page-" + (i + 2);
				page.AutoSize = true;
				// this.vPage.AutoSizeDrawing(); // this can make the page taller

				//Assuming 'No theme' is set for the page, no arrow will be shown so change theme to see connector arrow
				page.SetTheme("Office Theme");

				// Page 1 is Standard
				if (!SetupDiagramPage(page, diagramData.VisioPageOrientation, diagramData.VisioPageSize))
				{
					double xPosition = page.PageSheet.get_CellsU("PageWidth").ResultIU;
					double yPosition = page.PageSheet.get_CellsU("PageHeight").ResultIU;
					var pageOrientation = page.PageSheet.get_CellsU("PrintPageOrientation").ResultIU;
					ConsoleOut.writeLine(string.Format("page:{0}, Height:{1}, Width:{2} and Orientation:{3}", page.Name, yPosition, xPosition, diagramData.VisioPageOrientation));
				}
			}

			// Move the second page to the first position in the list of pages.
			//page1.Index = 1;
			//page2.Index = 2;
			//return page1.Index;

			// set the active page to the first page
			this.appVisio.ActiveWindow.Page = pagesObj[1];
			return pagesObj;
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
			}

			if (stnObj == null)
			{
				// else look to see if the Stencil is part of the added stincel files
				if (this.stencils.Count > 0)
				{
					foreach (Visio1.Document stencil in this.stencils)
					{
						try
						{
							stnObj = stencil.Masters[device.ShapeInfo.StencilImage];
						}
						catch(System.Runtime.InteropServices.COMException com)
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
				string sTmp = string.Format("ERROR::_drawShape - Can't find Stencil:{0}", device.ShapeInfo.StencilImage);
				MessageBox.Show(sTmp);
				Console.WriteLine(sTmp);
				return null;
			}

			Visio1.Pages pagesObj = this.appVisio.ActiveDocument.Pages;
			// switch Visio Diagram Page if needed based on the shape data VisioPage value
			this.appVisio.ActiveWindow.Page = pagesObj[device.ShapeInfo.VisioPage];

			// draw the shape on the document
			shpObj = this.appVisio.ActivePage.Drop(stnObj, device.ShapeInfo.Pos_x, device.ShapeInfo.Pos_y);
			shpObj.NameU = device.ShapeInfo.UniqueKey;
			if ("NetworkPipe".Equals(device.ShapeInfo.StencilImage))
			{
				// we need to resize the stencil NetworkPipe (remember this is rotated 90deg
				// so we need to go East for Height and South for Width
				if (device.ShapeInfo.Width > 0.0)
				{
					// we need to make wider (increase south)
					shpObj.Resize(VisResizeDirection.visResizeDirS, device.ShapeInfo.Height, VisUnitCodes.visDrawingUnits);
				}
				if (device.ShapeInfo.Height > 0.0)
				{
					// we need to make taller (increase east)
					shpObj.Resize(VisResizeDirection.visResizeDirE, device.ShapeInfo.Width, VisUnitCodes.visDrawingUnits);
				}
			}
			else
			{
				// normal stencils are normal (east-width and south-height)
				if (device.ShapeInfo.Width > 0.0)
				{
					// we need to make wider (increase east)
					shpObj.Resize(VisResizeDirection.visResizeDirE, device.ShapeInfo.Width, VisUnitCodes.visDrawingUnits);
				}
				if (device.ShapeInfo.Height > 0.0)
				{
					// we need to make taller (increase south)
					shpObj.Resize(VisResizeDirection.visResizeDirS, device.ShapeInfo.Height, VisUnitCodes.visDrawingUnits);
				}
			}

			//var linePatternCell = shpConn.get_CellsU("LinePattern");
			string rgb = string.Empty;
			switch (device.ShapeInfo.FillColor.Trim().ToUpper())
			{
				case "YELLOW":
					rgb = VisioVariables.COLOR_YELLOW;
					break;
				case "GREEN":
					rgb = VisioVariables.COLOR_GREEN;
					break;
				case "LIGHT GREEN":
					rgb = VisioVariables.COLOR_GREEN_LIGHT;
					break;
				case "RED":
					rgb = VisioVariables.COLOR_RED;
					break;
				case "GRAY":
					rgb = VisioVariables.COLOR_GRAY;
					break;

				case "BLUE":
					rgb = VisioVariables.COLOR_BLUE;
					break;
				case "LIGHT BLUE":
					rgb = VisioVariables.COLOR_BLUE_SERVER;
					break;
				case "CYAN":
					rgb = VisioVariables.COLOR_CYAN;
					break;

				case "ORANGE":
					rgb = VisioVariables.COLOR_ORANGE;
					break;
				case "LIGHT ORANGE":
					rgb = VisioVariables.COLOR_ORANGE_SERVER;
					break;
				default:
					// no fill
					break;
			}
			if (!string.IsNullOrEmpty(rgb))
			{
				//set the shape back color
				shpObj.get_CellsSRC(
					(short)VisSectionIndices.visSectionObject,
					(short)VisRowIndices.visRowFill,
					(short)VisCellIndices.visFillForegnd).FormulaU = rgb;

				shpObj.get_CellsSRC(
					 (short)Visio1.VisSectionIndices.visSectionObject,
					 (short)Visio1.VisRowIndices.visRowFill,
					 (short)Visio1.VisCellIndices.visFillBkgnd).FormulaU = VisioVariables.COLOR_BLACK;

				// for an shape to be filled this needs to be set
				shpObj.get_CellsSRC(
					 (short)Visio1.VisSectionIndices.visSectionObject,
					 (short)Visio1.VisRowIndices.visRowFill,
					 (short)Visio1.VisCellIndices.visFillPattern).FormulaU = "1";

				shpObj.get_Cells("LineColor").Formula = rgb;
			}

			// we want to keep the shape outline color Black for this Stencil
			if (device.ShapeInfo.UniqueKey.ToUpper().StartsWith("TABLECELL"))
			{
				shpObj.get_Cells("LineColor").Formula = VisioVariables.COLOR_BLACK;
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
					shpObj.get_Cells("Char.Size").Formula = "=" + device.ShapeInfo.StencilLabelFontSize + " pt";
					//shpObj.Cells("Char.Size").FormulaU = device.ShapeInfo.StencilLabelFontSize + " pt";
					//string fontSize = shpObj.get_Cells("Char.Size").Formula;
				}
				// check if we have an IP that needs to be displayed
				if (!string.IsNullOrEmpty(device.OmniIP))
				{
					shpObj.Text += " (" + device.OmniIP + ")";
				}
				if (!string.IsNullOrEmpty(device.OmniPorts))
				{
					shpObj.Text += " (" + device.OmniPorts + ")";
				}
				int textLen = shpObj.Text.Length;


				// dont resize the text for the Title and Footer stencils
				//				if (!"Title".Equals(device.ShapeInfo.StencilImage) && !"Footer".Equals(device.ShapeInfo.StencilImage))
				//				{
				//					if (textLen > 25)
				//					{
				//						double scale = shpObj.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionUser, (short)Visio1.VisRowIndices.visRowUser, (short)Visio1.VisCellIndices.visUserValue).ResultIU;
				//						scale = 0.5;
				//						//Then set the font, and the TextMargins(for any that are non - zero) with the following(assuming the normal font size is 12 and the left margin is 4pt.:
				//						shpObj.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionCharacter, 0, (short)Visio1.VisCellIndices.visCharacterSize).FormulaU = (scale * 12).ToString() + "pt";
				//						shpObj.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionObject, (short)Visio1.VisRowIndices.visRowText, (short)Visio1.VisCellIndices.visTxtBlkLeftMargin).FormulaU = (scale * 4).ToString() + "pt";
				//					}
				//				}
			}
			ConsoleOut.writeLine("draw Stencil: " + device.ShapeInfo.StencilImage);
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
		/// this will clear the stencils list for reuse
		/// </summary>
		public void ClearStencilList()
		{
			if (this.stencils != null)
			{
				// must clear this list otherwise an Exception will occur dealing with RPS miss leading error when app is ran again without closing
				stencils.Clear();
			}
		}
		/// <summary>
		/// VisioForceCloseAll
		/// Close all the Visio documents
		/// This will display the Save file dialog
		/// </summary>
		public void VisioForceCloseAll()
		{
			try
			{
				ClearStencilList();

				if (this.vDocuments != null)
				{
					while (this.vDocuments.Count > 0)
					{
						this.vDocuments.Application.ActiveDocument.Close();
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
				ConsoleOut.writeLine(string.Format("This exception is OK because the user closed the Visio document before exiting the application:  {0}", ex.Message));
			}
		}

		/// <summary>
		/// DrawAllShapes
		/// Draw all the visio stencils obtained from the data file
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
					MessageBox.Show(string.Format("Exception::setupVisioDiagram - Stencil Image:{0} not found.  Shape Key:{1}\n{2}", device.ShapeInfo.StencilImage, device.ShapeInfo.UniqueKey, ex.Message));
					Console.WriteLine(string.Format("Exception::setupVisioDiagram - Stencil Image:{0} not found.  Shape Key:{1}\n{2}", device.ShapeInfo.StencilImage, device.ShapeInfo.UniqueKey, ex.Message));
					return true;
				}
			}
			// use before saving AutoSizeDrawing
			appVisio.AutoLayout = true;
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

			try
			{
				// iterate over the ShapeConnectionsMap to determine if a connection shape is needed
				for (int nCnt = 0; nCnt < diagData.ShapeConnectionsMap.Count; nCnt++)
				{
					// nCnt is the key
					if (diagData.ShapeConnectionsMap.TryGetValue(nCnt, out lookupShapeConnection))
					{
						// Drop the built-in connector object on the lower left corner of the page:
						// need to drop on another page
						Visio1.Pages pagesObj = appVisio.ActiveDocument.Pages;

						// switch Visio Diagram Page if needed based on the shape data VisioPage value
						appVisio.ActiveWindow.Page = pagesObj[lookupShapeConnection.device.ShapeInfo.VisioPage];

						// draw the object on the Visio diagram
						shpConn = appVisio.ActivePage.Drop(pagesObj.Application.ConnectorToolDataObject, 0.0, 0.0);

						// Set the connector properties, using different arrows, colors, and patterns for many-to-many relationships.
						shpConn.get_CellsU("ShdwPattern").ResultIU = VisioVariables.SHDW_PATTERN;
						shpConn.get_CellsU("BeginArrow").ResultIU = VisioVariables.ARROW_NONE;
						shpConn.get_CellsU("EndArrow").ResultIU = VisioVariables.ARROW_NONE;
						shpConn.get_CellsU("LineColor").FormulaU = VisioVariables.COLOR_BLACK;
						shpConn.get_CellsU("Rounding").ResultIU = VisioVariables.ROUNDING;
						shpConn.get_CellsU("LinePattern").ResultIU = VisioVariables.LINE_PATTERN_SOLID;
						shpConn.get_CellsU("LineWeight").FormulaU = VisioVariables.LINE_WEIGHT_1;

						if (lookupShapeConnection.device.ShapeInfo.LineWeight != VisioVariables.LINE_WEIGHT_1)
						{
							shpConn.get_CellsU("LineWeight").FormulaU = lookupShapeConnection.device.ShapeInfo.LineWeight;
						}
						if (lookupShapeConnection.LinePattern > 0)
						{
							shpConn.get_CellsU("LinePattern").ResultIU = lookupShapeConnection.LinePattern;
						}

						switch ((string)lookupShapeConnection.ArrowType.Trim().ToUpper())
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
						switch (lookupShapeConnection.LineColor.Trim().ToUpper())
						{
							case "YELLOW":
								shpConn.get_CellsU("LineColor").FormulaU = VisioVariables.COLOR_YELLOW;
								break;
							case "LIGHT GREEN":
								shpConn.get_CellsU("LineColor").FormulaU = VisioVariables.COLOR_GREEN_LIGHT;
								break;
							case "GREEN":
								shpConn.get_CellsU("LineColor").FormulaU = VisioVariables.COLOR_GREEN;
								break;
							case "GRAY":
								shpConn.get_CellsU("LineColor").FormulaU = VisioVariables.COLOR_GRAY;
								break;
							case "RED":
								shpConn.get_CellsU("LineColor").FormulaU = VisioVariables.COLOR_RED;
								break;
							case "CYAN":
								shpConn.get_CellsU("LineColor").FormulaU = VisioVariables.COLOR_CYAN;
								break;
							case "LIGHT BLUE":
								shpConn.get_CellsU("LineColor").FormulaU = VisioVariables.COLOR_BLUE_SERVER;
								break;
							case "BLUE":
								shpConn.get_CellsU("LineColor").FormulaU = VisioVariables.COLOR_BLUE;
								break;
							case "LIGHT ORANGE":
								shpConn.get_CellsU("LineColor").FormulaU = VisioVariables.COLOR_ORANGE_SERVER;
								break;
							case "ORANGE":
								shpConn.get_CellsU("LineColor").FormulaU = VisioVariables.COLOR_ORANGE;
								break;
							default:
							case "BLACK":
								shpConn.get_CellsU("LineColor").FormulaU = VisioVariables.COLOR_BLACK;
								break;
						}

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
				throw new Exception(string.Format("Exception::ConnectShapes - {0}", ex.Message));
			}
			return false;
		}

		/// <summary>
		/// ListStencils
		/// List all the stencils in the master stencil document
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

			ArrayList masterArray_0 = new ArrayList();
			Visio1.Document doc_0 = vDocuments[1];    // Document stencil figures
			Visio1.Document doc_1 = vDocuments[2];    // Stencil figures
			Visio1.Masters masters_0 = doc_0.Masters;
			Visio1.Masters masters_1 = doc_1.Masters;
			foreach (Visio1.Master master in masters_0)
			{
				// Document stencil figures
				masterArray_0.Add(master.NameU);   // THIS WILL CONTAIN DIAGRAM FIGURES
				Console.WriteLine(string.Format("ListDocumentStencils - Master0 - ID:{0} Name:{1} NameU:{2}", master.ID, master.Name, master.NameU));
			}
			Console.WriteLine("\n\n");
			foreach (Visio1.Master master in masters_1)
			{
				// Document stencil figures
				masterArray_0.Add(master.NameU);   // THIS WILL CONTAIN DIAGRAM FIGURES
				Console.WriteLine(string.Format("ListDocumentStencils - Master1 - ID:{0} Name:{1} NameU:{2}", master.ID, master.Name, master.NameU));
			}
			this.VisioForceCloseAll();

			return false;
		}
	}
}

//		private void setShapeTextBottom()
//		{
//			Visio1.Selection sel = null;
//			
//			sel = Visio1.ActiveWindow.Selection;
//
//			foreach(Visio1.Shape shp in sel)
//			{ 
//				// '// 'Add' the Text Transfomrm section, if it's not there:
//
//				if (!shp.RowExists   ((short)Visio1.VisSectionIndices.visSectionObject,
//										 (short)Visio1.VisRowIndices.visRowTextXForm,
//										 (short)Visio1.VisExistsFlags.visExistsAnywhere ))
//				{
//					shp.AddRow((short)Visio1.VisSectionIndices.visSectionObject, (short)Visio1.VisRowIndices.visRowTextXForm, (short)Visio1.VisRowTags.visTagDefault);
//				}
//
//				//	Set the text transform formulas:
//				shp.get_CellsU("TxtHeight").FormulaForceU = "Height*0";
//				shp.get_CellsU("TxtPinY").FormulaForceU = "Height*0";
//
//				//	Set the paragraph alignment formula:
//				shp.get_CellsU("VerticalAlign").FormulaForceU = "0";
//			}
//		}


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
//# use one of docked stencils 
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