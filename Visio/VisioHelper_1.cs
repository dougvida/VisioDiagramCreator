using Microsoft.Office.Interop.Visio;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VisioAutomation.VDX.Elements;
using VisioDiagramCreator.Models;
using Visio1 = Microsoft.Office.Interop.Visio;
using Font = Microsoft.Office.Interop.Visio.Font;
using System.Windows.Forms;
using System.Security.Policy;
using System.Runtime.Remoting.Contexts;

namespace VisioDiagramCreator.Visio
{
	public class VisioHelper
	{
		public Visio1.Application appVisio;
		public Visio1.Documents vDocuments;
		public Visio1.Document vDocument;
		Visio1.Document stencils;
		public Visio1.Page pageObj;

		private Visio1.Pages _setupVisioDiagram(DiagramData allData)
		{
			// Start Visio
			appVisio = new Visio1.Application();
			vDocuments = appVisio.Documents;


			//var visioDocument = docsObj.Add(allData.TemplateFilePath);
			vDocument = vDocuments.Add(allData.TemplateFilePath);

			// this is only needed if the visio template file does not contain the stencil
			// Use this method if the stencil needs to be added to the Visio document
			// we should test if stencil is found to determine which one to use
			//Visio1.Document stincels = appVisio.Documents.Add(allData.StencilFilePath);

			// use this method if the template file already contains the stencil
			stencils = vDocuments[allData.StencilFilePath];

			//
			// this section is if we want to add more than one page to the document
			//
			Visio1.Pages pagesObj = appVisio.ActiveDocument.Pages;
			pageObj = pagesObj[1];
			pageObj.AutoSize = false;
			//pageObj.AutoSizeDrawing();


			// Create a new document.
			//Visio1.Document doc = vApp.Documents.Add("");

			// The new document will have one page, get the a reference to it.
			Visio1.Page page1 = vDocument.Pages[1];

			// Add a second page.
			Visio1.Page page2 = vDocument.Pages.Add();

			// Name the pages. This is what is shown in the page tabs.
			page1.Name = "Abc";
			page2.Name = "Def";

			// Move the second page to the first position in the list of pages.
			page2.Index = 1;
			//return page1.Index;

			return pagesObj;
		}


		public void DrawAllShapes(ref DiagramData allData)
		{
			Visio1.Pages vPages = _setupVisioDiagram(allData);

			int nShowSites = allData.sites.Count();
			int nSitesDroppedCnt = 0;
			int nSites = allData.sites.Count();

			// Add this back in if we are using an site stencil
//			if (nShowSites > 3)  // dont show more than 3 site stencils on the document
//			{
//				nShowSites = 1;
//			}
//			else
//			{
//				nShowSites = nSites;
//			}

			Visio1.Shape shpObj = null;
			foreach (Device device in allData.devices)
			{
				try
				{
					// we only want to draw one site if there are more than 3
					if (device.VisioInfo.UniqueKey.Contains("Site-"))
					{
						if (nSitesDroppedCnt < nShowSites)
						{
							// jsut show one site stencil and combind all the site text to one
							shpObj = _drawShape(ref allData, device);
							if (nShowSites == 1)
							{
								shpObj.Text = "Number of Sites: " + allData.sites.Count();
							}
							nSitesDroppedCnt++;
						}
					}
					else
					{
						// draw other shapes
						// add list of shaps to ignore
						if (!device.VisioInfo.UniqueKey.Contains("CTS"))
						{
							shpObj = _drawShape(ref allData, device);
						}
					}
				}
				catch (Exception ex)
				{
					throw new Exception(device.VisioInfo.UniqueKey + "::" + ex.Message);
				}
			}
			// use before saving AutoSizeDrawing
			appVisio.AutoLayout = true;
			pageObj.AutoSize = true;
			pageObj.AutoSizeDrawing();

		}

		/**
		 *	_drawShape - draw the stencil shape on the visio diagram
		 *
		 *	<param name="allData">Global variables object</param>
		 *	<param name="shapeObj">the information about the shape to draw</param>
		 **/
		private Visio1.Shape _drawShape(ref DiagramData data, Device device)
		{
			Visio1.Master stnObj = stencils.Masters[device.VisioInfo.StencilImage];
			if (stnObj == null)
			{
				MessageBox.Show("Can't find master stencil: " + device.VisioInfo.StencilImage);
				throw new Exception(device.VisioInfo.StencilImage + "Visio Helper PageStencil");
			}
			Visio1.Shape shpObj = appVisio.ActivePage.Drop(stnObj, device.VisioInfo.Pos_x, device.VisioInfo.Pos_y);
			//if ("GroupLarge".Equals(shapeObj.VisioInfo.StencilImage))
			//{  // resize the NetworkPipe stencil
			//	shpObj.Resize(VisResizeDirection.visResizeDirS, 0.2, VisUnitCodes.visDrawingUnits);
			//}
			if ("NetworkPipe".Equals(device.VisioInfo.StencilImage))
			{	// resize the NetworkPipe stencil
				shpObj.Resize(VisResizeDirection.visResizeDirE, 0.5, VisUnitCodes.visDrawingUnits);
			}
			if (!string.IsNullOrEmpty(device.VisioInfo.StencilLabel))
			{
				shpObj.Text = device.VisioInfo.StencilLabel;
				//shpObj.TextStyleKeepFmt = "Normal";		// Using this code would not allow the font size to be changed
				//shpObj.Cells("Char.Size").FormulaU = "8 pt";
				//string fontSize = shpObj.get_Cells("Char.Size").Formula;

				// check if we have an IP that needs to be displayed
				if (!string.IsNullOrEmpty(device.OmniIP))
				{
					shpObj.Text += " ("+device.OmniIP+")";
				}
				if (!string.IsNullOrEmpty(device.OmniPorts))
				{
					shpObj.Text += " (" + device.OmniPorts + ")";
				}
				int textLen = shpObj.Text.Length;

				// dont resize the text for the Title and Footer stencils
				if (!"Title".Equals(device.VisioInfo.StencilImage) && !"Footer".Equals(device.VisioInfo.StencilImage))
				{
					if (textLen > 25)
					{
						double scale = shpObj.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionUser, (short)Visio1.VisRowIndices.visRowUser, (short)Visio1.VisCellIndices.visUserValue).ResultIU;
						scale = 0.5;
						//Then set the font, and the TextMargins(for any that are non - zero) with the following(assuming the normal font size is 12 and the left margin is 4pt.:
						shpObj.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionCharacter, 0, (short)Visio1.VisCellIndices.visCharacterSize).FormulaU = (scale * 12).ToString() + "pt";
						shpObj.get_CellsSRC((short)Visio1.VisSectionIndices.visSectionObject, (short)Visio1.VisRowIndices.visRowText, (short)Visio1.VisCellIndices.visTxtBlkLeftMargin).FormulaU = (scale * 4).ToString() + "pt";
					}
				}
			}
			// the dictionary should contain 
			data.connectionMap.Add(device.VisioInfo.UniqueKey, shpObj);
			return shpObj;
		}


		public void ListStencils()
		{
			Visio1.Application app = new Visio1.Application();
			Visio1.Documents docs = app.Documents;

			ArrayList masterArray_0 = new ArrayList();
			ArrayList masterArray_1 = new ArrayList();
			Visio1.Document doc_0 = docs[1];    // HERE IS THE MAIN POINT
			Visio1.Document doc_1 = docs[2];    // HERE IS THE MAIN POINT
			Visio1.Masters masters_0 = doc_0.Masters;
			Visio1.Masters masters_1 = doc_1.Masters;
			foreach (Visio1.Master master in masters_0)
			{
				masterArray_0.Add(master.NameU);   // THIS WILL CONTAIN DIAGRAM FIGURES
			}
			foreach (Visio1.Master master in masters_1)
			{
				masterArray_1.Add(master.NameU);  // THIS WILL CONTAIN STENCIL FIGURES
			}
		}
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