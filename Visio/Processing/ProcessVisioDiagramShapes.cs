using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using OmnicellBlueprintingTool.Models;
using Visio1 = Microsoft.Office.Interop.Visio;
using Microsoft.Office.Core;
using System.Drawing;
using ColorV = Microsoft.Office.Interop.Visio.Color;
using System.Data.SqlTypes;
using System.Text.RegularExpressions;
using Color = System.Drawing.Color;
using System.Reflection.Emit;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using VisioAutomation.VDX.Elements;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using OmnicellBlueprintingTool.ExcelHelpers;
using Microsoft.Extensions.FileSystemGlobbing.Internal;
using Microsoft.Extensions.FileSystemGlobbing;
using static System.Net.Mime.MediaTypeNames;
using System.Xml.Linq;
using static OmnicellBlueprintingTool.Visio.VisioVariables;

namespace OmnicellBlueprintingTool.Visio
{
	public class ProcessVisioDiagramShapes
	{
		public Visio1.Application appVisio = null;
		public Visio1.Documents vDocuments = null;
		public Visio1.Document vDocument = null;

		Visio1.Page vPage = null;
		VisioHelper visHlpr = null;

		/// <summary>
		/// GetAllShapesProperties
		/// Get all the shape properties for each shape contained within the document
		/// </summary>
		/// <param name="diagamFilePathName"></param>
		/// <return>Dictionary<int, ShapeInformation> </return>
		public Dictionary<string, ShapeInformation> GetAllShapesProperties(VisioHelper visioHelper, string diagamFilePathName, VisioVariables.ShowDiagram dspMode)
		{
			// Open up one of Visio's sample drawings.
			appVisio = new Visio1.Application();

			visHlpr = new VisioHelper();
			visHlpr.ShowVisioDiagram(appVisio, dspMode);              // don't show the diagram

			vDocument = appVisio.Documents.Open(diagamFilePathName);
			vDocuments = appVisio.Documents;
			
			// The new document will have one page, get the a reference to it.
			vPage = vDocument.Pages[1];

			visHlpr.ShowVisioDiagram(appVisio, VisioVariables.ShowDiagram.Show);
			ConsoleOut.writeLine(string.Format("Active Document:{0,20}: Master in document:{1,20}", appVisio.ActiveDocument, appVisio.ActiveDocument.Masters));

			// get the connectors for each shape in the diagram 
			Dictionary<string, ShapeInformation> shpConn = getShapeInformation(visioHelper, appVisio, vDocument);

			try
			{
				if (this.vDocuments != null)
				{
					vDocument.Saved = true;
					vDocument.Close();
				}
				vDocuments = null;
				if (appVisio != null)
				{
					appVisio.Quit();
					appVisio = null;
				}
				GC.Collect();
				GC.WaitForPendingFinalizers();
				GC.Collect();
				GC.WaitForPendingFinalizers();
			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				ConsoleOut.writeLine(string.Format("This exception is OK because the user closed the Visio document before exiting the application:{0}", ex.Message));
			}
			return shpConn;
		}

		/// <summary>
		/// getShapeInformation
		/// get the visio diagram stencil connection information
		/// used to add to the Excel Data file
		/// </summary>
		/// <param name="doc">Visio document</param>
		/// <returns>"Dictionary<string, ShapeInformation>"</returns>
		private static Dictionary<string, ShapeInformation> getShapeInformation(VisioHelper visioHelper, Visio1.Application appVisio, Visio1.Document doc)
		{
			// Look at each shape in the collection.
			Visio1.Page page = null;
			Visio1.Pages pagesObj = doc.Pages;

			Dictionary<string, ShapeInformation> allPageShapesMap = null;	// changed the Key to a string because we need to use a GUID from the shape
			Dictionary<int, Visio1.Shape> connectorsMap = new Dictionary<int, Visio1.Shape>();
			ShapeInformation shpInfo = null;

			try
			{
				//allPageShapesMap = new Dictionary<int, ShapeInformation>();
				allPageShapesMap = new Dictionary<string, ShapeInformation>();
				string sColor = string.Empty;

				// start at 1 because Visio page indexing is not zero base
				for (int nCnt = 1; nCnt <= pagesObj.Count; nCnt++ )	// we need to loop through each tab in the Visio Diagram 
				{
					page = pagesObj[nCnt];
					appVisio.ActiveWindow.Page = pagesObj[nCnt]; 
					
					ConsoleOut.writeLine(string.Format("Gathering all shapes from page:'{0}'", page.Name));

					foreach (Visio1.Shape shape in page.Shapes)	// get the shapes on this page
					{
						// Use this index to look at each row in the properties section.
						shpInfo = new ShapeInformation();

						int nIdx = page.Name.IndexOf("Page-", StringComparison.OrdinalIgnoreCase);
						if (nIdx >= 0)
						{
							// found lets modify the Visio Page field
							// I.E.  Page-# is a default type of page so we will write just the # in the Excel file Visio Page column
							string[] saTmp = page.Name.Split('-');
							shpInfo.VisioPage = saTmp[1]; // just get the number when writing to the Excel file
						}
						else
						{	
							// use the fill page name
							shpInfo.VisioPage = page.Name;
						}

						// this will return a string name:ID
						shpInfo.StencilImage = getShapeName(shape.Name);
						shpInfo.StencilLabel = shape.Text.Trim();

						shpInfo.UniqueKey = getShapeUniqueKeyName(shape);

						double colorIdx = shape.CellsU["FillBkgnd"].ResultIU;
						Microsoft.Office.Interop.Visio.Color c = doc.Colors.Item16[(short)colorIdx];
						shpInfo.RGBFillColor = $"RGB({c.Red},{c.Green},{c.Blue})";
						sColor = visioHelper.GetColorValueFromRGB(shpInfo.RGBFillColor);
						shpInfo.FillColor = "";
						if (!string.IsNullOrEmpty(sColor))
						{
							shpInfo.FillColor = sColor;
							shpInfo.RGBFillColor = "";		// we dont need the rgb fill color if a color was found
						}
						if (shape.Name.IndexOf("OC_Dash", StringComparison.OrdinalIgnoreCase) >= 0)
						{
							// TODO-2 look at this later
							colorIdx = shape.CellsU["LineColor"].ResultIU;
							Microsoft.Office.Interop.Visio.Color c1 = doc.Colors.Item16[(short)colorIdx];
							shpInfo.RGBFillColor = $"RGB({c1.Red},{c1.Green},{c1.Blue})";
							sColor = visioHelper.GetColorValueFromRGB(shpInfo.RGBFillColor);
							shpInfo.FillColor = "";
							if (!string.IsNullOrEmpty(sColor))
							{
								shpInfo.FillColor = sColor;
								shpInfo.RGBFillColor = ""; // we dont need the rgb fill color if a color was found
							}						
						}

						if (shpInfo.RGBFillColor.IndexOf(VisioVariables.RGB_COLOR_2SKIP, StringComparison.OrdinalIgnoreCase) >= 0 || 
							 shpInfo.RGBFillColor.Equals("RGB(0,0,0)", StringComparison.OrdinalIgnoreCase) || 
							 shpInfo.FillColor.IndexOf("White", StringComparison.OrdinalIgnoreCase) >= 0 )
						{
							// the color is black for a shape so we don't need to set this field empty is default for black
							shpInfo.RGBFillColor = string.Empty;   // we dont need the rgb fill color if a color was found
							shpInfo.FillColor = string.Empty;
						}

						//short iRow = (short)VisRowIndices.visRowFirst;
						shpInfo.Pos_x = Math.Truncate(shape.Cells["PinX"].ResultIU * 1000) / 1000;
						shpInfo.Pos_y = Math.Truncate(shape.Cells["PinY"].ResultIU * 1000) / 1000;

						// set the shape width including the stencil width adjustment
						shpInfo.Width = (Math.Truncate(shape.Cells["Width"].ResultIU * 1000) / 1000);
						// set the shape height including the stencil height adjustment 
						shpInfo.Height = (Math.Truncate(shape.Cells["Height"].ResultIU * 1000) / 1000);

						// get the origional stencel width and height to use as an offset later
						var sizes = GetStencilSize(appVisio, shpInfo.StencilImage);
						shpInfo.StencilWidth = sizes.width;
						shpInfo.StencilHeight = sizes.height;

						if (shpInfo.Width <= (shpInfo.StencilWidth + VisioVariables.STENCIL_SIZE_BUFFER) &&
							 shpInfo.Width >= Math.Abs(shpInfo.StencilWidth - VisioVariables.STENCIL_SIZE_BUFFER))
						{
							// no size adjustment needed to be added to Excel Width cell
							shpInfo.StencilWidth = 0;
							shpInfo.Width = 0;
						}

						if (shpInfo.Height <= (shpInfo.StencilHeight + VisioVariables.STENCIL_SIZE_BUFFER) &&
							 shpInfo.Height >= Math.Abs(shpInfo.StencilHeight - VisioVariables.STENCIL_SIZE_BUFFER))
						{
							// no size adjustment needed to be added to Excel Height cell
							shpInfo.StencilHeight = 0;
							shpInfo.Height = 0;
						}

						// if the shpInfo.Weight and shpInfo.Height values are already 0 skip
						if (shpInfo.Width != 0.0 && shpInfo.Height != 0.0)
						{
							// need to manage some shapes differently
							if (shape.Name.IndexOf("OC_PortsLDAP_", StringComparison.OrdinalIgnoreCase) >= 0 ||
								 shape.Name.IndexOf("OC_IconKey", StringComparison.OrdinalIgnoreCase) >= 0)
							{
								shpInfo.Width = 0;
								shpInfo.Height = 0;
							}
							else if ((shape.Name.IndexOf("OC_Ethernet", StringComparison.OrdinalIgnoreCase)) >=0 )
							{
								// this stencil is vertical don't set height if shape is Ethernet type
								// stencil OC_Ethernet2H is Horizontial may need to reverse the width/height
								shpInfo.Height = 0;
							}
						}

						// use the connection shape to obtain what is connected to what
						if (shape.Style.Trim().IndexOf("Connector", StringComparison.OrdinalIgnoreCase) >= 0)
						{
							// get connection information add to the dictionary to be used later
							if (!connectorsMap.ContainsKey(shape.ID))
							{
								connectorsMap.Add(shape.ID, shape);                                           
							}
							else
							{
								ConsoleOut.writeLine(string.Format("ERROR::Failed to add connection UniqueKey:{0,30} GUID:{1,40} Text:'{2}' already exists in the Map", shpInfo.ID, shape.NameU, shape.UniqueID[(short)VisUniqueIDArgs.visGetGUID], shape.Text));
							}
							// we don't want to add this shape object to the allPageShapesMap
							continue;
						}

						shpInfo.ID = shape.ID;

						// we are using GUID which is Visio's way of making a global unique ID for a shape
						shpInfo.GUID = shape.UniqueID[(short)VisUniqueIDArgs.visGetOrMakeGUID];				
						if (!allPageShapesMap.ContainsKey(shpInfo.GUID))		// Using the GUID from the shape.  this is Unique through all of the documents in the file
						{
							allPageShapesMap.Add(shpInfo.GUID, shpInfo);  // shape.ID
							ConsoleOut.writeLine(string.Format("Added shape UniqueKey:{0} GUID:{1} Text:'{2}'", shpInfo.UniqueKey.PadRight(25), shpInfo.GUID.PadRight(40), shpInfo.StencilLabel));
						}
						else
						{
							ConsoleOut.writeLine(string.Format("ERROR::Failed to add shape UniqueKey:{0} GUID:{1} Text:'{2}'", shpInfo.UniqueKey.PadRight(25), shpInfo.GUID.PadRight(40), shpInfo.StencilLabel));
						}
					}

					// now let make the connections
					foreach (KeyValuePair<int, Visio1.Shape> shape in connectorsMap)
					{
						if (shape.Value.Connects.Count > 0)
						{
							getShapeConnections(visioHelper, doc, shape.Value, ref allPageShapesMap, ref shpInfo);
						}
					}
				}
			}
			catch (Exception ex)
			{
				string sTmp = string.Format("ProcessVisioDiagramShapes::getShapeInformation - Exception\n\nForeach loop:\nPage:'{0,25}' UniqueKey:'{1,30}' GUID:{3}\n\n{4}", page.Name.PadRight(20), shpInfo.UniqueKey.PadRight(25), shpInfo.GUID, ex.Message);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			return allPageShapesMap;
		}

		/// <summary>
		/// GetStencilSize
		/// Get the size width, height of the stencil based on the argument
		/// </summary>
		/// <param name="app">Visio1.Application</param>
		/// <param name="stencilName">search for this</param>
		/// <returns>(double width, double height)</returns>
		public static (double width, double height) GetStencilSize(Visio1.Application app, string name)
		{
			Visio1.Master stencil = null;

			double width = 0;
			double height = 0;

			// for some reason unknown at this time a stencil may not be found
			// but i know it does exists in the stencil file because the shape was drawn using the stencil file
			// in any case if not found just return width and height = 0
			try
			{
				if (string.IsNullOrEmpty(name))
				{
					return (width, height);
				}
				foreach (Visio1.Document doc in app.Documents)
				{
					try
					{
						stencil = doc.Masters.get_ItemU(name);
					}
					catch (Exception ex) 
					{
						// eat it and fall through
					}
					if (stencil != null)
					{
						foreach (Visio1.Shape shape in stencil.Shapes)
						{
							if (shape.Name.Trim().IndexOf(name) >= 0)
							{
								width = Math.Truncate(shape.get_Cells("Width").ResultIU * 1000) / 1000;
								height = Math.Truncate(shape.get_Cells("Height").ResultIU * 1000) / 1000;
							}
						}
					}
				}
			}
			catch (Exception ep)
			{
				ConsoleOut.writeLine(string.Format("ProcessVisioDiagramShapes::GetStencilSize - Exception\nStencil:'{0}' not found.\n\n{1}", name, ep.Message.ToString()));
				// fall through with both width and height set to 0
			}
			return (width, height);
		}

		/// <summary>
		/// IsShapeHeightSameAsStencilHeight
		/// compare the two shapes Heights
		/// use a gap of VisioVariables.STENCIL_SIZE_BUFFER to adjust for a little bit of size differences
		/// </summary>
		/// <param name="shpInfo"></param>
		/// <returns>bool</returns>
		private static bool IsShapeHeightSameAsStencilHeight(ShapeInformation shpInfo)
		{
			// TODO - 2 we need to put this in a common module
			if (shpInfo.Height <= (shpInfo.StencilHeight + VisioVariables.STENCIL_SIZE_BUFFER) &&
				 shpInfo.Height >= Math.Abs(shpInfo.StencilHeight - VisioVariables.STENCIL_SIZE_BUFFER))
			{
				// if within margin - good
				return true;
			}
			return false;
		}

		/// <summary>
		/// getShapeUniqueKeyName
		/// this will fix the shape name for Excel Data file
		/// the name will be formatted as just name (should match stencil Image name) with ":" followed by the shape ID
		/// </summary>
		/// <param name="shape"></param>
		/// <returns>string or null</returns>
		private static string getShapeUniqueKeyName(Visio1.Shape shape)
		{
			// this shouuld never happen
			if (shape == null)
			{
				return null;
			}
			string[] saTmp = shape.Name.Split('.');
			if (saTmp.Length > 0)
			{
				if (saTmp.Length > 3)
				{
					// this is special lets use the full name
					return string.Format("{0}:{1}", shape.Name.Trim(), shape.ID);

				}
				// lets get the first part as the shape name
				return string.Format("{0}:{1}", saTmp[0].Trim(), shape.ID);
			}
			return string.Format("{0}:{1}", shape.Name.Trim(), shape.ID);
		}
		
		/// <summary>
		/// getShapeName
		/// return the shape name should be same as shape Image name
		/// </summary>
		/// <param name="shape"></param>
		/// <returns>string</returns>
		private static string getShapeName(string name)
		{
			// this shouuld never happen
			if (string.IsNullOrEmpty(name))
			{
				return null;
			}
			string[] saTmp = name.Split('.');
			if (saTmp.Length > 0)
			{
				// lets get the first part as the shape name
				return saTmp[0].Trim().Trim();
			}
			saTmp = name.Split(':');
			if (saTmp.Length > 0)
			{
				// lets get the first part as the shape name
				return saTmp[0].Trim().Trim();
			}
			return string.Format("{0}", name.Trim());
		}

		/// <summary>
		/// getShapeConnections
		/// the shape passed in is a connection shape
		/// from the connector shape get the connector information (color, text, weight, etc.)
		/// from the connector shape get the toSheet that the connector is connected to (should be a stencil shape already in the AllPageShapesMap)
		/// the Key should be the shape GUID
		/// </summary>
		/// <param name="visioHelper"></param>
		/// <param name="doc">document object</param>
		/// <param name="connShp"> connection shape object</param>
		/// <param name="allPageShapesMap">dictionary that contains all the shapes in the diagram</param>
		/// <param name="shpInfo">dictionary to hold the modified shapes to be returned</param>
		private static void getShapeConnections(VisioHelper visioHelper, Visio1.Document doc, Visio1.Shape connShp, ref Dictionary<string, ShapeInformation> allPageShapesMap, ref ShapeInformation shpInfo)
		{
			// the shape passed in is the connection
			// the toSheet is what the connection is connected to.
			// in most cases the connections will be 2 or greater
			
			string sTmp = string.Empty;
			string sTmp2 = string.Empty;
			string lineWeight = String.Empty;

			Visio1.Shape toshape = null;
			Visio1.Connect visconnect = null;

			ShapeInformation lookupShapeMap = null;
			string lookupGUID = string.Empty;		//  use the shape GUID which is what was used as the key to appPageShapesMap
			
			string connectorLabel = string.Empty;
			string arrowType = VisioVariables.sARROW_NONE;
			string lineColor = VisioVariables.sCOLOR_BLACK;
			string rgbLineColor = visioHelper.GetRGBColor(VisioVariables.sCOLOR_BLACK);
			double linePattern = VisioVariables.LINE_PATTERN_SOLID;
			Visio1.Connects visconnects2 = connShp.Connects;

			// get the connector information
			connectorLabel = connShp.Text;

			try
			{
				int startArrow = 0;
				int endArrow = 0;
				string data = connShp.get_CellsU("BeginArrow").FormulaU;
				if (data.IndexOf("THEME", StringComparison.OrdinalIgnoreCase) < 0)
				{
					startArrow = int.Parse(connShp.get_CellsU("BeginArrow").FormulaU);
				}
				data = connShp.get_CellsU("EndArrow").FormulaU;
				if (data.IndexOf("THEME", StringComparison.OrdinalIgnoreCase) < 0)
				{
					endArrow = int.Parse(connShp.get_CellsU("EndArrow").FormulaU);
				}
				if (startArrow > 0 && endArrow > 0) // both
				{
					arrowType = VisioVariables.sARROW_BOTH;
				}
				else if (startArrow > 0 && endArrow == 0)
				{
					arrowType = VisioVariables.sARROW_START;
				}
				else if (startArrow == 0 && endArrow > 0)
				{
					arrowType = VisioVariables.sARROW_END;
				}

				var colorIdx = connShp.CellsU["LineColor"].ResultIU;
				Microsoft.Office.Interop.Visio.Color c = doc.Colors.Item16[(short)colorIdx];
				shpInfo.RGBFillColor = $"RGB({c.Red},{c.Green},{c.Blue})";
				string sColor = visioHelper.GetColorValueFromRGB(shpInfo.RGBFillColor);
				shpInfo.ToLineColor = "";
				if (!string.IsNullOrEmpty(sColor))
				{
					lineColor = sColor;
					shpInfo.RGBFillColor = ""; // we dont need the rgb fill color if a color was found
				}

				linePattern = connShp.get_CellsU("LinePattern").ResultIU;

				lineWeight = VisioVariables.sLINE_WEIGHT_1;	// set default
				sTmp = connShp.get_CellsU("LineWeight").FormulaU;
				if (sTmp.IndexOf("THEME", StringComparison.OrdinalIgnoreCase) < 0)
				{
					// we have a valid value so lets see if we support it
					if (visioHelper.IsConnectorLineWeight(sTmp))
					{
						lineWeight = sTmp;	// set new lineWeight
					}
				}

				int nFromCnt = 0;
				int nToCnt = 0;
				int ethernetID = 0;
				string ethernetUniqueKey = string.Empty;
				string ethernetUniqueGUID = string.Empty;

				for (int k = 1; k <= visconnects2.Count; k++)
				{
					// look through the connections to get the both ends
					visconnect = visconnects2[k];
					toshape = visconnect.ToSheet;

					if (k == 1)
					{
						// first end From
						sTmp = string.Empty;
						lookupGUID = toshape.UniqueID[(short)VisUniqueIDArgs.visGetGUID];
						//uniqueKey = getShapeUniqueKeyName(toshape);

						// the key must match the same key (UniqueKey or GUID)
						allPageShapesMap.TryGetValue(lookupGUID, out lookupShapeMap);

						// if this first one is the Ethernet stencil we need to do some special work
						if (toshape.Name.IndexOf("OC_Ethernet", StringComparison.OrdinalIgnoreCase) >= 0)
						{
							ethernetID = toshape.ID; // save this we need to do a trick later
							ethernetUniqueKey = getShapeUniqueKeyName(toshape);
							ethernetUniqueGUID = toshape.UniqueID[(short)VisUniqueIDArgs.visGetGUID];
							sTmp = string.Format("Connect From Shape:{0} ", ethernetUniqueKey);
						}
						else
						{
							// shape is not a Ethernet type shape
							sTmp = string.Format("Connect Shape:{0} GUID:{1} ", getShapeUniqueKeyName(toshape).PadRight(30), lookupGUID.PadRight(40));
							ethernetID = 0;      // not ethernet shpe type
							ethernetUniqueKey = string.Empty;
							ethernetUniqueGUID = string.Empty;
						}
					}
					else
					{
						// get the next stencil shape to connect to
						// we need to test this section I"m not sure this will occur because all the diagram shapes should be in the allPageShapesMap
						if (lookupShapeMap == null)
						{
							//lookupKey = toshape.ID;
							if (ethernetID > 0)
							{
								//allPageShapesMap.TryGetValue(ethernetID, out lookupShapeMap);
								allPageShapesMap.TryGetValue(ethernetUniqueGUID, out lookupShapeMap);
								ethernetID = 0;
							}
							else
							{
								//allPageShapesMap.TryGetValue(toshape.ID, out lookupShapeMap);
								allPageShapesMap.TryGetValue(lookupGUID, out lookupShapeMap);
							}
							if (lookupShapeMap != null)
							{
								if (string.IsNullOrEmpty(lookupShapeMap.ConnectFrom))
								{
									if (nFromCnt++ > 0)
									{
										lookupShapeMap.ConnectFrom += "," + getShapeUniqueKeyName(toshape); // lookupShape.NameU;
									}
									else
									{
										lookupShapeMap.ConnectFrom += getShapeUniqueKeyName(toshape); // lookupShape.NameU;
									}
									lookupShapeMap.ConnectFromID = toshape.ID;
								}
								lookupShapeMap.FromLineLabel = connectorLabel;
								lookupShapeMap.FromArrowType = arrowType;
								lookupShapeMap.FromLineColor = lineColor;
								lookupShapeMap.FromLinePattern = linePattern;
								lookupShapeMap.FromLineWeight = lineWeight;
							}
							//uniqueKey = getShapeUniqueKeyName(toshape);
							lookupGUID = toshape.UniqueID[(short)VisUniqueIDArgs.visGetGUID];
						}
						else
						{
							// if the first connector was Ethernet we need to make this a From type
							// this is when we need to do some special work
							if (!string.IsNullOrEmpty(ethernetUniqueKey))
							{
								// get the next shape to be populated using From..
								allPageShapesMap.TryGetValue(toshape.UniqueID[(short)VisUniqueIDArgs.visGetGUID], out lookupShapeMap);  // get the next connection shape
								if (string.IsNullOrEmpty(lookupShapeMap.ConnectFrom))
								{
									if (nToCnt++ > 0)
									{
										lookupShapeMap.ConnectFrom += "," + ethernetUniqueKey;   // append to existing value
									}
									else
									{
										lookupShapeMap.ConnectFrom += ethernetUniqueKey;
									}
									lookupShapeMap.ConnectFromID = ethernetID;
								}

								lookupShapeMap.FromLineLabel = connectorLabel;  // use the Text value from the connector shape
								lookupShapeMap.FromArrowType = arrowType;
								lookupShapeMap.FromLineColor = lineColor;
								lookupShapeMap.FromLinePattern = linePattern;
								lookupShapeMap.FromLineWeight = lineWeight;

								//uniqueKey = getShapeUniqueKeyName(toshape);
								lookupGUID = toshape.UniqueID[(short)VisUniqueIDArgs.visGetGUID];

								// clear these values
								ethernetID = 0;
								ethernetUniqueKey = string.Empty;
							}
							else
							{
								// shape is not an Ethernet shape so coontinue as normal
								if (string.IsNullOrEmpty(lookupShapeMap.ConnectTo))
								{
									if (nToCnt++ > 0)
									{
										lookupShapeMap.ConnectTo += "," + getShapeUniqueKeyName(toshape); // append to existing value;
									}
									else
									{
										lookupShapeMap.ConnectTo += getShapeUniqueKeyName(toshape);
									}
									lookupShapeMap.ConnectToID = toshape.ID;
								}

								lookupShapeMap.ToLineLabel = connectorLabel;  // use the Text value from the connector shape
								lookupShapeMap.ToArrowType = arrowType;
								lookupShapeMap.ToLineColor = lineColor;
								lookupShapeMap.ToLinePattern = linePattern;
								lookupShapeMap.ToLineWeight = lineWeight;
							}
						}
						sTmp += string.Format("to Shape:{0,10}, linelabel:'{1}'", getShapeUniqueKeyName(toshape), connShp.Text);
					}
				}
				if (lookupShapeMap != null)
				{
					allPageShapesMap[lookupGUID] = lookupShapeMap;
				}
			}
			catch (Exception ex)
			{
				sTmp = string.Format("ProcessVisioDiagramShapes::getShapeConnections - Exception\n\nshape ID:{0,10}, GUID:{1,30}, Label:'{2}'\n\n{3}", connShp.ID, lookupGUID, connShp.Text, ex.Message);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}

			ConsoleOut.writeLine(sTmp);
		}
		
		/*******************************************************************************************************************/

		/// <summary>
		/// GetShapeConnections
		/// this will attempt to get the connection information between stencilsList using stencilsList on the Visio Diagram
		/// it only works on stencil shapes not connector types
		/// problem is getting the connector information line pattern, color, label etc
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="shape"></param>
		/// <param name="allPageShapesMap"></param>
		/// <param name="shpInfo"></param>
		private static void getShapeConnectionsOld(VisioHelper visioHelper, Visio1.Document doc, Visio1.Shape shape, ref Dictionary<int, ShapeInformation> allPageShapesMap, ref ShapeInformation shpInfo)
		{
			// what we need to do is get the shape and determine if shape is connected.
			// we don't want to save the shape if it is a connector

			// Look at each shape in the collection.
			Visio1.Page page = doc.Pages[1];

			Dictionary<int, string> connectors = new Dictionary<int, string>();
			Dictionary<string, string> connectMap = new Dictionary<string, string>();

			//continue;
			string sTmp = string.Empty;
			string sTmp2 = string.Empty;
			string lineWeight = string.Empty;

			// get connections To
			var shpConnection = shape.ConnectedShapes(VisConnectedShapesFlags.visConnectedShapesOutgoingNodes, "");
			if (shpConnection != null && shpConnection.Length > 0)
			{
				try
				{
					string sColor = string.Empty;
					int nCnt = 0;
					foreach (int nIdx in shpConnection)
					{
						Visio1.Shape lookupShape = page.Shapes.ItemFromID[nIdx];
						string sKey = shape.ID + ":" + lookupShape.ID;
						string sKey2 = lookupShape.ID + ":" + shape.ID; // see if duplicate exists
						if (connectMap.ContainsKey(sKey) || connectMap.ContainsKey(sKey2))
						{
							//shpInfo.ConnectToID = 0;
							//shpInfo.ConnectTo = string.Empty;
							continue;      // we don't want to save this information
						}
						ConsoleOut.writeLine(string.Format("Connecting shapeID:{0,10}-{1,30} To shapeID:{2,30}-{3,30} in the diagram", shpInfo.ID, shpInfo.UniqueKey, lookupShape.ID, lookupShape.NameU));
						shpInfo.ConnectToID = lookupShape.ID;
						if (nCnt++ > 0)
						{
							shpInfo.ConnectTo += "," + getShapeUniqueKeyName(lookupShape);
						}
						else
						{
							shpInfo.ConnectTo += getShapeUniqueKeyName(lookupShape);
						}

						// new section
						int startArrow = int.Parse(lookupShape.get_CellsU("BeginArrow").FormulaU);
						int endArrow = int.Parse(lookupShape.get_CellsU("EndArrow").FormulaU);
						if (startArrow > 0 && endArrow > 0) // both
						{
							shpInfo.ToArrowType = VisioVariables.sARROW_BOTH;
						}
						else if (startArrow > 0 && endArrow == 0)
						{
							shpInfo.ToArrowType = VisioVariables.sARROW_START;
						}
						else if (startArrow == 0 && endArrow > 0)
						{
							shpInfo.ToArrowType = VisioVariables.sARROW_END;
						}
						else
						{
							shpInfo.ToArrowType = VisioVariables.sARROW_NONE;
						}
						var colorIdx = shape.CellsU["LineColor"].ResultIU;
						Microsoft.Office.Interop.Visio.Color c = doc.Colors.Item16[(short)colorIdx];
						shpInfo.RGBFillColor = $"RGB({c.Red},{c.Green},{c.Blue})";
						sColor = visioHelper.GetColorValueFromRGB(shpInfo.RGBFillColor);

						shpInfo.ToLineColor = "";
						if (!string.IsNullOrEmpty(sColor))
						{
							shpInfo.ToLineColor = sColor;
						}

						// end new section

						//shpInfo.ToLineColor = VisioVariables.sCOLOR_BLACK;
						shpInfo.ToLinePattern = VisioVariables.LINE_PATTERN_SOLID;
						shpInfo.ToLineLabel = shape.Text;
						connectMap.Add(sKey, sKey2);
					}
				}
				catch (Exception exp)
				{
					sTmp = String.Format("ProcessVisioDiagramShapes::getShapeConnections - Exception\n\nConnection:{0}\n{1}", shpInfo.ConnectFrom, exp.Message);
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}

			// get connections from
			shpConnection = shape.ConnectedShapes(VisConnectedShapesFlags.visConnectedShapesIncomingNodes, "");
			if (shpConnection != null && shpConnection.Length > 0)
			{
				try
				{
					string sColor = string.Empty;
					int nCnt = 0;
					foreach (int nIdx in shpConnection)
					{
						Visio1.Shape lookupShape = page.Shapes.ItemFromID[nIdx];

						string sKey = shape.ID + ":" + lookupShape.ID;
						string sKey2 = lookupShape.ID + ":" + shape.ID; // see if duplicate exists
						if (connectMap.ContainsKey(sKey) || connectMap.ContainsKey(sKey2))
						{
							//shpInfo.ConnectFromID = 0;
							//shpInfo.ConnectFrom = string.Empty;
							continue;      // we don't want to save this information
						}
						ConsoleOut.writeLine(string.Format("Connecting shapeID:{0,10}-{1,30} From shapeID:{2,30}-{3,30} in the diagram", shpInfo.ID, shpInfo.UniqueKey, lookupShape.ID, lookupShape.NameU));
						shpInfo.ConnectFromID = lookupShape.ID;
						if (nCnt++ > 0)
						{
							shpInfo.ConnectFrom += "," + getShapeUniqueKeyName(lookupShape);
						}
						else
						{
							shpInfo.ConnectFrom += getShapeUniqueKeyName(lookupShape);
						}

						// new section
						int startArrow = int.Parse(lookupShape.get_CellsU("BeginArrow").FormulaU);
						int endArrow = int.Parse(lookupShape.get_CellsU("EndArrow").FormulaU);
						if (startArrow > 0 && endArrow > 0) // both
						{
							shpInfo.FromArrowType = VisioVariables.sARROW_BOTH;
						}
						else if (startArrow > 0 && endArrow == 0)
						{
							shpInfo.FromArrowType = VisioVariables.sARROW_START;
						}
						else if (startArrow == 0 && endArrow > 0)
						{
							shpInfo.FromArrowType = VisioVariables.sARROW_END;
						}
						else
						{
							shpInfo.ToArrowType = VisioVariables.sARROW_NONE;
						}
						var colorIdx = shape.CellsU["LineColor"].ResultIU;
						Microsoft.Office.Interop.Visio.Color c = doc.Colors.Item16[(short)colorIdx];
						shpInfo.RGBFillColor = $"RGB({c.Red},{c.Green},{c.Blue})";
						sColor = visioHelper.GetColorValueFromRGB(shpInfo.RGBFillColor);
						shpInfo.FromLineColor = "";
						if (!string.IsNullOrEmpty(sColor))
						{
							shpInfo.FromLineColor = sColor;
							shpInfo.RGBFillColor = ""; // we dont need the rgb fill color if a color was found
						}

						//shpInfo.FromLineColor = VisioVariables.sCOLOR_BLACK;
						shpInfo.FromLinePattern = VisioVariables.LINE_PATTERN_SOLID;
						shpInfo.FromLineLabel = shape.Text;
						connectMap.Add(sKey, sKey2);
					}
				}
				catch (Exception exp)
				{
					sTmp = string.Format("ProcessVisioDiagramShapes::getShapeConnections - Exception\n\nConnection:{0}\n{1}", shpInfo.ConnectFrom, exp.Message);
					MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		/// <summary>
		/// GetShapeConnections2
		/// this will attempt to get the connection information for each connection object
		/// this method should be able to provide more information about the connect line (Color, Pattern, etc)
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="connShp"></param>
		/// <param name="allPageShapesMap"></param>
		/// <param name="shpInfo"></param>
		private static void getShapeConnections2Old(VisioHelper visioHelper, Visio1.Document doc, Visio1.Shape connShp, ref Dictionary<int, ShapeInformation> allPageShapesMap, ref ShapeInformation shpInfo, ref Dictionary<int, string> connectorsMap)
		{
			// what we need to do is get the shape and determine if shape is connected.
			// we don't want to save the shape if it is a connector

			// Look at each shape in the collection.
			//Visio1.Page page = doc.Pages[1];

			//Dictionary<int, string> connectors = new Dictionary<int, string>();
			//Dictionary<string, string> connectMap = new Dictionary<string, string>();

			//continue;
			string sTmp = string.Empty;
			string sTmp2 = string.Empty;
			string lineWeight = String.Empty;

			ShapeInformation lookupFromShapeMap = null;
			ShapeInformation lookupShapeMap = null;
			int lookupKey = 0;
			int lookupFromKey = 0;
			string arrowType = VisioVariables.sARROW_NONE;
			string lineColor = VisioVariables.sCOLOR_BLACK;
			string rgbLineColor = visioHelper.GetRGBColor(VisioVariables.sCOLOR_BLACK);
			double linePattern = VisioVariables.LINE_PATTERN_SOLID;
			Visio1.Connects visconnects2 = connShp.Connects;

			for (int k = 1; k <= visconnects2.Count; k++)
			{
				// look through the connections to get the both ends
				Visio1.Connect visconnect = visconnects2[k];
				Visio1.Shape fromshape = visconnect.FromSheet;

				// first end From
				lookupFromKey = fromshape.ID;
				allPageShapesMap.TryGetValue(fromshape.ID, out lookupFromShapeMap);

				Visio1.Shape toshape = visconnect.ToSheet;
				if (k == 1)
				{
					// first end From
					lookupKey = toshape.ID;
					allPageShapesMap.TryGetValue(toshape.ID, out lookupShapeMap);

					sTmp = string.Empty;
					sTmp2 = string.Empty;
					sTmp = string.Format("Connector ID:'{0,10}' Shape ID:'{1,10}'-'{2,30}' LineLabel:'{3}'", connShp.ID, toshape.ID, getShapeUniqueKeyName(toshape), connShp.Text);
					sTmp2 = string.Format("id:'{0,10}';name:'{1,30}';label:'{2}'", toshape.ID, toshape.Name, connShp.Text);

					int startArrow = int.Parse(connShp.get_CellsU("BeginArrow").FormulaU);
					int endArrow = int.Parse(connShp.get_CellsU("EndArrow").FormulaU);
					if (startArrow > 0 && endArrow > 0) // both
					{
						arrowType = VisioVariables.sARROW_BOTH;
					}
					else if (startArrow > 0 && endArrow == 0)
					{
						arrowType = VisioVariables.sARROW_START;
					}
					else if (startArrow == 0 && endArrow > 0)
					{
						arrowType = VisioVariables.sARROW_END;
					}
					var colorIdx = connShp.CellsU["LineColor"].ResultIU;
					Microsoft.Office.Interop.Visio.Color c = doc.Colors.Item16[(short)colorIdx];
					shpInfo.RGBFillColor = $"RGB({c.Red},{c.Green},{c.Blue})";
					string sColor = visioHelper.GetColorValueFromRGB(shpInfo.RGBFillColor);
					shpInfo.ToLineColor = "";
					if (!string.IsNullOrEmpty(sColor))
					{
						lineColor = sColor;
						shpInfo.RGBFillColor = ""; // we dont need the rgb fill color if a color was found
					}

					//rgbLineColor = connShp.get_CellsU("LineColor").FormulaU;      // RGB color value
					//if (rgbLineColor.IndexOf("THEME", StringComparison.OrdinalIgnoreCase) >= 0)
					//{
					//	// we need to parse out the RGB value
					//	int nStart = rgbLineColor.IndexOf("RGB");
					//	rgbLineColor = rgbLineColor.Substring(nStart, (rgbLineColor.Length - nStart - 1));
					//
					//	//Color c = doc.Colors.Item16[(short)rgbLineColor];
					//	//shpInfo.RGBFillColor = $"RGB({c.Red},{c.Green},{c.Blue})";
					//
					//}
					//lineColor = String.Empty;
					//sColor = visioHelper.GetColorValueFromRGB(rgbLineColor);		// will be a color word or null if not found
					//if (string.IsNullOrEmpty(sColor))
					//{
					//	lineColor = VisioVariables.sCOLOR_BLACK;		// connector line color
					//}

					linePattern = connShp.get_CellsU("LinePattern").ResultIU;

					lineWeight = VisioVariables.sLINE_WEIGHT_1;	// set default
					sTmp = connShp.get_CellsU("LineWeight").FormulaU;
					if (sTmp.IndexOf("THERM", StringComparison.OrdinalIgnoreCase) < 0)
					{
						// we have a valid value so lets see if we support it
						if (visioHelper.IsConnectorLineWeight(sTmp))
						{
							lineWeight = sTmp;	// set value
						}
					}
				}
				else
				{
					// second end To
					if (lookupShapeMap == null)
					{
						// the shape was not found so we are lookup up the From shape
						// fill in the connectFrom fields if this has occurred
						//lookupKey = toshape.ID;
						allPageShapesMap.TryGetValue(toshape.ID, out lookupShapeMap);
						if (lookupShapeMap != null)
						{
							if (string.IsNullOrEmpty(lookupShapeMap.ConnectFrom))
							{
								lookupShapeMap.ConnectFrom = getShapeUniqueKeyName(toshape);
								lookupShapeMap.ConnectFromID = lookupKey;
							}
							lookupShapeMap.FromLineLabel = connShp.Text;
							lookupShapeMap.FromArrowType = arrowType;
							lookupShapeMap.FromLineColor = lineColor;
							lookupShapeMap.FromLinePattern = linePattern;
							lookupShapeMap.FromLineWeight = lineWeight;
						}
						lookupKey = toshape.ID; // keep in this order.  we use this for update the object
					}
					else
					{
						if (string.IsNullOrEmpty(lookupShapeMap.ConnectTo))
						{
							lookupShapeMap.ConnectTo = getShapeUniqueKeyName(toshape);
							lookupShapeMap.ConnectToID = toshape.ID;
						}
						lookupShapeMap.ToLineLabel = connShp.Text;  // use the Text value from the connector shape
						lookupShapeMap.ToArrowType = arrowType;
						lookupShapeMap.ToLineColor = lineColor;
						lookupShapeMap.ToLinePattern = linePattern;
						lookupShapeMap.ToLineWeight = lineWeight;
					}

					sTmp += string.Format(" - '{0,10}' To Shape ID:'{1,10}'-'{2,30}' LineLabel:'{3}'", connShp.ID, toshape.ID, getShapeUniqueKeyName(toshape), connShp.Text);
					sTmp2 += string.Format("|id:'{0,10}';name:'{1,30}';label:'{2}'", toshape.ID, toshape.Name, connShp.Text);
				}
			}
			if (lookupShapeMap != null)
			{
				allPageShapesMap[lookupKey] = lookupShapeMap;
			}

			connectorsMap.Add(connShp.ID, sTmp2);
			ConsoleOut.writeLine(sTmp);
			ConsoleOut.writeLine(string.Format("Found shape ID:'{0,10}'-'{1,30}' in the diagram", shpInfo.ID, shpInfo.UniqueKey));
		}


	}
}
