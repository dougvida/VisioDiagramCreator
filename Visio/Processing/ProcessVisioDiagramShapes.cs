using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using OmnicellBlueprintingTool.Models;
using Visio1 = Microsoft.Office.Interop.Visio;
using Microsoft.Office.Core;
using System.Drawing;
using Color = Microsoft.Office.Interop.Visio.Color;

namespace OmnicellBlueprintingTool.Visio
{
	public class ProcessVisioDiagramShapes
	{
		Visio1.Application appVisio = null;
		Visio1.Document vDocument = null;
		Visio1.Page vPage = null;
		VisioHelper visHlpr = null;

		/// <summary>
		/// GetAllShapesProperties
		/// Get all the shape properties for each shape contained within the document
		/// </summary>
		/// <param name="diagamFilePathName"></param>
		/// <return>Dictionary<int, ShapeInformation> </return>
		public Dictionary<int, ShapeInformation> GetAllShapesProperties(string diagamFilePathName, VisioVariables.ShowDiagram dspMode)
		{
			// Open up one of Visio's sample drawings.
			appVisio = new Visio1.Application();
			visHlpr = new VisioHelper();
			visHlpr.ShowVisioDiagram(appVisio, dspMode);              // don't show the diagram

			this.vDocument = appVisio.Documents.Open(diagamFilePathName);
			// The new document will have one page, get the a reference to it.
			vPage = this.vDocument.Pages[1];

			visHlpr.ShowVisioDiagram(appVisio, VisioVariables.ShowDiagram.Show);
			ConsoleOut.writeLine(string.Format("Active Document:{0}: Master in document:{1}", appVisio.ActiveDocument, appVisio.ActiveDocument.Masters));

			// get the connectors for each shape in the diagram 
			Dictionary<int, ShapeInformation> shpConn = getShapeInformation(this.vDocument);

			try
			{
				if (this.vDocument != null)
				{
					this.vDocument.Application.ActiveDocument.Close();
				}
				if (appVisio != null)
				{
					appVisio.Quit();
				}
			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				ConsoleOut.writeLine(string.Format("This exception is OK because the user closed the Visio document before exiting the application:  {0}", ex.Message));
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
		private static Dictionary<int, ShapeInformation> getShapeInformation(Visio1.Document doc)
		{
			// Look at each shape in the collection.
			Visio1.Page page = doc.Pages[1];

			Dictionary<int, ShapeInformation> allPageShapesMap = null;
			ShapeInformation shpInfo = null;

			try
			{
				allPageShapesMap = new Dictionary<int, ShapeInformation>();
				string sColor = string.Empty;

				foreach (Visio1.Shape shape in page.Shapes)
				{
					// Use this index to look at each row in the properties section.
					shpInfo = new ShapeInformation();

					shpInfo.ID = shape.ID;
					shpInfo.UniqueKey = shape.NameU.Trim();

					// get shape fillForgnd and FillBkgnd colors
					//var fillForeColor = shape.Cells["FillForegnd"].ResultIU;
					//var fillBkColor = shape.Cells["FillBkgnd"].ResultIU;
					Color c = doc.Colors.Item16[(short)shape.Cells["FillBkgnd"].ResultIU];
					shpInfo.rgbFillColor = $"RGB({c.Red},{c.Green},{c.Blue})";
					sColor = VisioVariables.GetColorValueFromRGB(shpInfo.rgbFillColor);
					shpInfo.FillColor = sColor;
					if (string.IsNullOrEmpty(sColor))
					{
						shpInfo.FillColor = "";	// color not found so don't set it
					}

					//short iRow = (short)VisRowIndices.visRowFirst;
					shpInfo.Pos_x = Math.Truncate(shape.Cells["PinX"].ResultIU * 1000) / 1000;
					shpInfo.Pos_y = Math.Truncate(shape.Cells["PinY"].ResultIU * 1000) / 1000;
					if (shape.Name.IndexOf("Ethernet") > 0)
					{
						shpInfo.Width = Math.Truncate(shape.Cells["Width"].ResultIU * 1000) / 1000;
					}
					else
					{
						shpInfo.Width = 0;
						shpInfo.Height = 0;
					}

					string[] saStr = shape.NameU.Split(':');
					shpInfo.StencilImage = saStr[0].Trim();
					saStr = shape.Name.Split('.');
					if (saStr.Length > 1)
					{
						shpInfo.StencilImage = saStr[0].Trim();
					}
					else
					{
						shpInfo.StencilImage = shape.Name.Trim();
					}
					shpInfo.StencilLabel = shape.Text.Trim();

					if (shape.Style.ToUpper().IndexOf("CONNECTOR") >= 0)
					{
						// get connection information
						getShapeConnections2(doc, shape, ref allPageShapesMap, ref shpInfo);

						// we don't want to add this shape object to the allPageShapesMap dictionary
						continue;
					}

					// if shape is Ethernet type don't get the connections
					if (shpInfo.StencilImage.ToUpper().IndexOf("ETHERNET") <= 0)
					{
						// get shape connections
						// getShapeConnections(doc, shape, ref allPageShapesMap, ref shpInfo);
					}

					ConsoleOut.writeLine(string.Format("Stencil ID:{0} Key:{1}", shpInfo.ID, shpInfo.UniqueKey));
					if (!allPageShapesMap.ContainsKey(shape.ID))		// && !allPageShapesMap.ContainsKey(sKey2)) // cnnShape.ID
					{
						allPageShapesMap.Add(shape.ID, shpInfo);	// shape.ID
					}
				}
			}
			catch (Exception ex)
			{
				string sTmp = string.Format("ProcessVisioDiagramShapes::getShapeInformation - Exception\n\nForeach loop:\n{0}", ex.Message);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			return allPageShapesMap;
		}

		/// <summary>
		/// GetShapeConnections
		/// this will attempt to get the connection information between stencils using stencils on the Visio Diagram
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="shape"></param>
		/// <param name="allPageShapesMap"></param>
		/// <param name="shpInfo"></param>
		private static void getShapeConnections(Visio1.Document doc, Visio1.Shape shape, ref Dictionary<int, ShapeInformation> allPageShapesMap, ref ShapeInformation shpInfo)
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
						ConsoleOut.writeLine(string.Format("Connecting shapeID:{0}-{1} To shapeID:{2}-{3}in the diagram", shpInfo.ID, shpInfo.UniqueKey, lookupShape.ID, lookupShape.NameU));
						shpInfo.ConnectToID = lookupShape.ID;
						if (nCnt++ > 0)
						{
							shpInfo.ConnectTo += "," + lookupShape.NameU;
						}
						else
						{
							shpInfo.ConnectTo += lookupShape.NameU;
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
						var colorIdx = shape.CellsU["FillBkgnd"].ResultIU;
						var c = doc.Colors.Item16[(short)colorIdx];
						shpInfo.rgbFillColor = $"RGB({c.Red},{c.Green},{c.Blue})";
						sColor = VisioVariables.GetColorValueFromRGB(shpInfo.rgbFillColor);
						//shpInfo.ToLineColor = "";
						if (!string.IsNullOrEmpty(sColor))
						{
							shpInfo.ToLineColor = sColor;
						}


						// end new section

						//shpInfo.ToLineColor = "Black";
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
						ConsoleOut.writeLine(string.Format("Connecting shapeID:{0}-{1} From shapeID:{2}-{3}in the diagram", shpInfo.ID, shpInfo.UniqueKey, lookupShape.ID, lookupShape.NameU));
						shpInfo.ConnectFromID = lookupShape.ID;
						if (nCnt++ > 0)
						{
							shpInfo.ConnectFrom += "," + lookupShape.NameU;
						}
						else
						{
							shpInfo.ConnectFrom += lookupShape.NameU;
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
						var colorIdx = shape.CellsU["FillBkgnd"].ResultIU;
						var c = doc.Colors.Item16[(short)colorIdx];
						shpInfo.rgbFillColor = $"RGB({c.Red},{c.Green},{c.Blue})";
						sColor = VisioVariables.GetColorValueFromRGB(shpInfo.rgbFillColor);
						//shpInfo.FromLineColor = "";
						if (!string.IsNullOrEmpty(sColor))
						{
							shpInfo.FromLineColor = sColor;
						}
						// end new section

						//shpInfo.FromLineColor = "Black";
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
		/// <param name="shape"></param>
		/// <param name="allPageShapesMap"></param>
		/// <param name="shpInfo"></param>
		private static void getShapeConnections2(Visio1.Document doc, Visio1.Shape shape, ref Dictionary<int, ShapeInformation> allPageShapesMap, ref ShapeInformation shpInfo)
		{
			// what we need to do is get the shape and determine if shape is connected.
			// we don't want to save the shape if it is a connector

			// Look at each shape in the collection.
			//Visio1.Page page = doc.Pages[1];

			Dictionary<int, string> connectors = new Dictionary<int, string>();
			Dictionary<string, string> connectMap = new Dictionary<string, string>();

			//continue;
			string sTmp = string.Empty;
			string sTmp2 = string.Empty;
			string lineWeight = string.Empty;

			ShapeInformation lookupShapeMap = null;
			int lookupKey = 0;
			string arrowType = VisioVariables.sARROW_NONE;
			string lineColor = "Black";
			string rgbLineColor = VisioVariables.GetRGBColor("Black");
			double linePattern = VisioVariables.LINE_PATTERN_SOLID;
			Visio1.Connects visconnects2 = shape.Connects;

			for (int k = 1; k <= visconnects2.Count; k++)
			{
				// look through the connections to get the both ends
				Visio1.Connect visconnect = visconnects2[k];
				Visio1.Shape toshape = visconnect.ToSheet;
				if (k == 1)
				{
					// first end From
					lookupKey = toshape.ID;
					allPageShapesMap.TryGetValue(toshape.ID, out lookupShapeMap);

					sTmp = string.Empty;
					sTmp2 = string.Empty;
					sTmp = string.Format("Connector ID:{0} Shape ID:{1}-{2} LineLabel:{3}", shape.ID, toshape.ID, toshape.NameU, shape.Text);
					sTmp2 = string.Format("id:{0};name:{1};label:{2}", toshape.ID, toshape.Name, shape.Text);

					int startArrow = int.Parse(shape.get_CellsU("BeginArrow").FormulaU);
					int endArrow = int.Parse(shape.get_CellsU("EndArrow").FormulaU);
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
					rgbLineColor = shape.get_CellsU("LineColor").FormulaU;      // RGB color value
					lineColor = String.Empty;
					string sColor = VisioVariables.GetColorValueFromRGB(rgbLineColor);		// will be a color word or null if not found
					if (!string.IsNullOrEmpty(sColor))
					{
						lineColor = "Black";		// connector line color
					}

					linePattern = double.Parse(shape.get_CellsU("LinePattern").FormulaU);
					lineWeight = shape.get_CellsU("LineWeight").FormulaU;
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
								lookupShapeMap.ConnectFrom = toshape.NameU;
								lookupShapeMap.ConnectFromID = lookupKey;
							}
							lookupShapeMap.FromLineLabel = shape.Text;
							lookupShapeMap.FromArrowType = arrowType;
							lookupShapeMap.FromLineColor = lineColor;
							lookupShapeMap.FromLinePattern = linePattern;
						}
						lookupKey = toshape.ID; // keep in this order.  we use this for update the object
					}
					else
					{
						if (string.IsNullOrEmpty(lookupShapeMap.ConnectTo))
						{
							lookupShapeMap.ConnectTo = toshape.NameU;
							lookupShapeMap.ConnectToID = toshape.ID;
						}
						lookupShapeMap.ToLineLabel = shape.Text;  // use the Text value from the connector shape
						lookupShapeMap.ToArrowType = arrowType;
						lookupShapeMap.ToLineColor = lineColor;
						lookupShapeMap.ToLinePattern = linePattern;
					}

					sTmp += string.Format(" - {0} To Shape ID:{1}-{2} LineLabel:{3}", shape.ID, toshape.ID, toshape.NameU, shape.Text);
					sTmp2 += string.Format("|id:{0};name:{1};label:{2}", toshape.ID, toshape.Name, shape.Text);
				}
			}
			if (lookupShapeMap != null)
			{
				allPageShapesMap[lookupKey] = lookupShapeMap;
			}

			connectors.Add(shape.ID, sTmp2);
			ConsoleOut.writeLine(sTmp);
			ConsoleOut.writeLine(string.Format("Found shape ID:{0}-{1} in the diagram", shpInfo.ID, shpInfo.UniqueKey));
		}
	}
}
