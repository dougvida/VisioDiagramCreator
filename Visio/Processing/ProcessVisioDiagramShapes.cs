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
		public Dictionary<int, ShapeInformation> GetAllShapesProperties(VisioHelper visioHelper, string diagamFilePathName, VisioVariables.ShowDiagram dspMode)
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
			Dictionary<int, ShapeInformation> shpConn = getShapeInformation(visioHelper, this.vDocument);

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
		private static Dictionary<int, ShapeInformation> getShapeInformation(VisioHelper visioHelper, Visio1.Document doc)
		{
			// Look at each shape in the collection.
			Visio1.Page page = doc.Pages[1];

			Dictionary<int, ShapeInformation> allPageShapesMap = null;
			Dictionary<int, Visio1.Shape> connectorsMap = new Dictionary<int, Visio1.Shape>();
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
					shpInfo.UniqueKey = getShapeUniqueKeyName(shape);

					// get shape fillForgnd and FillBkgnd colors
					//var fillForeColor = shape.Cells["FillForegnd"].ResultIU;
					//var fillBkColor = shape.Cells["FillBkgnd"].ResultIU;
					Microsoft.Office.Interop.Visio.Color c =  doc.Colors.Item16[(short)shape.Cells["FillBkgnd"].ResultIU];
					shpInfo.rgbFillColor = $"RGB({c.Red},{c.Green},{c.Blue})";
					sColor = visioHelper.GetColorValueFromRGB(shpInfo.rgbFillColor);
					if (string.IsNullOrEmpty(sColor))
					{
						// no color found lets try to find the best match
						sColor = getColorNameFromRGB(visioHelper, c.Red, c.Green, c.Blue);
					}
					shpInfo.FillColor = "";
					if (!string.IsNullOrEmpty(sColor))
					{
						shpInfo.FillColor = sColor;
					}
					if (shpInfo.rgbFillColor.IndexOf(VisioVariables.RGB_COLOR_2SKIP) >= 0)  // found
					{
						shpInfo.rgbFillColor = "";	// we don't want to right this color value to the Excel file
					}

					//short iRow = (short)VisRowIndices.visRowFirst;
					shpInfo.Pos_x = Math.Truncate(shape.Cells["PinX"].ResultIU * 1000) / 1000;
					shpInfo.Pos_y = Math.Truncate(shape.Cells["PinY"].ResultIU * 1000) / 1000;
					if (shape.Name.IndexOf("OC_Ethernet", StringComparison.OrdinalIgnoreCase) >= 0 || 
						shape.Name.IndexOf("OC_Group", StringComparison.OrdinalIgnoreCase) >= 0 ||
						shape.Name.IndexOf("OC_Footer", StringComparison.OrdinalIgnoreCase) >= 0 ||
						shape.Name.IndexOf("OC_Dash", StringComparison.OrdinalIgnoreCase) >= 0)
					{
						shpInfo.Width = Math.Truncate(shape.Cells["Width"].ResultIU * 1000) / 1000;
						if ((shape.Name.IndexOf("OC_Ethernet", StringComparison.OrdinalIgnoreCase) >= 0))
						{
							// don't set height if shape is Ethernet type
							shpInfo.Height = 0;
						}
						else
						{
							double dHeight = Math.Truncate(shape.Cells["Height"].ResultIU * 1000) / 1000;
							if (shape.Name.IndexOf("OC_Footer", StringComparison.OrdinalIgnoreCase) >= 0)
							{
								dHeight = dHeight - 0.25;
								shpInfo.Width = 0;		// we don't want to save the width because it's already a page with in size
							}
							shpInfo.Height = dHeight;
						}
					}
					else
					{
						shpInfo.Width = 0;
						shpInfo.Height = 0;
					}

					// this will return a string name:ID
					string shpName = getShapeUniqueKeyName(shape);
					string[] saStr = shpName.Split(':');		// need to just get the Name
					shpInfo.StencilImage = saStr[0].Trim();

					shpInfo.StencilLabel = shape.Text.Trim();

					// use the connection shape to obtain what is connected to what
					if (shape.Style.Trim().IndexOf("Connector", StringComparison.OrdinalIgnoreCase) >= 0)
					{
						// get connection information add to the dictionary to be used later
						connectorsMap.Add(shape.ID, shape);

						// we don't want to add this shape object to the allPageShapesMap dictionary
						continue;
					}

					ConsoleOut.writeLine(string.Format("Stencil ID:{0} Key:{1}", shpInfo.ID, shpInfo.UniqueKey));
					if (!allPageShapesMap.ContainsKey(shape.ID))
					{
						allPageShapesMap.Add(shape.ID, shpInfo);	// shape.ID
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
			catch (Exception ex)
			{
				string sTmp = string.Format("ProcessVisioDiagramShapes::getShapeInformation - Exception\n\nForeach loop:\nWorking on {0}\n\n{1}", shpInfo.UniqueKey, ex.Message);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			return allPageShapesMap;
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
		private static string getShapeName(Visio1.Shape shape)
		{
			// this shouuld never happen
			if (shape == null)
			{
				return null;
			}
			string[] saTmp = shape.Name.Split('.');
			if (saTmp.Length > 0)
			{
				// lets get the first part as the shape name
				return saTmp[0].Trim().Trim();
			}
			saTmp = shape.Name.Split(':');
			if (saTmp.Length > 0)
			{
				// lets get the first part as the shape name
				return saTmp[0].Trim().Trim();
			}
			return string.Format("{0}", shape.Name.Trim());
		}
		/// <summary>
		/// getShapeConnections
		/// the shape passed in is a connection shape
		/// from the connector shape get the connector information (color, text, weight, etc.)
		/// from the connector shape get the toSheet that the connector is connected to (should be a stencil shape already in the AllPageShapesMap)
		/// 
		/// </summary>
		/// <param name="visioHelper"></param>
		/// <param name="doc">document object</param>
		/// <param name="connShp"> connection shape object</param>
		/// <param name="allPageShapesMap">dictionary that contains all the shapes in the diagram</param>
		/// <param name="shpInfo">dictionary to hold the modified shapes to be returned</param>
		private static void getShapeConnections(VisioHelper visioHelper, Visio1.Document doc, Visio1.Shape connShp, ref Dictionary<int, ShapeInformation> allPageShapesMap, ref ShapeInformation shpInfo)
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
			int lookupKey = 0;
			string connectorLabel = string.Empty;
			string arrowType = VisioVariables.sARROW_NONE;
			string lineColor = VisioVariables.sCOLOR_BLACK;
			string rgbLineColor = visioHelper.GetRGBColor(VisioVariables.sCOLOR_BLACK);
			double linePattern = VisioVariables.LINE_PATTERN_SOLID;
			Visio1.Connects visconnects2 = connShp.Connects;

			// get the connector information
			connectorLabel = connShp.Text;

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
			shpInfo.rgbFillColor = $"RGB({c.Red},{c.Green},{c.Blue})";
			string sColor = visioHelper.GetColorValueFromRGB(shpInfo.rgbFillColor);
			if (string.IsNullOrEmpty(sColor))
			{
				// no color found lets try to find the best match
				sColor = getColorNameFromRGB(visioHelper, c.Red, c.Green, c.Blue);
			}
			shpInfo.ToLineColor = "";
			if (!string.IsNullOrEmpty(sColor))
			{
				lineColor = sColor;
			}

			linePattern = connShp.get_CellsU("LinePattern").ResultIU;
			
			lineWeight = connShp.get_CellsU("LineWeight").FormulaU;
			if (lineWeight.IndexOf("THERM", StringComparison.OrdinalIgnoreCase) >= 0)
			{
				lineWeight = VisioVariables.sLINE_WEIGHT_1;
			}
			else
			{
				// we have a valid value so lets see if we support it
				lineWeight = visioHelper.FindConnectorLineWeight(lineWeight);
				if (string.IsNullOrEmpty(lineWeight))
				{
					lineWeight = VisioVariables.sLINE_WEIGHT_1;
				}
			}

			int nFromCnt = 0;
			int nToCnt = 0;
			int ethernetID = 0;
			string ethernetUniqueKey = string.Empty;

			for (int k = 1; k <= visconnects2.Count; k++)
			{
				// look through the connections to get the both ends
				visconnect = visconnects2[k];
				toshape = visconnect.ToSheet;

				if (k == 1)
				{
					// first end From
					sTmp = string.Empty;
					lookupKey = toshape.ID;
					allPageShapesMap.TryGetValue(toshape.ID, out lookupShapeMap);

					// if this first one is the Ethernet stencil we need to do some special work
					if (toshape.Name.IndexOf("OC_Ethernet") >= 0)
					{
						ethernetID = toshape.ID; // save this we need to do a trick later
						ethernetUniqueKey = getShapeUniqueKeyName(toshape);
						sTmp = string.Format("From Shape:{0} ", ethernetUniqueKey);
					}
					else
					{
						sTmp = string.Format("Connect Shape:{0} ", ethernetUniqueKey);
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
							allPageShapesMap.TryGetValue(ethernetID, out lookupShapeMap);
							ethernetID = 0;
						}
						else
						{
							allPageShapesMap.TryGetValue(toshape.ID, out lookupShapeMap);
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
								lookupShapeMap.ConnectFromID = lookupKey;
							}
							lookupShapeMap.FromLineLabel = connectorLabel;
							lookupShapeMap.FromArrowType = arrowType;
							lookupShapeMap.FromLineColor = lineColor;
							lookupShapeMap.FromLinePattern = linePattern;
							lookupShapeMap.FromLineWeight = lineWeight;
						}
						lookupKey = toshape.ID; // keep in this order.  we use this for update the object
					}
					else
					{
						// if the first connector was Ethernet we need to make this a From type
						// this is when we need to do some special work
						if (!string.IsNullOrEmpty(ethernetUniqueKey))
						{
							// get the next shape to be populated using From..
							allPageShapesMap.TryGetValue(toshape.ID, out lookupShapeMap);	// get the next connection shape
							if (string.IsNullOrEmpty(lookupShapeMap.ConnectFrom))
							{
								if (nToCnt++ > 0)
								{
									lookupShapeMap.ConnectFrom += "," + ethernetUniqueKey;	// append to existing value
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

							lookupKey = toshape.ID;		// use the ID for this shape not the Ethernet shape

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
					sTmp += string.Format("to Shape:{0}, linelabel:'{1}'", getShapeUniqueKeyName(toshape), connShp.Text);
				}
			}
			if (lookupShapeMap != null)
			{
				allPageShapesMap[lookupKey] = lookupShapeMap;
			}
			ConsoleOut.writeLine(sTmp);
		}

		/// <summary>
		/// getColorNameFromRGB
		/// Search if the RGB value is a color match to what the application supports.
		/// Return the closest color found based on the RGB value
		/// No match will return an empty string
		/// </summary>
		/// <param name="visHlper">VisioHelper</param>
		/// <param name="red">int</param>
		/// <param name="green">int</param>
		/// <param name="blue">int</param>
		/// <returns>color string if found otherwise empty string</returns>
		private static string getColorNameFromRGB(VisioHelper visHlper, int red, int green, int blue)
		{
			Color lookupColor = Color.FromArgb(255, red, green, blue);
			ConsoleOut.writeLine(lookupColor.Name);

			Dictionary<string,Color> appColorsMap = visHlper.GetColorNameColorsMap();
			List<string> matches = new List<string>();
			foreach (KeyValuePair<string,Color> appColor in appColorsMap)
			{
				// want to start narrow than widen to find the closest color to the app's supporrted color
				for (int threshold = 5; threshold <= 75; threshold += 10)
				{
					if (IsColorsClose(lookupColor, appColor.Value, threshold))
					{
						return appColor.Key;	// color match found
					}
				}
			}
			return "";	// color match not found
		}

		/// <summary>
		/// IsColorsClose
		/// this will compare two rgb color
		/// the color from the Visio diagram will be check if exists or a close respersentation supported by the application default colors (best match)
		/// the threshold is used to widen the search.  most cases an exact match will not occur so we need to widen the R,G,B values
		/// </summary>
		/// <param name="a">Color to look up</param>
		/// <param name="z">Color supported by application</param>
		/// <param name="threshold">default is 75</param>
		/// <returns>bool - true found else false</returns>
		private static bool IsColorsClose(Color a, Color z, int threshold = 75)
		{
			int r = (int)a.R - z.R,
					g = (int)a.G - z.G,
					b = (int)a.B - z.B;
			return (r * r + g * g + b * b) <= threshold * threshold;
		}



		/// <summary>
		/// GetShapeConnections
		/// this will attempt to get the connection information between stencils using stencils on the Visio Diagram
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
						ConsoleOut.writeLine(string.Format("Connecting shapeID:{0}-{1} To shapeID:{2}-{3} in the diagram", shpInfo.ID, shpInfo.UniqueKey, lookupShape.ID, lookupShape.NameU));
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
						shpInfo.rgbFillColor = $"RGB({c.Red},{c.Green},{c.Blue})";
						sColor = visioHelper.GetColorValueFromRGB(shpInfo.rgbFillColor);

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
						ConsoleOut.writeLine(string.Format("Connecting shapeID:{0}-{1} From shapeID:{2}-{3} in the diagram", shpInfo.ID, shpInfo.UniqueKey, lookupShape.ID, lookupShape.NameU));
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
						shpInfo.rgbFillColor = $"RGB({c.Red},{c.Green},{c.Blue})";
						sColor = visioHelper.GetColorValueFromRGB(shpInfo.rgbFillColor);
						if (string.IsNullOrEmpty(sColor))
						{
							// no color found lets try to find the best match
							sColor = getColorNameFromRGB(visioHelper, c.Red, c.Green, c.Blue);
						}
						shpInfo.FromLineColor = "";
						if (!string.IsNullOrEmpty(sColor))
						{
							shpInfo.FromLineColor = sColor;
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
					sTmp = string.Format("Connector ID:'{0}' Shape ID:'{1}'-'{2}' LineLabel:'{3}'", connShp.ID, toshape.ID, getShapeUniqueKeyName(toshape), connShp.Text);
					sTmp2 = string.Format("id:'{0}';name:'{1}';label:'{2}'", toshape.ID, toshape.Name, connShp.Text);

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
					shpInfo.rgbFillColor = $"RGB({c.Red},{c.Green},{c.Blue})";
					string sColor = visioHelper.GetColorValueFromRGB(shpInfo.rgbFillColor);
					if (string.IsNullOrEmpty(sColor))
					{
						// no color found lets try to find the best match
						sColor = getColorNameFromRGB(visioHelper, c.Red, c.Green, c.Blue);
					}
					shpInfo.ToLineColor = "";
					if (!string.IsNullOrEmpty(sColor))
					{
						lineColor = sColor;
					}

					//rgbLineColor = connShp.get_CellsU("LineColor").FormulaU;      // RGB color value
					//if (rgbLineColor.IndexOf("THEME") >= 0)
					//{
					//	// we need to parse out the RGB value
					//	int nStart = rgbLineColor.IndexOf("RGB");
					//	rgbLineColor = rgbLineColor.Substring(nStart, (rgbLineColor.Length - nStart - 1));
					//
					//	//Color c = doc.Colors.Item16[(short)rgbLineColor];
					//	//shpInfo.rgbFillColor = $"RGB({c.Red},{c.Green},{c.Blue})";
					//
					//}
					//lineColor = String.Empty;
					//sColor = visioHelper.GetColorValueFromRGB(rgbLineColor);		// will be a color word or null if not found
					//if (string.IsNullOrEmpty(sColor))
					//{
					//	lineColor = VisioVariables.sCOLOR_BLACK;		// connector line color
					//}

					linePattern = connShp.get_CellsU("LinePattern").ResultIU;
					lineWeight = connShp.get_CellsU("LineWeight").FormulaU;
					if (lineWeight.IndexOf("THERM", StringComparison.OrdinalIgnoreCase) >= 0)
					{
						lineWeight = VisioVariables.sLINE_WEIGHT_1;
					}
					else
					{
						// we have a valid value so lets see if we support it
						lineWeight = visioHelper.FindConnectorLineWeight(lineWeight);
						if (string.IsNullOrEmpty(lineWeight))
						{
							lineWeight = VisioVariables.sLINE_WEIGHT_1;
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

					sTmp += string.Format(" - '{0}' To Shape ID:'{1}'-'{2}' LineLabel:'{3}'", connShp.ID, toshape.ID, getShapeUniqueKeyName(toshape), connShp.Text);
					sTmp2 += string.Format("|id:'{0}';name:'{1}';label:'{2}'", toshape.ID, toshape.Name, connShp.Text);
				}
			}
			if (lookupShapeMap != null)
			{
				allPageShapesMap[lookupKey] = lookupShapeMap;
			}

			connectorsMap.Add(connShp.ID, sTmp2);
			ConsoleOut.writeLine(sTmp);
			ConsoleOut.writeLine(string.Format("Found shape ID:'{0}'-'{1}' in the diagram", shpInfo.ID, shpInfo.UniqueKey));
		}


	}
}
