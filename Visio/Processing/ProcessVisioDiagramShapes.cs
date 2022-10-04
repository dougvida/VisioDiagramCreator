using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using OmnicellBlueprintingTool.Models;
using Visio1 = Microsoft.Office.Interop.Visio;

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
			Dictionary<int, ShapeInformation> shpConn = getShapeConnections(this.vDocument);

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
		/// getShapeConnections
		/// get the visio diagram stencil connection information
		/// used to add to the Excel Data file
		/// </summary>
		/// <param name="doc">Visio document</param>
		/// <returns>"Dictionary<string, ShapeInformation>"</returns>
		private static Dictionary<int, ShapeInformation> getShapeConnections(Visio1.Document doc)
		{
			// Look at each shape in the collection.
			Visio1.Page page = doc.Pages[1];

			Dictionary<int, ShapeInformation> allPageShapesMap = null;
			Dictionary<int, string> connectors = null;
			Dictionary<string, string> connectMap = null;
			ShapeInformation shpInfo = null;

			try
			{
				connectMap = new Dictionary<string, string>();
				allPageShapesMap = new Dictionary<int, ShapeInformation>();

				connectors = new Dictionary<int, string>();
				foreach (Visio1.Shape shape in page.Shapes)
				{
					// Use this index to look at each row in the properties section.
					shpInfo = new ShapeInformation();

					shpInfo.ID = shape.ID;
					shpInfo.UniqueKey = shape.NameU.Trim();

					//short iRow = (short)VisRowIndices.visRowFirst;
					shpInfo.Pos_x = Math.Truncate(shape.Cells["PinX"].ResultIU * 1000) / 1000;
					shpInfo.Pos_y = Math.Truncate(shape.Cells["PinY"].ResultIU * 1000) / 1000;
					shpInfo.Width = Math.Truncate(shape.Cells["Width"].ResultIU * 1000) / 1000;
					shpInfo.Height = Math.Truncate(shape.Cells["Height"].ResultIU * 1000) / 1000;

					string[] saStr = shape.NameU.Split(':');
					shpInfo.StencilImage = saStr[0].Trim();
					shpInfo.StencilLabel = shape.Text.Trim();

					if (shpInfo.StencilImage.ToUpper().IndexOf("NETWORKPIPE") >= 0)
					{
						// skip this shape
						ConsoleOut.writeLine(string.Format("Skip this ID:{0}; shapeKey:{1} - Stencil Image:{2}", shape.ID, shpInfo.UniqueKey, shpInfo.StencilImage));
						continue;
					}
					if (shape.Style.Equals("Connector"))
					{
						ShapeInformation lookupShapeMap = null;
						Visio1.Connects visconnects2 = shape.Connects;
						int lookupKey = 0;
						string sTmp = string.Empty;
						string sTmp2 = string.Empty;
						string arrowType = VisioVariables.sARROW_NONE;
						string lineColor = VisioVariables.COLOR_BLACK;
						double linePattern = VisioVariables.LINE_PATTERN_SOLID;
						string lineWeight = string.Empty;

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
								sTmp = string.Format("Connector ID:{0} Shape ID:{1}-{2} LineLabel:{3}", shape.ID, toshape.ID, toshape.Name, shape.Text);
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
								lineColor = shape.get_CellsU("LineColor").FormulaU;
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
											lookupShapeMap.ConnectFrom = toshape.Name;
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
										lookupShapeMap.ConnectTo = toshape.Name;
										lookupShapeMap.ConnectToID = toshape.ID;
									}
									lookupShapeMap.ToLineLabel = shape.Text;  // use the Text value from the connector shape
									lookupShapeMap.ToArrowType = arrowType;
									lookupShapeMap.ToLineColor = lineColor;
									lookupShapeMap.ToLinePattern = linePattern;
								}

								sTmp += string.Format(" - {0} To Shape ID:{1}-{2} LineLabel:{3}", shape.ID, toshape.ID, toshape.Name, shape.Text);
								sTmp2 += string.Format("|id:{0};name:{1};label:{2}", toshape.ID, toshape.Name, shape.Text);
							}
						}
						if (lookupShapeMap != null)
						{
							allPageShapesMap[lookupKey] = lookupShapeMap;
						}

						connectors.Add(shape.ID, sTmp2);
						ConsoleOut.writeLine(sTmp);
						continue;
					}

					ConsoleOut.writeLine(string.Format("Found shape ID:{0}-{1} in the diagram", shpInfo.ID, shpInfo.UniqueKey));

					// get connections To
					var shpConnection = shape.ConnectedShapes(VisConnectedShapesFlags.visConnectedShapesOutgoingNodes, "");
					if (shpConnection != null && shpConnection.Length > 0)
					{
						try
						{
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
								shpInfo.ConnectToID = lookupShape.ID;
								if (nCnt++ > 0)
								{
									shpInfo.ConnectTo += "," + lookupShape.NameU;
								}
								else
								{
									shpInfo.ConnectTo += lookupShape.NameU;
								}
								shpInfo.ToLineColor = "BLACK";
								shpInfo.ToLinePattern = VisioVariables.LINE_PATTERN_SOLID;
								connectMap.Add(sKey, sKey2);
							}
						}
						catch (Exception exp)
						{
							MessageBox.Show(String.Format("getShapeConnections - Connection:{0} - {1}", shpInfo.ConnectFrom, exp.Message));
						}
					}
					// get connections from
					shpConnection = shape.ConnectedShapes(VisConnectedShapesFlags.visConnectedShapesAllNodes, "");
					if (shpConnection != null && shpConnection.Length > 0)
					{
						try
						{
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
								shpInfo.ConnectFromID = lookupShape.ID;
								if (nCnt++ > 0)
								{
									shpInfo.ConnectFrom += "," + lookupShape.NameU;
								}
								else
								{
									shpInfo.ConnectFrom += lookupShape.NameU;
								}
								shpInfo.FromLineColor = "BLACK";
								shpInfo.FromLinePattern = VisioVariables.LINE_PATTERN_SOLID;
								connectMap.Add(sKey, sKey2);
							}
						}
						catch (Exception exp)
						{
							MessageBox.Show(String.Format("getShapeConnections - Connection:{0} - {1}", shpInfo.ConnectFrom, exp.Message));
						}
					}
					if (!allPageShapesMap.ContainsKey(shape.ID)) // && !allPageShapesMap.ContainsKey(sKey2)) // cnnShape.ID
					{
						allPageShapesMap.Add(shape.ID, shpInfo);   // shape.ID
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Exception::getShapeConnections - Foreach loop:\n{0}", ex.Message));
			}
			return allPageShapesMap;
		}
	}
}
