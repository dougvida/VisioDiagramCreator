using System;
using System.Collections.Generic;
using Visio1 = Microsoft.Office.Interop.Visio;
using Microsoft.Office.Interop.Visio;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Linq;
using VisioDiagramCreator.Visio.Models;
using System.Drawing.Design;
using System.Text;
using System.Collections;
using VisioAutomation.VDX.Elements;

namespace VisioDiagramCreator.Visio
{
	public class ProcessVisioDiagramShapes
	{
		/// <summary>
		/// GetAllShapesProperties
		/// Get all the shape properties for each shape contained within the document
		/// </summary>
		/// <param name="diagamFilePathName"></param>
		/// <return>Dictionary<int, ShapeInformation> </return>
		public Dictionary<int, ShapeInformation> GetAllShapesProperties(string diagamFilePathName, VisioVariables.ShowDiagram dspMode)
		{
			// Open up one of Visio's sample drawings.
			Visio1.Application appVisio = new Visio1.Application();
			new VisioHelper().ShowVisioDiagram(appVisio, dspMode);              // don't show the diagram

			Visio1.Document vDocuments = appVisio.Documents.Open(diagamFilePathName);
			// The new document will have one page, get the a reference to it.
			Visio1.Page page = vDocuments.Pages[1];

			new VisioHelper().ShowVisioDiagram(appVisio, VisioVariables.ShowDiagram.Show);
			Console.WriteLine("Active Document:{0}: Master in document:{1}", appVisio.ActiveDocument, appVisio.ActiveDocument.Masters);

			// get the connectors for each shape in the diagram 
			return getShapeConnections(vDocuments);
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
			Dictionary<string, string> connectMap = null;

			ShapeInformation shpInfo = null;
			try
			{
				connectMap = new Dictionary<string, string>();
				allPageShapesMap = new Dictionary<int, ShapeInformation>();

				bool bFound = false;

				foreach (Visio1.Shape shape in page.Shapes)
				{
					bFound = false;

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
						//Console.WriteLine(string.Format("Skip this ID:{0}; shapeKey:{1} - Stencil Image:{2}", shape.ID, shpInfo.UniqueKey, shpInfo.StencilImage));
						continue;
					}
					if (shape.Style.Equals("Connector"))
					{
						continue;
					}

					if (shpInfo.ID == 21)
					{
						int x = 0;
					}
					// get connections To
					var shpConnection = shape.ConnectedShapes(VisConnectedShapesFlags.visConnectedShapesOutgoingNodes, "");
					if (shpConnection != null && shpConnection.Length > 0)
					{
						Console.WriteLine("");
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
								bFound = true;
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
						Console.WriteLine("");
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
								bFound = true;
							}
						}
						catch (Exception exp)
						{
							MessageBox.Show(String.Format("getShapeConnections - Connection:{0} - {1}", shpInfo.ConnectFrom, exp.Message));
						}
					}

					// we only want to save the shape object once for all the connections
					if (bFound)
					{
						if (!allPageShapesMap.ContainsKey(shape.ID)) // && !allPageShapesMap.ContainsKey(sKey2)) // cnnShape.ID
						{
							//dConnectMap.Add(sKey, cnnShape.NameU);		// cnnShape.ID
							allPageShapesMap.Add(shape.ID, shpInfo);   // shape.ID
						}
					}
				}
			}
			catch(Exception ex)
			{
				MessageBox.Show(string.Format("Exception::getShapeConnections - Foreach loop:\n{0}", ex.Message));
			}
			return allPageShapesMap;
		}		
	} 
}
