﻿using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using VisioDiagramCreator.Models;
using Visio1 = Microsoft.Office.Interop.Visio;

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
			ConsoleOut.writeLine(string.Format("Active Document:{0}: Master in document:{1}", appVisio.ActiveDocument, appVisio.ActiveDocument.Masters));

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
						ConsoleOut.writeLine(string.Format("Skip this ID:{0}; shapeKey:{1} - Stencil Image:{2}", shape.ID, shpInfo.UniqueKey, shpInfo.StencilImage));
						continue;
					}
					if (shape.Style.Equals("Connector"))
					{
						Visio1.Connects visconnects2 = shape.Connects;
						ConsoleOut.writeLine(string.Format("Connector ID:{0}-{1}",shape.ID, shape.Name));
						for (int k = 1; k <= visconnects2.Count; k++)
						{
							Visio1.Connect visconnect = visconnects2[k];
							Visio1.Shape toshape = visconnect.ToSheet;
							ConsoleOut.writeLine(string.Format("To Shape ID:{0}-{1} LineLabel:{2}",toshape.ID, toshape.Name, shape.Text));
						}
						//ConsoleOut.writeLine(string.Format("Connector ID:{0} - {1}.  connect To/From {2}", shape.ID, shape.Text, ID));
						continue;
					}

					ConsoleOut.writeLine(string.Format("shape ID:{0}-{1}",shpInfo.ID, shpInfo.UniqueKey));

					if (shpInfo.ID == 21)
					{
						int x = 0;
					}

					Visio1.Connects visconnects = shape.Connects;
					ConsoleOut.writeLine(string.Format("Connector ID:{0}-{1}", shape.ID, shape.Name));
					for (int k = 1; k <= visconnects.Count; k++)
					{
						Visio1.Connect visconnect = visconnects[k];
						Visio1.Shape toshape = visconnect.ToSheet;
						ConsoleOut.writeLine(string.Format("To Shape ID:{0}-{1} LineLabel:{2}", toshape.ID, toshape.Name, shape.Text));
					}
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


		private bool ShapesAreConnected(Visio1.Shape shape1, Visio1.Shape shape2)
		{
			// in Visio our 2 shapes will each be connected to a connector, not to each other
			// so we need to see if theyare both connected to the same connector

			bool Connected = false;
			// since we are pinning the connector to each shape, we only need to check
			// the fromshapes attribute on each shape
			Visio1.Connects shape1FromConnects = shape1.FromConnects;
			Visio1.Connects shape2FromConnects = shape2.FromConnects;

			foreach (Visio1.Shape connect in shape1FromConnects)
			{
				// first method
				// for each shape shape 1 is connected to, see if shape2 is connected 
				var shape = from Visio1.Shape cs in shape2FromConnects where cs == connect select cs;
				if (shape.FirstOrDefault() != null) Connected = true;

				// second method, convert shape2's connected shapes to an IEnumerable and
				// see if it contains any shape1 shapes  
				IEnumerable<Visio1.Shape> shapesasie = (IEnumerable<Visio1.Shape>)shape2FromConnects;

				if (shapesasie.Contains(connect))
				{
					return true;
				}
			}

			return Connected;

			//third method
			//convert both to IEnumerable and check if they intersect
			IEnumerable<Visio1.Shape> shapes1asie = (IEnumerable<Visio1.Shape>)shape1FromConnects;
			IEnumerable<Visio1.Shape> shapes2asie = (IEnumerable<Visio1.Shape>)shape2FromConnects;
			var shapes = shapes1asie.Intersect(shapes2asie);
			if (shapes.Count() > 0) 
			{ 
				return true; 
			}
			else 
			{ 
				return false; 
			}
		}
	}
}
