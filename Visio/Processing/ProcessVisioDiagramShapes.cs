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

		private static Dictionary<int, ShapeInformation> getShapeConnections(Visio1.Document doc)
		{
			// Look at each shape in the collection.
			Visio1.Page page = doc.Pages[1];

			Dictionary<int, ShapeInformation> AllPageShapesMap = new Dictionary<int, ShapeInformation>();
			Dictionary<int, ConnectShpExcelData> dExcelConnectShapeMap = new Dictionary<int, ConnectShpExcelData>();

			Dictionary<int, string> dConnectTo = null;
			Dictionary<int, string> dConnectFrom = null;
			HashSet<int> connFromHS = null;
			HashSet<int> connToHS = null;

			ShapeInformation shpInfo = null;
			try
			{
				ConnectShpExcelData cntShpData = null;

				foreach (Visio1.Shape shape in page.Shapes)
				{
					// Use this index to look at each row in the properties section.
					shpInfo = new ShapeInformation();
					cntShpData = new ConnectShpExcelData();

					dConnectTo = new Dictionary<int, string>();
					dConnectFrom = new Dictionary<int, string>();
					connFromHS = new HashSet<int>();
					connToHS = new HashSet<int>();

					shpInfo.ID = shape.ID;
					shpInfo.UniqueKey = shape.NameU.Trim();

					//short iRow = (short)VisRowIndices.visRowFirst;
					shpInfo.Pos_x = Math.Truncate(shape.Cells["PinX"].ResultIU * 1000) / 1000;
					shpInfo.Pos_y = Math.Truncate(shape.Cells["PinY"].ResultIU * 1000) / 1000;
					shpInfo.Width = Math.Truncate(shape.Cells["Width"].ResultIU * 1000) / 1000;
					shpInfo.Height = Math.Truncate(shape.Cells["Height"].ResultIU * 1000) / 1000;

					string sConnFrom = string.Empty;
					string sConnTo = string.Empty;

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
						int x = 0;
						continue;
					}

					var shpFromConnected = shape.ConnectedShapes(VisConnectedShapesFlags.visConnectedShapesIncomingNodes, "");
					if (shpFromConnected != null && shpFromConnected.Length > 0)
					{
						Console.WriteLine("");
						try
						{
							sConnFrom = string.Empty;
							int Id = 0;
							foreach (int nIdx in shpFromConnected)
							{
								dConnectFrom = new Dictionary<int, string>();
					
								Id = nIdx;
								Visio1.Shape cnnShape = page.Shapes.ItemFromID[nIdx];
								shpInfo.ConnectFromID = cnnShape.ID;
								shpInfo.ConnectFrom = cnnShape.NameU;
								shpInfo.FromLineColor = "BLACK";
								shpInfo.FromLinePattern = VisioVariables.LINE_PATTERN_SOLID;
					
								cntShpData.ID = shpInfo.ID;
								cntShpData.UniqueKey = shpInfo.UniqueKey;
					
								connFromHS.Add(cnnShape.ID);
					
								if (!dConnectFrom.ContainsKey(cnnShape.ID))
								{
									dConnectFrom.Add(cnnShape.ID, cnnShape.NameU);
									if (dConnectFrom.Count != dConnectFrom.Distinct().Count())
									{
										Console.WriteLine("Contains duplicates");
									}
								}
								//bool keyExists = AllPageShapesMap.ContainsKey(shape.ID);
								//if (!keyExists)
								//{
									AllPageShapesMap.Add(shape.ID, shpInfo);
								//}
								//else
								//{
								//	ShapeInformation value = new ShapeInformation();
								//	AllPageShapesMap.TryGetValue(shape.ID, out value);
								//	if (string.IsNullOrEmpty(value.ConnectFrom))
								//	{
								//		value.ConnectFromID = shape.ID;
								//		value.ConnectFrom = cnnShape.NameU;
								//		AllPageShapesMap[shape.ID] = value;	// update the value
								//	}
								//}
							}
						}
						catch (Exception exp)
						{
							MessageBox.Show(String.Format("ConnectFrom:{0} - {1}", shpInfo.ConnectFrom, exp.Message));
						}
					}

					var shpToConnected = shape.ConnectedShapes(VisConnectedShapesFlags.visConnectedShapesOutgoingNodes, "");
					if (shpToConnected != null && shpToConnected.Length > 0)
					{
						try
						{
							sConnTo = string.Empty;
							int Id = 0;
							foreach (int nIdx in shpToConnected)
							{
								dConnectTo = new Dictionary<int, string>();

								Id = nIdx;
								Visio1.Shape cnnShape = page.Shapes.ItemFromID[nIdx];

								//Console.WriteLine("Shapes that are To / Outgoing connections: {0}", sp.NameU);
								shpInfo.ConnectToID = cnnShape.ID;
								shpInfo.ConnectTo = cnnShape.NameU;
								shpInfo.ToLineColor = "BLACK";
								shpInfo.ToLinePattern = VisioVariables.LINE_PATTERN_SOLID;

								cntShpData.ID = shpInfo.ID;
								cntShpData.UniqueKey = shpInfo.UniqueKey;

								connToHS.Add(cnnShape.ID);

								if (!dConnectTo.ContainsKey(cnnShape.ID))
								{
									dConnectTo.Add(cnnShape.ID, cnnShape.NameU);
									if (dConnectTo.Count != dConnectTo.Distinct().Count())
									{
										Console.WriteLine("Contains duplicates");
									}
								}
								//bool keyExists = AllPageShapesMap.ContainsKey(shape.ID);
								//if (!keyExists)
								//{
									AllPageShapesMap.Add(shape.ID, shpInfo);
								//}
								//else
								//{
								//	ShapeInformation value = new ShapeInformation();
								//	AllPageShapesMap.TryGetValue(shape.ID, out value);
								//	if (string.IsNullOrEmpty(value.ConnectTo))
								//	{
								//		value.ConnectFromID = shape.ID;
								//		value.ConnectTo = cnnShape.NameU;
								//		AllPageShapesMap[shape.ID] = value; // update the value
								//	}
								//}
							}
						}
						catch (Exception exp)
						{
							MessageBox.Show(string.Format("ConnectTo:{0} - {1}", shpInfo.ConnectTo, exp.Message));
						}
					}
					else
					{
						if (!"Connector".Equals(shape.Style))
						{
							// Print the results.
							//	Console.WriteLine(string.Format("Shape:{0}; ShapKey:{1}; StencilImage:{2}; ShapeLabel:{3}; IP; Ports; Devices; PosX:{4}; PoxY:{5}; Width:{6}; Height:{7}",
							//												ShapeType, shpInfo.UniqueKey, shpInfo.StencilImage, shpInfo.StencilLabel, shpInfo.Pos_x, shpInfo.Pos_y, shpInfo.Width, shpInfo.Height));
						}
					}

					//if (dConnectFrom != null && dConnectFrom.Count < 1)
					//{
					//	dConnectFrom.Clear();
					//}
					//if (dConnectTo != null && dConnectTo.Count < 1)
					//{
					//	dConnectTo.Clear();
					//}

					cntShpData.connFromHS = connFromHS;
					cntShpData.connToHS = connToHS;
					cntShpData.CntFrom = dConnectFrom;
					cntShpData.CntTo = dConnectTo;
					if ( (cntShpData.CntFrom != null && cntShpData.CntFrom.Count > 0) || (cntShpData.CntTo != null && cntShpData.CntTo.Count > 0))
					{
						dExcelConnectShapeMap.Add(cntShpData.ID, cntShpData);
					}
				}
				int y = 0;
			}
			catch(Exception ex)
			{
				MessageBox.Show(string.Format("Foreach issue: {0}", ex.Message));
			}


			//			dExcelConnectShapeMap = doesKeyExists(dExcelConnectShapeMap);
			Console.WriteLine("\n\n\n");

			foreach (var item in dExcelConnectShapeMap.Values)
			{
				StringBuilder sbStr = new StringBuilder();
				sbStr.Append(String.Format("Master Shape:{0}:{1}",item.ID, item.UniqueKey));
				if (item.CntFrom != null && item.CntFrom.Count > 0)
				{
					foreach(var from in item.CntFrom)
					{
						sbStr.Append(string.Format(" To:{0}:{1}", from.Key.ToString(), from.Value.ToString()));
					}
				}
				if (item.CntTo != null && item.CntTo.Count > 0)
				{
					foreach (var to in item.CntTo)
					{
						sbStr.Append(string.Format(" To:{0}:{1}",to.Key.ToString(), to.Value.ToString()));
					}
				}
				Console.WriteLine(sbStr.ToString());
			}
			return AllPageShapesMap;
		}
				
		private static Dictionary<int, ConnectShpExcelData> doesKeyExists(Dictionary<int, ConnectShpExcelData> connMap)
		{
			Dictionary<int, ConnectShpExcelData> newCnnMap = new Dictionary<int, ConnectShpExcelData>();
			foreach (var first in connMap)
			{
				foreach(var second in connMap)
				{
					// if second.ID = first.ID skip
					if (first.Key != second.Key)
					{
						if (second.Value.CntFrom != null && second.Value.CntFrom.Count > 0)
						{
							// iterate over the CntFrom.   if Key = first.ID remove it
							foreach (var value in second.Value.CntFrom)
							{
								if (value.Key != first.Key)
								{
									// keep this entry
									newCnnMap.Add(value.Key, second.Value);
								}
							}
						}
						// iterate over the cntTo.  if key = first.ID remove it
						if (second.Value.CntTo != null && second.Value.CntTo.Count > 0)
						{
							// iterate over the CntFrom.   if Key = first.ID remove it
							foreach (var value in second.Value.CntTo)
							{
								if (value.Key != first.Key)
								{
									// keep this entry
									newCnnMap.Add(value.Key, second.Value);
								}
							}
						}
					}
				}
			}
			return newCnnMap;
		}


				// While there are stil rows to look at.
				//while (shape.get_CellsSRCExists((short)VisSectionIndices.visSectionProp, iRow, (short)VisCellIndices.visCustPropsValue, (short)0) != 0)
				//{
				//	// Get the label and value of the current property.
				//	string label = shape.get_CellsSRC(
				//			  (short)VisSectionIndices.visSectionProp,
				//			  iRow,
				//			  (short)VisCellIndices.visCustPropsLabel
				//		 ).get_ResultStr(VisUnitCodes.visNoCast);
				//
				//	string value = shape.get_CellsSRC(
				//			  (short)VisSectionIndices.visSectionProp,
				//			  iRow,
				//			  (short)VisCellIndices.visCustPropsValue
				//		 ).get_ResultStr(VisUnitCodes.visNoCast);
				//
				//	// Print the results.
				//	Console.WriteLine(string.Format(
				//		 "Connection - Shape={0}; Label={1}; Value={2}",
				//		 shape.Name, label, value));
				//
				//	// Move to the next row in the properties section.
				//	iRow++;
				//}

				// Now look at child shapes in the collection.
				//if (shape.Master == null && shape.Shapes.Count > 0)
				//	getShapeConnections(doc);
			//}
		//}

		//public void ReadShapes(Microsoft.Office.Core.IRibbonControl control)
		//{
		//	ExportElement exportElement;
		//	ArrayList exportElements = new ArrayList();
		//
		//	Visio.Document currentDocument = Visio1.ActiveDocument;
		//	Visio.Pages Pages = currentDocument.Pages;
		//	Visio.Shapes Shapes;
		//
		//	foreach (Visio.Page Page in Pages)
		//	{
		//		Shapes = Page.Shapes;
		//		foreach (Visio.Shape Shape in Shapes)
		//		{
		//			exportElement = new ExportElement();
		//			exportElement.Name = Shape.Master.NameU;
		//			exportElement.ID = Shape.ID;
		//			exportElement.Text = Shape.Text;
		//			...
       //        // and any other properties you'd like
		 //
      //         exportElements.Add(exportElement);
		//		}
		//	}
		//}
	} 
}
