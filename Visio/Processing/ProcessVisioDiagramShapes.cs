using System;
using System.Collections.Generic;
using Visio1 = Microsoft.Office.Interop.Visio;
using Microsoft.Office.Interop.Visio;
using System.Windows.Forms;

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
		public Dictionary<int, ShapeInformation> GetAllShapesProperties(string diagamFilePathName)
		{
			// Open up one of Visio's sample drawings.
			Visio1.Application app = new Visio1.Application();
			Visio1.Document doc = app.Documents.Open(diagamFilePathName);

			// Get the first page in the sample drawing.
			Visio1.Page page = doc.Pages[1];

			Console.WriteLine("Active Document:{0}: Master in document:{1}", app.ActiveDocument, app.ActiveDocument.Masters);

			// Start with the collection of shapes on the page and 
			// print the properties we find,
			return printProperties(doc);
		}

		/* This function will travel recursively through a collection of 
		 * shapes and print the custom properties in each shape. 
		 * 
		 * The reason I don't simply look at the shapes in Page.Shapes is 
		 * that when you use the Group command the shapes you group become 
		 * child shapes of the group shape and are no longer one of the 
		 * items in Page.Shapes.
		 * 
		 * This function will not recursive into shapes which have a Master. 
		 * This means that shapes which were created by dropping from stencils 
		 * will have their properties printed but properties of child shapes 
		 * inside them will be ignored. I do this because such properties are 
		 * not typically shown to the user and are often used to implement 
		 * features of the shapes such as data graphics.
		 * 
		 * An alternative halting condition for the recursion which may be 
		 * sensible for many drawing types would be to stop when you 
		 * find a shape with custom properties.
		 */
		public static Dictionary<int, ShapeInformation> printProperties(Visio1.Document doc)
		{
			// Look at each shape in the collection.
			Visio1.Page page = doc.Pages[1];
			Dictionary<int, ShapeInformation> AllPageShapesMap = new Dictionary<int, ShapeInformation>();
			ShapeInformation shpInfo = null;
			try
			{
				foreach (Visio1.Shape shape in page.Shapes)
				{
					// Use this index to look at each row in the properties 
					// section.
					shpInfo = new ShapeInformation();

					shpInfo.ID = shape.ID;

					short iRow = (short)VisRowIndices.visRowFirst;
					shpInfo.Pos_x = Math.Truncate(shape.Cells["PinX"].ResultIU * 1000) / 1000;
					shpInfo.Pos_y = Math.Truncate(shape.Cells["PinY"].ResultIU * 1000) / 1000;
					shpInfo.Width = Math.Truncate(shape.Cells["Width"].ResultIU * 1000) / 1000;
					shpInfo.Height = Math.Truncate(shape.Cells["Height"].ResultIU * 1000) / 1000;

					string sConnFrom = string.Empty;
					string sConnTo = string.Empty;

					string ShapeType = "Shape";
					shpInfo.UniqueKey = shape.NameU.Trim();

					string[] saStr = shape.NameU.Split(':');
					shpInfo.StencilImage = saStr[0].Trim();
					shpInfo.StencilLabel = shape.Text.Trim();

					AllPageShapesMap.Add(shape.ID, shpInfo);

					if (shpInfo.StencilImage.ToUpper().IndexOf("NETWORKPIPE") >= 0)
					{
						// skip this shape
						Console.WriteLine(string.Format("Skip this ID:{0}; shapeKey:{1} - Stencil Image:{2}", shape.ID, shpInfo.UniqueKey, shpInfo.StencilImage));
						continue;
					}
					if (shape.Style.Equals("Connector"))
					{
						continue;
					}
					continue;
					var shpFromConnected = shape.ConnectedShapes(VisConnectedShapesFlags.visConnectedShapesOutgoingNodes, "");
					if (shpFromConnected != null && shpFromConnected.Length > 0)
					{
						Console.WriteLine("");
						try
						{
							sConnFrom = string.Empty;
							int Id = 0;
							foreach (int nIdx in shpFromConnected)
							{
								Id = nIdx;
								Visio1.Shape cnnShape = page.Shapes.ItemFromID[nIdx];
								shpInfo.ConnectFromID = cnnShape.ID;
								shpInfo.ConnectFrom = cnnShape.NameU;
								shpInfo.FromLineColor = "BLACK";
								shpInfo.FromLinePattern = VisioVariables.LINE_PATTERN_SOLID;
					
								bool keyExists = AllPageShapesMap.ContainsKey(shape.ID);
								if (!keyExists)
								{
									AllPageShapesMap.Add(shape.ID, shpInfo);
								}
								else
								{
									ShapeInformation value = new ShapeInformation();
									AllPageShapesMap.TryGetValue(shape.ID, out value);
									if (string.IsNullOrEmpty(value.ConnectFrom))
									{
										value.ConnectFromID = shape.ID;
										value.ConnectFrom = cnnShape.NameU;
										AllPageShapesMap[shape.ID] = value;	// update the value
									}
								}
					
								// Print the results.
								if (!string.IsNullOrEmpty(sConnFrom))
								{
									sConnFrom += "\n";
								}
								sConnFrom += (string.Format("Connect From: ID:{0};ShapeType:{1}; LookupId:{2}; UniqueKey:{3}; StencilLabel:{4}; PosX:{5}; PoxY:{6}; ConnectFromKey:{7}; ConnecID:{8}",
																			shape.ID, ShapeType, Id, shpInfo.UniqueKey, shpInfo.StencilLabel, shpInfo.Pos_x, shpInfo.Pos_y, shpInfo.ConnectFrom, cnnShape.ID));
							}
						}
						catch (Exception exp)
						{
							MessageBox.Show(String.Format("ConnectFrom:{0} - {1}", shpInfo.ConnectFrom, exp.Message));
						}
					}

					var shpToConnected = shape.ConnectedShapes(VisConnectedShapesFlags.visConnectedShapesIncomingNodes, "");
					if (shpToConnected != null && shpToConnected.Length > 0)
					{
						try
						{
							sConnTo = string.Empty;
							int Id = 0;
							foreach (int nIdx in shpToConnected)
							{
								Id = nIdx;
								Visio1.Shape cnnShape = page.Shapes.ItemFromID[nIdx];

								//Console.WriteLine("Shapes that are To / Outgoing connections: {0}", sp.NameU);
								shpInfo.ConnectToID = cnnShape.ID;
								shpInfo.ConnectTo = cnnShape.NameU;
								shpInfo.ToLineColor = "BLACK";
								shpInfo.ToLinePattern = VisioVariables.LINE_PATTERN_SOLID;

								bool keyExists = AllPageShapesMap.ContainsKey(shape.ID);
								if (!keyExists)
								{
									AllPageShapesMap.Add(shape.ID, shpInfo);
								}
								else
								{
									ShapeInformation value = new ShapeInformation();
									AllPageShapesMap.TryGetValue(shape.ID, out value);
									if (string.IsNullOrEmpty(value.ConnectTo))
									{
										value.ConnectFromID = shape.ID;
										value.ConnectTo = cnnShape.NameU;
										AllPageShapesMap[shape.ID] = value; // update the value
									}
								}

								// Print the results.
								if (!string.IsNullOrEmpty(sConnTo))
								{
									sConnTo += "\n";
								}
								sConnTo += (string.Format("Connect To: ID:{0}; ShapeType:{1}; LookupId:{2}; UniqueKey:{3}; StencilLabel:{4}; PosX:{5}; PoxY:{6}; ConnectToKey:{7}; ConnectID:{8}",
																			shape.ID, ShapeType, Id, shpInfo.UniqueKey, shpInfo.StencilLabel, shpInfo.Pos_x, shpInfo.Pos_y, shpInfo.ConnectTo, cnnShape.ID));
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

					if (!string.IsNullOrEmpty(sConnFrom) || !string.IsNullOrEmpty(sConnTo))
					{
						Console.WriteLine("{0}\n{1}", sConnFrom, sConnTo);
					}
				}
			}
			catch(Exception ex)
			{
				MessageBox.Show(string.Format("Foreach issue: {0}", ex.Message));
			}
			return AllPageShapesMap;
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
				//	printProperties(doc);
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
