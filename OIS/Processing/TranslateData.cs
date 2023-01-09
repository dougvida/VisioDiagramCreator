using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using OIS.Models;
using OmnicellBlueprintingTool.Models;
using OmnicellBlueprintingTool.Visio;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition.Primitives;
using System.Linq;
using VisioAutomation.VDX.Elements;

namespace OmnicellOISNodes.Processing
{
	internal class TranslateData
	{

		public static Dictionary<string, ShapeInformation> ConvertData(Dictionary<string, List<OISSetupData>> dataMap)
		{
			Dictionary<string, ShapeInformation> shapeInfoMap = new Dictionary<string, ShapeInformation>();
			ShapeInformation shapeInfo = null;

			List<OISSetupData> nodes = null;

			int nCounter = 1;

			const double dColumnGap = 1.225;
			const double Column1 = 0.250;
			const double Column2 = 0.900;
			const double Column3 = Column2 + dColumnGap;

			double dRootPosX = 0.250;	// Left most column position
			double dTopPosY = 20.000;  // Top Row position

			double dSkipX = 0.375;     // Column gap
			double dSkipY = 0.600;     // Row gap

//			double dSquareWidth = 0.500;		// width of Square shape
//			double dRectangleWidth = 1.000;	// width of Rectangle shape
//			double dFileWidth = 1.020;			// width of File shape
//			double dDatabaseWidth = 1.500;	// width of Database shape

			double dPosX = dRootPosX;
			double dPosY =  dTopPosY;

			string sFillColor = string.Empty;
			string sStencilImage = string.Empty;
			// lets iterate through the dictionary starting with the first one and write the output
			foreach (KeyValuePair<string, List<OISSetupData>> item in dataMap)
			{
				string sConnectFrom = string.Empty;		// this will contain the previous node for connection

				// iterate over each node from the base node
				for (int nCnt = 0; nCnt < item.Value.Count; nCnt++)
				{
					nodes = parseNode(item.Value[nCnt]);

					for (int nIdx = 0; nIdx < nodes.Count; nIdx++) // foreach (OISSetupData node in nodes)
					{					
						sStencilImage = string.Empty;
						sFillColor = string.Empty;

						if (nodes[nIdx].Node.Length == 3 )
						{
							sStencilImage = "OC_Square2";
							sFillColor = "Green Light";
							dPosX = Column2;
							if (nIdx == 1 && (nodes[nIdx].Details.IndexOf("MSMQ") >= 0) ) // second shape
							{
								// first item was a OC_Database2 (MSMQ) so lets set shape position to left
								dPosX = Column1;
							}
						}
						else if (nodes[nIdx].Type.IndexOf("[X~Y]", StringComparison.OrdinalIgnoreCase) >= 0)
						{
							sStencilImage = "OC_Rectangle2";
							sFillColor = "Blue Light";
						}
						else if (nodes[nIdx].Node.IndexOf("Input", StringComparison.OrdinalIgnoreCase) >= 0)
						{
							dPosX = Column1;     // shape is first one set to Left

							sStencilImage = "OC_Rectangle2";
							sFillColor = "Orange";

							if (nodes[nIdx].Label.IndexOf("File", StringComparison.OrdinalIgnoreCase) >= 0)
							{
								sStencilImage = "OC_File2";
								sFillColor = "Silver";
								if (nIdx == 0)  // is First object (Input)
								{
//									dPosX = Column1;     // shape is first one set to Left
									dPosY += 0.075;      // we need to shift the object up
								}
							}
							else if (nodes[nIdx].Label.IndexOf("MSMQ", StringComparison.OrdinalIgnoreCase) >= 0)
							{
								sStencilImage = "OC_Database2";
								sFillColor = "Silver";
								if (nIdx == 0)	// Is first object (Input)
								{
									// MSMQ is the first input shape, shift it to 3 columns
									dPosX = Column3;
									dPosY -= 0.150;		// add more space between previous row and this shape
								}
							}
						}
						else if (nodes[nIdx].Node.IndexOf("Output", StringComparison.OrdinalIgnoreCase) >= 0)
						{
							sStencilImage = "OC_Rectangle2";
							sFillColor = "Orange";

							dPosX += dSkipX; // output device so give a bit more space after left shape

							if (nodes[nIdx].Label.IndexOf("File", StringComparison.OrdinalIgnoreCase) >= 0)
							{
								sStencilImage = "OC_File2";
								sFillColor = "Cyan";
								dPosY += 0.065;      // ship up to aline center
							}
							else if (nodes[nIdx].Label.IndexOf("MSMQ", StringComparison.OrdinalIgnoreCase) >= 0)
							{
								sStencilImage = "OC_Database2";
								sFillColor = "Cyan";
								dPosY -= 0.036;		// shift down to align center
							}
						}
						else if (nodes[nIdx].Node.Length == 5)
						{
							sStencilImage = "OC_Square2";
						}
						else
						{
							// the sStencilImage may have been set already if null than this is an ERROR
							if (string.IsNullOrEmpty(sStencilImage))
							{
								Console.WriteLine(string.Format("*** ERROR *** Selecting a StencilImage: '{0}' '{1}'", nodes[nIdx].Node, nodes[nIdx].Label));
								sStencilImage = "OC_Process2";  // this is an error 
								sFillColor = "red";
							}
						}

						// update the second shape to add
						shapeInfo = new ShapeInformation();
						shapeInfo.VisioPage = "1";
						shapeInfo.ShapeType = "Shape";
						shapeInfo.UniqueKey = string.Format("{0}:{1}", sStencilImage, nCounter++);
						shapeInfo.StencilImage = sStencilImage;
						shapeInfo.StencilLabel = string.Format("{0}\n{1}", nodes[nIdx].Node, nodes[nIdx].Label);
						shapeInfo.FillColor = sFillColor;
						shapeInfo.Pos_x = (Math.Truncate(dPosX * 1000) / 1000);
						shapeInfo.Pos_y = (Math.Truncate(dPosY * 1000) / 1000);

						if (shapeInfo.StencilLabel.IndexOf("Input", StringComparison.OrdinalIgnoreCase) >= 0)
						{
							// the node was an Input so we want all other nodes to be shifted down half row

							if ( (nodes[nIdx].Label.IndexOf("MSMQ", StringComparison.OrdinalIgnoreCase) >= 0 ||
								   nodes[nIdx].Label.IndexOf("File", StringComparison.OrdinalIgnoreCase) >= 0 ) && nIdx == 0)
							{
								// Current shape is an Input shape of (MSMQ or File) so set X position to Left for next shape
								dPosX = dRootPosX;
								dPosY -= dSkipY;
							}
							else
							{
								dPosX += dColumnGap;
								dPosY -= dSkipY;
							}
						}
						else
						{
							// set Position X for next node
							if (shapeInfo.StencilImage.IndexOf("OC_Square") >= 0)
							{
								dPosX += 0.700;
							}
							else
							{
								dPosX += dColumnGap;
							}
						}

						// make a connection
						if (!string.IsNullOrEmpty(sConnectFrom))
						{
							shapeInfo.ConnectFrom = sConnectFrom;
						}
						sConnectFrom = shapeInfo.UniqueKey;		// update the variable to hold this node for a connection with the next node

						Console.WriteLine(string.Format("\nTranslateData:: UniqueKey:'{0}' Label:'{1}'", shapeInfo.UniqueKey, shapeInfo.StencilLabel));

						if (!shapeInfoMap.ContainsKey(nodes[nIdx].Node))
						{
							shapeInfoMap.Add(nodes[nIdx].Node, shapeInfo);
						}
						else
						{
							Console.WriteLine(string.Format("*** ERROR *** Key :'{0}' already exists in the Map", nodes[nIdx].Node));
						}
					}
				}

				// set the starting position for next row
				dPosX = dRootPosX;

				// set next row position
				dPosY -= dSkipY;
			}
			// write to the file each line in the StringBuilder
			Console.WriteLine(string.Format("\nNumber of entries in the dictionary: {0}", shapeInfoMap.Count));

			return shapeInfoMap;
		}

		private static List<OISSetupData> parseNode(OISSetupData nodeArg)
		{
			List<OISSetupData> nodes = new List<OISSetupData>();
			OISSetupData node1 = null;
			OISSetupData node2 = null;

			node1 = getNodeData(nodeArg);
			if (node1 != null)
			{
				if (node1.Type.Equals("[X~Y]"))
				{
					nodes.Add(node1);
					return nodes;              // there is only one for this type of Node
				}

				// get the second part
				if (nodeArg.Node.Length == 3 || nodeArg.Node.Length >= 5)
				{
					node2 = new OISSetupData();
					node2.Index = nodeArg.Index;
					node2.Type = nodeArg.Type;
					node2.Node = nodeArg.Node;
					node2.Path = nodeArg.Path;
					node2.Desc = nodeArg.Desc;
					node2.Details= nodeArg.Details;
					node2.Label = nodeArg.Desc;
					if (string.IsNullOrEmpty(nodeArg.Desc))
					{
						node2.Label = nodeArg.Type; // nodeArg.Details;
					}
				}
				else
				{
					node2 = getNodeData(nodeArg);
				}

				if (node1.Node.IndexOf("Input", StringComparison.OrdinalIgnoreCase) >= 0)
				{
					nodes.Add(node1);
					nodes.Add(node2);
				}
				else if (node1.Node.IndexOf("Output", StringComparison.OrdinalIgnoreCase) >= 0)
				{
					nodes.Add(node2);
					nodes.Add(node1);
				}
			}
			return nodes;
		}

		private static OISSetupData getNodeData(OISSetupData nodeArg)
		{
			string sIP = string.Empty;
			string sPort = string.Empty;

			OISSetupData node = new OISSetupData();
			node.Index = nodeArg.Index;
			node.Type = nodeArg.Type;
			node.Node = nodeArg.Node;
			node.Path = nodeArg.Path;
			node.Desc = nodeArg.Desc;
			node.Details = nodeArg.Details;

			// the input node needs to be the first one of two
			// the output node needs to be the last one of two
			switch(nodeArg.Type)
			{
				case "[X~Y]":
					if (string.IsNullOrEmpty(node.Desc))
					{
						node.Label = string.Format("{0}\n{1}", node.Type, node.Details);
					}
					else
					{
						node.Label = string.Format("{0}\n{1}", node.Type, node.Desc);
					}
					break;
			//	case "ADOComm":
			//		break;
			//	case "KafkaComm":
			//		break;
			//	case "MSMQComm":
			//		break;
			//	case "NullWriterComm":
			//		break;
			//	case "FileReaderComm":
			//		break;
			//	case "":
			//		break;
			//	case "ClientSocketComm":
			//		break;
			//	case "ServerSocketComm":
			//		break;
			//	case "SFTPFileReaderComm":
			//		break;
			//	case "FTPFileWriterComm":
			//		break;
			//	case "DataBaseComm":
			//		break;
			//	case "MOVEitDMZWriterComm":
			//		break;
				default:
					int nInputNode = nodeArg.Details.IndexOf("Input Node:", StringComparison.OrdinalIgnoreCase);
					if (nInputNode >= 0)
					{
						node.Node += ":Input";

						// found now parse  for Port
						string sInputNode = nodeArg.Details.Substring(nInputNode);
						int nHost = nodeArg.Details.Substring(nInputNode).IndexOf("Host IP:", StringComparison.OrdinalIgnoreCase);
						if (nHost >= 0)
						{
							string sHost = sInputNode.Substring(nHost);

							// lets see if the name Port exists
							int nPort = sHost.IndexOf("Port:", StringComparison.OrdinalIgnoreCase);
							if (nPort >= 0)
							{
								string[] sTmp2 = sHost.Split(' ');
								sIP = sTmp2[1].Substring(3);
								if (sTmp2[3].Equals("Port:", StringComparison.OrdinalIgnoreCase))
								{
									sPort = sTmp2[4];
								}
							}
							node.Label = string.Format("Host:\n{0}:{1}", sIP, sPort);
						}
						else
						{
							node.Label = nodeArg.Details.Substring(nInputNode);

							int nMSMQ = nodeArg.Details.Substring(nInputNode).IndexOf("MSMQ:", StringComparison.OrdinalIgnoreCase);
							if (nMSMQ >= 0)
							{
								node.Label = nodeArg.Details.Substring(nMSMQ).Trim();
							}
							else
							{
								int nFile = nodeArg.Details.Substring(nInputNode).IndexOf("[File]", StringComparison.OrdinalIgnoreCase);
								if (nFile >= 0)
								{
									node.Label = nodeArg.Details.Substring(nFile).Trim();
								}
								else
								{
									node.Label = string.Format("UNKNOWN\n{0}", nodeArg.Details.Substring(nInputNode));
								}
							}
						}
					}
					else
					{
						int nOutputNode = nodeArg.Details.IndexOf("Output Node:", StringComparison.OrdinalIgnoreCase);
						if (nOutputNode >= 0)
						{
							node.Node += ":Output";

							node.Label = nodeArg.Details.Substring(nOutputNode);

							int nMSMQ = nodeArg.Details.Substring(nOutputNode).IndexOf("MSMQ:", StringComparison.OrdinalIgnoreCase);
							if (nMSMQ >= 0)
							{
								node.Label = nodeArg.Details.Substring(nMSMQ).Trim();
							}
							else
							{
								int nFile = nodeArg.Details.Substring(nOutputNode).IndexOf("[File]", StringComparison.OrdinalIgnoreCase);
								if (nFile >= 0)
								{
									node.Label = nodeArg.Details.Substring(nFile).Trim();
								}
								else
								{
									node.Label = string.Format("UNKNOWN\n{0}", nodeArg.Details.Substring(nOutputNode));
								}
							}
						}
						else
						{
							node.Label = string.Format("{0}\n{1}", node.Type, node.Desc);
							if (string.IsNullOrEmpty(node.Desc))
							{
								node.Label = string.Format("{0}", node.Type);
							}
						}
					}
					break;
			}
			return node;
		}
	}
}
