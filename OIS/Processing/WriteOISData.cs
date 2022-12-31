using OIS.Models;
using System;
using System.Collections.Generic;

namespace OmnicellOISNodes.Processing
{
	internal class WriteOISData
	{
		public static int WriteAllData(Dictionary<string, List<OISSetupData>> dataMap, string fileNamePath)
		{
			int nReturn = 0;  // no error
			List<string> sb = new List<String>();

			int nIndex = 1;
			string sHeader = "Visio Page, Shape Type, Unique Key, Stencil Image, Stencil Label Position, Stencil Label Font Size, Mach_name, Mach_Id, Site_Id,Sitename, Site_Address, Omnis_name,Omnis_Id,SiteId_Omni_Id,IP,Ports,Devices Count,PosX,PosY,Width,Height,Fill Color, RGB Fill Color,Connect from, From Line Label,From Line Pattern,From Arrow Type, From Line Color, From Line Weight,Connect To,To Line label,To Line Pattern, To Arrow Type, To Line Color, To Line Weight";
			sb.Add(sHeader);

			string sTmp, sLabel;

			string sIP = string.Empty;
			string sPort = string.Empty;

			double dRootPosX = 0.250;			// root node
			double dTopPosY = 16.000;			// top of page
			double dPosX = dRootPosX;
			double dPosY =  dTopPosY;

			// lets iterate through the dictionary starting with the first one and write the output
			foreach (KeyValuePair<string, List<OISSetupData>> item in dataMap)
			{
				// iterate over each node from the base node
				for (int nCnt = 0; nCnt < item.Value.Count; nCnt++)
				{
					OISSetupData node = item.Value[nCnt];
					if (nCnt == 0)    // first base node
					{
						Console.WriteLine(string.Format("\nPrimary Node:'{0}' Type:'{1}' Desc:'{2}' Details:'{3}'", node.Node, node.Type, node.Desc, node.Details));
						sLabel = string.Format("{0}:{1}\n{2}\n{3}", node.Node, node.Type, node.Desc, node.Details);
						dPosX = dRootPosX;
					}
					else if (nCnt == 1)
					{
						// normally a tanslator node
						Console.WriteLine(string.Format("[X~Y]   Node:'{0}' Desc:'{1}' - Translation", node.Node, node.Desc));
						sLabel = string.Format("[X~Y]\n{0}\n{1}", node.Node, node.Desc);
						dPosX += 0.600;
						dPosY -= 0.400;
					}
					else
					{
						Console.WriteLine(string.Format("\nPrimary Node:'{0}' Type:'{1}' Desc:'{2}' Details:'{3}'", node.Node, node.Type, node.Desc, node.Details));
						sLabel = string.Format("{0}:{1}\n{2}\n{3}", node.Node, node.Type, node.Desc, node.Details);
						dPosX += 0.600;
						dPosY -= 0.400;
					}
					sTmp = string.Format("{0},Shape,OC_Rectangle1:{1},OC_Rectangle1,{2},,,,,,,,,,,{3},{4},{5},{6},,,,," +
						"ConnectFrom,FromLabel,,,,," +
						"Connect To,ToLineLabel,,,,", 1, nIndex, sLabel, sIP, sPort, (Math.Truncate(dPosX * 1000) / 1000), (Math.Truncate(dPosY * 1000) / 1000));
					sb.Add(sTmp);
				}
			}
			// write to the file each line in the StringBuilder
			Console.WriteLine(string.Format("\nNumber of entries in the StringBuilder: {0}", sb.Count));

			// write the data to the Excel file


			return nReturn;
		}
	}
}
