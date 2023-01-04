using LumenWorks.Framework.IO.Csv;
using Microsoft.Office.Interop.Excel;
using OIS.Models;
using OmnicellBlueprintingTool.Visio;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;

namespace OmnicellOISNodes.Processing
{
	internal class ParseOISSetup
	{
		public static Dictionary<string, List<OISSetupData>> ParseOISSetupFile(string OISSetupFilePath)
		{
			Dictionary<string, List<OISSetupData>> oisDataMap = new Dictionary<string, List<OISSetupData>>();

			List<OISSetupData> oisDataList = new List<OISSetupData>();	// contains the list of like nodes
			OISSetupData oisData = null;                                // data for each node

			bool bHeader = true;
			using (CsvReader csv = new CsvReader(new StreamReader(OISSetupFilePath), bHeader))
			{
				int fieldCount = csv.FieldCount;
				string[] headers = csv.GetFieldHeaders();

				while (csv.ReadNextRecord())
				{
					oisData = new OISSetupData();
					for (int i = 0; i < fieldCount; i++)
					{
						// we need to group the same like nodes together
						// I.E.  ABC is group with ABC1 and ABC11 not ADT
						switch(i)
						{
							case 0:  //Index
								oisData.Index = Convert.ToInt32(csv[i].Trim());
								break;
							case 1:	// Type
								oisData.Type = csv[i].Trim();
								break;
							case 2:  // Path
								oisData.Path = csv[i].Trim();
								break;
							case 3:  // Node
								oisData.Node = csv[i].Trim();
								break;	
							case 4:  // Desc
								oisData.Desc = csv[i].Trim();
								break;
							case 5:	// Details
								oisData.Details = csv[i].Trim();
								break;
							case 6:	// RegExpr
								oisData.RegExpr = csv[i].Trim();
								break;

							default: // error
								Console.WriteLine(string.Format("Error with parsing data.  Index:{0}, value:{1}", i, csv[i]));
								break;
						}
					}
					if (oisDataList.Count > 0)
					{
						if (oisData.Node.StartsWith(oisDataList[0].Node))
						{
							oisDataList.Add(oisData);
						}
						else
						{
							// we have a different node so start a new list
							oisDataMap.Add(oisDataList[0].Node, oisDataList);
							oisDataList = new List<OISSetupData>();
							oisDataList.Add(oisData);
						}
					}
					else
					{
						oisDataList.Add(oisData);
					}
				}
				// we need to write the last one to the dictionary
				oisDataMap.Add(oisDataList[0].Node, oisDataList);
			}
			return oisDataMap;
		}
	}
}
