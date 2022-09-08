using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.WebSockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;
//using Visio1 = Microsoft.Office.Interop.Visio;
using VisioDiagramCreator.Models;

namespace VisioDiagramCreator.Visio
{
	public class ProcessVisioShapeConnections
	{
		/// <summary>
		/// BuildShapeConnections
		/// This function will parse through all the Devices found in the 'AllShapesMap' Dictionary
		/// to get the unique key names to use as the connectionFrom and connectionTo.
		/// The unique key names are used to look up in the connectionMap dictionary to obtain the visioShape object that 
		/// was used when the object was dropped on the page
		/// </summary>
		/// <param name="diagData">Object</param>
		/// <returns>Dictionary<int,ShapeConnection></int></returns>
		public Dictionary<int, ShapeConnection> BuildShapeConnections(DiagramData diagData)
		{
			string[] saStr = null;
			string sKey = string.Empty;

			Dictionary<int, ShapeConnection> vConnectMap = new Dictionary<int, ShapeConnection>();

			ShapeConnection vsConnect = null;
			Device device = new Device();
			Device lookUpDevice = new Device();

			int nCnt = 0;
			foreach (KeyValuePair<string, Device> item in diagData.AllShapesMap)
			{
				if (!string.IsNullOrEmpty(item.Key.ToString()))
				{
					sKey = item.Key.ToString();
					device = item.Value;

					// check if a ConnectFrom is populated
					// if yes than do the connection map
					if (!string.IsNullOrEmpty(device.ShapeInfo.ConnectFrom)) // this is a stencil wanting to connect back to a shape object
					{
						// check if we have more than one connect tion make
						saStr = device.ShapeInfo.ConnectFrom.Split(',');
						foreach (string str in saStr)
						{
							if (string.IsNullOrEmpty(str))
								continue;

							// lookup the connectFrom
							if (diagData.AllShapesMap.TryGetValue(str.Trim(), out lookUpDevice))
							{
								// device object we want all the info about the connection
								// lookUpDevice we only need the shapObj and UniqueKey
								vsConnect = new ShapeConnection();
								vsConnect.ShpObj = device.ShapeInfo.ShpObj;

								vsConnect.LineLabel = device.ShapeInfo.FromLineLabel;
								vsConnect.ArrowType = device.ShapeInfo.FromArrowType;
								vsConnect.LinePattern = device.ShapeInfo.FromLinePattern;
								vsConnect.LineColor = device.ShapeInfo.FromLineColor;

								vsConnect.UniqueFromKey = lookUpDevice.ShapeInfo.UniqueKey;
								vsConnect.ShpFromObj = lookUpDevice.ShapeInfo.ShpObj;			// lookUpDevice shape Object (ConnectFrom)

								vsConnect.UniqueToKey = device.ShapeInfo.UniqueKey;
								vsConnect.ShpToObj = device.ShapeInfo.ShpObj;					// current shape Object
								vsConnect.device = device;
								vConnectMap.Add(nCnt++, vsConnect);
							}
							Console.WriteLine("ConnectFrom - Found UniqueFromKey: {0}", str);
						}
					}

					// next check if a ConnectTo is populated
					if (!string.IsNullOrEmpty(device.ShapeInfo.ConnectTo))
					{
						// check if we have more than one connect tion make
						saStr = device.ShapeInfo.ConnectTo.Split(',');
						foreach (string str in saStr)
						{
							if (string.IsNullOrEmpty(str))
								continue;

							// lookup the connectTo object
							if (diagData.AllShapesMap.TryGetValue(str.Trim(), out lookUpDevice))
							{
								// device object we want all the info about the connection
								// lookUpDevice we only need the shapObj and UniqueKey
								vsConnect = new ShapeConnection();
								vsConnect.ShpObj = lookUpDevice.ShapeInfo.ShpObj;

								vsConnect.LineLabel = device.ShapeInfo.ToLineLabel;
								vsConnect.ArrowType = device.ShapeInfo.ToArrowType;
								vsConnect.LinePattern = device.ShapeInfo.ToLinePattern;
								vsConnect.LineColor = device.ShapeInfo.ToLineColor;

								vsConnect.UniqueFromKey = device.ShapeInfo.UniqueKey;
								vsConnect.ShpFromObj = device.ShapeInfo.ShpObj;

								vsConnect.UniqueToKey = lookUpDevice.ShapeInfo.UniqueKey;
								vsConnect.ShpToObj = lookUpDevice.ShapeInfo.ShpObj;
								vsConnect.device = device;
								vConnectMap.Add(nCnt++, vsConnect);
							}
							Console.WriteLine("ConnectTo - Found UniqueToKey: {0}", str);
						}
					}
				}
			}
			return vConnectMap;
		}
	}
}
