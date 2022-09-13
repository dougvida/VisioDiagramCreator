using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Visio1 = Microsoft.Office.Interop.Visio;
using VisioDiagramCreator.Models;
using VisioDiagramCreator.Visio;

namespace VisioDiagramCreator.Models
{
	public class DiagramData
	{
		public string VisioPageOrientation { get; set; } = "Portrait";
		public string VisioPageSize { get; set; } = "Letter";

		// This map contains all the shapes from the Excel Data file
		// it will be used to build up the connections with other shapes if needed
		public Dictionary<string, Device> AllShapesMap = new Dictionary<string, Device>();

		// This map will contain all the connections to and from shaps based on the Excel Data 
		public Dictionary<int, ShapeConnection> ShapeConnectionsMap = new Dictionary<int, ShapeConnection>();

		public string TemplateFilePath { get; set; }
		public string StencilFilePath { get; set; }

		public int MaxVisioPages { get; set; } = 1;		// default is 1 page visio document

		public List<Device> Devices { get; set; }

		public void Reset()
		{
			AllShapesMap.Clear();
			ShapeConnectionsMap.Clear();
			MaxVisioPages = 1;
			Devices.Clear();
		}
	}
}
