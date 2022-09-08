using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VisioDiagramCreator.Visio;
using Visio1 = Microsoft.Office.Interop.Visio;

namespace VisioDiagramCreator.Visio
{
	public class ShapeInformation
	{
		public int ID { get; set; } = 0;	// shape ID
		public string UniqueKey { get; set; }     // this is used to uniquly identify each shape that has been droped in the document.   used as key for Dictionary
		public string StencilLabel { get; set; } = String.Empty;  // Text to add to the stencil image
		public string StencilImage { get; set; } = String.Empty;  // Must be the exact name for the stencil image in the Visio template/document
		public string FillColor { get; set; } = "None"; // stincel fill color (default is none)	
		public double Pos_x { get; set; } = 0.0;         // X position to place the image
		public double Pos_y { get; set; } = 0.0;         // Y position to place the image
		public double Width { get; set; } = 0.0;         // width size of the image
		public double Height { get; set; } = 0.0;        // height size of the image

		public int ConnectFromID { get; set; } = 0;
		public string ConnectFrom { get; set; } = String.Empty;       // UniqueKey value for Lookukp
		public string FromLineLabel { get; set; } = String.Empty;
		public double FromLinePattern { get; set; } = VisioVariables.LINE_PATTERN_SOLID;	// solid line
		public string FromLineColor { get; set; } = VisioVariables.COLOR_BLACK;
		public string FromArrowType { get; set; } = VisioVariables.sARROW_NONE;

		public int ConnectToID { get; set; } = 0;
		public string ConnectTo { get; set; } = String.Empty;        // UniqueKey value for Lookup
		public string ToLineLabel { get; set; } = String.Empty;
		public double ToLinePattern { get; set; } = VisioVariables.LINE_PATTERN_SOLID;	// solid line
		public string ToLineColor { get; set; } = VisioVariables.COLOR_BLACK;
		public string ToArrowType { get; set; } = VisioVariables.sARROW_NONE;

		public Visio1.Shape ShpObj { get; set; }		// this shape object
	}
}
