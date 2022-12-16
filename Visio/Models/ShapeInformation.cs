using System;
using Visio1 = Microsoft.Office.Interop.Visio;

namespace OmnicellBlueprintingTool.Visio
{
	public class ShapeInformation
	{
		public string VisioPage { get; set; } = "1";   // Visio Page can be just 1 or an actual text value 'Page-1'
		public int ID { get; set; } = 0;          // shape ID - used when reading a Visio diagram obtain the shape information
		public string GUID { get; set; }				// this is the uniqueID for a visio shape   note shape.ID is not unique accoross documents so you need to use the GUID wich is
		public string ShapeType { get; set; }     // 
		public string UniqueKey { get; set; }     // this is used to uniquly identify each shape that has been droped in the document.   used as key for Dictionary
		public string StencilImage { get; set; } = String.Empty;  // Must be the exact name for the stencil image in the Visio template/document
		public string StencilLabel { get; set; } = String.Empty;  // Text to add to the stencil image
		public string StencilLabelPosition { get; set; } = String.Empty;		// default Text box on Top
		public string StencilLabelFontSize { get; set; } = String.Empty;  // Text size if different than default (I.E.  12:B is 12 pt. Bold else 12 is 12 pt.)
		public bool IsStencilLabelFontBold { get; set; } = false;
		public string FillColor { get; set; } = String.Empty; // stincel fill color (default is none)
		public string RGBFillColor { get; set; } = String.Empty;	// will ge populated when reading a visio diagram.
																					// Try to match up to a preset color we are using	
		public double Pos_x { get; set; } = 0.0;				// X position to place the image
		public double Pos_y { get; set; } = 0.0;				// Y position to place the image
		public double Width { get; set; } = 0.0;				// width size of the image
		public double Height { get; set; } = 0.0;				// height size of the image
		public double StencilWidth { get; set; } = 0.0;		// Width of the shape from the stencil (used as an offset when writing to the Excel file)
		public double StencilHeight { get; set; } = 0.0;	// Height of the shape from the stencil (used as an offset when writing to the Excel file)

		public int ConnectFromID { get; set; } = 0;
		public string ConnectFrom { get; set; } = String.Empty;       // UniqueKey value for Lookukp
		public string FromLineLabel { get; set; } = String.Empty;
		public double FromLinePattern { get; set; } = VisioVariables.LINE_PATTERN_SOLID; // solid line
		public string FromLineColor { get; set; } = string.Empty; // empty is same as color "Black"
		public string FromArrowType { get; set; } = string.Empty; // empty is same as "None"
		public string FromLineWeight { get; set; } = VisioVariables.sLINE_WEIGHT_1;

		public int ConnectToID { get; set; } = 0;
		public string ConnectTo { get; set; } = String.Empty;        // UniqueKey value for Lookup
		public string ToLineLabel { get; set; } = String.Empty;
		public double ToLinePattern { get; set; } = VisioVariables.LINE_PATTERN_SOLID;   // solid line
		public string ToLineColor { get; set; } = string.Empty;  // empty is same as color "Black"
		public string ToArrowType { get; set; } = string.Empty;  // empty is same as "None"

		public string ToLineWeight { get; set; } = VisioVariables.sLINE_WEIGHT_1;

		public Visio1.Shape ShpObj { get; set; }     // this shape object
	}
}
