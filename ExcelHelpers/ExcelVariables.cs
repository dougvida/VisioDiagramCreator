using System;
using System.Collections.Generic;

namespace OmnicellBlueprintingTool.ExcelHelpers
{
	public class ExcelVariables
	{

		public static int GetHeaderCount()
		{
			return (Enum.GetNames(typeof(CellIndex)).Length);
		}

		public enum CellIndex
		{
			// NOTE ****
			// the order of this enum must match the column order in the Excel file
			VisioPage = 1,       // Page indicator to place this shape
			ShapeType,           // key
			UniqueKey,				// device unique Key used for connecting visio shapes
			StencilImage,			// device visio image name
			StencilLabel,			// device label
			StencilLabelPosition,// label position default is top.   Bottom
			StencilLabelFontSize,// default is what the stencil font size is   (use 12:B for 12 pt. Bold or 12 for 12 pt)
			
			Mach_name,           // device machine Name
			Mach_id,             // device machine Id
			Site_id,             // device site Id
			Site_name,           // deivce site name
			Site_address,        // device site address
			Omnis_name,          // device name
			Omnis_id,            // device Id
			SiteIdOmniId,        // site_id+omni_id
			IP,                  // device IP address
			Ports,               // device Ports
			DevicesCount,        // number of Devices for this type (part of a group)

			PosX,                // Shape position X
			PosY,                // shape position Y
			Width,               // shape width
			Height,              // shape height
			FillColor,           // color to fill stincel
			
			ConnectFrom,         // used to link this visio shape to another visio shape
			FromLineLabel,       // Arrow Text
			FromLinePattern,     // Line pattern solid = 1
			FromArrowType,       // Can contain one of these [None, Start, End, Both]
			FromLineColor,       // Arrow Color

			ConnectTo,           // used to link this visio shape to another visio shape
			ToLineLabel,         // Arrow Text
			ToLinePattern,       // Line pattern default (solid);  Solid=1, Dash=2, Dotted=3, Dash_Dot=4
			ToArrowType,         // Can contain one of these [None, Start, End, Both]
			ToLineColor,         // Arrow Color
		}

		public static Dictionary<int, string> GetExcelHeaderNames()
		{
			Dictionary<int, string> excelHeaderNames = new Dictionary<int, string>
			{
				// Excel data file header.  Must be in this sequence
				{ 0, "Visio Page"},		      // Page indicator to place this shape
				{ 1, "Shape Type"},           // key
				{ 2, "Unique Key"},           // unique Key used for connecting visio shapes
				{ 3, "Stencil Image"},        // stencil visio image name
				{ 4, "Stencil Label"},        // stencil label
				{ 5, "Stencil Label Position"}, // stencil label position
				{ 6, "Stencil Label Font Size"}, // default is what the stencil font size is   (use 12:B for 12 pt. Bold or 12 for 12 pt)

				{ 7, "Mach_name"},				// device machine Name
				{ 8, "Mach_id"},					// device machine Id
				{ 9, "Site_id"},					// device site Id
				{ 10, "Site_name"},				// deivce site name
				{ 11, "Site_address"},			// device site address
				{ 12, "Omnis_name"},				// device name
				{ 13, "Omnis_id"},				// device Id
				{ 14, "SiteId_OmniId"},			// site_id+omni_id
				{ 15, "IP"},						// device IP address
				{ 16, "Ports"},					// device Ports
				{ 17, "Devices Count"},			// number of Devices for this type (part of a group)

				{ 18, "PosX"},						// Shape position X
				{ 19, "PosY"},						// shape position Y
				{ 20, "Width"},					// shape width
				{ 21, "Height"},					// shape height
				{ 22, "Fill Color"},          // color to fill stincel

				{ 23, "Connect From"},        // used to link this visio shape to another visio shape
				{ 24, "From Line Label"},      // Arrow Text
				{ 25, "From Line Pattern"},    // Line pattern solid = 1, 2 = dashed, 3=Dotted, 4=Dash_Dot
				{ 26, "From Arrow Type"},      // Can contain one of these [None, Start, End, Both]
				{ 27, "From Line Color"},      // Arrow Color

				{ 28, "Connect To"},          // used to link this visio shape to another visio shape
				{ 29, "To Line Label"},        // Arrow Text
				{ 30, "To Line Pattern"},      // Line pattern solid = 1
				{ 31, "To Arrow Type"},        // Can contain one of these [None, Start, End, Both]
				{ 32, "To Line Color"}			// Arrow Color
			};
			return excelHeaderNames;
		}
	}
}
