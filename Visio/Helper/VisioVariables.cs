using OmnicellBlueprintingTool.Models;
using System;
using System.Collections.Generic;
using System.Net.Http.Headers;

namespace OmnicellBlueprintingTool.Visio
{
	public class VisioVariables
	{
		public static string DefaultBlueprintingTemplateFile = "OC_BlueprintingTemplate.vstx";
		public static string DefaultBlueprintingStencilFile = "OC_BlueprintingStencils.vssx";
		private static StringComparer comparer = StringComparer.OrdinalIgnoreCase;

		private static List<string> _shapeTypes = null;
		private static List<string> _connectorArrows = null;
		private static List<string> _connectorLinePatterns = null;
		private static List<string> _stencilLabelPositions = null;
		private static List<string> _stencilLabelFontSizes = null;
		private static List<string> _connectorLineWeights = null;
		private static List<string> _defaultStencilNames = null;

		private static Dictionary<string, string> _visioColorsMap = null; // new Dictionary<string, string>(comparer); 

		public VisioVariables()
		{
			setupVisioColorsMap();
		}

		public const double HEIGHT = 0.25;

		// connector ends
		public const double BEGIN_ARROW = 4;      // Filled arrow
		public const double END_ARROW = 4;        // Filled arrow
		public const double ARROW_NONE = 0;       // None
		public const string sARROW_NONE = "None";
		public const string sARROW_START = "Start";
		public const string sARROW_END = "End";
		public const string sARROW_BOTH = "Both";

		// connector weight (default is LINE_WEIGHT_1
		public const string sLINE_WEIGHT_1 = "1 pt";
		//public const string sLINE_WEIGHT_1_5 = "1.5 pt";
		//public const string sLINE_WEIGHT_2 = "2 pt";

		// connector corner 
		public const double ROUNDING = 0.0625;

		// color string names
		public const string sCOLOR_BLACK = "Black";

		// connector line pattern
		public const double SHDW_PATTERN = 0;           // None

		public const string sLINE_PATTERN_SOLID = "Solid";
		public const double LINE_PATTERN_SOLID = 1;     // ______ solid line
		public const string sLINE_PATTERN_DASHED = "Dashed";
		public const double LINE_PATTERN_DASH = 2;      // _ _ _ _ dashed lined
		public const string sLINE_PATTERN_DOTTED = "Dotted";
		public const double LINE_PATTERN_DOTTED = 3;    // . . . . dotted line
		public const string sLINE_PATTERN_DASHDOT = "Dash_Dot";
		public const double LINE_PATTERN_DASHDOT = 4;   // _ . _ . _ dash/dot-(t)ed line

		public const string STINCEL_LABEL_POSITION_TOP = "Top";
		public const string STINCEL_LABEL_POSITION_BOTTOM = "Bottom";

		public enum FormulaUse
		{
			Value,
			Error,
			Hide,
			Match
		}

		public enum ShowDiagram
		{
			NoShow = 0,
			Show = 1
		}

		public enum VisioPageOrientation
		{
			Landscape,
			Portrait
		}

		public enum VisioPageSize
		{
			Letter,
			Tabloid,
			Ledger,
			Legal,
			A3,
			A4
		}


		public static VisioPageOrientation GetVisioPageOrientation(string pgOr)
		{
			switch (pgOr.Trim().ToUpper())
			{
				case "Portrait":
					return VisioPageOrientation.Landscape;

				case "Landscape":
				default:
					return VisioPageOrientation.Portrait;
			}
		}

		public static VisioPageSize GetVisioPageSize(string pgSz)
		{
			switch (pgSz.Trim().ToUpper())
			{
				case "TABLOID":
					return VisioPageSize.Tabloid;
				case "LEDGER":
					return VisioPageSize.Ledger;
				case "LEGAL":
					return VisioPageSize.Legal;
				case "A3":
					return VisioPageSize.A3;
				case "A4":
					return VisioPageSize.A4;

				case "LETTER":
				default:
					return VisioPageSize.Letter;
			}
		}

		public static List<string> GetShapeTypes()
		{
			if (_shapeTypes == null)
			{
				_shapeTypes = new List<string>();

				_shapeTypes.Add("");
				_shapeTypes.Add("Template");
				_shapeTypes.Add("Stencil");
				_shapeTypes.Add("Page Setup");
				_shapeTypes.Add("Shape");
			}
			return _shapeTypes;
		}
		public static List<string> GetConnectorArrows()
		{
			if (_connectorArrows == null)
			{
				_connectorArrows = new List<string>();

				_connectorArrows.Add("");
				_connectorArrows.Add("None");
				_connectorArrows.Add("Start");
				_connectorArrows.Add("End");
				_connectorArrows.Add("Both");
			}
			return _connectorArrows;
		}

		public static List<string> GetConnectorLinePatterns()
		{
			if (_connectorLinePatterns == null)
			{
				_connectorLinePatterns = new List<string>();

				_connectorLinePatterns.Add("");
				_connectorLinePatterns.Add("Solid");
				_connectorLinePatterns.Add("Dashed");
				_connectorLinePatterns.Add("Dotted");
				_connectorLinePatterns.Add("Dash_Dot");
			}
			return _connectorLinePatterns;
		}

		public static List<string> GetStencilLabelPositions()
		{
			if (_stencilLabelPositions == null)
			{
				_stencilLabelPositions = new List<string>();

				_stencilLabelPositions.Add("");
				_stencilLabelPositions.Add("Top");
				_stencilLabelPositions.Add("Bottom");
			}
			return _stencilLabelPositions;
		}

		public static List<string> GetStencilLabelFontSize()
		{
			if (_stencilLabelFontSizes == null)
			{
				_stencilLabelFontSizes = new List<string>();
				_stencilLabelFontSizes.Add("");
				_stencilLabelFontSizes.Add("6");
				_stencilLabelFontSizes.Add("6:B");
				_stencilLabelFontSizes.Add("8");
				_stencilLabelFontSizes.Add("8:B");
				_stencilLabelFontSizes.Add("9");
				_stencilLabelFontSizes.Add("9:B");
				_stencilLabelFontSizes.Add("10");
				_stencilLabelFontSizes.Add("10:B");
				_stencilLabelFontSizes.Add("11");
				_stencilLabelFontSizes.Add("11:B");
				_stencilLabelFontSizes.Add("12");
				_stencilLabelFontSizes.Add("12:B");
				_stencilLabelFontSizes.Add("14");
				_stencilLabelFontSizes.Add("14:B");
			}
			return _stencilLabelFontSizes;
		}
		public static List<string> GetConnectorLineWeights()
		{
			if (_connectorLineWeights == null)
			{
				_connectorLineWeights = new List<string>();

				_connectorLineWeights.Add("");
				_connectorLineWeights.Add("1 pt");
				_connectorLineWeights.Add("1.5 pt");
				_connectorLineWeights.Add("2 pt");
				_connectorLineWeights.Add("2.25 pt");
				_connectorLineWeights.Add("2.5 pt");
				_connectorLineWeights.Add("3 pt");
				_connectorLineWeights.Add("4 pt");
				_connectorLineWeights.Add("6 pt");
			}
			return _connectorLineWeights;
		}

		public static List<string> GetDefaultStencilNames()
		{
			if (_defaultStencilNames == null)
			{
				_defaultStencilNames = new List<string>();

				_defaultStencilNames.Add("");
				_defaultStencilNames.Add("OC_AIO");
				_defaultStencilNames.Add("OC_AWS");
				_defaultStencilNames.Add("OC_BlisterPackager");
				_defaultStencilNames.Add("OC_Carousel");
				_defaultStencilNames.Add("OC_Carousel2");
				_defaultStencilNames.Add("OC_CPM");
				_defaultStencilNames.Add("OC_CSM");
				_defaultStencilNames.Add("OC_CustomerService");
				_defaultStencilNames.Add("OC_DashOutline");
				_defaultStencilNames.Add("OC_Database");
				_defaultStencilNames.Add("OC_Database2");
				_defaultStencilNames.Add("OC_DB");
				_defaultStencilNames.Add("OC_Devices");
				_defaultStencilNames.Add("OC_EHRSystems");
				_defaultStencilNames.Add("OC_Ethernet");
				_defaultStencilNames.Add("OC_Ethernet2");
				_defaultStencilNames.Add("OC_Ethernet3");
				_defaultStencilNames.Add("OC_File");
				_defaultStencilNames.Add("OC_File2");
				_defaultStencilNames.Add("OC_Firewall");
				_defaultStencilNames.Add("OC_Footer");
				_defaultStencilNames.Add("OC_Group1");
				_defaultStencilNames.Add("OC_Group2");
				_defaultStencilNames.Add("OC_Group3");
				_defaultStencilNames.Add("OC_Group4");
				_defaultStencilNames.Add("OC_IconKey");
				_defaultStencilNames.Add("OC_IconKey2");
				_defaultStencilNames.Add("OC_IVX");
				_defaultStencilNames.Add("OC_Label1");
				_defaultStencilNames.Add("OC_LineLedgen");
				_defaultStencilNames.Add("OC_Logo");
				_defaultStencilNames.Add("OC_OISInterfaces");
				_defaultStencilNames.Add("OC_PC");
				_defaultStencilNames.Add("OC_PortsLDAP_info");
				_defaultStencilNames.Add("OC_Process");
				_defaultStencilNames.Add("OC_Rectangle");
				_defaultStencilNames.Add("OC_Rectangle2");
				_defaultStencilNames.Add("OC_Rectangle3R");
				_defaultStencilNames.Add("OC_SeeEmbedded");
				_defaultStencilNames.Add("OC_Server");
				_defaultStencilNames.Add("OC_Server2");
				_defaultStencilNames.Add("OC_Server3");
				_defaultStencilNames.Add("OC_Site");
				_defaultStencilNames.Add("OC_Square");
				_defaultStencilNames.Add("OC_Square2");
				_defaultStencilNames.Add("OC_Supply");
				_defaultStencilNames.Add("OC_TabelCell");
				_defaultStencilNames.Add("OC_TableCell2");
				_defaultStencilNames.Add("OC_TagLabeler");
				_defaultStencilNames.Add("OC_Title");
				_defaultStencilNames.Add("OC_vSuite_EMM");
				_defaultStencilNames.Add("OC_vSuiteCloud");
				_defaultStencilNames.Add("OC_XR2");
				_defaultStencilNames.Add("OC_XT");
				_defaultStencilNames.Add("OC_YuyamaPackager");
			}
			return _defaultStencilNames;
		}

		private static void setupVisioColorsMap()
		{
			_visioColorsMap = new Dictionary<string, string>(comparer); 
			//_visioColorsMap = new Dictionary<string, string>();

			// Visio colors
			//_visioColorsMap.Add("", "RGB(0,0,0)");
			_visioColorsMap.Add("Beige", "RGB(245,245,220)");
			_visioColorsMap.Add("Black", "RGB(0,0,0)");
			_visioColorsMap.Add("Blue Light", "RGB(220,234,252)");
			_visioColorsMap.Add("Blue Medium", "RGB(71,141,233)");  // groupbox1 color
			_visioColorsMap.Add("Blue Bright", "RGB(71,150,255)");
			_visioColorsMap.Add("Blue", "RGB(17, 117, 245)");
			_visioColorsMap.Add("Cyan", "RGB(0,255,255)");
			_visioColorsMap.Add("Gold", "RGB(255,215,0)");
			_visioColorsMap.Add("Silver", "RGB(232,232,232)");
			_visioColorsMap.Add("Gray Light", "RGB(216,216,216)");
			_visioColorsMap.Add("Gray Medium", "RGB(216,216,216)");
			_visioColorsMap.Add("Gray", "RGB(165,165,165)");
			_visioColorsMap.Add("Mint", "RGB(198,224,180)");
			_visioColorsMap.Add("Green Light", "RGB(240,247,236)");
			_visioColorsMap.Add("Green LightMedium", "RGB(162,226,171)");
			_visioColorsMap.Add("Green Medium", "RGB(88,183,100)");
			_visioColorsMap.Add("Green Lime", "RGB(50,205,50)");
			_visioColorsMap.Add("Lime", "RGB(0,255,0)");
			_visioColorsMap.Add("Green Sea", "RGB(60,179,113)");     // Omnicell header green
			_visioColorsMap.Add("Green Dark", "RGB(67,169,72)");
			_visioColorsMap.Add("Green", "RGB(13,164,33)");
			_visioColorsMap.Add("Peach", "RGB(255,242,204)");        // BD groupbox4 color
			_visioColorsMap.Add("Tan", "RGB(210,180,140)");
			_visioColorsMap.Add("Khaki", "RGB(240,230,140)");
			_visioColorsMap.Add("Khaki Dark", "RGB(189,183,107)");
			_visioColorsMap.Add("Brown", "RGB(210,105,30)");
			_visioColorsMap.Add("Magenta", "RGB(255,0,255)");
			_visioColorsMap.Add("Navy", "RGB(60,80,160)");
			_visioColorsMap.Add("Olive", "RGB(120,120,0)");
			_visioColorsMap.Add("Olive Drab", "RGB(100,140,30)");
			_visioColorsMap.Add("Orange Light", "RGB(255,210,120)");
			_visioColorsMap.Add("Orange Medium", "RGB(250,190,80)");
			_visioColorsMap.Add("Orange", "RGB(255,172,4)");
			_visioColorsMap.Add("Purple", "RGB(54,54,222)");
			_visioColorsMap.Add("Pink Light", "RGB(255,182,193)");
			_visioColorsMap.Add("Red", "RGB(255,0,0)");
			_visioColorsMap.Add("Salmon", "RGB(251,229,213)");
			_visioColorsMap.Add("Teal", "RGB(0,128,128)");
			_visioColorsMap.Add("White", "RGB(255,255,255)");
			_visioColorsMap.Add("White Smoke", "RGB(245,245,245)");
			_visioColorsMap.Add("Yellow", "RGB(255,255,0)");
		}

		/// <summary>
		/// GetRGBColor
		/// return the RGB color value based on the color string argument
		/// color argument "Black" will return "RGB(0,0,0)"
		/// </summary>
		/// <param name="color">search value</param>
		/// <returns>"RGB(???,???,???)"</returns>
		public static string GetRGBColor(string color)
		{
			string value = string.Empty;
			if (_visioColorsMap == null)
			{
				setupVisioColorsMap();
			}
			if (string.IsNullOrEmpty(color))
			{
				return null;
			}

			foreach(KeyValuePair<string, string> kvp in _visioColorsMap)
			{
				if (string.Equals(kvp.Key, color, StringComparison.OrdinalIgnoreCase))
				{
					return kvp.Value.Trim().ToString();
				}
			}
			return value;
		}

		/// <summary>
		/// GetColorValueFromDB
		/// return the color string based on the rgb value argument
		/// search "RGB(0,0,0)" will return "Black"
		/// </summary>
		/// <param name="rgb"></param>
		/// <returns>string</returns>
		/// <text>color name</text>
		public static string GetColorValueFromRGB(string rgb)
		{
			if (_visioColorsMap == null)
			{
				setupVisioColorsMap();
			}
			if (string.IsNullOrEmpty(rgb))
			{
				return null; // default to Black
			}
			foreach (KeyValuePair<string, string> item in _visioColorsMap)
			{
				if (item.Value.Equals(rgb.Trim()))
				{
					return item.Key;
				}
			}
			return null;
		}

		public static string[] GetAllColorKeyValues()
		{
			int nIdx = 0;
			string[] saTmp = new string[_visioColorsMap.Count+1];
			saTmp[nIdx++] = "";	// we need to add a blank as the first entry
			foreach (KeyValuePair<string, string> keyValue in _visioColorsMap)
			{
				// adjust the index to be minus 1 bacause we added a row outside the array
				saTmp[nIdx++] = keyValue.Key.Trim();
			}
			return saTmp;
		}

		/// <summary>
		/// GetConnectorLineWeight
		/// search the list for the paramater
		/// if found use that value as the To or From Line Weight value as a string
		/// if not found null will be returned so use the default value
		/// ignore case
		/// </summary>
		/// <param name="value">lookup</param>
		/// <returns>Found value or null</returns>
		public static string GetConnectorLineWeight(string value)
		{
			if (_connectorLineWeights == null)
			{
				GetConnectorLineWeights();
			}
			if (string.IsNullOrEmpty(value))
			{
				return ""; // Use default line weight.   "1 pt"
			}
			foreach (string item in _connectorLineWeights)
			{
				if (item.Equals(value.Trim(),StringComparison.OrdinalIgnoreCase))
				{
					return item;
				}
			}
			return "";
		}

		/// <summary>
		/// GetDefaultStencilName
		/// Search if stencil name exists as a default stencil
		/// ignore case
		/// </summary>
		/// <param name="name">stencil name</param>
		/// <returns>"" if not found else the stencil name</returns>
		public static string GetDefaultStencilName(string name)
		{
			if (_defaultStencilNames == null)
			{
				GetDefaultStencilNames();
			}
			if (string.IsNullOrEmpty(name))
			{
				return "";
			}
			foreach (string item in _defaultStencilNames)
			{
				if (item.Equals(name.Trim(), StringComparison.OrdinalIgnoreCase))
				{
					return item;
				}
			}
			return "";
		}

	}
}
