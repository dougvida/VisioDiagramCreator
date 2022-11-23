//using Newtonsoft.Json;
//using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OmnicellBlueprintingTool.Models
{
	public class AppConfiguration
	{

		public string Version { get; set; }

		public Dictionary<string,string> ColorsMap { get; set; }
		public List<string> Colors { get; set; }
		public List<string> Arrows { get; set; }
		public List<string> LinePatterns { get; set; }
		public List<string> StencilLabelPosition { get; set; }
		public List<string> ShapeTypes { get; set; }
		public List<string> LabelFontSizes { get; set; }
		public List<string> StencilNames { get; set; }
	}
}
