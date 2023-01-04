using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OIS.Models
{
	public class OISSetupData
	{
		// Index,Type,Path,Node,Desc,Details,RegExpr
		public int Index { get; set; } = -1;
		public string Type { get; set; } = string.Empty;
		public string Path { get; set; } = string.Empty;
		public string Node { get; set; } = string.Empty;
		public string Desc { get; set; } = string.Empty;
		public string Details { get; set; } = string.Empty;
		public string RegExpr { get; set; } = string.Empty;
		public string Label { get; set; } = string.Empty;

		//public string Direction { get; set; } = string.Empty;
		//public string Stencil { get; set; } = string.Empty;
	}
}
