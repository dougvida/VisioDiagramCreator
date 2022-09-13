using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VisioDiagramCreator.Visio.Models
{
	public class ConnectShpExcelData
	{
		public string UniqueKey { get; set; } = string.Empty;
		public int ID { get; set; }

		public Dictionary<int, string> CntFrom { get; set; }
		public Dictionary<int, string> CntTo { get; set; }

		public HashSet<int> connFromHS = new HashSet<int>();
		public HashSet<int> connToHS = new HashSet<int>();
	}
}
