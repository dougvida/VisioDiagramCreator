using VisioDiagramCreator.Visio;

namespace VisioDiagramCreator.Models
{
	public class Device
	{
		public string MachineName { get; set; }
		public string MachineId { get; set; }

		public string SiteId { get; set; }
		public string SiteName { get; set; }
		public string SiteAddress { get; set; }
		public string SiteType { get; set; }

		public string OmniType { get; set; }
		public string OmniId { get; set; }
		public string SiteId_OmniId { get; set; }
		public string OmniArea { get; set; }
		public string OmniName { get; set; }
		public string OmniVersion { get; set; }
		public string OmniIP { get; set; }
		public string OmniPorts { get; set; }
		public string OmniStatus { get; set; }

		// Visio diagram section
		public ShapeInformation ShapeInfo { get; set; }  // Visio diagram Shape info/and connections
	}
}
