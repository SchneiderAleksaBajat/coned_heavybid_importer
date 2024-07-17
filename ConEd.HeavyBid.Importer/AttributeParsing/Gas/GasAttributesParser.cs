using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SE.PS.Azure.Data.Messages;

namespace ConEd.HeavyBid.Importer.AttributeParsing.Gas
{
	public class GasAttributesParser : ICuAttributesParser
	{
		private Dictionary<string, string> operations;
		private Dictionary<string, string> materials;
		private Dictionary<string, double> sizes;
		private Dictionary<string, string> covers;
		private Dictionary<string, string> breakPoints;
		private Dictionary<string, string> pressures;
		private Dictionary<string, string> weldingFusion;
		public GasAttributesParser()
		{
			InitializeAttributesLibrary();
		}

		public CuAttribute CreateTableNameAttribute(List<string> values)
		{
			throw new NotImplementedException();
		}

		public List<CuAttribute> ParseAttributes(string description)
		{
			List<CuAttribute> cuAttributes = new List<CuAttribute>();

			string[] attributes = description.Split('-');

			foreach (string attribute in attributes)
			{
				if (attribute == "M")
				{
					continue;
				}

				if (attribute == "MAIN" && description.Contains("CORROSION"))
				{
					CuAttribute cuAttribute = new CuAttribute
					{
						Key = "METHODINSTALLED",
						Value = "Corrosion Work on a Main Install"
					};

					cuAttributes.Add(cuAttribute);
					continue;
				}

				if (attribute == "SVC" && description.Contains("CORROSION"))
				{
					CuAttribute cuAttribute = new CuAttribute
					{
						Key = "INSTALLATIONMETHOD",
						Value = "Corrosion Work on a Service Install"
					};

					cuAttributes.Add(cuAttribute);
					continue;
				}

				string value;
				if (operations.TryGetValue(attribute, out value))
				{
					string name = IsServiceOperation(attribute) ? "INSTALLATIONMETHOD" : "METHODINSTALLED";
					CuAttribute cuAttribute = new CuAttribute
					{
						Key = name,
						Value = value
					};

					cuAttributes.Add(cuAttribute);
					continue;
				}

				double size;
				if (double.TryParse(attribute, out size) || sizes.TryGetValue(attribute, out size))
				{
					string unit = size == 1 ? "INCH" : "INCHES";
					CuAttribute cuAttribute = SetSize(cuAttributes, size, description);

					cuAttributes.Add(cuAttribute);
					continue;
				}


				//if (sizes.TryGetValue(attribute, out value))
				//{
				//	CuAttribute cuAttribute = SetSize(cuAttributes, value);

				//	cuAttributes.Add(cuAttribute);
				//	continue;
				//}

				if (materials.TryGetValue(attribute, out value))
				{
					CuAttribute cuAttribute = SetMaterial(cuAttributes, attribute, description);

					cuAttributes.Add(cuAttribute);
					continue;
				}

				if (covers.TryGetValue(attribute, out value))
				{
					string name = IsServiceOperation(description) ? "COVERTYPE" : "PIPECOVER";
					CuAttribute cuAttribute = new CuAttribute
					{
						Key = name,
						Value = value
					};

					cuAttributes.Add(cuAttribute);
					continue;
				}

				if (breakPoints.TryGetValue(attribute, out value))
				{
					CuAttribute cuAttribute = new CuAttribute
					{
						Key = "Break Point",
						Value = value
					};

					cuAttributes.Add(cuAttribute);
					continue;
				}

				if (pressures.TryGetValue(attribute, out value))
				{
					string name = description.Contains("MVIS") || description.Contains("TPT") || description.Contains("MSET") ? "PRESSURECLASS" : "PIPEPRESSURE";
					CuAttribute cuAttribute = new CuAttribute
					{
						Key = name,
						Value = value
					};

					cuAttributes.Add(cuAttribute);
					continue;
				}

				if (weldingFusion.TryGetValue(attribute, out value))
				{
					CuAttribute cuAttribute = new CuAttribute
					{
						Key = "Weld/Fuse",
						Value = value
					};

					cuAttributes.Add(cuAttribute);
					continue;
				}

				continue;
			}

			return cuAttributes;
		}

		private bool IsServiceOperation(string id)
		{
			return id.Contains("SIS") || id.Contains("SIR") || id.Contains("SX") || id.Contains("SSLV") || id.Contains("SVPGCTCP");
		}

		private CuAttribute SetSize(List<CuAttribute> attributes, double size, string id)
		{
			CuAttribute operation = attributes.Where(x => x.Key == "Operation").FirstOrDefault();
			string name = id.Contains("MVIS") ? "VALVESIZE" : "PIPESIZE";

			if (operation == null || operation.Value != "Tie In")
			{
				return new CuAttribute
				{
					Key = name,
					Value = size
				};
			}

			CuAttribute mainSize = attributes.Where(x => x.Key == "Existing Main Size").FirstOrDefault();
			if (mainSize == null)
			{
				return new CuAttribute
				{
					Key = "Existing Main Size",
					Value = size
				};
			}

			return new CuAttribute
			{
				Key = "Tie In Main Size",
				Value = size
			};
		}

		private CuAttribute SetMaterial(List<CuAttribute> attributes, string material, string id)
		{
			CuAttribute operation = attributes.Where(x => x.Key == "Operation").FirstOrDefault();
			string name = id.Contains("MVIS") ? "VALVEMATERIAL" : "PIPEMATERIAL";

			if (operation == null || operation.Value != "Tie In")
			{
				return new CuAttribute
				{
					Key = name,
					Value = material
				};
			}

			CuAttribute mainSize = attributes.Where(x => x.Key == "Existing Main Material").FirstOrDefault();
			if (mainSize == null)
			{
				return new CuAttribute
				{
					Key = "Existing Main Material",
					Value = material
				};
			}

			return new CuAttribute
			{
				Key = "Tie In Main Material",
				Value = material
			};
		}

		private void InitializeOperationDict()
		{
			operations = new Dictionary<string, string>();

			operations.Add("SIS", "Service installation");
			operations.Add("SIR", "Service Insertion");
			operations.Add("MIS", "Main Installation");
			operations.Add("MIR", "Main Insertion");
			operations.Add("MVIS", "Main Valve Installation");
			operations.Add("TPT", "Tapping Tee");
			operations.Add("TI", "Tie In");
			operations.Add("MNOFFSET", "Main Offset");
			operations.Add("SXNMN", "Service Tranfer on New Main because of Main Replacement");
			operations.Add("SXEMN", "Service Transfer on Existing Main because of a Main Cut out or a Main Insertion");
			operations.Add("WELD", "Weld");
			operations.Add("FUSE", "Fuse");
			operations.Add("MSLV", "Main Repair Sleeve Installation");
			operations.Add("SSLV", "Service Repair Sleeve Installation");
			operations.Add("MSET", "Meter/Regulator Set Installation");
			operations.Add("SVPGCTCP", "Service Purge Cut and Cap");
			operations.Add("MNPGCTCP", "Main Purge Cut and Cap");
			operations.Add("MAIN-CORROSION", "Corrosion Work on a Main Install");
			operations.Add("SVC-CORROSION", "Corrosion Work on a Service Install");
		}

		private void InitializeMaterialsDict()
		{
			materials = new Dictionary<string, string>();

			materials.Add("PE", "Polyethylene");
			materials.Add("ST", "Steel");
			materials.Add("CT", "Copper Tubing");
		}

		private void InitializeSizesDict()
		{
			sizes = new Dictionary<string, double>();

			sizes.Add("1.25CTS", 1.25);
			sizes.Add("1.25IPS", 1.25);
		}

		private void InitializeCoversDict()
		{
			covers = new Dictionary<string, string>();

			covers.Add("ASP", "Asphalt");
			covers.Add("CON", "Concrete");
			covers.Add("EARTH", "Earth");
			covers.Add("SDW", "Sidewalk");
			covers.Add("RDW", "RoadWay");
			covers.Add("OT", "Open Trench, i.e., no Excavation / Backfill required");
		}

		private void InitializeBreakPointsDict()
		{
			breakPoints = new Dictionary<string, string>();

			breakPoints.Add("ML", "Manhattan Less than or Equal to breakpoint of 30 feet");
			breakPoints.Add("MG", "Manhattan Greater than breakpoint of 30 feet");
			breakPoints.Add("QXL", "Queens or Bronx Less than or Equal to breakpoint of 50 feet");
			breakPoints.Add("QXG", "Queens or Bronx Greater than breakpoint of 50 feet");
			breakPoints.Add("WL", "Westchester Less than or Equal to breakpoint of 65 feet");
			breakPoints.Add("WG", "Westchester Greater than breakpoint of 65 feet");
		}

		private void InitializePressuresDict()
		{
			pressures = new Dictionary<string, string>();

			pressures.Add("LP", "2");
			pressures.Add("HP", "5");
			pressures.Add("MP", "4");
		}

		private void InitializeWeldingFusionDict()
		{
			weldingFusion = new Dictionary<string, string>();

			weldingFusion.Add("WLDMN", "Weld on Main");
			weldingFusion.Add("WLDSLV", "Weld on Sleeve");
			weldingFusion.Add("FUSEMN", "Fuse on Main");
			weldingFusion.Add("FUSESLV", "Fuse on Sleeve");
		}


		private void InitializeAttributesLibrary()
		{
			InitializeOperationDict();
			InitializeMaterialsDict();
			InitializeSizesDict();
			InitializeCoversDict();
			InitializeBreakPointsDict();
			InitializePressuresDict();
			InitializeWeldingFusionDict();
		}

	}
}
