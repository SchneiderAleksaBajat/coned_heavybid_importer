using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using ConEd.HeavyBid.Importer.AttributeParsing;
using ConEd.HeavyBid.Importer.AttributeParsing.Gas;
using ConEd.HeavyBid.Importer.Functionality;
using SE.Geospatial.Designer.CatalogService.Contracts;
using SE.PS.Azure.Data.Messages;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConEd.HeavyBid.Importer.CuLibraryProcessing
{
	public class ActivityCodeCuLibraryProcessor : CuLibraryProcessor
	{
        private readonly Dictionary<string, CuMessage> wmsCodeToCu;
        private readonly ConfigReader configReader;
        private readonly ICuAttributesParser gasAttributesParser;        
        private readonly ResourceDomain domain;
		public ActivityCodeCuLibraryProcessor() : base()
		{
			wmsCodeToCu = new Dictionary<string, CuMessage>();
            configReader = new ConfigReader();
			domain = ResourceDomain.Gas;
            gasAttributesParser = new GasAttributesParser();            
        }

		public override List<CuMessage> CreateCus(List<string> typesForProcessing, List<string> tables, string path, bool attachCodes)
		{
			Excel.Application xlApp = new Excel.Application();
			Excel.Workbooks wBooks = xlApp.Workbooks;			
			Excel.Workbook xlWorkbook = wBooks.Open(path);
			Excel.Sheets workSheets = xlApp.Worksheets;
			Excel.Worksheet cuWorksheet = (Excel.Worksheet)workSheets[1];//Data
			Excel.Range cuWorksheetXlRange = cuWorksheet.Rows;
			Excel.Range cuWorksheetCells = cuWorksheet.Cells;

			List<CuMessage> cus = ProcessDataSpreadsheet(cuWorksheetXlRange, cuWorksheetCells, attachCodes);

			Marshal.ReleaseComObject(cuWorksheetCells);
			Marshal.ReleaseComObject(cuWorksheetXlRange);
			Marshal.ReleaseComObject(cuWorksheet);

			Marshal.ReleaseComObject(workSheets);

			xlWorkbook.Close();
			Marshal.ReleaseComObject(xlWorkbook);

			wBooks.Close();
			Marshal.ReleaseComObject(wBooks);

			xlApp.Quit();
			Marshal.ReleaseComObject(xlApp);

			GC.Collect();
			GC.WaitForPendingFinalizers();

			return cus;
		}

		private List<CuMessage> ProcessDataSpreadsheet(Excel.Range rowRange, Excel.Range colRange, bool attachCodes)
		{
			for (int i = 2; i <= rowRange.Count; i++)
			{
				string muId = GetRowValue(colRange, 1, i);//mu id*

                if (muId == null)
				{
					break;
				}

                ICuAttributesParser attributesParser;

                attributesParser = gasAttributesParser;                                

                CuMessage cu = GetCu(muId, attributesParser);
				Console.WriteLine($"{i}. Loading: {muId}.");

                if (attachCodes)
                {
					string heavyBidUoM = GetRowValue(colRange, 12, i);
                    if (heavyBidUoM.ToLower() == "fot")
                    {
                        heavyBidUoM = "LF";
                    }
                    string heavyBidActivityCode = GetRowValue(colRange, 26, i);

                    CuAttribute hbActivityCode = new CuAttribute
                    {
                        Key = "HeavyBid Activity Code",
                        Value = heavyBidActivityCode
                    };

                    CuAttribute hbUoM = new CuAttribute
                    {
                        Key = "HeavyBid Activity Unit of Measure",
                        Value = heavyBidUoM
                    };
                    cu.AttributeArray.Add(hbActivityCode);
                    cu.AttributeArray.Add(hbUoM);
                    Console.WriteLine($"{i}-2. Assigned activity codes for {muId}: {heavyBidUoM} - {heavyBidActivityCode}.");
                }
            }

			return wmsCodeToCu.Select(x => x.Value).ToList();
		}

		private CuMessage GetCu(string wmsCode, ICuAttributesParser attributesParser)
		{
			CuMessage cu;
			if (wmsCodeToCu.TryGetValue(wmsCode, out cu))
			{
				return cu;
			}

            cu = new CuMessage();

            cu.WmsCode = wmsCode;
            cu.Description = cu.WmsCode;
			cu.Name = cu.WmsCode;
            cu.Domain = domain;
            cu.AvailableWorkFunctions = new string[] { "Install" };
			cu.Status = "Active";
			cu.IsMacro = true;

            cu.AttributeArray = attributesParser.ParseAttributes(wmsCode);

            wmsCodeToCu.Add(wmsCode, cu);

			return cu;
		}
	}
}
