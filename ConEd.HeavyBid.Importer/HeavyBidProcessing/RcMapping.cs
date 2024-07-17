using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using ConEd.HeavyBid.Importer.Functionality;
using ConEd.HeavyBid.Importer.Utility;
using SE.Geospatial.Designer.CatalogService.Contracts;
using SE.PS.Azure.Data.Messages;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConEd.HeavyBid.Importer.HeavyBidProcessing
{
    public enum Placeholder
    {
        X,
        HashTag
    }
    public enum ReplacementType
    {
        TPT,
        Null
    }
    public class TapCode
    {
        public double TapSize { get; set; }

        public string Code { get; set; }
    }

    public class RcMapping
    {
        public Dictionary<double, List<TapCode>> TptMapping;
        public Dictionary<double, string> ServiceReconnect;

        private ConfigReader configReader;

        public RcMapping()
        {
            configReader = new ConfigReader();
            InitTptMapping();
            InitServiceReconnect();
        }

        public string ReplaceRcPlaceholder(Placeholder placeholder, string cuId, string rc, List<CuAttribute> attributes)
        {
            if (placeholder == Placeholder.HashTag)
            {
                ReplacementType replacementType = FindReplacementType(cuId);

                switch (replacementType)
                {
                    case ReplacementType.TPT:
                        return TransformTPT(rc, attributes);
                    case ReplacementType.Null:
                        return rc;
                }
            }

            if (placeholder == Placeholder.X)
            {
                var attr = attributes;
            }

            return rc;
        }

        private ReplacementType FindReplacementType(string cuId)
        {
            if (cuId.Contains("TPT"))
            {
                return ReplacementType.TPT;
            }
            return ReplacementType.Null;
        }

        private string TransformTPT(string rc, List<CuAttribute> attributes)
        {
            return rc;
        }

        private void InitServiceReconnect()
        {
            ServiceReconnect = new Dictionary<double, string>
            {
                { 0, "A" },
                { 3, "B" },
                { 6, "C" }
            };
        }

        private void InitTptMapping()
        {
            TptMapping = new Dictionary<double, List<TapCode>>();
            using (ComReleaser releaser = new ComReleaser())
            {
                var resourceCodespath = configReader.GetResourceCodesPath();

                Tuple<Excel.Range, Excel.Range> ranges = ExcelHelper.GetTable(releaser, resourceCodespath, 10);

                for (int i = 2; i < ranges.Item1.Count; i++)
                {
                    string mainSize = ExcelHelper.GetCellValue(ranges.Item2, i, ExcelHelper.GetColumnNumber("A"));
                    if (mainSize == null)
                    {
                        break;
                    }
                    string tapSize = ExcelHelper.GetCellValue(ranges.Item2, i, ExcelHelper.GetColumnNumber("B"));
                    string code = ExcelHelper.GetCellValue(ranges.Item2, i, ExcelHelper.GetColumnNumber("C"));

                    string[] splitResult = tapSize.Split('\"', StringSplitOptions.RemoveEmptyEntries);                    
                    if (double.TryParse(splitResult[0], out double iTapSize) && double.TryParse(splitResult[0], out double iMainSize))
                    {
                        TapCode tapCode = new TapCode
                        {
                            Code = code,
                            TapSize = iTapSize
                        };

                        if (TptMapping.TryGetValue(iMainSize, out List<TapCode> tcList))
                        {
                            tcList.Add(tapCode);
                        }
                        else
                        {
                            TptMapping[iMainSize] = new List<TapCode> { tapCode };
                        }
                    }
                }
            }
        }
    }
}
