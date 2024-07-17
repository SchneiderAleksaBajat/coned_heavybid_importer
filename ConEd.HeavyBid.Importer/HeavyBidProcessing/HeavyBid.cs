using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using SE.PS.Azure.Data.Messages;
using Excel = Microsoft.Office.Interop.Excel;
using ConEd.HeavyBid.Importer.Utility;



namespace ConEd.HeavyBid.Importer.HeavyBidProcessing
{
    public class HeavyBid
    {
        private CuToSpec CuToSpec { get; set; } = new CuToSpec();
        public List<CuMessage> InjectResourceCodes(List<CuMessage> messages)
        {
            Dictionary<string, List<CuResourceCode>> resourceMapping;
            Dictionary<string, string> muToCuMapping;
            RcMapping rcMapping = new RcMapping();
            CuToSpec cuToSpec = new CuToSpec();

            using (ComReleaser comReleaser = new ComReleaser())
            {
                var resourceCodesTable = ExcelHelper.GetTable(comReleaser, ConfigurationManager.AppSettings["ResourceCodes"]);
                if (resourceCodesTable == null) return new List<CuMessage>();
                var (rowsRCT, colsRCT) = resourceCodesTable;
                resourceMapping = GetCuToResourceMapping(rowsRCT, colsRCT, cuToSpec);

                var cuIdToMuId = ExcelHelper.GetTable(comReleaser, ConfigurationManager.AppSettings["MuToCu"]);
                if (cuIdToMuId == null) return new List<CuMessage>();
                var (rowsCT, colsCT) = cuIdToMuId;
                muToCuMapping = GetMuToCuMapping(rowsCT, colsCT, cuToSpec);
            }

            int numberOfUpdates = 0;
            int numberOfTotal = 0;

            foreach (CuMessage msg in messages)
            {
                if (muToCuMapping.ContainsKey(msg.WmsCode))
                {
                    string cuId = muToCuMapping[msg.WmsCode];
                    if (resourceMapping.TryGetValue(cuId, out List<CuResourceCode> rcList))
                    {
                        CuResourceCode rc = TakeTheCorrectRc(rcList);
                        string resourceCode = TransformRc(rc, rcMapping, msg.AttributeArray);
                        if (resourceCode != null)
                        {
                            msg.AttributeArray.Add(new CuAttribute
                            {
                                Key = "HeavyBid Resource Code",
                                Value = resourceCode
                            });
                            numberOfUpdates++;
                        }
                        numberOfTotal++;
                    }
                }
            }

//            Console.WriteLine($"Processed: '{numberOfTotal}' . Updated: '{numberOfUpdates}'");

            return messages;
        }

        private Dictionary<string, string> GetMuToCuMapping(Excel.Range rowRange, Excel.Range colRange, CuToSpec cuToSpec)
        {
            Dictionary<string, string> mapping = new Dictionary<string, string>();

            for (int i = 2; i < rowRange.Count; i++)
            {                
                string muId = ExcelHelper.GetCellValue(colRange, i, 1);
                string cuId = ExcelHelper.GetCellValue(colRange, i, 2);

                Console.WriteLine($"{i}. Creating macro to cu mapping: {muId} - {cuId}.");

                if (muId == null || cuId == null) break;

                if (cuToSpec.IsDrivingCu(cuId))
                {
                    mapping[muId] = cuId;
                }
            }

            return mapping;
        }

        private Dictionary<string, List<CuResourceCode>> GetCuToResourceMapping(Excel.Range rowRange, Excel.Range colRange, CuToSpec cuToSpec)
        {
            Dictionary<string, List<CuResourceCode>> mapping = new Dictionary<string, List<CuResourceCode>>();

            bool doesTptExist = false;

            for (int i = 2; i < rowRange.Count; i++)
            {                
                string cuId = ExcelHelper.GetCellValue(colRange, i, 2);
                Console.WriteLine($"{i}. Creating cu to resource mapping for {cuId}.");
                if (cuId == null)
                {
                    break;
                }
                string lpResourceCode = ExcelHelper.GetCellValue(colRange, i, ExcelHelper.GetColumnNumber("AJ"));
                string hpResourceCode = ExcelHelper.GetCellValue(colRange, i, ExcelHelper.GetColumnNumber("AK"));
                string lpColor = ExcelHelper.GetCellColor(colRange, i, ExcelHelper.GetColumnNumber("AJ"));
                string hpColor = ExcelHelper.GetCellColor(colRange, i, ExcelHelper.GetColumnNumber("AK"));
                CuResourceCode resourceCode = new CuResourceCode(cuId, lpResourceCode, hpResourceCode, lpColor, hpColor);

                if (mapping.TryGetValue(cuId, out var rcList))
                {
                    bool unique = true;
                    foreach (CuResourceCode rc in rcList)
                    {
                        if (resourceCode.Equals(rc))
                        {
                            unique = false;
                            break;
                        }
                    }
                    if (unique && resourceCode.Relevant() && cuToSpec.IsDrivingCu(cuId)) rcList.Add(resourceCode);
                }
                else
                {
                    if (resourceCode.Relevant() && cuToSpec.IsDrivingCu(cuId))
                    {
                        mapping.Add(cuId, new List<CuResourceCode>
                        {
                            resourceCode
                        });
                    }
                }
            }

            return mapping;
        }


        private CuResourceCode TakeTheCorrectRc(List<CuResourceCode> rcList)
        {
            return rcList[0];
        }

        private string TransformRc(CuResourceCode rc, RcMapping rcMapping, List<CuAttribute> attributes)
        {

            var rcData = rc.GetRcAndColor();

            string rcString = rcData.Item1;
            string rcColor = rcData.Item2;

            if (rcColor == ExcelColors.White)
            {
                return rcString;
            }

            if (rcColor == ExcelColors.Yellow)
            {
                if (rcString.Last() == '*')
                {
                    return rcString.Split('-')[0];
                }

                if (rcString.Last() == '#')
                {
                    return rcMapping.ReplaceRcPlaceholder(Placeholder.HashTag, rc.CuId, rcString, attributes);
                }

                if (rcString.Contains("X"))
                {
                    return rcMapping.ReplaceRcPlaceholder(Placeholder.X, rc.CuId, rcString, attributes);
                }
            }

            return null;
        }
    }
}
