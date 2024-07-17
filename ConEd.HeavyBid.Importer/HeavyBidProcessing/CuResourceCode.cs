using System;
using System.Collections.Generic;
using System.Text;
using ConEd.HeavyBid.Importer.Utility;

namespace ConEd.HeavyBid.Importer.HeavyBidProcessing
{
    public class CuResourceCode
    {
        public string CuId { get; set; }
        public string LpResource { get; set; }
        public string HpResource { get; set; }
        public string LpColor { get; set; }
        public string HpColor { get; set; }

        public CuResourceCode(string cuId, string lpResource, string hpResource, string lpColor, string hpColor)
        {
            CuId = cuId;
            LpResource = lpResource;
            HpResource = hpResource;
            LpColor = lpColor;
            HpColor = hpColor;
        }

        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType())
            {
                return false;
            }

            CuResourceCode other = (CuResourceCode)obj;
            return (LpResource == other.LpResource) && (HpResource == other.HpResource);
        }

        public Tuple<string, string> GetRcAndColor()
        {
            if (this.HpResource == this.LpResource)
            {
                return new Tuple<string, string>(this.HpResource, this.HpColor);
            }

            if (this.HpColor == this.LpColor)
            {
                return new Tuple<string, string>(this.HpResource, this.HpColor);
            }

            if (this.HpColor != ExcelColors.Pink)
            {
                return new Tuple<string, string>(this.HpResource, this.HpColor);
            }

            return new Tuple<string, string>(this.HpResource, this.LpColor);
        }

        public bool Relevant()
        {
            if (this.LpColor == ExcelColors.Pink && this.HpColor == ExcelColors.Pink)
            {
                return false;
            }

            return true;
        }
    }
}
