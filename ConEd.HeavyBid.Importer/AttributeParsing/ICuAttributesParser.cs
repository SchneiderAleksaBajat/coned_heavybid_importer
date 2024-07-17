using System;
using System.Collections.Generic;
using System.Text;
using SE.PS.Azure.Data.Messages;

namespace ConEd.HeavyBid.Importer.AttributeParsing
{
    public interface ICuAttributesParser
    {
        List<CuAttribute> ParseAttributes(string description);

        CuAttribute CreateTableNameAttribute(List<string> values);
    }
}
