using System;
using System.Collections.Generic;
using System.Text;
using SE.PS.Azure.Data.Messages;

namespace ConEd.HeavyBid.Importer.CuLibraryProcessing
{
    public interface ICuLibraryProcessor
    {
        List<CuMessage> GenerateCuSyncMessage(string path, List<string> typesForProcessing, List<string> tables, bool attachCodes);
    }
}
