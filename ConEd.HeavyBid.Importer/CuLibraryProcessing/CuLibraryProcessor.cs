using SE.PS.Azure.Data.Data;
using SE.PS.Azure.Data.Messages;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using SE.Geospatial.Designer.CatalogService.Contracts;
using Excel = Microsoft.Office.Interop.Excel;


namespace ConEd.HeavyBid.Importer.CuLibraryProcessing
{
    public abstract class CuLibraryProcessor : ICuLibraryProcessor
    {
        public CuLibraryProcessor()
        {

        }

        public List<CuMessage> GenerateCuSyncMessage(string path, List<string> typesForProcessing, List<string> tables, bool attachCodes)
        {
            if (!File.Exists(path))
            {
                throw new IOException($"{path} does not exist.");
            }

            return CreateCus(typesForProcessing, tables, path, attachCodes);
        }

        public abstract List<CuMessage> CreateCus(List<string> typesForProcessing, List<string> tables, string path, bool attachCodes);

        protected string GetRowValue(Excel.Range cells, int columnIndex, int rowIndex)
        {
            Excel.Range row = (Excel.Range)cells[rowIndex, columnIndex];
            var value = row.Value;

            if (value != null)
            {
                return value.ToString();
            }

            Marshal.ReleaseComObject(row);

            return (string)value;
        }

    }
}
