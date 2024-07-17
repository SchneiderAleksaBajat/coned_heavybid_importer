using System;
using System.Configuration;
using System.Threading.Tasks;
using ConEd.HeavyBid.Importer.Functionality;
using SE.Geospatial.Designer.CatalogService.Contracts;

namespace ConEd.HeavyBid.Importer
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("HB & Activity Code Importer V1");

            try
            {
                IHandler handler = new HbResourceImportHandler();
                await handler.Do();                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            
            Console.ReadLine();
        }
    }
}
