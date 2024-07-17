using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using ConEd.HeavyBid.Importer.CuLibraryProcessing;
using Newtonsoft.Json;
using SE.PS.Azure.Data.Messages;
using SE.PS.WmsIntegration.Contracts.Parameters;
using ConEd.HeavyBid.Importer.Utility;
using SE.Geospatial.Designer.CatalogService.Contracts;

namespace ConEd.HeavyBid.Importer.Functionality
{
    public class HbResourceImportHandler : IHandler
    {
        private readonly ConfigReader _configReader;
        private readonly CuLibraryProcessor _processor;
        private readonly HeavyBidProcessing.HeavyBid _heavyBid;
        private readonly WmsiClientFactory _wmsiClientFactory;

        public HbResourceImportHandler()
        {
            _configReader = new ConfigReader();
            _processor = new ActivityCodeCuLibraryProcessor();
            _wmsiClientFactory = new WmsiClientFactory();
            _heavyBid = new HeavyBidProcessing.HeavyBid();
        }

        public async Task Do()
        {
            string input;
            bool validInput;
            do
            {
                Console.WriteLine("Choose import type:");
                Console.WriteLine("'1': 'Activity Codes'\n'2': 'Resource Codes'\n'3': 'Resource Codes' and 'Activity Codes'");
                input = Console.ReadLine();

                validInput = (input == "1" || input == "2" || input == "3");
                if (!validInput)
                {
                    Console.WriteLine("Invalid input. Try again.");
                }

            } while (!validInput);


            bool attachCodes = input == "1" || input == "3";
            var path = _configReader.GetActivityCodesPath();
            List<CuMessage> messages = _processor.GenerateCuSyncMessage(path, new List<string> {"*"}, new List<string> { "*" }, attachCodes);            

            if (input == "2" || input == "3") _heavyBid.InjectResourceCodes(messages);

            BulkCuUploadParameters parameters = new BulkCuUploadParameters
            {
                SessionId = Guid.NewGuid().ToString(),
                Cus = messages
            };
            
            string output = JsonConvert.SerializeObject(parameters);
            await File.WriteAllTextAsync("./output.json", output);
            Console.WriteLine("Output JSON is generated. Please check if it matches expected results and press enter if it does.");
            Console.ReadLine();

            var _wmsiClient = _wmsiClientFactory.Get();
            await _wmsiClient.BulkUpload.CreateAsync(parameters).ConfigureAwait(false);
            Console.WriteLine("Import finished.");
        }
    }
}
