using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Text;
using SE.Geospatial.Designer.CatalogService.Contracts;

namespace ConEd.HeavyBid.Importer.Functionality
{
    public class ConfigReader
    {
        public string GetServiceName()
        {
            var key = ConfigurationManager.AppSettings["ServiceName"];

            if(key == null)
            {
                return "WmsIntegration";
            }

            return key;
        }


        public List<string> GetTypesForProcessing()
        {
            var key = ConfigurationManager.AppSettings["TypesForProcessing"];
            if (key == null)
            {
                return new List<string>();
            }

            return new List<string>(ConfigurationManager.AppSettings["TypesForProcessing"].Split(new char[] { ',' }));
        }

        public List<string> GetTablesForProcessing()
        {
            var key = ConfigurationManager.AppSettings["GisTableName"];
            if (key == null)
            {
                return new List<string>();
            }

            return new List<string>(ConfigurationManager.AppSettings["GisTableName"].Split(new char[] { ',' }));
        }

        public string GetActivityCodesPath()
        {
            var path = ConfigurationManager.AppSettings["ActivityCodes"];               

            if (!File.Exists(path))
            {
                throw new FileNotFoundException($"'Activity Codes' file is not found in path: {path}");
            }

            return path;
        }

        public string GetResourceCodesPath()
        {
            var path = ConfigurationManager.AppSettings["ResourceCodes"];            

            if (!File.Exists(path))
            {
                throw new FileNotFoundException($"'Resource Codes' file is not found in path: {path}");
            }

            return path;
        }

        public string GetMuToCuPath(ResourceDomain domain)
        {
            var path = ConfigurationManager.AppSettings["MuToCu"];
                
            if (!File.Exists(path))
            {
                throw new FileNotFoundException($"Cu Library file is not found in path: {path}");
            }

            return path;
        }

        public ResourceDomain GetDomain()
        {
            var domain = ConfigurationManager.AppSettings["Domain"];

            if (domain == null)
            {
                throw new KeyNotFoundException("Domain was not found.");
            }

            domain = domain.ToLower();

            if (domain == "gas")
            {
                return ResourceDomain.Gas;
            }
            if (domain == "electric")
            {
                return ResourceDomain.Electric;
            }
            

            throw new Exception("Invalid key");
        }
    }
}
