using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

using ConEd.HeavyBid.Importer.Functionality;

using SE.Geospatial.Services.Common.SDK;
using SE.Geospatial.Services.Common.SDK.Authentication;
using SE.Geospatial.Services.Tenancy.Contracts.Resources;
using SE.Geospatial.Services.Tenancy.SDK;
using SE.Geospatial.Services.Tenancy.SDK.Authentication;
using SE.Geospatial.Services.Tenancy.SDK.Operations;
using SE.PS.WmsIntegration.SDK;

namespace ConEd.HeavyBid.Importer.Utility
{
    public class WmsiClientFactory
    {
        private readonly ConfigReader _configReader;
        public WmsiClientFactory()
        {
            _configReader = new ConfigReader();
        }

        public WmsIntegrationServicesClient Get()
        {
            TenancyServiceClient tenancyClient = new TenancyServiceClient(null, new Config
            {
                BaseUrl = WellknownTenancyServiceUrl.Production
            });

            tenancyClient.UseAnonymous();

            TenantInfoResource tenantInfo = tenancyClient.Tenants.GetPublic(ConfigurationManager.AppSettings["TenantId"]);
            var serviceName = _configReader.GetServiceName();
            var wmsiService = tenantInfo.Services.FirstOrDefault(x => x.Name == serviceName);

            if(wmsiService == null)
            {
                var tenantService = string.Join(',', tenantInfo.Services.Select(x => x.Name).ToArray());
                throw new Exception($"Tenant does not have a service called '{serviceName}'. Here's the list of possible entries: {tenantService}.");
            }

            WmsIntegrationServicesClient wmsiClient = new WmsIntegrationServicesClient(new Config { BaseUrl = wmsiService.Url.ToString() });

            Console.WriteLine($"Authenticating to {wmsiService.Url}");

            SE.Geospatial.Services.Common.Security.IAuthenticator authenticator = tenantInfo.CreateAuthenticator(ConfigurationManager.AppSettings["ApiKeyClientId"], ConfigurationManager.AppSettings["ApiKeySecret"]);
            authenticator.Authenticate();

            wmsiClient.UseAuthenticator(authenticator);

            return wmsiClient;
        }
    }
}
