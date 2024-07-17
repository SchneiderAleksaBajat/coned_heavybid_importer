using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SE.Geospatial.Designer.CatalogService.Contracts;
using SE.Geospatial.Designer.CatalogService.SDK;
using SE.Geospatial.Services.Common.Extensions;
using SE.Geospatial.Services.Common.SDK;
using SE.Geospatial.Services.Common.SDK.Authentication;
using SE.Geospatial.Services.Tenancy.Contracts.Resources;
using SE.Geospatial.Services.Tenancy.SDK;
using SE.Geospatial.Services.Tenancy.SDK.Authentication;
using SE.Geospatial.Services.Tenancy.SDK.Operations;

namespace ConEd.HeavyBid.Importer.HeavyBidProcessing
{
    public class CuToSpec
    {
        private CatalogServiceClient _catalogServiceClient;
        private TenantInfoResource _tenantInfoResource;
        private Dictionary<Guid, Guid> _cuToSpec;
        private Dictionary<string, Guid> _cuToGuidMapping;

        private TenantInfoResource GetTenantInfoResource()
        {
            if (_tenantInfoResource == null)
            {
                var tenancyServiceClient = new TenancyServiceClient(new Config
                {
                    BaseUrl = WellknownTenancyServiceUrl.Production
                });
                tenancyServiceClient.UseAnonymous();
                string tenantId = ConfigurationManager.AppSettings["TenantId"];
                _tenantInfoResource = tenancyServiceClient.Tenants.GetPublic(tenantId);
            }

            return _tenantInfoResource;
        }

        private CatalogServiceClient GetCatalogServiceClient()
        {
            if (_catalogServiceClient == null)
            {
                var appSettings = ConfigurationManager.AppSettings;
                var tenantInfoResource = GetTenantInfoResource();
                var catalogService = tenantInfoResource.ServiceResource("DesignerEquipmentCatalog");
                _catalogServiceClient = new CatalogServiceClient(new Config
                {
                    BaseUrl = catalogService.Url.ToString()
                });
                //Use client credentials aka API key to connect
                var cca = new ClientCredentialsAuthenticator(
                    appSettings["ApiKeyClientId"],
                    appSettings["ApiKeySecret"],
                    catalogService.Auth0ApiIdentifier,
                    tenantInfoResource.Auth0Domain);
                _catalogServiceClient.UseAuthenticator(cca);
                cca.Authenticate();
            }

            return _catalogServiceClient;
        }

        public Dictionary<Guid, Guid> GetCuToSpec(ResourceDomain domain)
        {
            if (_cuToSpec == null)
            {
                _cuToSpec = new Dictionary<Guid, Guid>();
                var tenantInfoResource = GetTenantInfoResource();
                string tenantId = tenantInfoResource.TenantId;
                var catalog = GetCatalogServiceClient();

                var spec2Cus = catalog.Workflow.SelectAll<SpecToCuRelationshipResource>(domain, tenantId).GetAwaiter().GetResult();
                foreach (var rel in spec2Cus)
                {
                    if (!_cuToSpec.ContainsKey(rel.CuId))
                    {
                        _cuToSpec.Add(rel.CuId, rel.SpecId);
                    }
                }
            }

            return _cuToSpec;
        }

        private Dictionary<string, Guid> GetCuToGuidMapping()
        {
            if (_cuToGuidMapping == null)
            {
                _cuToGuidMapping = new Dictionary<string, Guid>();
                var catalog = GetCatalogServiceClient();
                var tenantInfo = GetTenantInfoResource();
                var response = catalog.Workflow.GetCache<CompatibleUnitResource>(ResourceDomain.Gas, null, tenantInfo.TenantId).ConfigureAwait(false).GetAwaiter().GetResult();
                var cuResources = response.Values.ToList();

                foreach (var resource in cuResources)
                {
                    _cuToGuidMapping.Add(resource.WmsCode, resource.Id);
                }
            }

            return _cuToGuidMapping;
        }

        public bool IsDrivingCu(string cuId)
        {
            var cuToSpec = GetCuToSpec(ResourceDomain.Gas);
            var cuToGuid = GetCuToGuidMapping();

            if (cuToGuid.TryGetValue(cuId, out Guid cuGuid))
            {
                if (cuToSpec.ContainsKey(cuGuid))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
