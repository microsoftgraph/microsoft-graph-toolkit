param resourceBaseName string
param frontendHosting_storageName string = 'frontendstg${uniqueString(resourceBaseName)}'
@secure()
param m365ClientId string
@secure()
param m365ClientSecret string
param m365TenantId string
param m365OauthAuthorityHost string
param simpleAuth_sku string = 'F1'
param simpleAuth_serverFarmsName string = '${resourceBaseName}-simpleAuth-serverfarms'
param simpleAuth_webAppName string = '${resourceBaseName}-simpleAuth-webapp'
param simpleAuth_packageUri string = 'https://github.com/OfficeDev/TeamsFx/releases/download/simpleauth@0.1.0/Microsoft.TeamsFx.SimpleAuth_0.1.0.zip'

var m365ApplicationIdUri = 'api://${frontendHostingProvision.outputs.domain}/${m365ClientId}'

module frontendHostingProvision './modules/frontendHostingProvision.bicep' = {
  name: 'frontendHostingProvision'
  params: {
    frontendHostingStorageName: frontendHosting_storageName
  }
}
module simpleAuthProvision './modules/simpleAuthProvision.bicep' = {
  name: 'simpleAuthProvision'
  params: {
    simpleAuthServerFarmsName: simpleAuth_serverFarmsName
    simpleAuthWebAppName: simpleAuth_webAppName
    sku: simpleAuth_sku
  }
}
module simpleAuthConfiguration './modules/simpleAuthConfiguration.bicep' = {
  name: 'simpleAuthConfiguration'
  dependsOn: [
    simpleAuthProvision
  ]
  params: {
    simpleAuthWebAppName: simpleAuth_webAppName
    m365ClientId: m365ClientId
    m365ClientSecret: m365ClientSecret
    m365ApplicationIdUri: m365ApplicationIdUri
    frontendHostingStorageEndpoint: frontendHostingProvision.outputs.endpoint
    m365TenantId: m365TenantId
    oauthAuthorityHost: m365OauthAuthorityHost
    simpelAuthPackageUri: simpleAuth_packageUri
  }
}

output frontendHosting_storageResourceId string = frontendHostingProvision.outputs.resourceId
output frontendHosting_endpoint string = frontendHostingProvision.outputs.endpoint
output frontendHosting_domain string = frontendHostingProvision.outputs.domain
output simpleAuth_skuName string = simpleAuthProvision.outputs.skuName
output simpleAuth_endpoint string = simpleAuthProvision.outputs.endpoint
output simpleAuth_webAppName string = simpleAuthProvision.outputs.webAppName
output simpleAuth_appServicePlanName string = simpleAuthProvision.outputs.appServicePlanName