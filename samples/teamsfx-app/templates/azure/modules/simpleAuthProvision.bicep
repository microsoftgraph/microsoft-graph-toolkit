param sku string
param simpleAuthServerFarmsName string
param simpleAuthWebAppName string

resource simpleAuthServerFarms 'Microsoft.Web/serverfarms@2020-06-01' = {
  name: simpleAuthServerFarmsName
  location: resourceGroup().location
  sku: {
    name: sku
  }
  kind: 'app'
  properties: {
    reserved: false
  }
}

resource simpleAuthWebApp 'Microsoft.Web/sites@2020-06-01' = {
  kind: 'app'
  name: simpleAuthWebAppName
  location: resourceGroup().location
  properties: {
    reserved: false
    serverFarmId: simpleAuthServerFarms.id
    siteConfig: {
      alwaysOn: false
      http20Enabled: false
      numberOfWorkers: 1
    }
  }
}

output webAppName string = simpleAuthWebAppName
output skuName string = sku
output endpoint string = 'https://${simpleAuthWebApp.properties.hostNames[0]}'
output appServicePlanName string = simpleAuthServerFarmsName
