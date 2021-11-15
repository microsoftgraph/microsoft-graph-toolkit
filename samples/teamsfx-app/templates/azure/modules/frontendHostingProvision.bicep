@minLength(3)
@maxLength(24)
@description('Name of Storage Accounts for frontend hosting.')
param frontendHostingStorageName string

var siteDomain = replace(replace(frontendHostingStorage.properties.primaryEndpoints.web, 'https://', ''), '/', '')

resource frontendHostingStorage 'Microsoft.Storage/storageAccounts@2021-04-01' = {
  kind: 'StorageV2'
  location: resourceGroup().location
  name: frontendHostingStorageName
  properties: {
    accessTier: 'Hot'
    supportsHttpsTrafficOnly: true
    isHnsEnabled: false
  }
  sku: {
    name: 'Standard_LRS'
  }
}

output resourceId string = frontendHostingStorage.id
output endpoint string = 'https://${siteDomain}'
output domain string = siteDomain
