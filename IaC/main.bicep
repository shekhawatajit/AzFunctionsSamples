targetScope = 'subscription'
@description('Name of the resourceGroup to create.')
param resourceGroupName string
@description('The Azure region into which the resources should be deployed.')
param resourceGroupLocation string
@description('Prefix string to use with resources.')
param appNamePrefix string
@description('The type of environment. This must be nonprod or prod.')
@allowed([
  'nonprod'
  'prod'
])
param environment string
@description('Hub Site Url to add in Azure Function Configuration.')
param HubSite string
@description('Guid of Request SharePoint list.')
param RequestListId string

// resource group created in target subscription
resource resourceGroup 'Microsoft.Resources/resourceGroups@2021-04-01' = {
  name: resourceGroupName
  location: resourceGroupLocation
}


module FunctionResources 'FunctionResources.bicep' = {
  name: 'FunctionResources'
  scope: resourceGroup
  params: {
    location: resourceGroupLocation
    appNamePrefix: appNamePrefix
    environmentType: environment
    RequestListId: RequestListId
    HubSite: HubSite
  }
}
