@description('The Azure region into which the resources should be deployed.')
param location string = resourceGroup().location

@description('Guid of Request SharePoint list.')
param RequestListId string

@description('Prefix string to use with resources.')
param appNamePrefix string

@description('Hub Site Url to add in Azure Function Configuration.')
param HubSite string

@description('The type of environment. This must be nonprod or prod.')
@allowed([
  'nonprod'
  'prod'
])
param environmentType string
// Define the SKUs for each component based on the environment type.
var environmentConfigurationMap = {
  nonprod: {
    appServicePlan: {
      sku: {
        name: 'Y1'
        capacity: 1
      }
    }
    contentGovernanceStorageAccount: {
      sku: {
        name: 'Standard_LRS'
      }
    }
  }
  prod: {
    appServicePlan: {
      sku: {
        name: 'S1'
        capacity: 2
      }
    }
    contentGovernanceStorageAccount: {
      sku: {
        name: 'Standard_ZRS'
      }
    }
  }
}

var functionAppName = '${appNamePrefix}-functionapp'
var appServiceName = '${appNamePrefix}-appservice'
var appInsightsName = '${appNamePrefix}-appinsights'
var workspaceName = '${appNamePrefix}-workspace'
var storageAccountName = format('{0}sta', replace(appNamePrefix, '-', ''))
var queueName = '${appNamePrefix}-queue'
resource storageaccount 'Microsoft.Storage/storageAccounts@2021-02-01' = {
  name: storageAccountName
  location: location
  kind: 'StorageV2'
  sku: environmentConfigurationMap[environmentType].contentGovernanceStorageAccount.sku
}
resource queueservice 'Microsoft.Storage/storageAccounts/queueServices@2021-02-01' = {
  name: 'default'
  parent: storageaccount
}
resource queue 'Microsoft.Storage/storageAccounts/queueServices/queues@2021-02-01' = {
  name: queueName
  parent: queueservice
}
resource appServicePlan 'Microsoft.Web/serverfarms@2020-12-01' = {
  name: appServiceName
  location: location
  kind: 'linux'
  sku: environmentConfigurationMap[environmentType].appServicePlan.sku
  properties: {
    reserved: true
  }
}

resource logAnalyticsWorkspace 'Microsoft.OperationalInsights/workspaces@2020-10-01' = {
  name: workspaceName
  location: location
  properties: {
    sku: {
      name: 'PerGB2018'
    }
  }
}
resource appInsightsComponents 'Microsoft.Insights/components@2020-02-02-preview' = {
  name: appInsightsName
  location: location
  kind: 'web'
  properties: {
    Application_Type: 'web'
    WorkspaceResourceId: logAnalyticsWorkspace.id
  }
}

resource azureFunction 'Microsoft.Web/sites@2021-03-01' = {
  name: functionAppName
  location: location
  kind: 'functionapp,linux'
  properties: {
    serverFarmId: appServicePlan.id
    siteConfig: {
      linuxFxVersion: 'DOTNET|6.0'
    }
  }
  identity: {
    type: 'SystemAssigned'
  }
}

resource functionalAppSettings 'Microsoft.Web/sites/config@2021-03-01' = {
  name: '${functionAppName}/appsettings'
  properties: {
    APPINSIGHTS_INSTRUMENTATIONKEY: appInsightsComponents.properties.InstrumentationKey
    APPLICATIONINSIGHTS_CONNECTION_STRING: appInsightsComponents.properties.ConnectionString
    AzureWebJobsStorage: 'DefaultEndpointsProtocol=https;AccountName=${storageaccount.name};AccountKey=${storageaccount.listKeys().keys[0].value};EndpointSuffix=core.windows.net'
    FUNCTIONS_EXTENSION_VERSION: '~4'
    FUNCTIONS_WORKER_RUNTIME: 'dotnet'
    linuxFxVersion: 'DOTNET|6.0'
    WEBSITE_LOAD_USER_PROFILE: 1
    
    CertificateName: 'oipdevelopment'
    ClientId: '3764eb06-7e9d-408b-83e3-2d1982ac5707'
    KeyVaultName: 'oipkv'
    TenantId: '51575b39-28de-4120-94c6-af4c743f70f1'
    QueueName: queue.name
    RequestListId: RequestListId
    HubSite: HubSite
  }
}
