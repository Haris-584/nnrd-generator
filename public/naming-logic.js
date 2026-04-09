(function(exports) {
  // Official Microsoft CAF abbreviations:
  // https://learn.microsoft.com/en-us/azure/cloud-adoption-framework/ready/azure-best-practices/resource-abbreviations
  const resourcePatterns = {
    // Compute & Web
    'Virtual Machine':            'vm',
    'VM Jump Server':             'vm',       // VM type — name tag differentiates (jumpbox)
    'VM Self-Hosted Agent':       'vm',       // VM type — name tag differentiates (agent)
    'VM Scale Set':               'vmss',
    'App Service':                'app',
    'Static Web App':             'stapp',
    'Function App':               'func',
    'Windows Virtual Desktop':    'vdws',
    // Containers
    'Azure Kubernetes Service':   'aks',
    'Azure Container Registry':   'cr',
    'Container Instances':        'ci',
    // Databases
    'Azure SQL Database':         'sqldb',
    'Azure SQL Managed Instance': 'sqlmi',
    'MySQL Database':             'mysql',
    'PostgreSQL Database':        'psql',
    'Cosmos DB':                  'cosmos',
    'Redis Cache':                'redis',
    // Storage
    'Storage Blob':               'st',
    // Networking
    'Virtual Network':            'vnet',
    'Subnet':                     'snet',
    'Network Security Group':     'nsg',
    'Public IP':                  'pip',
    'Network Interface':          'nic',
    'Application Gateway':        'agw',
    'Azure Front Door':           'afd',
    'Bastion':                    'bas',
    'ExpressRoute':               'erc',
    'VPN Gateway':                'vpng',
    'Private Endpoint':           'pep',
    'Route Table':                'rt',
    'Azure Firewall':             'afw',
    'Load Balancer':              'lbi',
    // AI & ML
    'AI Search':                  'srch',
    'OpenAI':                     'oai',
    'Cognitive Services':         'ais',
    'Machine Learning Workspace': 'mlw',
    'Bot Function':               'azurebot',
    // Data & Analytics
    'Data Factory':               'adf',
    'Databricks':                 'dbw',
    'Azure Synapse':              'synw',
    'Event Hubs':                 'evh',
    'Service Bus':                'sbns',
    'Event Grid':                 'evgd',
    'Stream Analytics':           'asa',
    // Management, Security & Monitoring
    'Key Vault':                  'kv',
    'Log Analytics Workspace':    'log',
    'Application Insights':       'appi',
    'Azure Monitor':              'dcr',
    'APIM':                       'apim',
  };

  const envMap = {
    'development': 'dev',
    'staging': 'stg',
    'production': 'prod',
    'hub': 'hub'
  };

  const regionMap = {
    'Qatar Central': 'qc',
    'Sweden Central': 'swc',
    'West Europe': 'weu',
    'UAE North': 'uaen'
  };

  exports.resourcePatterns = resourcePatterns;

  exports.generateResourceName = function(resourceType, resourceName, project, env, region, number) {
    const isDr = env.includes('(DR)');
    const baseEnv = env.replace(' (DR)', '');
    
    const prefix = resourcePatterns[resourceType] || resourceType.replace(/\s+/g, '-').toLowerCase();
    const cleanProj = project.replace(/\s+/g, '-').toLowerCase();
    
    let shortEnv = envMap[baseEnv.toLowerCase()] || baseEnv.substring(0, 3).toLowerCase();
    if (isDr) shortEnv += '-dr';
    
    const shortReg = regionMap[region] || region.substring(0, 2).toLowerCase();
    const nameTag = resourceName ? resourceName.replace(/\s+/g, '-').toLowerCase().trim() : '';

    // Only insert the nameTag segment if the user provided one
    let result = nameTag
      ? `${prefix}-${nameTag}-${cleanProj}-${shortEnv}-${shortReg}-${String(number).padStart(2, '0')}`
      : `${prefix}-${cleanProj}-${shortEnv}-${shortReg}-${String(number).padStart(2, '0')}`;

    // Fix for Azure Storage Accounts (no hyphens, max 24 chars)
    if (resourceType === 'Storage Blob') {
      result = result.replace(/-/g, '').substring(0, 24);
    }
    
    // Fix for Azure Key Vaults (max 24 chars)
    if (resourceType === 'Key Vault') {
      if (result.length > 24) {
        result = result.substring(0, 24);
        if (result.endsWith('-')) {
            result = result.substring(0, result.length - 1);
        }
      }
    }

    return result;
  };

  exports.generateResourceGroup = function(project, env, region) {
    const isDr = env.includes('(DR)');
    const baseEnv = env.replace(' (DR)', '');
    
    const cleanProj = project.replace(/\s+/g, '-').toLowerCase();
    let shortEnv = envMap[baseEnv.toLowerCase()] || baseEnv.substring(0, 3).toLowerCase();
    if (isDr) shortEnv += '-dr';
    
    const shortReg = regionMap[region] || region.substring(0, 2).toLowerCase();
    return `rg-${cleanProj}-${shortEnv}-${shortReg}-01`;
  };

  exports.generateNetworkName = function(type, purposeTag, project, env, region, number = '01') {
    const isDr = env.includes('(DR)');
    const baseEnv = env.replace(' (DR)', '');
    
    const cleanProj = project.replace(/\s+/g, '-').toLowerCase();
    let shortEnv = envMap[baseEnv.toLowerCase()] || baseEnv.substring(0, 3).toLowerCase();
    if (isDr) shortEnv += '-dr';
    
    const shortReg = regionMap[region] || (region ? region.substring(0, 2).toLowerCase() : 'reg');
    
    if (type === 'vnet') {
      return `vnet-${cleanProj}-${shortEnv}-${shortReg}-${number}`;
    } else {
      return `snet-${purposeTag}-${cleanProj}-${shortEnv}-${shortReg}-${number}`;
    }
  };

})(typeof exports === 'undefined' ? (this.NamingLogic = {}) : exports);
