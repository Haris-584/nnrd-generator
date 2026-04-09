const form = document.getElementById('nnrd-form');
const previewTable = document.getElementById('preview-tbody');
const envCheckboxes = document.querySelectorAll('input[name="environments"]');
const networkConfigsContainer = document.getElementById('network-configs');
const enableDrCheckbox = document.getElementById('enable-dr');
const drRegionSelect = document.getElementById('dr-region');

let selectedSpokeResources = [];
let selectedHubResources = [];
const networkConfigs = {};

// --- Integrated Searchable Select ---
function setupSearchableSelect(selectId) {
  const select = document.getElementById(selectId);
  if (!select) return;

  const container = document.createElement('div');
  container.className = 'combo-box-container';
  select.parentNode.insertBefore(container, select);
  
  const input = document.createElement('input');
  input.className = 'combo-box-input';
  input.placeholder = select.options[0].textContent;
  input.type = 'text';
  input.autocomplete = 'off';
  
  const dropdown = document.createElement('div');
  dropdown.className = 'combo-box-dropdown';
  
  container.appendChild(input);
  container.appendChild(dropdown);
  container.appendChild(select);
  select.style.display = 'none';

  // Extract data once
  const items = Array.from(select.querySelectorAll('optgroup, option')).map(node => {
     if (node.tagName === 'OPTGROUP') {
       return { 
         type: 'group', 
         label: node.label, 
         options: Array.from(node.querySelectorAll('option')).map(o => ({ text: o.textContent, value: o.value })) 
       };
     }
     if (node.parentNode.tagName !== 'OPTGROUP' && node.value) {
       return { type: 'option', text: node.textContent, value: node.value };
     }
     return null;
  }).filter(n => n);

  const renderDropdown = (term = '') => {
    dropdown.innerHTML = '';
    const filtered = items.map(item => {
      if (item.type === 'group') {
        const matching = item.options.filter(o => o.text.toLowerCase().includes(term.toLowerCase()));
        return matching.length > 0 ? { ...item, options: matching } : null;
      }
      return item.text.toLowerCase().includes(term.toLowerCase()) ? item : null;
    }).filter(n => n);

    if (filtered.length === 0) {
      dropdown.style.display = 'none';
      container.closest('.form-section')?.classList.remove('has-open-dropdown');
      return;
    }

    filtered.forEach(item => {
      if (item.type === 'group') {
        const header = document.createElement('div');
        header.className = 'combo-box-group';
        header.textContent = item.label;
        dropdown.appendChild(header);
        item.options.forEach(opt => createItem(opt));
      } else {
        createItem(item);
      }
    });
    
    dropdown.style.display = 'block';
    container.closest('.form-section')?.classList.add('has-open-dropdown');

    // Scroll selected item into view if it exists
    const selected = dropdown.querySelector('.combo-box-item.selected');
    if (selected) {
      selected.scrollIntoView({ block: 'nearest' });
    }
  };

  const createItem = (item) => {
    const div = document.createElement('div');
    div.className = 'combo-box-item';
    if (select.value === item.value) {
      div.classList.add('selected');
    }
    div.textContent = item.text;
    div.addEventListener('click', (e) => {
      e.stopPropagation();
      input.value = item.text;
      select.value = item.value;
      dropdown.style.display = 'none';
      container.closest('.form-section')?.classList.remove('has-open-dropdown');
      select.dispatchEvent(new Event('change'));
    });
    dropdown.appendChild(div);
  };

  // Focus and Click handlers
  input.addEventListener('focus', () => {
    input.select();
    renderDropdown(''); // Always show all on focus
  });

  input.addEventListener('click', (e) => {
    e.stopPropagation();
    if (dropdown.style.display !== 'block') {
      renderDropdown('');
    }
  });

  input.addEventListener('input', () => {
    renderDropdown(input.value);
  });

  // Global click to close
  document.addEventListener('mousedown', (e) => {
    if (!container.contains(e.target)) {
      dropdown.style.display = 'none';
      container.closest('.form-section')?.classList.remove('has-open-dropdown');
    }
  });

  // Reset function to be called after adding a resource
  select.resetCustom = () => {
    input.value = '';
    select.value = '';
    dropdown.style.display = 'none';
    container.closest('.form-section')?.classList.remove('has-open-dropdown');
  };
}

window.addEventListener('DOMContentLoaded', () => {
  setupSearchableSelect('spoke-resource-type-select');
  setupSearchableSelect('hub-resource-type-select');
});

// Handle environment selection changes
envCheckboxes.forEach((checkbox) => {
  checkbox.addEventListener('change', () => {
    updateNetworkConfigSections();
    updatePreviewTable();
  });
});

enableDrCheckbox.addEventListener('change', (e) => {
  drRegionSelect.disabled = !e.target.checked;
  if (e.target.checked) {
    const prRegion = document.querySelector('select[name="region"]').value;
    if (prRegion) drRegionSelect.value = prRegion;
  }
  updateNetworkConfigSections();
  updatePreviewTable();
});

drRegionSelect.addEventListener('change', () => {
  updateNetworkConfigSections();
  updatePreviewTable();
});

const envOrder = {
  'Hub': 1,
  'Development': 2,
  'Staging': 3,
  'Production': 4,
  'Hub (DR)': 5,
  'Production (DR)': 6
};

function getSelectedEnvironments() {
  const envs = [];
  const enableDr = enableDrCheckbox.checked;
  
  document.querySelectorAll('input[name="environments"]:checked').forEach((input) => {
    envs.push(input.value);
    if (enableDr && (input.value === 'Production' || input.value === 'Hub')) {
      envs.push(`${input.value} (DR)`);
    }
  });

  envs.sort((a, b) => (envOrder[a] || 99) - (envOrder[b] || 99));

  return envs;
}

function getRegionForEnv(env) {
  if (env.includes('(DR)')) {
    return drRegionSelect.value || 'region';
  }
  return document.querySelector('select[name="region"]').value || 'region';
}

function updateNetworkConfigSections() {
  const envs = getSelectedEnvironments();
  const project = document.querySelector('input[name="project"]').value || 'name';

  networkConfigsContainer.innerHTML = '';
  
  envs.forEach((env) => {
    if (!networkConfigs[env]) {
      networkConfigs[env] = {
        vnet_name: '',
        address_space: '',
        subnets: [],
      };
    }
    
    const region = getRegionForEnv(env);
    const defaultVnetName = NamingLogic.generateNetworkName('vnet', '', project, env, region);

    const section = document.createElement('div');
    section.className = 'network-env-section';
    section.innerHTML = `
      <h4>${env}</h4>
      <div class="form-grid">
        <div>
          <label>VNet Name (Blank for default):</label>
          <input type="text" class="env-vnet-name" data-env="${env}" placeholder="${defaultVnetName}" value="${networkConfigs[env].vnet_name}" />
        </div>
        <div>
          <label>Requested Size:</label>
          <input type="text" class="env-address-space" data-env="${env}" placeholder="e.g., /24" value="${networkConfigs[env].address_space}" />
        </div>
      </div>
      <div class="subnet-header">
        <h4>Subnets</h4>
        <button type="button" class="add-subnet-btn btn-secondary" data-env="${env}">+ Add Subnet</button>
      </div>
      <div class="subnet-list" data-env="${env}"></div>
    `;
    
    networkConfigsContainer.appendChild(section);
    
    section.querySelector('.env-vnet-name').addEventListener('change', (e) => {
      networkConfigs[env].vnet_name = e.target.value;
    });
    
    section.querySelector('.env-address-space').addEventListener('change', (e) => {
      networkConfigs[env].address_space = e.target.value;
    });
    
    section.querySelector('.add-subnet-btn').addEventListener('click', (e) => {
      e.preventDefault();
      createSubnetRow(env);
    });
    
    renderSubnets(env);
  });
}

function createSubnetRow(env, name = '', cidr = '', purpose = '') {
  const subnetList = networkConfigsContainer.querySelector(`.subnet-list[data-env="${env}"]`);
  const row = document.createElement('div');
  row.className = 'subnet-row';
  
  const nameInput = document.createElement('input');
  nameInput.className = 'subnet-name';
  nameInput.placeholder = 'Subnet name';
  nameInput.value = name;
  nameInput.required = true;

  const cidrInput = document.createElement('input');
  cidrInput.className = 'subnet-cidr';
  cidrInput.placeholder = 'Size (e.g., /27)';
  cidrInput.value = cidr;
  cidrInput.required = true;

  const purposeInput = document.createElement('input');
  purposeInput.className = 'subnet-purpose';
  purposeInput.placeholder = 'Purpose';
  purposeInput.value = purpose;

  const removeBtn = document.createElement('button');
  removeBtn.type = 'button';
  removeBtn.className = 'remove-row';
  removeBtn.textContent = 'Remove';

  row.appendChild(nameInput);
  row.appendChild(cidrInput);
  row.appendChild(purposeInput);
  row.appendChild(removeBtn);
  
  removeBtn.addEventListener('click', () => {
    row.remove();
    saveSubnets(env);
  });
  
  const inputsListener = () => saveSubnets(env);
  row.querySelectorAll('input').forEach((input) => input.addEventListener('change', inputsListener));
  
  subnetList.appendChild(row);
}

function saveSubnets(env) {
  const subnetList = networkConfigsContainer.querySelector(`.subnet-list[data-env="${env}"]`);
  networkConfigs[env].subnets = Array.from(subnetList.querySelectorAll('.subnet-row')).map((row) => ({
    name: row.querySelector('.subnet-name').value.trim(),
    cidr: row.querySelector('.subnet-cidr').value.trim(),
    purpose: row.querySelector('.subnet-purpose').value.trim(),
  }));
  updatePreviewTable();
}

function enforceRequiredSubnets(env) {
  const project = document.querySelector('input[name="project"]').value || 'name';
  const region = getRegionForEnv(env);

  const isHub = env.toLowerCase().includes('hub');
  const activeResources = isHub ? selectedHubResources : selectedSpokeResources;

  const requiresAppSvc = activeResources.some(r => r.type === 'App Service');
  const requiresFunc = activeResources.some(r => r.type === 'Function App');
  const requiresMySQL = activeResources.some(r => r.type === 'MySQL Database');
  const requiresAPIM = activeResources.some(r => r.type === 'APIM' || r.type === 'Azure API Management (APIM)');
  const requiresBastion = activeResources.some(r => r.type === 'Bastion' || r.type === 'Azure Bastion');
  const requiresAppGW = activeResources.some(r => r.type === 'Application Gateway');

  let subnets = networkConfigs[env].subnets || [];
  let changed = false;

  const intSize = env.toLowerCase().includes('development') ? '/29' : '/28';

  const ensureSubnet = (exactName, defaultSize, purposeTxt) => {
    if (!subnets.some(s => s.name === exactName || s.name.includes(exactName))) {
      createSubnetRow(env, exactName, defaultSize, purposeTxt);
      changed = true;
    }
  };

  const ensureStdSubnet = (purposeTag, defaultSize, purposeTxt) => {
    const snetName = NamingLogic.generateNetworkName('snet', purposeTag, project, env, region);
    if (!subnets.some(s => s.name === snetName || s.name.includes(`snet-${purposeTag}`))) {
      createSubnetRow(env, snetName, defaultSize, purposeTxt);
      changed = true;
    }
  };

  // Base
  ensureStdSubnet('pep', '/27', 'Private Endpoints');

  // Integrations
  if (requiresAppSvc) ensureStdSubnet('app', intSize, 'App Service Integration');
  if (requiresFunc) ensureStdSubnet('func', intSize, 'Function App Integration');
  if (requiresMySQL) ensureStdSubnet('mysql', intSize, 'MySQL Subnet');
  if (requiresAPIM) ensureStdSubnet('apim', intSize, 'APIM Subnet');
  if (requiresAppGW) ensureStdSubnet('agw', '/24', 'Application Gateway Subnet');

  // Bastion is unique, must follow Microsoft exact name without prefix 'AzureBastionSubnet'
  if (requiresBastion) ensureSubnet('AzureBastionSubnet', '/26', 'Azure Bastion Service');

  if (changed) {
    saveSubnets(env);
  }
}

function renderSubnets(env) {
  const subnetList = networkConfigsContainer.querySelector(`.subnet-list[data-env="${env}"]`);
  subnetList.innerHTML = '';
  
  const subnets = networkConfigs[env].subnets || [];
  subnets.forEach((subnet) => {
    createSubnetRow(env, subnet.name, subnet.cidr, subnet.purpose);
  });
  
  enforceRequiredSubnets(env);
}

// Spoke Handlers
const addSpokeResourceBtn = document.getElementById('add-spoke-resource-btn');
const spokeResourceTypeSelect = document.getElementById('spoke-resource-type-select');
const spokeResourceNameInput = document.getElementById('spoke-resource-name-input');
const spokeResourceQtyInput = document.getElementById('spoke-resource-qty-input');
const selectedSpokeResourcesList = document.getElementById('selected-spoke-resources');

addSpokeResourceBtn.addEventListener('click', () => {
  const type = spokeResourceTypeSelect.value.trim();
  const name = spokeResourceNameInput.value.trim();
  const qty = parseInt(spokeResourceQtyInput.value) || 1;
  
  if (!type) {
    alert('Please select a generic Spoke resource type');
    return;
  }
  
  selectedSpokeResources.push({ type, name, qty });
  if (spokeResourceTypeSelect.resetCustom) spokeResourceTypeSelect.resetCustom();
  else spokeResourceTypeSelect.value = '';
  spokeResourceNameInput.value = '';
  spokeResourceQtyInput.value = '1';
  renderSelectedResources();
});

spokeResourceNameInput.addEventListener('keypress', (e) => {
  if (e.key === 'Enter') addSpokeResourceBtn.click();
});

// Hub Handlers
const addHubResourceBtn = document.getElementById('add-hub-resource-btn');
const hubResourceTypeSelect = document.getElementById('hub-resource-type-select');
const hubResourceNameInput = document.getElementById('hub-resource-name-input');
const hubResourceQtyInput = document.getElementById('hub-resource-qty-input');
const selectedHubResourcesList = document.getElementById('selected-hub-resources');

addHubResourceBtn.addEventListener('click', () => {
  const type = hubResourceTypeSelect.value.trim();
  const name = hubResourceNameInput.value.trim();
  const qty = parseInt(hubResourceQtyInput.value) || 1;
  
  if (!type) {
    alert('Please select a generic Hub resource type');
    return;
  }
  
  selectedHubResources.push({ type, name, qty });
  if (hubResourceTypeSelect.resetCustom) hubResourceTypeSelect.resetCustom();
  else hubResourceTypeSelect.value = '';
  hubResourceNameInput.value = '';
  hubResourceQtyInput.value = '1';
  renderSelectedResources();
});

hubResourceNameInput.addEventListener('keypress', (e) => {
  if (e.key === 'Enter') addHubResourceBtn.click();
});

function renderSelectedResources() {
  selectedSpokeResourcesList.innerHTML = '';
  selectedSpokeResources.forEach((res, idx) => {
    const tag = document.createElement('div');
    tag.className = 'resource-tag';
    tag.innerHTML = `
      <span>${res.type}${res.name ? ` (${res.name})` : ''} x${res.qty}</span>
      <button type="button" class="remove-tag" data-idx="${idx}">×</button>
    `;
    tag.querySelector('.remove-tag').addEventListener('click', () => {
      selectedSpokeResources.splice(idx, 1);
      renderSelectedResources();
    });
    selectedSpokeResourcesList.appendChild(tag);
  });

  selectedHubResourcesList.innerHTML = '';
  selectedHubResources.forEach((res, idx) => {
    const tag = document.createElement('div');
    tag.className = 'resource-tag';
    tag.innerHTML = `
      <span>${res.type}${res.name ? ` (${res.name})` : ''} x${res.qty}</span>
      <button type="button" class="remove-tag" data-idx="${idx}">×</button>
    `;
    tag.querySelector('.remove-tag').addEventListener('click', () => {
      selectedHubResources.splice(idx, 1);
      renderSelectedResources();
    });
    selectedHubResourcesList.appendChild(tag);
  });

  updatePreviewTable();
  const envs = getSelectedEnvironments();
  envs.forEach(env => enforceRequiredSubnets(env));
}

// Global state for manual overrides
let manualOverrides = {};

function updatePreviewTable() {
  const project = document.querySelector('input[name="project"]').value || 'proj';
  const envs = getSelectedEnvironments();
  
  previewTable.innerHTML = '';
  
  if (selectedSpokeResources.length === 0 && selectedHubResources.length === 0) {
    previewTable.innerHTML = '<tr><td colspan="4" class="empty">Add resources to preview naming</td></tr>';
    return;
  }
  
  const prEnvs = envs.filter(e => !e.includes('(DR)'));
  const drEnvs = envs.filter(e => e.includes('(DR)'));

  const renderEnvGroupToPreview = (envList, groupName) => {
    if (envList.length === 0) return;

    let hasAnyResources = false;
    envList.forEach(env => {
      const isHub = env.toLowerCase().includes('hub');
      if ((isHub ? selectedHubResources : selectedSpokeResources).length > 0) hasAnyResources = true;
    });

    if (!hasAnyResources) return;

    const groupRow = document.createElement('tr');
    groupRow.innerHTML = `<td colspan="4" style="background: var(--primary); color: #fff; font-weight: 600; text-align: center; font-size: 15px;">${groupName}</td>`;
    previewTable.appendChild(groupRow);

    envList.forEach((env) => {
      const isHub = env.toLowerCase().includes('hub');
      const activeResources = isHub ? selectedHubResources : selectedSpokeResources;
      const region = getRegionForEnv(env);
      
      if (activeResources.length > 0) {
        const envRow = document.createElement('tr');
        envRow.innerHTML = `<td colspan="4" style="background: #e0e7ff; color: var(--primary); font-weight: 600; text-align: center;">${env}</td>`;
        previewTable.appendChild(envRow);

        activeResources.forEach((res, resIdx) => {
          for (let i = 1; i <= res.qty; i++) {
            const row = document.createElement('tr');
            const defaultRG = NamingLogic.generateResourceGroup(project, env, region);
            const defaultName = NamingLogic.generateResourceName(res.type, res.name, project, env, region, i);
            
            // Unified override key logic
            const overrideKey = `${env.trim()}_${res.type.trim()}_${resIdx}_${i}`;
            const currentRG = manualOverrides[overrideKey]?.rg || defaultRG;
            const currentName = manualOverrides[overrideKey]?.name || defaultName;

            row.classList.add('naming-row');
            row.dataset.env = env;
            row.dataset.resType = res.type;
            row.dataset.instance = i;

            row.innerHTML = `
              <td>${res.type}</td>
              <td class="editable-rg" contenteditable="true" style="font-family: monospace; color: var(--text-muted);">${currentRG}</td>
              <td style="font-weight: 500;">${env}</td>
              <td class="editable-name" contenteditable="true" style="font-family: monospace; font-weight: 600; color: var(--primary-hover);">${currentName}</td>
            `;

            // Keep manualOverrides updated for persistence during re-renders
            const rgCell = row.querySelector('.editable-rg');
            const nameCell = row.querySelector('.editable-name');

            const saveEdit = () => {
              manualOverrides[overrideKey] = {
                rg: (rgCell.innerText || rgCell.textContent || '').trim(),
                name: (nameCell.innerText || nameCell.textContent || '').trim()
              };
            };

            rgCell.addEventListener('input', saveEdit);
            nameCell.addEventListener('input', saveEdit);
            
            previewTable.appendChild(row);
          }
        });
      }
    });
  };

  renderEnvGroupToPreview(prEnvs, 'Primary Region');
  renderEnvGroupToPreview(drEnvs, 'Disaster Recovery Region');
}

// Update preview when region inputs change
document.querySelector('input[name="project"]').addEventListener('input', () => {
    updatePreviewTable();
    updateNetworkConfigSections();
});
document.querySelector('select[name="region"]').addEventListener('change', (e) => {
    if (enableDrCheckbox.checked) {
      drRegionSelect.value = e.target.value;
    }
    updatePreviewTable();
    updateNetworkConfigSections();
});

// Initialize
updateNetworkConfigSections();

form.addEventListener('submit', async (event) => {
  event.preventDefault();
  
  const formData = new FormData(form);
  const envs = getSelectedEnvironments();
  
  // Build resources using the EXACT SAME looping logic as updatePreviewTable
  const envResources = {};
  envs.forEach((env) => {
    envResources[env] = [];
    const isHub = env.toLowerCase().includes('hub');
    const activeResources = isHub ? selectedHubResources : selectedSpokeResources;

    activeResources.forEach((res, resIdx) => {
      for (let i = 1; i <= res.qty; i++) {
        const overrideKey = `${env.trim()}_${res.type.trim()}_${resIdx}_${i}`;
        const override = manualOverrides[overrideKey];

        envResources[env].push({
          type: res.type,
          name: res.name,
          instance: i,
          manualName: override?.name || null,
          manualRG: override?.rg || null
        });
      }
    });
  });
  
  const projectValue = formData.get('project').trim();
  const regionValue = formData.get('region').trim();
  const drRegionValue = drRegionSelect.value.trim();

  const networksToSend = JSON.parse(JSON.stringify(networkConfigs));
  envs.forEach((env) => {
    const envRegion = getRegionForEnv(env);
    if (!networksToSend[env].vnet_name) {
      networksToSend[env].vnet_name = NamingLogic.generateNetworkName('vnet', '', projectValue, env, envRegion);
    }
  });

  const payload = {
    project: projectValue,
    region: regionValue,
    dr_region: drRegionValue,
    environments: envs,
    resources: envResources,
    networks: networksToSend,
  };
  
  try {
    const response = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(errorText || 'Failed to generate file');
    }
    
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${payload.project.toUpperCase()}-NNRD.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  } catch (error) {
    alert('Error: ' + error.message);
  }
});
