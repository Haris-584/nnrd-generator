const path = require('path');
const express = require('express');
const ExcelJS = require('exceljs');
const cors = require('cors');
const NamingLogic = require('./public/naming-logic.js');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.get('/nnrd', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'nnrd.html'));
});

app.get('/accesssheet', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'accesssheet.html'));
});
function applyHeaderStyle(cell) {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0078D4' } }; // Azure Blue
  cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: 'Calibri' };
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  cell.border = {
    left: { style: 'thin', color: { argb: 'FF005A9E' } },
    right: { style: 'thin', color: { argb: 'FF005A9E' } },
    top: { style: 'thin', color: { argb: 'FF005A9E' } },
    bottom: { style: 'thin', color: { argb: 'FF005A9E' } },
  };
}

function applyDataStyle(cell, isEven) {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
  cell.font = { size: 11, color: { argb: 'FF000000' }, name: 'Calibri' };
  cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  cell.border = {
    left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
    right: { style: 'thin', color: { argb: 'FFCCCCCC' } },
    top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
    bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
  };
}

app.post('/api/generate', async (req, res) => {
  try {
    const {
      project,
      region,
      dr_region: drRegion = '',
      environments = [],
      resources = {},
      networks = {},
    } = req.body;

    if (!project || !region || environments.length === 0) {
      return res.status(400).send('Project, region, and at least one environment are required.');
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'NNRD Generator';
    workbook.title = `${project} - Network and Naming Documentation`;

    const deepBlueBg = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF002060' } };
    const brightGreenBg = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B050' } };
    const greyBg = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF808080' } };

    function applyCorporateHeader(cell) {
      cell.fill = deepBlueBg;
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: 'Calibri' };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = { left: { style: 'thin' }, right: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } };
    }

    const namingSheet = workbook.addWorksheet('Naming Reference');
    namingSheet.columns = [
      { key: 'type', width: 35 },
      { key: 'rg', width: 40 },
      { key: 'envCol', width: 20 },
      { key: 'name', width: 50 },
      { key: 'regionCol', width: 20 },
      { key: 'tags', width: 45 },
      { key: 'remarks', width: 30 },
      { key: 'scope', width: 10 },
    ];
    
    // Row 1: Naming Reference Master Header
    const namingTitle = namingSheet.addRow(['Naming Reference']);
    namingTitle.getCell(1).fill = deepBlueBg;
    namingTitle.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 14, name: 'Calibri' };
    namingTitle.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    namingSheet.mergeCells('A1:H1');
    namingTitle.height = 30;

    // Row 2: Standard Columns
    const namingHeaders = namingSheet.addRow(['Resource Type', 'Resource group', 'Environment', 'Resource Name', 'Region', 'Tags', 'Remarks', 'Scope']);
    namingHeaders.eachCell(applyCorporateHeader);
    namingHeaders.height = 25;

    const networkSheet = workbook.addWorksheet('Network Requirements');
    networkSheet.columns = [
      { key: 'env', width: 20 },
      { key: 'vnet', width: 35 },
      { key: 'vnet_size', width: 25 },
      { key: 'snet', width: 35 },
      { key: 'snet_size', width: 20 },
      { key: 'purpose', width: 45 },
      { key: 'remarks', width: 30 },
      { key: 'scope', width: 10 },
    ];
    
    const networkTitle = networkSheet.addRow(['Network Requirements']);
    networkTitle.getCell(1).fill = deepBlueBg;
    networkTitle.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 14, name: 'Calibri' };
    networkTitle.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    networkSheet.mergeCells('A1:H1');
    networkTitle.height = 30;

    const networkHeaders = networkSheet.addRow(['Environment', 'VNet Name', 'VNet Size', 'Subnet Name', 'Requested Size', 'Purpose', 'Remarks', 'Scope']);
    networkHeaders.eachCell(applyCorporateHeader);
    networkHeaders.height = 25;

    const projectSheet = workbook.addWorksheet('Project Info');
    projectSheet.columns = [
      { header: 'Property', key: 'property', width: 30 },
      { header: 'Value', key: 'value', width: 50 },
    ];
    projectSheet.getRow(1).eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0078D4' } };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: 'Calibri' };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    
    const infoRows = [
      ['Project Name', project],
      ['Environments', environments.join(', ')],
      ['Region', region],
      ['Generated On', new Date().toISOString()],
    ];
    
    infoRows.forEach((row, idx) => {
      projectSheet.addRow({ property: row[0], value: row[1] });
      projectSheet.getRow(idx + 2).eachCell((cell) => applyDataStyle(cell, idx % 2 === 0));
    });

    let namingRowIdx = 3;
    let networkRowIdx = 3;

    const prEnvs = environments.filter(e => !e.includes('(DR)'));
    const drEnvs = environments.filter(e => e.includes('(DR)'));

    const renderEnvsGroup = (envsGroup, groupTitle, scopeLabel, groupBgColor) => {
      if (envsGroup.length === 0) return;

      const startScopeNamingRow = namingRowIdx;
      const startScopeNetworkRow = networkRowIdx;

      let namingRowsGeneratedInGroup = 0;
      let networkRowsGeneratedInGroup = 0;

      envsGroup.forEach((env) => {
        const envResources = resources[env] || [];
        const envNetwork = networks[env] || {};
        const vnetName = envNetwork.vnet_name || '';

        const envRegion = env.includes('(DR)') ? drRegion : region;
        const envBase = env.replace(' (DR)', '');
        const mappedEnvText = envBase.toUpperCase() === 'HUB' ? 'HUB' : envBase;
        
        const tagsText = `Project: ${project.toUpperCase()}\nEnvironment: ${envBase}\nScope: ${groupTitle}\nCreated By: (Applab)`;
        const remarksText = `Tags and Naming of Resources can be changed based on Client IT requirements.`;

        // === Naming Reference ===
        if (envResources.length > 0) {
          const envNamingSubHeader = namingSheet.addRow([`${envBase} - ${groupTitle}`, '', '', '', '', '', '', '']);
          envNamingSubHeader.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: 'Calibri' };
          envNamingSubHeader.getCell(1).fill = groupBgColor;
          envNamingSubHeader.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
          namingSheet.mergeCells(`A${namingRowIdx}:G${namingRowIdx}`); // Merge A to G, leaving H (Scope) free
          namingRowIdx++;
          namingRowsGeneratedInGroup++;
        }

        const startNamingRow = namingRowIdx;
        let addedNamingRows = 0;

        envResources.forEach((res) => {
          const rg = NamingLogic.generateResourceGroup(project, env, envRegion);
          const resourceName = NamingLogic.generateResourceName(res.type, res.name, project, env, envRegion, res.instance);
          
          const row = namingSheet.addRow({
            type: res.type,
            rg: rg,
            envCol: mappedEnvText,
            name: resourceName,
            regionCol: envRegion,
            tags: tagsText,
            remarks: remarksText,
            scope: '' // Leave empty for vertical merger later
          });
          
          row.eachCell((cell) => applyDataStyle(cell));
          row.getCell(3).font = { bold: true, size: 11, color: { argb: 'FF000000' }, name: 'Calibri' };
          row.height = 45; // Taller row to comfortably fit multi-line tags
          namingRowIdx++;
          addedNamingRows++;
          namingRowsGeneratedInGroup++;
        });
        
        if (addedNamingRows > 1) {
          namingSheet.mergeCells(`B${startNamingRow}:B${namingRowIdx - 1}`);
          namingSheet.mergeCells(`C${startNamingRow}:C${namingRowIdx - 1}`);
          namingSheet.mergeCells(`E${startNamingRow}:E${namingRowIdx - 1}`);
          namingSheet.mergeCells(`F${startNamingRow}:F${namingRowIdx - 1}`);
          namingSheet.mergeCells(`G${startNamingRow}:G${namingRowIdx - 1}`);
        }

        // === Network Requirements ===
        const subnets = Array.isArray(envNetwork.subnets) ? envNetwork.subnets : [];
        if (subnets.length > 0 || vnetName) {
          const envNetworkSubHeader = networkSheet.addRow([`${envBase} - ${groupTitle}`, '', '', '', '', '', '', '']);
          envNetworkSubHeader.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: 'Calibri' };
          envNetworkSubHeader.getCell(1).fill = groupBgColor;
          envNetworkSubHeader.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
          networkSheet.mergeCells(`A${networkRowIdx}:G${networkRowIdx}`);
          networkRowIdx++;
          networkRowsGeneratedInGroup++;
        }

        const startNetworkRow = networkRowIdx;
        let addedNetworkRows = 0;
        
        if (subnets.length === 0) {
          const row = networkSheet.addRow({
            env: mappedEnvText,
            vnet: vnetName,
            vnet_size: envNetwork.address_space || '',
            snet: '',
            snet_size: '',
            purpose: '',
            remarks: remarksText,
            scope: ''
          });
          row.eachCell((cell) => applyDataStyle(cell));
          row.getCell(1).font = { bold: true, size: 11, color: { argb: 'FF000000' }, name: 'Calibri' };
          row.height = 25;
          networkRowIdx++;
          addedNetworkRows++;
          networkRowsGeneratedInGroup++;
        } else {
          subnets.forEach((subnet) => {
            const row = networkSheet.addRow({
              env: mappedEnvText,
              vnet: vnetName,
              vnet_size: envNetwork.address_space || '',
              snet: subnet.name || '',
              snet_size: subnet.cidr || '',
              purpose: subnet.purpose || '',
              remarks: remarksText,
              scope: ''
            });
            row.eachCell((cell) => applyDataStyle(cell));
            row.getCell(1).font = { bold: true, size: 11, color: { argb: 'FF000000' }, name: 'Calibri' };
            row.height = 25;
            networkRowIdx++;
            addedNetworkRows++;
            networkRowsGeneratedInGroup++;
          });
        }

        if (addedNetworkRows > 1) {
          networkSheet.mergeCells(`A${startNetworkRow}:A${networkRowIdx - 1}`);
          networkSheet.mergeCells(`B${startNetworkRow}:B${networkRowIdx - 1}`);
          networkSheet.mergeCells(`C${startNetworkRow}:C${networkRowIdx - 1}`);
          networkSheet.mergeCells(`G${startNetworkRow}:G${networkRowIdx - 1}`);
        }
      });

      // Apply the massive right-hand vertical SCOPE merge across the entire group block
      if (namingRowsGeneratedInGroup > 0) {
        namingSheet.mergeCells(`H${startScopeNamingRow}:H${namingRowIdx - 1}`);
        const nameScopeCell = namingSheet.getCell(`H${startScopeNamingRow}`);
        nameScopeCell.value = scopeLabel;
        nameScopeCell.fill = groupBgColor;
        nameScopeCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 14, name: 'Calibri' };
        nameScopeCell.alignment = { horizontal: 'center', vertical: 'middle', textRotation: 90 };
        nameScopeCell.border = { left: { style: 'thin' }, right: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } };
      }

      if (networkRowsGeneratedInGroup > 0) {
        networkSheet.mergeCells(`H${startScopeNetworkRow}:H${networkRowIdx - 1}`);
        const netScopeCell = networkSheet.getCell(`H${startScopeNetworkRow}`);
        netScopeCell.value = scopeLabel;
        netScopeCell.fill = groupBgColor;
        netScopeCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 14, name: 'Calibri' };
        netScopeCell.alignment = { horizontal: 'center', vertical: 'middle', textRotation: 90 };
        netScopeCell.border = { left: { style: 'thin' }, right: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } };
      }
    };

    renderEnvsGroup(prEnvs, 'Primary Region', 'PR Setup', brightGreenBg);
    renderEnvsGroup(drEnvs, 'Disaster Recovery Region', 'DR Setup', greyBg);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${project}-NNRD.xlsx"`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).send('Error generating document: ' + err.message);
  }
});

app.post('/api/generate-access', async (req, res) => {
  try {
    const { members = [] } = req.body;
    if (!members || members.length === 0) {
      return res.status(400).send('No members provided.');
    }

    // Sort heavily to group merges smoothly
    members.sort((a, b) => {
      if (a.org !== b.org) return a.org.localeCompare(b.org);
      if (a.team !== b.team) return a.team.localeCompare(b.team);
      if (a.status !== b.status) return b.status.localeCompare(a.status); // Granted > Pending
      return 0;
    });

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Access Generator';
    
    // Shared Styling
    const deepBlueBg = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF002060' } };
    const lightBlueHeaderBg = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } }; // Light blue per screenshot
    const yellowBg = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };

    function applySubHeader(cell) {
      cell.fill = lightBlueHeaderBg;
      cell.font = { bold: true, color: { argb: 'FF000000' }, size: 11, name: 'Calibri' };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = { left:{style:'thin'}, right:{style:'thin'}, top:{style:'thin'}, bottom:{style:'thin'} };
    }
    
    function applyDataCell(cell, overrideAlign) {
      cell.font = { size: 10, color: { argb: 'FF000000' }, name: 'Calibri' };
      cell.alignment = overrideAlign || { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = { left:{style:'thin'}, right:{style:'thin'}, top:{style:'thin'}, bottom:{style:'thin'} };
    }

    // Helper for aggressive column vertical merges based on same-values
    function mergeByValue(sheet, colLetter, rowsStart, rowsEnd) {
      if (rowsEnd <= rowsStart) return;
      let startMergeRow = rowsStart;
      let currentValue = sheet.getCell(colLetter + startMergeRow).value;
      
      for (let r = rowsStart + 1; r <= rowsEnd; r++) {
        let cell = sheet.getCell(colLetter + r);
        let cellValue = cell.value;
        if (cellValue !== currentValue) {
          if (r - 1 > startMergeRow && currentValue) {
            sheet.mergeCells(`${colLetter}${startMergeRow}:${colLetter}${r - 1}`);
            sheet.getCell(`${colLetter}${startMergeRow}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          }
          startMergeRow = r;
          currentValue = cellValue;
        }
      }
      if (rowsEnd > startMergeRow && currentValue) {
        sheet.mergeCells(`${colLetter}${startMergeRow}:${colLetter}${rowsEnd}`);
        sheet.getCell(`${colLetter}${startMergeRow}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      }
    }

    // -------------------------------------------------------------
    // SHEET 1: PROJECT MEMBERS
    // -------------------------------------------------------------
    const projSheet = workbook.addWorksheet('Project Members');
    projSheet.columns = [
      { key: 'name', width: 25 },
      { key: 'org', width: 20 },
      { key: 'team', width: 30 },
      { key: 'email1', width: 35 },
      { key: 'email2', width: 35 },
      { key: 'status', width: 25 },
      { key: 'date', width: 20 },
      { key: 'remarks', width: 20 },
    ];
    
    const projTitle = projSheet.addRow(['Azure Cloud Project Members and their Entra ID Accounts']);
    projTitle.getCell(1).fill = deepBlueBg;
    projTitle.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: 'Calibri' };
    projTitle.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    projSheet.mergeCells('A1:H1');
    projTitle.height = 25;

    const projHeaders = projSheet.addRow(['Name', 'Organization', 'Designation', 'E-mail Address', 'User Account (Email ID)', 'Status', 'Date of Request', 'Remarks']);
    projHeaders.eachCell(applySubHeader);
    
    let currentRow = 3;
    members.forEach(m => {
      const row = projSheet.addRow({
        name: m.name, org: m.org, team: m.team, 
        email1: m.email, email2: m.email, 
        status: m.status === 'Granted' ? 'Granted' : 'to be Granted',
        date: m.date, remarks: ''
      });
      row.eachCell(cell => applyDataCell(cell));
      row.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
      row.getCell(4).alignment = { horizontal: 'left', vertical: 'middle' };
      row.getCell(5).alignment = { horizontal: 'left', vertical: 'middle' };
      
      // Status formatting explicitly matched
      if (m.status !== 'Granted') {
        row.getCell(6).fill = yellowBg;
        row.getCell(6).font = { bold: true, size: 10, name: 'Calibri' };
      }
      currentRow++;
    });

    const endProjRow = currentRow - 1;
    mergeByValue(projSheet, 'B', 3, endProjRow);
    mergeByValue(projSheet, 'C', 3, endProjRow);
    mergeByValue(projSheet, 'F', 3, endProjRow);
    mergeByValue(projSheet, 'G', 3, endProjRow);
    mergeByValue(projSheet, 'H', 3, endProjRow);


    // -------------------------------------------------------------
    // SHEET 2: CLOUD ACCESS
    // -------------------------------------------------------------
    const cloudSheet = workbook.addWorksheet('Cloud Access');
    cloudSheet.columns = [
      { key: 'name', width: 25 },
      { key: 'email', width: 35 },
      { key: 'status', width: 15 },
      { key: 'date', width: 20 },
      { key: 'dev', width: 30 },
      { key: 'stg', width: 30 },
      { key: 'prod', width: 30 },
      { key: 'remarks', width: 45 },
    ];

    const cloudTitle = cloudSheet.addRow(['Azure Cloud Access']);
    cloudTitle.getCell(1).fill = deepBlueBg;
    cloudTitle.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: 'Calibri' };
    cloudTitle.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    cloudSheet.mergeCells('A1:H1');
    cloudTitle.height = 25;

    const cloudHeaders = cloudSheet.addRow(['Name', 'User Account (Email ID)', 'Status', 'Date of Request', 'Dev Role', 'Staging Role', 'Production Role', 'Remarks']);
    cloudHeaders.eachCell(applySubHeader);

    const cloudRemarksText = "Contributor role on resource group level will grant permission to configure related dependencies and manage the workloads according to project requirements.";
    
    let cRow = 3;
    members.forEach(m => {
      const row = cloudSheet.addRow({
        name: m.name, email: m.email, status: m.status, date: m.date,
        dev: m.cloudRole, stg: m.cloudRole, prod: m.cloudRole, remarks: cloudRemarksText
      });
      row.eachCell(cell => applyDataCell(cell));
      row.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
      row.getCell(2).alignment = { horizontal: 'left', vertical: 'middle' };
      
      if (m.status !== 'Granted') {
        row.getCell(1).fill = yellowBg;
        row.getCell(2).fill = yellowBg;
        row.getCell(3).fill = yellowBg;
        row.getCell(4).fill = yellowBg;
        row.getCell(3).font = { bold: true };
      }
      cRow++;
    });

    const endCloudRow = cRow - 1;
    mergeByValue(cloudSheet, 'C', 3, endCloudRow);
    mergeByValue(cloudSheet, 'D', 3, endCloudRow);
    mergeByValue(cloudSheet, 'E', 3, endCloudRow);
    mergeByValue(cloudSheet, 'F', 3, endCloudRow);
    mergeByValue(cloudSheet, 'G', 3, endCloudRow);
    mergeByValue(cloudSheet, 'H', 3, endCloudRow);

    // -------------------------------------------------------------
    // SHEET 3: DEVOPS ACCESS
    // -------------------------------------------------------------
    const devopsSheet = workbook.addWorksheet('DevOps Access');
    devopsSheet.columns = [
      { key: 'name', width: 25 },
      { key: 'email', width: 35 },
      { key: 'status', width: 15 },
      { key: 'proj', width: 25 },
      { key: 'date', width: 20 },
      { key: 'role', width: 35 },
      { key: 'remarks', width: 45 },
    ];

    const devopsTitle = devopsSheet.addRow(['Azure DevOps Access']);
    devopsTitle.getCell(1).fill = deepBlueBg;
    devopsTitle.getCell(1).font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: 'Calibri' };
    devopsTitle.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    devopsSheet.mergeCells('A1:G1');
    devopsTitle.height = 25;

    const devopsHeaders = devopsSheet.addRow(['Name', 'User Account (Email ID)', 'Status', 'Projects Level Access', 'Date of Request', 'Role', 'Remarks']);
    devopsHeaders.eachCell(applySubHeader);

    const devopsRemarksText = "Basic role is required to allow the engineers to work on with pipelines and related configurations without any hurdles.";

    let dRow = 3;
    members.forEach(m => {
      const row = devopsSheet.addRow({
        name: m.name, email: m.email, status: m.status, 
        proj: m.devopsProj, date: m.date, role: m.devopsRole, remarks: devopsRemarksText
      });
      row.eachCell(cell => applyDataCell(cell));
      row.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
      row.getCell(2).alignment = { horizontal: 'left', vertical: 'middle' };
      
      if (m.status !== 'Granted') {
        const rowCells = ['A','B'];
        rowCells.forEach(l => devopsSheet.getCell(`${l}${dRow}`).fill = yellowBg);
        row.getCell(3).font = { bold: true };
      }
      dRow++;
    });

    const endDevopsRow = dRow - 1;
    mergeByValue(devopsSheet, 'C', 3, endDevopsRow);
    mergeByValue(devopsSheet, 'D', 3, endDevopsRow);
    mergeByValue(devopsSheet, 'E', 3, endDevopsRow);
    mergeByValue(devopsSheet, 'F', 3, endDevopsRow);
    mergeByValue(devopsSheet, 'G', 3, endDevopsRow);


    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="Azure_Access_Matrix.xlsx"`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).send('Error generating access matrix: ' + err.message);
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});