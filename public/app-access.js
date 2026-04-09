let members = [];

// --- Helper: format today's date as M/D/YYYY ---
function getTodayDate() {
  const d = new Date();
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}

// Set today's date on page load into the date field
window.addEventListener('DOMContentLoaded', () => {
  const dateField = document.getElementById('m-date');
  if (dateField) dateField.value = getTodayDate();
  renderMembers();
});

// --- Form submit: add a member ---
document.getElementById('add-member-form').addEventListener('submit', function(e) {
  e.preventDefault();

  const team = document.getElementById('m-team').value;
  const computedDevOpsRole = (team === 'Cloud and DevOps Team')
    ? 'Basic & Project adminstrator'
    : 'Basic & Project Contributor';

  // Read project-level fields (outside the per-member form)
  const devopsProj = document.getElementById('project-name').value || 'Project';

  const member = {
    id: Date.now(),
    name: document.getElementById('m-name').value.trim(),
    email: document.getElementById('m-email').value.trim(),
    org: document.getElementById('m-org').value.trim(),
    team: team,
    status: document.getElementById('m-status').value,
    date: document.getElementById('m-date').value || getTodayDate(),
    cloudRole: document.getElementById('m-cloud-role').value,
    devopsProj: devopsProj,
    devopsRole: computedDevOpsRole
  };

  members.push(member);
  renderMembers();

  // Reset only per-member fields, keep org & date sticky
  const org = document.getElementById('m-org').value;
  const date = document.getElementById('m-date').value;
  this.reset();
  document.getElementById('m-org').value = org;
  document.getElementById('m-date').value = date;
  document.getElementById('m-cloud-role').value = 'Contributor';
});

// --- Render the member table ---
function renderMembers() {
  const tbody = document.getElementById('members-list');
  tbody.innerHTML = '';

  if (members.length === 0) {
    tbody.innerHTML = '<tr><td colspan="6" style="text-align: center; color: #94a3b8; padding: 3rem;">No members yet. Fill in the form on the left to get started.</td></tr>';
    return;
  }

  members.forEach(m => {
    const statusClass = m.status === 'Granted' ? 'badge-granted' : 'badge-pending';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>
        <div style="font-weight: 600; color: #0f172a">${m.name}</div>
        <div style="font-size: 0.8rem; color: #64748b">${m.email}</div>
      </td>
      <td>
        <div style="font-weight: 500">${m.team}</div>
        <div style="font-size: 0.8rem; color: #64748b">${m.org}</div>
      </td>
      <td><span class="badge ${statusClass}">${m.status}</span></td>
      <td><span style="font-size: 0.85rem">${m.cloudRole}</span></td>
      <td><div style="font-size: 0.85rem">${m.devopsRole}</div></td>
      <td><button class="del-btn" onclick="deleteMember(${m.id})">Remove</button></td>
    `;
    tbody.appendChild(tr);
  });
}

window.deleteMember = function(id) {
  members = members.filter(m => m.id !== id);
  renderMembers();
};

// --- Download Excel ---
document.getElementById('generate-excel-btn').addEventListener('click', async () => {
  if (members.length === 0) {
    alert('Please add at least one member before generating the document.');
    return;
  }

  const projectName = document.getElementById('project-name').value.trim() || 'Project';
  const devopsProj  = projectName;

  const btn = document.getElementById('generate-excel-btn');
  const originalText = btn.innerText;
  btn.innerText = 'Compiling...';
  btn.disabled = true;

  try {
    const payload = { members, projectName, devopsProj };

    const response = await fetch('/api/generate-access', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    if (!response.ok) throw new Error('Generation failed');

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${projectName.replace(/\s+/g, '_')}_Access_Matrix.xlsx`;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    a.remove();
  } catch (err) {
    console.error(err);
    alert('Error generating Excel file. Please check the server console.');
  } finally {
    btn.innerText = originalText;
    btn.disabled = false;
  }
});
