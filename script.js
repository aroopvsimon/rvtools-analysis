const fileInput = document.getElementById('fileInput');
const stats = document.getElementById('stats');
const sheetStatus = document.getElementById('sheetStatus');

const requiredViews = ['vInfo', 'vHost', 'vCluster', 'vDatastore'];
let latestClusterRows = [];
let latestHostRows = [];
let latestSummary = '';
let latestObservations = [];

fileInput.addEventListener('change', handleFile);
document.getElementById('copySummaryBtn').addEventListener('click', () => copyText(latestSummary));
document.getElementById('copyObsBtn').addEventListener('click', () => copyText(latestObservations.map(o => `${o.severity}: ${o.text}`).join('\n')));
document.getElementById('downloadClusterCsvBtn').addEventListener('click', () => downloadCsv('cluster-summary.csv', latestClusterRows));
document.getElementById('downloadHostCsvBtn').addEventListener('click', () => downloadCsv('host-summary.csv', latestHostRows));

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result);
    const wb = XLSX.read(data, { type: 'array' });
    const sheets = wb.SheetNames;
    sheetStatus.textContent = `Detected sheets: ${sheets.join(', ')}`;

    const vInfo = sheetJson(wb, 'vInfo');
    const vHost = sheetJson(wb, 'vHost');
    const vCluster = sheetJson(wb, 'vCluster');
    const vDatastore = sheetJson(wb, 'vDatastore');

    const model = analyze(vInfo, vHost, vCluster, vDatastore, sheets);
    render(model);
  };
  reader.readAsArrayBuffer(file);
}

function sheetJson(wb, name) {
  const ws = wb.Sheets[name];
  return ws ? XLSX.utils.sheet_to_json(ws, { defval: null }) : [];
}

function analyze(vInfo, vHost, vCluster, vDatastore, sheets) {
  const totalVMs = vInfo.length;
  const poweredOnVMs = vInfo.filter(r => String(r['Powerstate'] || '').toLowerCase() === 'poweredon').length;
  const poweredOffVMs = vInfo.filter(r => String(r['Powerstate'] || '').toLowerCase() === 'poweredoff').length;
  const totalvCPU = sum(vInfo, ['CPUs', 'NumCPU']);
  const totalMemGB = sum(vInfo, ['Memory', 'Memory size MB']) / 1024;
  const hostCount = vHost.length;
  const clusterCount = unique(vHost.map(r => pick(r, ['Cluster', 'Cluster Name']))).filter(Boolean).length || vCluster.length;
  const datastoreCount = vDatastore.length;

  const clusterMap = {};
  vInfo.forEach(vm => {
    const cluster = pick(vm, ['Cluster', 'Cluster Name']) || 'Unassigned';
    clusterMap[cluster] ||= { cluster, vmCount: 0, configuredvCPU: 0, configuredMemoryGB: 0 };
    clusterMap[cluster].vmCount += 1;
    clusterMap[cluster].configuredvCPU += num(pick(vm, ['CPUs', 'NumCPU']));
    clusterMap[cluster].configuredMemoryGB += num(pick(vm, ['Memory', 'Memory size MB'])) / 1024;
  });

  vHost.forEach(host => {
    const cluster = pick(host, ['Cluster', 'Cluster Name']) || 'Unassigned';
    clusterMap[cluster] ||= { cluster, vmCount: 0, configuredvCPU: 0, configuredMemoryGB: 0 };
    clusterMap[cluster].hosts = (clusterMap[cluster].hosts || 0) + 1;
  });

  vDatastore.forEach(ds => {
    const cluster = pick(ds, ['Cluster', 'Cluster Name']) || 'Unassigned';
    clusterMap[cluster] ||= { cluster, vmCount: 0, configuredvCPU: 0, configuredMemoryGB: 0 };
    clusterMap[cluster].storageCapacityGB = (clusterMap[cluster].storageCapacityGB || 0) + num(pick(ds, ['Capacity MB', 'Capacity MiB'])) / 1024;
    clusterMap[cluster].storageFreeGB = (clusterMap[cluster].storageFreeGB || 0) + num(pick(ds, ['Free MB', 'Free MiB'])) / 1024;
  });

  const clusterRows = Object.values(clusterMap).map(r => {
    const cap = round(r.storageCapacityGB || 0);
    const free = round(r.storageFreeGB || 0);
    const used = round(Math.max(cap - free, 0));
    const usedPct = cap ? round((used / cap) * 100) : 0;
    return {
      Cluster: r.cluster,
      'VM Count': r.vmCount || 0,
      Hosts: r.hosts || 0,
      'Configured vCPU': r.configuredvCPU || 0,
      'Configured Memory (GB)': round(r.configuredMemoryGB || 0),
      'Storage Capacity (GB)': cap,
      'Storage Used (GB)': used,
      'Free Capacity (GB)': free,
      'Storage Used %': usedPct
    };
  }).sort((a, b) => b['Configured Memory (GB)'] - a['Configured Memory (GB)']);

  const hostRows = vHost.map(h => ({
    Host: pick(h, ['Host', 'Name']) || '',
    Cluster: pick(h, ['Cluster', 'Cluster Name']) || '',
    'CPU Sockets': num(pick(h, ['CPUs', 'Packages'])),
    Cores: num(pick(h, ['Cores', 'NumCpuCores'])),
    'Memory (GB)': round(num(pick(h, ['Memory', 'Memory size MB'])) / 1024),
    'ESXi Version': pick(h, ['Version', 'ESX Version']) || ''
  })).sort((a, b) => String(a.Cluster).localeCompare(String(b.Cluster)) || String(a.Host).localeCompare(String(b.Host)));

  const observations = [];
  clusterRows.forEach(r => {
    if (r['Storage Used %'] >= 80) {
      observations.push({ severity: 'High', text: `${r.Cluster} storage utilization is at ${r['Storage Used %']}%, which is above the normal comfort threshold and should be reviewed for capacity expansion or cleanup.` });
    } else if (r['Storage Used %'] >= 70) {
      observations.push({ severity: 'Medium', text: `${r.Cluster} storage utilization is at ${r['Storage Used %']}%, indicating moderate capacity pressure that should be tracked.` });
    }
    if (r.Hosts > 0 && r['VM Count'] / r.Hosts > 60) {
      observations.push({ severity: 'Medium', text: `${r.Cluster} has approximately ${round(r['VM Count'] / r.Hosts, 1)} VMs per host, which suggests comparatively dense host consolidation.` });
    }
  });

  const oldHosts = hostRows.filter(h => {
    const version = String(h['ESXi Version']);
    const major = parseInt(version.split('.')[0], 10);
    return Number.isFinite(major) && major < 8;
  });
  if (oldHosts.length) {
    observations.push({ severity: 'High', text: `${oldHosts.length} host(s) appear to be on ESXi versions below 8.x, which may warrant lifecycle and supportability review.` });
  }

  if (!observations.length) {
    observations.push({ severity: 'Low', text: 'No major risk flags were detected from the limited sheets parsed. A deeper review can still be done for snapshots, tools status, hardware versions, and configuration best practices.' });
  }

  const execSummary = [
    `The RVTools dataset shows ${totalVMs} virtual machines across ${clusterCount} cluster(s), ${hostCount} host(s), and ${datastoreCount} datastore(s).`,
    `${poweredOnVMs} VMs are powered on and ${poweredOffVMs} are powered off.`,
    `The environment has ${Math.round(totalvCPU).toLocaleString()} configured vCPU and ${round(totalMemGB).toLocaleString()} GB of configured VM memory based on the available sheets.`,
    observations[0] ? `Primary observation: ${observations[0].text}` : ''
  ].filter(Boolean).join('\n\n');

  return {
    detected: requiredViews.map(name => `${name}: ${sheets.includes(name) ? 'yes' : 'no'}`),
    stats: [
      ['Total VMs', totalVMs],
      ['Powered On', poweredOnVMs],
      ['Powered Off', poweredOffVMs],
      ['Configured vCPU', Math.round(totalvCPU)],
      ['Configured Memory (GB)', round(totalMemGB)],
      ['Hosts', hostCount],
      ['Clusters', clusterCount],
      ['Datastores', datastoreCount]
    ],
    summary: execSummary,
    observations,
    clusterRows,
    hostRows
  };
}

function render(model) {
  latestClusterRows = model.clusterRows;
  latestHostRows = model.hostRows;
  latestSummary = model.summary;
  latestObservations = model.observations;

  stats.innerHTML = model.stats.map(([label, value]) => `
    <div class="stat"><div class="label">${label}</div><div class="value">${Number(value).toLocaleString?.() || value}</div></div>
  `).join('');

  document.getElementById('execSummaryCard').classList.remove('hidden');
  document.getElementById('execSummary').textContent = model.summary;

  document.getElementById('obsCard').classList.remove('hidden');
  document.getElementById('observations').innerHTML = model.observations.map(o => `
    <li><span class="badge ${o.severity.toLowerCase()}">${o.severity}</span>${escapeHtml(o.text)}</li>
  `).join('');

  document.getElementById('clusterCard').classList.toggle('hidden', !model.clusterRows.length);
  buildTable('clusterTable', model.clusterRows);

  document.getElementById('hostCard').classList.toggle('hidden', !model.hostRows.length);
  buildTable('hostTable', model.hostRows);
}

function buildTable(id, rows) {
  const table = document.getElementById(id);
  const thead = table.querySelector('thead');
  const tbody = table.querySelector('tbody');
  if (!rows.length) {
    thead.innerHTML = '';
    tbody.innerHTML = '';
    return;
  }
  const headers = Object.keys(rows[0]);
  thead.innerHTML = `<tr>${headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}</tr>`;
  tbody.innerHTML = rows.map(r => `<tr>${headers.map(h => `<td>${escapeHtml(String(r[h] ?? ''))}</td>`).join('')}</tr>`).join('');
}

function downloadCsv(filename, rows) {
  if (!rows.length) return;
  const headers = Object.keys(rows[0]);
  const csv = [headers.join(',')]
    .concat(rows.map(r => headers.map(h => csvCell(r[h])).join(',')))
    .join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  link.click();
}

function copyText(text) {
  navigator.clipboard.writeText(text).then(() => alert('Copied'));
}

function sum(rows, keys) { return rows.reduce((acc, row) => acc + num(pick(row, keys)), 0); }
function pick(row, keys) { return keys.map(k => row?.[k]).find(v => v !== undefined && v !== null && v !== ''); }
function num(v) { return Number(String(v ?? '').replace(/,/g, '')) || 0; }
function unique(arr) { return [...new Set(arr)]; }
function round(v, d = 0) { const p = 10 ** d; return Math.round((v + Number.EPSILON) * p) / p; }
function csvCell(v) { const s = String(v ?? ''); return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; }
function escapeHtml(s) { return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;'); }
