const uploadInput = document.getElementById("upload");
const statusEl = document.getElementById("status");
const summaryStatsEl = document.getElementById("summaryStats");
const detectedSheetsEl = document.getElementById("detectedSheets");
const executiveSummaryEl = document.getElementById("executiveSummary");
const observationsEl = document.getElementById("observations");
const clusterSummaryEl = document.getElementById("clusterSummary");
const hostSummaryEl = document.getElementById("hostSummary");
const esxiVersionEl = document.getElementById("esxiVersion");
const portgroupTableEl = document.getElementById("portgroupTable");
const hardwareAgeEl = document.getElementById("hardwareAge");

let hardwareLaunchYears = {};
let cachedExports = {
  clusters: [],
  hosts: [],
  portgroups: [],
  versions: [],
  hardware: []
};
let cachedObservationText = "";

fetch("hardware-launch-years.json")
  .then((res) => res.json())
  .then((data) => {
    hardwareLaunchYears = data;
  })
  .catch(() => {
    hardwareLaunchYears = {};
  });

uploadInput.addEventListener("change", handleFile);
document.getElementById("copyObservationsBtn").addEventListener("click", copyObservations);
document.getElementById("downloadClusterCsvBtn").addEventListener("click", () => downloadCsv("cluster-summary.csv", cachedExports.clusters));
document.getElementById("downloadHostCsvBtn").addEventListener("click", () => downloadCsv("host-summary.csv", cachedExports.hosts));
document.getElementById("downloadPortgroupCsvBtn").addEventListener("click", () => downloadCsv("portgroup-summary.csv", cachedExports.portgroups));
document.getElementById("downloadVersionCsvBtn").addEventListener("click", () => downloadCsv("esxi-versions.csv", cachedExports.versions));
document.getElementById("downloadHardwareCsvBtn").addEventListener("click", () => downloadCsv("hardware-age.csv", cachedExports.hardware));

function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  statusEl.textContent = `Reading ${file.name}...`;

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheets = workbook.SheetNames || [];
      const sheetMap = Object.fromEntries(sheets.map((name) => [name.toLowerCase(), name]));

      const vinfo = getSheetRows(workbook, sheetMap, "vinfo");
      const vhost = getSheetRows(workbook, sheetMap, "vhost");
      const vcluster = getSheetRows(workbook, sheetMap, "vcluster");
      const vdatastore = getSheetRows(workbook, sheetMap, "vdatastore");
      const vnic = getSheetRows(workbook, sheetMap, "vnic");

      renderDetectedSheets(sheets);
      renderAll(vinfo, vhost, vcluster, vdatastore, vnic);
      statusEl.textContent = `Loaded ${file.name} successfully.`;
    } catch (err) {
      console.error(err);
      statusEl.textContent = "Unable to parse the file. Please upload a valid RVTools Excel export.";
    }
  };
  reader.readAsArrayBuffer(file);
}

function getSheetRows(workbook, sheetMap, key) {
  const actualName = sheetMap[key];
  if (!actualName || !workbook.Sheets[actualName]) return [];
  return XLSX.utils.sheet_to_json(workbook.Sheets[actualName], { defval: "" });
}

function renderDetectedSheets(sheets) {
  if (!sheets.length) {
    detectedSheetsEl.innerHTML = '<div class="empty-state">No sheets detected.</div>';
    return;
  }
  detectedSheetsEl.innerHTML = sheets.map((sheet) => `<span class="badge badge-low" style="margin:4px 6px 0 0;">${escapeHtml(sheet)}</span>`).join("");
}

function renderAll(vinfo, vhost, vcluster, vdatastore, vnic) {
  const stats = buildTopStats(vinfo, vhost, vcluster, vdatastore);
  renderStats(stats);

  const clusterRows = buildClusterSummary(vinfo, vhost, vcluster, vdatastore);
  const hostRows = buildHostSummary(vhost);
  const versionRows = buildEsxiVersionSummary(vhost);
  const portgroupRows = buildPortgroupSummary(vnic);
  const hardwareRows = buildHardwareAgeSummary(vhost);
  const observations = buildObservations(stats, versionRows, portgroupRows, hardwareRows);

  cachedExports = {
    clusters: clusterRows,
    hosts: hostRows,
    portgroups: portgroupRows,
    versions: versionRows,
    hardware: hardwareRows
  };

  renderTable(clusterSummaryEl, clusterRows, [
    ["Cluster", "cluster"],
    ["VM Count", "vmCount"],
    ["Hosts", "hostCount"],
    ["Configured vCPU", "configuredCpu"],
    ["Configured Memory (GB)", "configuredMemoryGb"],
    ["Storage Capacity (TB)", "storageCapacityTb"],
    ["Storage Used (TB)", "storageUsedTb"],
    ["Free Capacity (TB)", "freeCapacityTb"]
  ]);

  renderTable(hostSummaryEl, hostRows, [
    ["Host", "host"],
    ["Cluster", "cluster"],
    ["Model", "model"],
    ["CPU Sockets", "cpuSockets"],
    ["Cores", "cores"],
    ["Memory (GB)", "memoryGb"],
    ["ESXi Version", "version"]
  ]);

  renderTable(esxiVersionEl, versionRows, [
    ["ESXi Version", "version"],
    ["Host Count", "hostCount"]
  ]);

  renderTable(portgroupTableEl, portgroupRows, [
    ["Port Group", "portgroup"],
    ["VM Count", "vmCount"]
  ]);

  renderTable(hardwareAgeEl, hardwareRows, [
    ["Host", "host"],
    ["Model", "model"],
    ["Launch Year", "launchYear"],
    ["Hardware Age", "hardwareAge"],
    ["Lifecycle Note", "note"]
  ]);

  renderExecutiveSummary(stats, versionRows, portgroupRows, hardwareRows, observations);
  renderObservations(observations);
}

function buildTopStats(vinfo, vhost, vcluster, vdatastore) {
  const totalVms = vinfo.length;
  const poweredOnVms = vinfo.filter((r) => toText(getField(r, ["Powerstate", "Power State"])) === "poweredon").length;
  const configuredVcpu = sumBy(vinfo, ["NumCPU", "vCPUs", "CPUs"]);
  const configuredMemoryMb = sumBy(vinfo, ["MemoryMB", "Memory", "Configured Memory", "vMemory"]);
  return {
    totalVms,
    poweredOnVms,
    poweredOffVms: Math.max(totalVms - poweredOnVms, 0),
    totalHosts: vhost.length,
    totalClusters: uniqueCount(vcluster.map((r) => getField(r, ["Cluster", "Name"]))),
    totalDatastores: uniqueCount(vdatastore.map((r) => getField(r, ["Datastore", "Name"]))),
    configuredVcpu,
    configuredMemoryGb: round(configuredMemoryMb / 1024)
  };
}

function renderStats(stats) {
  const cards = [
    ["Total VMs", stats.totalVms],
    ["Powered On", stats.poweredOnVms],
    ["Powered Off", stats.poweredOffVms],
    ["Hosts", stats.totalHosts],
    ["Clusters", stats.totalClusters],
    ["Datastores", stats.totalDatastores],
    ["Configured vCPU", stats.configuredVcpu],
    ["Configured Memory (GB)", stats.configuredMemoryGb]
  ];

  summaryStatsEl.innerHTML = cards.map(([label, value]) => `
    <div class="stat-card">
      <div class="stat-label">${escapeHtml(String(label))}</div>
      <div class="stat-value">${escapeHtml(String(value))}</div>
    </div>
  `).join("");
}

function buildClusterSummary(vinfo, vhost, vcluster, vdatastore) {
  const clusterMap = {};

  vinfo.forEach((row) => {
    const cluster = getField(row, ["Cluster", "Cluster Name"]) || "Unassigned";
    ensureCluster(clusterMap, cluster);
    clusterMap[cluster].vmCount += 1;
    clusterMap[cluster].configuredCpu += toNumber(getField(row, ["NumCPU", "vCPUs", "CPUs"]));
    clusterMap[cluster].configuredMemoryGb += toNumber(getField(row, ["MemoryMB", "Memory", "Configured Memory", "vMemory"])) / 1024;
  });

  vhost.forEach((row) => {
    const cluster = getField(row, ["Cluster", "Cluster Name"]) || "Unassigned";
    ensureCluster(clusterMap, cluster);
    clusterMap[cluster].hostCount += 1;
  });

  vcluster.forEach((row) => {
    const cluster = getField(row, ["Cluster", "Name"]) || "Unassigned";
    ensureCluster(clusterMap, cluster);
  });

  vdatastore.forEach((row) => {
    const cluster = getField(row, ["Cluster", "Cluster Name"]) || "Unassigned";
    ensureCluster(clusterMap, cluster);
    const capacityGb = toNumber(getField(row, ["CapacityGB", "Capacity", "Capacity GB"]));
    const freeGb = toNumber(getField(row, ["FreeGB", "Free", "Free GB"]));
    const usedGb = capacityGb && freeGb >= 0 ? capacityGb - freeGb : 0;
    clusterMap[cluster].storageCapacityTb += capacityGb / 1024;
    clusterMap[cluster].storageUsedTb += usedGb / 1024;
    clusterMap[cluster].freeCapacityTb += freeGb / 1024;
  });

  return Object.values(clusterMap)
    .map((row) => ({
      ...row,
      configuredMemoryGb: round(row.configuredMemoryGb),
      storageCapacityTb: round(row.storageCapacityTb),
      storageUsedTb: round(row.storageUsedTb),
      freeCapacityTb: round(row.freeCapacityTb)
    }))
    .sort((a, b) => b.configuredMemoryGb - a.configuredMemoryGb);
}

function ensureCluster(clusterMap, cluster) {
  if (!clusterMap[cluster]) {
    clusterMap[cluster] = {
      cluster,
      vmCount: 0,
      hostCount: 0,
      configuredCpu: 0,
      configuredMemoryGb: 0,
      storageCapacityTb: 0,
      storageUsedTb: 0,
      freeCapacityTb: 0
    };
  }
}

function buildHostSummary(vhost) {
  return vhost.map((row) => ({
    host: getField(row, ["Host", "Name"]) || "Unknown",
    cluster: getField(row, ["Cluster", "Cluster Name"]) || "Unassigned",
    model: getField(row, ["Model", "Hardware Model"]) || "Unknown",
    cpuSockets: toNumber(getField(row, ["CPUs", "CpuSockets", "CPU Sockets"])),
    cores: toNumber(getField(row, ["Cores", "CpuCores", "CPU Cores"])),
    memoryGb: round(toNumber(getField(row, ["MemoryGB", "Memory", "Memory GB"]))),
    version: getField(row, ["Version", "ESX Version", "ESXi Version"]) || "Unknown"
  })).sort((a, b) => a.cluster.localeCompare(b.cluster) || a.host.localeCompare(b.host));
}

function buildEsxiVersionSummary(vhost) {
  const counts = {};
  vhost.forEach((row) => {
    const version = getField(row, ["Version", "ESX Version", "ESXi Version"]) || "Unknown";
    counts[version] = (counts[version] || 0) + 1;
  });
  return Object.entries(counts)
    .map(([version, hostCount]) => ({ version, hostCount }))
    .sort((a, b) => b.hostCount - a.hostCount || b.version.localeCompare(a.version));
}

function buildPortgroupSummary(vnic) {
  const map = {};
  vnic.forEach((row) => {
    const portgroup = getField(row, ["Portgroup", "Network", "Port Group"]) || "Unknown";
    const vm = getField(row, ["VM", "VM Name", "VMName", "Name"]) || "Unknown";
    if (!map[portgroup]) map[portgroup] = new Set();
    map[portgroup].add(vm);
  });
  return Object.entries(map)
    .map(([portgroup, vmSet]) => ({ portgroup, vmCount: vmSet.size }))
    .sort((a, b) => b.vmCount - a.vmCount || a.portgroup.localeCompare(b.portgroup));
}

function buildHardwareAgeSummary(vhost) {
  const currentYear = new Date().getFullYear();
  return vhost.map((row) => {
    const host = getField(row, ["Host", "Name"]) || "Unknown";
    const model = getField(row, ["Model", "Hardware Model"]) || "Unknown";
    const launchYear = hardwareLaunchYears[model] || "Unknown";
    const hardwareAge = typeof launchYear === "number" ? currentYear - launchYear : "Unknown";
    let note = "Model launch year not found in lookup list";
    if (typeof hardwareAge === "number") {
      note = hardwareAge >= 7 ? "Approaching or within refresh window" : hardwareAge >= 5 ? "Mid-life hardware" : "Relatively current hardware";
    }
    return { host, model, launchYear, hardwareAge, note };
  }).sort((a, b) => {
    const ageA = typeof a.hardwareAge === "number" ? a.hardwareAge : -1;
    const ageB = typeof b.hardwareAge === "number" ? b.hardwareAge : -1;
    return ageB - ageA || a.host.localeCompare(b.host);
  });
}

function buildObservations(stats, versionRows, portgroupRows, hardwareRows) {
  const observations = [];

  const oldHardware = hardwareRows.filter((r) => typeof r.hardwareAge === "number" && r.hardwareAge >= 7);
  const unknownHardware = hardwareRows.filter((r) => r.launchYear === "Unknown");
  const topPortgroup = portgroupRows[0];
  const legacyVersions = versionRows.filter((r) => /^6\./.test(String(r.version)) || /^5\./.test(String(r.version)));

  if (oldHardware.length) {
    observations.push({ severity: "High", text: `${oldHardware.length} host(s) appear to be based on hardware models launched 7 or more years ago. These systems may be nearing refresh consideration for performance, supportability, and lifecycle alignment.` });
  }

  if (legacyVersions.length) {
    const totalLegacyHosts = legacyVersions.reduce((sum, row) => sum + row.hostCount, 0);
    observations.push({ severity: "Critical", text: `${totalLegacyHosts} host(s) are running legacy ESXi versions in the 6.x or earlier family. These should be reviewed for supportability, security posture, and upgrade planning.` });
  }

  if (topPortgroup && topPortgroup.vmCount >= 50) {
    observations.push({ severity: "Medium", text: `The busiest port group is ${topPortgroup.portgroup} with ${topPortgroup.vmCount} VM(s). Review network segmentation, broadcast domain sizing, and dependency concentration on this network.` });
  }

  if (stats.poweredOffVms > 0) {
    observations.push({ severity: "Low", text: `${stats.poweredOffVms} VM(s) are currently powered off. These can be reviewed for reclamation, archival, or cleanup opportunities if no longer required.` });
  }

  if (unknownHardware.length) {
    observations.push({ severity: "Medium", text: `${unknownHardware.length} host model(s) could not be matched to the local hardware launch-year lookup list. Add these models into hardware-launch-years.json to improve lifecycle reporting accuracy.` });
  }

  if (!observations.length) {
    observations.push({ severity: "Low", text: "No major red flags were automatically detected from the currently parsed sheets. A deeper review of HA, DRS, snapshots, tools status, and datastore conditions can be added next." });
  }

  cachedObservationText = observations.map((o, i) => `${i + 1}. [${o.severity}] ${o.text}`).join("\n\n");
  return observations;
}

function renderExecutiveSummary(stats, versionRows, portgroupRows, hardwareRows, observations) {
  const dominantVersion = versionRows[0] ? `${versionRows[0].version} (${versionRows[0].hostCount} host(s))` : "Not available";
  const busiestPortgroup = portgroupRows[0] ? `${portgroupRows[0].portgroup} (${portgroupRows[0].vmCount} VM(s))` : "Not available";
  const agedHosts = hardwareRows.filter((r) => typeof r.hardwareAge === "number" && r.hardwareAge >= 7).length;

  executiveSummaryEl.textContent =
    `This environment contains ${stats.totalVms} VM(s) across ${stats.totalHosts} host(s) and ${stats.totalClusters} cluster(s). ` +
    `The dominant ESXi version identified is ${dominantVersion}. ` +
    `The busiest port group is ${busiestPortgroup}. ` +
    `Aged hardware count (7+ years from model launch) is ${agedHosts}. ` +
    `A total of ${observations.length} automated observation(s) were generated for quick review.`;
}

function renderObservations(observations) {
  observationsEl.innerHTML = observations.map((item) => {
    const severityClass = `badge-${item.severity.toLowerCase()}`;
    return `
      <div class="observation-item">
        <span class="badge ${severityClass}">${escapeHtml(item.severity)}</span>
        <div class="observation-text">${escapeHtml(item.text)}</div>
      </div>
    `;
  }).join("");
}

function renderTable(container, rows, columns) {
  if (!rows.length) {
    container.innerHTML = '<div class="empty-state">No data available from the uploaded file for this section.</div>';
    return;
  }
  const header = columns.map(([label]) => `<th>${escapeHtml(label)}</th>`).join("");
  const body = rows.map((row) => `<tr>${columns.map(([, key]) => `<td>${escapeHtml(String(row[key] ?? ""))}</td>`).join("")}</tr>`).join("");
  container.innerHTML = `<div class="table-wrap"><table><thead><tr>${header}</tr></thead><tbody>${body}</tbody></table></div>`;
}

function sumBy(rows, keys) {
  return rows.reduce((sum, row) => sum + toNumber(getField(row, keys)), 0);
}

function getField(row, candidates) {
  for (const key of candidates) {
    if (row[key] !== undefined && row[key] !== null && row[key] !== "") return row[key];
  }
  return "";
}

function uniqueCount(items) {
  return new Set(items.filter(Boolean)).size;
}

function toNumber(value) {
  if (typeof value === "number") return value;
  if (typeof value === "string") {
    const normalized = value.replace(/,/g, "").trim();
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : 0;
  }
  return 0;
}

function toText(value) {
  return String(value || "").toLowerCase().replace(/\s+/g, "");
}

function round(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function copyObservations() {
  if (!cachedObservationText) return;
  navigator.clipboard.writeText(cachedObservationText)
    .then(() => { statusEl.textContent = "Observations copied to clipboard."; })
    .catch(() => { statusEl.textContent = "Could not copy observations on this device/browser."; });
}

function downloadCsv(filename, rows) {
  if (!rows || !rows.length) {
    statusEl.textContent = "No data available to export for this section.";
    return;
  }
  const headers = Object.keys(rows[0]);
  const csv = [
    headers.join(","),
    ...rows.map((row) => headers.map((h) => csvEscape(row[h])).join(","))
  ].join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function csvEscape(value) {
  const text = String(value ?? "");
  if (text.includes(",") || text.includes("\n") || text.includes('"')) {
    return '"' + text.replace(/"/g, '""') + '"';
  }
  return text;
}
