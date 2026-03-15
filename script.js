
document.getElementById("upload").addEventListener("change", handleFile);

let hardwareLaunchYears = {};

fetch("hardware-launch-years.json")
.then(res => res.json())
.then(data => hardwareLaunchYears = data);

function handleFile(e){

const file = e.target.files[0];
const reader = new FileReader();

reader.onload = function(evt){

const data = new Uint8Array(evt.target.result);
const workbook = XLSX.read(data, {type:'array'});

const vinfo = XLSX.utils.sheet_to_json(workbook.Sheets["vInfo"] || []);
const vhost = XLSX.utils.sheet_to_json(workbook.Sheets["vHost"] || []);
const vnic = XLSX.utils.sheet_to_json(workbook.Sheets["vNIC"] || []);

generateSummary(vinfo,vhost);
generateEsxiTable(vhost);
generatePortgroupTable(vnic);
generateHardwareAge(vhost);

};

reader.readAsArrayBuffer(file);

}

function generateSummary(vinfo,vhost){

let totalVM=vinfo.length;
let totalHosts=vhost.length;

let html = `<table>
<tr><th>Total VMs</th><th>Total Hosts</th></tr>
<tr><td>${totalVM}</td><td>${totalHosts}</td></tr>
</table>`;

document.getElementById("summary").innerHTML = html;

}

function generateEsxiTable(vhost){

const counts={};

vhost.forEach(row=>{
let version=row["ESX Version"] || row["Version"] || "Unknown";
counts[version]=(counts[version]||0)+1;
});

let html="<table><tr><th>ESXi Version</th><th>Host Count</th></tr>";

Object.entries(counts).forEach(([v,c])=>{
html+=`<tr><td>${v}</td><td>${c}</td></tr>`;
});

html+="</table>";

document.getElementById("esxiTable").innerHTML=html;

}

function portgroupComment(count){

if(count==0) return "No VM attachment observed";
if(count<=20) return "Limited VM footprint observed";
if(count<=50) return "Moderate VM concentration observed";
if(count<=100) return "High VM concentration observed";
return "Very high VM concentration observed; segmentation and resiliency should be reviewed";

}

function generatePortgroupTable(vnic){

const groups={};

vnic.forEach(row=>{

let pg=row["Portgroup"] || row["Network"] || "Unknown";
let vm=row["VM"] || row["VM Name"] || row["VMName"];

if(!groups[pg]) groups[pg]=new Set();
if(vm) groups[pg].add(vm);

});

let html="<table><tr><th>Port Group</th><th>VM Count Observed</th><th>Assessment Comment</th></tr>";

Object.entries(groups).forEach(([pg,vms])=>{

let count=vms.size;
let comment=portgroupComment(count);

html+=`<tr>
<td>${pg}</td>
<td>${count}</td>
<td>${comment}</td>
</tr>`;

});

html+="</table>";

document.getElementById("portgroupTable").innerHTML=html;

generateObservation(groups);

}

function generateHardwareAge(vhost){

const currentYear=new Date().getFullYear();

let html="<table><tr><th>Host</th><th>Model</th><th>Launch Year</th><th>Hardware Age</th></tr>";

vhost.forEach(row=>{

let host=row["Host"] || row["Name"] || "Unknown";
let model=row["Model"] || "Unknown";

let launchYear=hardwareLaunchYears[model];
let age=launchYear ? currentYear-launchYear : "Unknown";

html+=`<tr>
<td>${host}</td>
<td>${model}</td>
<td>${launchYear || "Unknown"}</td>
<td>${age}</td>
</tr>`;

});

html+="</table>";

document.getElementById("hardwareTable").innerHTML=html;

}

function generateObservation(groups){

let highNetworks=[];

Object.entries(groups).forEach(([pg,vms])=>{
if(vms.size>100) highNetworks.push(pg);
});

let text="";

if(highNetworks.length>0){

text = "Certain port groups host a very high concentration of workloads (" + highNetworks.join(", ") + "). These networks should be reviewed for segmentation, traffic isolation, and resiliency considerations.";

}else{

text = "VM workloads appear reasonably distributed across available port groups based on the uploaded RVTools dataset.";

}

document.getElementById("observations").innerText=text;

}
