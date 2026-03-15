# RVTools Analyzer

Static browser-based RVTools analyzer for GitHub Pages.

## Included features
- Upload RVTools Excel export locally in the browser
- Summary stats for VMs, hosts, clusters, and datastores
- Cluster summary table
- Host summary table
- ESXi version count table
- VM count by port group
- Hardware age estimation using local launch-year lookup
- Executive summary and quick assessment observations
- CSV export buttons for key tables

## Files to upload
- index.html
- styles.css
- script.js
- hardware-launch-years.json
- README.md

## Notes
Hardware age depends on the values in `hardware-launch-years.json`.
Add more server models there as needed.
