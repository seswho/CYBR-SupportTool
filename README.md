# CYBR-SupportTool
The predecessor of CyberArk xRay to collect the logs and config files from the CyberArk EPV Windows components.

## Current Version
v2.7

## Capabilities
- Has the ability to determine the version of EPV (Enterprise Password Vault) on the Vault to find the proper locations of the logs and config files, knows if the Vault is a Standalone, HA Cluster, Distributed Vaults, or a DR Vault
- Can detect if the component is a PVWA (Password Vault Web Access), CPM (Central Policy Manager), or PSM (Privelege Session Manager) and collect the logs and config files from the proper locations, including any AppLocker events from the PSM
- Able to detect if the Windows Credential Provider or Central Credential Provider is installed and running to collect the logs and config files from the proper locations
- Collect the logs from a specific timeframe or the last 72 hours
- Will collect logs and configs from EPM (Endpoint Protection Manager) servers

## TODO
- Current documentation on how to use CYBR SupportTool
- If the Windows server is running Replicate, collect the logs and config files

## Change Log
