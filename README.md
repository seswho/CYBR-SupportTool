# CYBR-SupportTool
The predecessor of CyberArk xRay to collect the logs and config files from the CyberArk PAS Windows components.

# Capabilities
- Has the ability to determine the version of PAS on the Vault to find the proper locations of the logs and config files
- Knows if the Vault is a Standalone, HA Cluster, Distributed Vaults, or a DR Vault
- Can detect if the component is a PVWA, CPM, or PSM and collect the logs and config files from the proper locations
- Able to detect if the Windows Credential Provider or Central Credential Provider is installed and running to collect the logs and config files from the proper locations
- If the Windows server is running PA Replicate, the logs and configs can be collected
- Collect the logs from a specific timeframe or the last 72 hours

# TODO
- Documentation on how to use CYBR SupportTool

# Current Version
v2.7

# Change Log
