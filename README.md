# Azure Discovery Report Generator

This utility performs read-only Azure CLI discovery against the current Azure subscription context and generates a client-facing DOCX in the same meeting-report structure and visual treatment used for the Acceleeration Holdings discovery document.

## Prerequisites

- macOS, Linux, or Windows
- Azure CLI installed
- An authenticated Azure CLI session
- Python 3.10+
- Python packages in `requirements.txt`

## Install

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
az login
```

## Run against the current Azure CLI subscription

```bash
python3 azure_discovery_report.py \
  --output Azure_Current_State_and_Project_Discovery.docx \
  --region "East US"
```

## Run against an explicit subscription

```bash
python3 azure_discovery_report.py \
  --subscription "44574a66-06fe-440f-a547-e3e27267674b" \
  --output Azure_Current_State_and_Project_Discovery.docx \
  --region "East US"
```

The script uses read-only Azure CLI queries. It does not create, update, move, delete, or modify Azure resources.

## What it discovers

The report gathers:

- Tenant and subscription context
- Resource groups
- Subscription resource inventory
- VNets and address spaces
- Subnets, NSGs, and route-table associations
- VNet DNS settings
- NSG rules
- NIC/private IP/public IP associations
- Public IPs
- Private endpoints
- Private DNS zones
- Route tables
- Virtual network gateways
- Azure Firewall
- NAT gateways
- Application Gateways
- ExpressRoute circuits
- VPN/connection objects
- Local Network Gateways
- Load balancers
- Resource-type counts

Several discovery objects use `az resource list --resource-type` rather than `az network ... list` because the Azure CLI version used in the project required a resource group for some network list commands.

## Output behavior

The generated report has the same major sections used for the client meeting document:

1. What exists today
2. Current environment at a glance
3. Existing services discovered
4. Network discovery findings
5. Existing security exposure observed
6. Proposed direction for the new project
7. Conceptual target state
8. Decisions required from the client
9. Immediate next steps
10. Appendix A - Current Azure facts from discovery

The report intentionally distinguishes successful negative discovery from failed discovery. A failed command is recorded as a discovery exception rather than being interpreted as "nothing exists."

## Notes

The proposed architecture text defaults to East US because that is the current project direction. The IP plan remains a client input and is not invented by the tool.

The report includes a generated conceptual architecture figure. It is a working-state diagram, not an automated network topology export.

## Virtual machine discovery

The report now includes **all Azure VMs returned by `az vm list -d`**. For each VM it reports:

- Resource group and region
- Power state
- Operating system
- VM size
- Private IP address(es)
- Public IP address(es)
- VNet/subnet placement
- NIC name(s)
- NSG association(s)
- Attached OS/data disks with size and SKU
- VM extensions discovered in the subscription
- Availability zone when specified

The VM discovery uses `az vm show`, `az network nic show`, and `az disk show` with the IDs returned by the VM configuration. All of these operations are read-only.
