# Azure Risky User to Azure Group Syncronization

This repository deploys a Cloudflare Scheduled Worker which synchronises users flagged by Azure's Risky User API into Groups based on risk level. These groups can be applied to Cloudflare Zero Trust policies to isolate application access.

## Group names
- IdentityProtection-RiskyUser-RiskLevel-high
- IdentityProtection-RiskyUser-RiskLevel-medium
- IdentityProtection-RiskyUser-RiskLevel-low

# Installation

- Install Wrangler CLI: https://developers.cloudflare.com/workers/wrangler/install-and-update/
- `wrangler login` 
- `wrangler publish --var AZURE_AD_TENANT_ID:YOUR_AZURE_TENANT_ID --var AZURE_AD_CLIENT_ID:YOUR_AZURE_CLIENT_ID`
- `wrangler secret put AZURE_AD_CLIENT_SECRET`