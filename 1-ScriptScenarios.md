# Exchange Migration Script Reference

This document maps the required PowerShell scripts and automation tools to their respective Exchange migration scenarios. To ensure a comprehensive analytical approach, each migration type is broken down into three strict phases: **Discovery & Assessment**, **Execution**, and **Validation & Cleanup**.

---

## 1. Exchange On-Premises to Exchange Online (Native/Hybrid)

### Phase 1: Discovery & Assessment
**Objective:** Identify environment blockers, stale objects, directory synchronization issues, and invalid attributes prior to migration.

#### Script: `Get-PreFlightHybridChecks.ps1`
* **Repository Path:** `\ExchangeOnline\Migrations\`
* **Description:** Scans for invalid characters in proxy addresses, duplicate UPNs, and ensures DirSync requirements are met.
* **Execution:**
```powershell
.\Get-PreFlightHybridChecks.ps1 -TargetOU "OU=Users,DC=domain,DC=com" -ExportPath "C:\Temp\PreFlight.csv"
```

### Phase 2: Execution
**Objective:** Initiate, monitor, and complete remote move requests.

#### Script: `Invoke-BatchedRemoteMove.ps1`
* **Repository Path:** `\ExchangeOnline\Migrations\`
* **Description:** Reads a CSV of targets, constructs migration batches, and initiates the remote move request to EXO.
* **Execution:**
```powershell
.\Invoke-BatchedRemoteMove.ps1 -UserList "C:\Temp\Batch1.csv" -TargetDeliveryDomain "tenant.mail.onmicrosoft.com" -SuspendWhenReadyToComplete
```

### Phase 3: Validation & Cleanup
**Objective:** Verify mailbox accessibility, mail flow routing, and update on-premises AD objects to RemoteMailbox.

#### Script: `Verify-HybridMailFlow.ps1`
* **Repository Path:** `\ExchangeOnline\Validation\`
* **Description:** Injects synthetic test messages between on-premises and EXO mailboxes, validating header paths.

---

## 2. Exchange On-Premises to Exchange On-Premises
*(e.g., Cross-Forest, Cross-Hardware, or Legacy to Modern)*

### Phase 1: Discovery & Assessment
**Objective:** Map existing database distribution, storage latency, and identify configuration mismatches (e.g., ghost AD paths).

#### Script: `Get-DatabaseCapacityAndHealth.ps1`
* **Repository Path:** `\OnPrem\Database\`
* **Description:** Calculates required physical hardware footprint for database seeding, factoring in whitespace and log volume.
* **Execution:**
```powershell
.\Get-DatabaseCapacityAndHealth.ps1 -Server "EXCH01" -IncludeLogDrives
```

### Phase 2: Execution (Migration & Seeding)
**Objective:** Execute local move requests, handle database mounting, and replication topology configuration.

#### Script: `Invoke-LocalMailboxMove.ps1`
* **Repository Path:** `\OnPrem\Migrations\`
* **Description:** Moves mailboxes to target databases while throttling based on target server resource health.
* **Execution:**
```powershell
.\Invoke-LocalMailboxMove.ps1 -Database "DB01_Legacy" -TargetDatabase "DB02_New" -BadItemLimit 10
```

### Phase 3: Validation & Cleanup
**Objective:** Confirm database replication health, remove legacy configurations, and clear legacy attributes.

#### Script: `Remove-LegacyExchangeArtifacts.ps1`
* **Repository Path:** `\OnPrem\Cleanup\`
* **Description:** Identifies and cleans up legacy SystemFolderPath attributes or arbitrary objects post-migration.

---

## 3. Exchange On-Premises to Exchange Online (MigrationWiz)

### Phase 1: Discovery & Assessment
**Objective:** Provision target identities, apply licenses, and extract on-premises statistics for BitTitan project creation.

#### Script: `Export-MWizUserRoster.ps1`
* **Repository Path:** `\MigrationWiz\`
* **Description:** Dumps primary SMTP, alias lists, and mailbox sizes into the required MWiz CSV import format.
* **Execution:**
```powershell
.\Export-MWizUserRoster.ps1 -OU "OU=Migrate,DC=domain,DC=com" -OutFile "MWiz_Import.csv"
```

### Phase 2: Execution (Pre-Stage & Delta)
**Objective:** Control the automated provisioning of the routing domain and manage the cutover pipeline.

#### Script: `Set-TargetRoutingAddresses.ps1`
* **Repository Path:** `\MigrationWiz\`
* **Description:** Stamps the `.onmicrosoft.com` routing address on target mailboxes prior to the MX cutover to ensure coexistence routing during the delta sync.

### Phase 3: Validation & Cleanup
**Objective:** Finalize MX/Autodiscover cutover and dismantle the source environment access.

#### Script: `Convert-SourceToMailUser.ps1`
* **Repository Path:** `\MigrationWiz\`
* **Description:** Strips the on-premises mailbox attributes, converting the user to a Mail-Enabled User (MEU) pointing to the external routing address.
* **Execution:**
```powershell
.\Convert-SourceToMailUser.ps1 -Identity "user@domain.com" -ExternalEmailAddress "user@tenant.mail.onmicrosoft.com"
```
