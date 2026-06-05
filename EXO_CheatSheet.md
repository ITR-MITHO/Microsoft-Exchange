# Microsoft Exchange & Microsoft 365 Administration Reference

A consolidated reference guide for Exchange Online migrations, mail flow configurations, and essential administrative commands.

## 🔗 Helpful Links

| Resource | URL |
| :--- | :--- |
| **Exchange Admin Center** | [admin.exchange.microsoft.com](https://admin.exchange.microsoft.com) |
| **Security: Policies & Rules** | [security.microsoft.com/threatpolicy](https://security.microsoft.com/threatpolicy) |
| **Security: Threat Explorer** | [security.microsoft.com/threatexplorer](https://security.microsoft.com/threatexplorer) |
| **Security: Quarantine** | [security.microsoft.com/quarantine](https://security.microsoft.com/quarantine) |
| **Security: User Submissions** | [security.microsoft.com/reportsubmission](https://security.microsoft.com/reportsubmission) |
| **Purview: Legacy Retention** | [purview.microsoft.com/datalifecyclemanagement/exchange/retentionpolicies](https://purview.microsoft.com/datalifecyclemanagement/exchange/retentionpolicies) |
| **Exchange Online Limits** | [Microsoft Learn: Exchange Limits](https://learn.microsoft.com/en-us/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits) |
| **Exchange URLs & IPs** | [Microsoft Learn: URLs & IPs](https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide#exchange-online) |
| **Inbound DANE** | [Tech Community: Inbound SMTP DANE](https://techcommunity.microsoft.com/blog/exchange/implementing-inbound-smtp-dane-with-dnssec-for-exchange-online-mail-flow/3939694) |
| **M365 Licenses** | [m365maps.com](https://m365maps.com/) |
| **Decomission the last Exchange Server** | [Decomission the last Exchange Server](https://learn.microsoft.com/en-us/exchange/hybrid-deployment/decommission-last-exchange-server) |

> **Note:** Quarantine notifications are sent from `Quarantine@messaging.microsoft.com`, notifications are NOT sent by the default quarantine policies.

---

## ✉️ Mail Flow & DNS Records

| Record Type | Configuration |
| :--- | :--- |
| **SPF (Standard)** | `v=spf1 spf.protection.outlook.com -all` |
| **SPF (Redirect)** | `v=spf1 redirect=domain.com` |
| **DMARC** | `v=DMARC1; p=reject; pct=100; adkim=s; aspf=s` |
| **MX (Standard)** | `domain-com.mail.protection.outlook.com` |
| **MX-DANE** | `domain-com.l-v1.mx.microsoft` |

---

## 💻 Essential PowerShell Commands

### General Administration

| Description | Command |
| :--- | :--- |
| **Permanently clear previous mailbox info** | `Set-User <MAILBOX> -PermanentlyClearPreviousMailboxInfo -Confirm:$false` |
| **Apply mandatory properties** | `Set-Mailbox <MAILBOX> -ApplyMandatoryProperties` |
| **Configure inbound connector to skip IPs** | `Set-InboundConnector "Relay" -EFSkipIPS 127.0.0.1,127.0.0.2` |

### Migration & Move Requests

| Description | Command |
| :--- | :--- |
| **Approve migration with bad items immediately** | `Set-MoveRequest <MAILBOX> -SkippedItemApprovalTime $(Get-Date).ToUniversalTime()` |
| **Complete move request immediately** | `Set-MoveRequest <MAILBOX> -CompleteAfter 1` |

### Permissions

| Description | Command |
| :--- | :--- |
| **Set calendar permissions (LimitedDetails)** | `Set-MailboxFolderPermission <ALIAS>:\Calendar -User Default -AccessRights LimitedDetails` |
| **Grant Send As permission** | `Add-RecipientPermission <MAILBOX> -Trustee <USERNAME> -AccessRights SendAs -Confirm:$false` |
| **Grant Full Access with Automapping** | `Add-MailboxPermission <MAILBOX> -User <USERNAME> -AccessRights FullAccess -AutoMapping $true` |

### Diagnostics

| Description | Command |
| :--- | :--- |
| **10-day message trace (max 5000 results)** | `Get-MessageTraceV2 -ResultSize 5000 -StartDate (Get-Date).AddDays(-10) -EndDate (Get-Date)` |

---

## ⚙️ Client Configuration (Registry)

| Action | Registry Path | Key (DWORD) | Value |
| :--- | :--- | :--- | :--- |
| **Hide Outlook New Button** | `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\General` | `HideNewOutlookToggle` | `1` |
| **Prevent Outlook NEW auto install** | `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Options\General` | `DoNewOutlookAutoMigration` | `0` |
| **Enable Modern Authentication** | `HKEY_CURRENT_USER\Software\Microsoft\Exchange` | `EnableADAL` | `1` |
| **Always use Exchange Online Autodiscover** | `HKEY_CURRENT_USER\Software\Microsoft\Exchange` | `AlwaysUseMSOAuthForAutoDiscover` | `1` |
| **Allow Outlook to useExchange Online Autodiscover** | HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\AutoDiscover` | `ExcludeExplicitO365Endpoint` | `0` |
