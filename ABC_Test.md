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

> **Note:** Quarantine notifications are sent from `Quarantine@messaging.microsoft.com`.

---

## ✉️ Mail Flow & DNS Records

Standard baseline configurations for routing and authentication.

```text
# SPF Records
Standard:   v=spf1 spf.protection.outlook.com -all
Redirection: v=spf1 redirect=domain.com

# DMARC Record
v=DMARC1; p=reject; pct=100; adkim=s; aspf=s

# MX Records
Standard: domain-com.mail.protection.outlook.com
MX-DANE:  domain-com.l-v1.mx.microsoft
