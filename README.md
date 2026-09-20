# Exchange Server Field Notes

Microsoft Exchange Server field notes, deep dives, troubleshooting, and practical findings.

This repository contains concise technical notes based on Exchange Server field experience and troubleshooting. Full articles, screenshots, command output, and references are published on [ceyhunkirmizitas.net](https://ceyhunkirmizitas.net/).

## Topics

- [Authentication](authentication/)
  - [OWA Authentication](authentication/owa/)
  - [Exchange Server Auth Certificate](authentication/auth-certificate/)
- [Exchange Hybrid](hybrid/)
- [Migration and Upgrade](migration-upgrade/)
- [High Availability](high-availability/)
- [Security and Hardening](security-hardening/)
- [Troubleshooting](troubleshooting/)

## Scripts

Public PowerShell scripts are maintained in separate repositories:

- [GetExchangeURLs-v2.ps1](https://github.com/Ceyhun-Kirmizitas/GetExchangeURLs-v2.ps1)
- [MonitorExchangeAuthCertificate-TimeZoneAware.ps1](https://github.com/Ceyhun-Kirmizitas/MonitorExchangeAuthCertificate-TimeZoneAware.ps1)

## OWA Authentication Deep Dive

- [Part 1 — Exchange Server OWA Authentication Request Flow: Frontend, HttpProxy and Backend](authentication/owa/part-1-how-owa-authentication-really-works.md)
- [Part 2 — Exchange OWA Forms-Based vs Basic Authentication: HTTP 401, Browser Prompt and IIS Changes](authentication/owa/part-2-forms-based-and-basic-authentication.md)
- [Part 3 — Exchange OWA Frontend vs Backend Authentication: Backend 401 After Successful FBA Login](authentication/owa/part-3-frontend-vs-backend-authentication.md)

## Exchange Server Auth Certificate

- [Exchange Server Auth Certificate Field Guide: Validation, Rotation, Recovery, and Time Zone Issues](https://ceyhunkirmizitas.net/exchange-server-auth-certificate-renewal-recovery-timezone/)

More Exchange Server field notes will be added as the series continues.
