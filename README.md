# spo-admin

## Summary

A SharePoint Online administration tool built with SharePoint Framework (SPFx) that provides a user-friendly interface to manage site-selected permissions for applications using Microsoft Graph API. This web part allows administrators to grant, view, and revoke permissions for apps on specific SharePoint sites.

<img width="1387" height="634" alt="image" src="https://github.com/user-attachments/assets/91545b3b-5625-4bd4-b0a3-c48e65475e9b" />


![SharePoint Framework](https://img.shields.io/badge/SharePoint%20Framework-1.22.1-green.svg)
![React](https://img.shields.io/badge/React-17.0.1-blue.svg)
![TypeScript](https://img.shields.io/badge/TypeScript-4.7.4-blue.svg)

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.22.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)


## Prerequisites

- Microsoft 365 developer tenant
- Node.js (version 22)
- SharePoint Framework development environment set up v1.22
- Azure AD app registration (for app mode authentication)
- Appropriate permissions in Microsoft Graph (Sites.ReadWrite.All, etc.)
-Tenant Admin Rights for adding App Permissions to a site.

## Solution

| Solution    | Author(s)                               |
| ----------- | --------------------------------------- |
| spo-admin   | Sagar Pandey (sagarpandey88)            |

## Version history

| Version | Date             | Comments                          |
| ------- | ---------------- | --------------------------------- |
| 1.0     | December 18, 2025| Added tabs for add/list permissions, display name support |


## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Navigate to the client/spfx-spo-admin folder
- Run `npm install`
- Run `npm run start` to start the local development server
- Add the web part to a SharePoint page to test

## Features

This SharePoint Framework web part demonstrates the following concepts:

- **Authentication Modes**: Support for both user-mode (delegated) and app-mode (application) authentication with Microsoft Graph
- **Permission Management**: Grant site-selected permissions (Read, Write, Full Control) to applications for specific SharePoint sites
- **Permission Listing**: View current permissions granted to a specific application on a site
- **Permission Revocation**: Remove permissions from applications
- **Responsive UI**: Built with Fluent UI components for a consistent Microsoft 365 experience
- **Error Handling**: Comprehensive error handling and user feedback

### Key Components

- **ManageSiteSelected Web Part**: Main component with tabbed interface for Adding  App permissions for a site  and listing all the permissions
-

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [Heft Documentation](https://heft.rushstack.io/)
