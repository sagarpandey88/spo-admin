# Plan: Build Site-Selected Permissions Manager (Revised)

Transform the ManageSiteSelected webpart into a comprehensive interface for managing SharePoint site-selected app permissions using Microsoft Graph API, PnPjs for authentication, and Fluent UI components with functional React patterns.

## Steps

1. **Install dependencies** — Add `@pnp/sp@^4.0.0`, `@pnp/graph@^4.0.0`, and `@microsoft/microsoft-graph-client@^3.0.0` to [package.json](client/spfx-spo-admin/package.json), then run npm install.

2. **Create authentication service** — Build [src/services/GraphAuthService.ts](client/spfx-spo-admin/src/services/GraphAuthService.ts) with methods: `getUserModeClient(context)` using SPFx MSGraphClientFactory, and `getAppModeClient(appId, secret, tenantId)` using client credentials flow with @microsoft/microsoft-graph-client.

3. **Create utility helpers** — Build [src/utils/SiteUrlHelper.ts](client/spfx-spo-admin/src/utils/SiteUrlHelper.ts) with `getSiteIdFromUrl(graphClient, siteUrl)` that parses URL and calls `GET /sites/{hostname}:/{serverRelativePath}` to resolve site ID.

4. **Build permission mapping** — Create [src/models/PermissionTypes.ts](client/spfx-spo-admin/src/models/PermissionTypes.ts) with enum/constants mapping UI dropdown values (Read → "read", Write → "write", Full Control → "write.all") and TypeScript interfaces for form state and permission response objects.

5. **Refactor to functional component** — Convert [ManageSiteSelected.tsx](client/spfx-spo-admin/src/webparts/manageSiteSelected/components/ManageSiteSelected.tsx) to functional component using useState for form fields (siteUrl, authType, appId, secret, registryAppId, permissionType), loading states, permissions array, and error/success messages.

6. **Build form UI with Fluent UI** — Add TextField components (Site URL, App ID, Secret, Registry App ID), Dropdown components (Authentication Type with options "User"/"App", Permission Type with "Read"/"Write"/"Full Control"), conditional rendering to show/hide app credentials fields, and Stack/StackItem for layout with proper spacing per [ManageSiteSelected.module.scss](client/spfx-spo-admin/src/webparts/manageSiteSelected/components/ManageSiteSelected.module.scss).

7. **Implement "Show Permissions" handler** — Create async function that: resolves site ID from URL, initializes Graph client based on auth mode, calls `GET /sites/{siteId}/permissions`, filters results for the registry app ID, updates state with permissions array, and handles errors with MessageBar display.

8. **Implement "Add Permission" handler** — Create async function that: resolves site ID, initializes Graph client, calls `POST /sites/{siteId}/permissions` with payload `{roles: [mappedPermission], grantedToIdentities: [{application: {id: registryAppId}}]}`, shows success MessageBar, and optionally refreshes permissions list.

9. **Build permissions DetailsList** — Add DetailsList component displaying columns: App Display Name, App ID, Granted Roles (array), Granted Date/Time; make sortable and filterable; add command bar with Revoke action button calling `DELETE /sites/{siteId}/permissions/{permissionId}`.

10. **Add validation and error handling** — Validate Site URL format, require all fields based on auth mode, wrap Graph calls in try-catch with specific error messages (invalid URL, auth failure, Graph API errors), use MessageBar with MessageBarType.error/success/info for user feedback.

## Implementation Notes

- Use `useCallback` for button handlers to prevent unnecessary re-renders
- Store tenant ID in webpart properties or extract from SPFx context for app-mode authentication
- Graph API permission scope required: `Sites.FullControl.All` (must be approved in SharePoint Admin Center API Management)
- Reference endpoint documentation: https://learn.microsoft.com/en-us/graph/permissions-selected-overview?tabs=http
