import * as React from 'react';
import { useState, useCallback } from 'react';
import styles from './ManageSiteSelected.module.scss';
import type { IManageSiteSelectedProps } from './IManageSiteSelectedProps';
import {
  TextField,
  PrimaryButton,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  Stack,
  Label,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  CommandBar,
  ICommandBarItemProps,
  Spinner,
  SpinnerSize,
  Pivot,
  PivotItem
} from '@fluentui/react';
import { GraphAuthService } from '../../../services/GraphAuthService';
import { SiteUrlHelper } from '../../../utils/SiteUrlHelper';
import {
  AuthMode,
  PermissionTypeLabel,
  PERMISSION_MAPPING,
  ISitePermission,
  IPermissionDisplay,
  IAddPermissionRequest
} from '../../../models/PermissionTypes';

const ManageSiteSelected: React.FC<IManageSiteSelectedProps> = (props) => {
  // Form state
  const [siteUrl, setSiteUrl] = useState<string>('');
  const [authType, setAuthType] = useState<AuthMode>(AuthMode.User);
  const [appId, setAppId] = useState<string>('');
  const [secret, setSecret] = useState<string>('');
  const [registryAppId, setRegistryAppId] = useState<string>('');
  const [displayName, setDisplayName] = useState<string>('');
  const [permissionType, setPermissionType] = useState<PermissionTypeLabel>(PermissionTypeLabel.Read);

  // UI state
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string>('');
  const [successMessage, setSuccessMessage] = useState<string>('');
  const [permissions, setPermissions] = useState<IPermissionDisplay[]>([]);
  const [selectedPermission, setSelectedPermission] = useState<IPermissionDisplay | null>(null);

  // Dropdown options
  const authTypeOptions: IDropdownOption[] = [
    { key: AuthMode.User, text: 'User Mode' },
    { key: AuthMode.App, text: 'App Mode (Client Credentials)' }
  ];

  const permissionTypeOptions: IDropdownOption[] = [
    { key: PermissionTypeLabel.Read, text: 'Read' },
    { key: PermissionTypeLabel.Write, text: 'Write' },
    { key: PermissionTypeLabel.FullControl, text: 'Full Control' }
  ];

  // Validation helper
  const validateForm = useCallback((forAddPermission: boolean = false): string | null => {
    if (!siteUrl || !SiteUrlHelper.isValidSharePointUrl(siteUrl)) {
      return 'Please enter a valid SharePoint site URL (https://...)';
    }

    if (!registryAppId) {
      return 'Please enter the Registry App ID';
    }

    if (forAddPermission && !displayName) {
      return 'Please enter the Registry App Display Name';
    }

    if (authType === AuthMode.App) {
      if (!appId) return 'App ID is required for App Mode authentication';
      if (!secret) return 'Client Secret is required for App Mode authentication';
    }

    if (forAddPermission && !permissionType) {
      return 'Please select a permission type';
    }

    return null;
  }, [siteUrl, authType, appId, secret, registryAppId, displayName, permissionType]);

  // Show Permissions Handler
  const handleShowPermissions = useCallback(async () => {
    // Clear previous messages
    setErrorMessage('');
    setSuccessMessage('');

    // Validate form
    const validationError = validateForm(false);
    if (validationError) {
      setErrorMessage(validationError);
      return;
    }

    setIsLoading(true);

    try {
      // Initialize Graph client based on auth mode
      let graphClient;
      if (authType === AuthMode.User) {
        graphClient = await GraphAuthService.getUserModeClient(props.context);
      } else {
        const tenantId = props.context.pageContext.aadInfo?.tenantId?.toString() || '';
        graphClient = await GraphAuthService.getAppModeClient(appId, secret, tenantId);
      }

      // Resolve site ID from URL
      const siteId = await SiteUrlHelper.getSiteIdFromUrl(graphClient, siteUrl);

      // Get site permissions
      const response = await graphClient
        .api(`/sites/${siteId}/permissions`)
        .get();

      const sitePermissions: ISitePermission[] = response.value || [];

      // Filter permissions for the registry app ID
      const filteredPermissions = sitePermissions.filter((perm: ISitePermission) => {
        const identities = perm.grantedToIdentitiesV2 || perm.grantedToIdentities || [];
        return identities.some(identity => 
          identity.application?.id?.toLowerCase() === registryAppId.toLowerCase()
        );
      });

      // Transform to display format
      const displayPermissions: IPermissionDisplay[] = filteredPermissions.map((perm: ISitePermission) => {
        const identities = perm.grantedToIdentitiesV2 || perm.grantedToIdentities || [];
        const appIdentity = identities.find(id => id.application);
        
        return {
          id: perm.id,
          appDisplayName: appIdentity?.application?.displayName || 'Unknown',
          appId: appIdentity?.application?.id || registryAppId,
          roles: perm.roles || [],
          grantedDateTime: new Date().toLocaleString() // Note: API might not return this
        };
      });

      setPermissions(displayPermissions);
      setSuccessMessage(`Found ${displayPermissions.length} permission(s) for the specified app`);
    } catch (error) {
      console.error('Error fetching permissions:', error);
      setErrorMessage(error instanceof Error ? error.message : 'Failed to fetch permissions');
    } finally {
      setIsLoading(false);
    }
  }, [authType, siteUrl, appId, secret, registryAppId, props.context, validateForm]);

  // Add Permission Handler
  const handleAddPermission = useCallback(async () => {
    // Clear previous messages
    setErrorMessage('');
    setSuccessMessage('');

    // Validate form
    const validationError = validateForm(true);
    if (validationError) {
      setErrorMessage(validationError);
      return;
    }

    setIsLoading(true);

    try {
      // Initialize Graph client based on auth mode
      let graphClient;
      if (authType === AuthMode.User) {
        graphClient = await GraphAuthService.getUserModeClient(props.context);
      } else {
        const tenantId = props.context.pageContext.aadInfo?.tenantId?.toString() || '';
        graphClient = await GraphAuthService.getAppModeClient(appId, secret, tenantId);
      }

      // Resolve site ID from URL
      const siteId = await SiteUrlHelper.getSiteIdFromUrl(graphClient, siteUrl);

      // Map permission type to Graph API roles
      const mappedPermission = PERMISSION_MAPPING[permissionType];

      // Prepare request payload
      const requestPayload: IAddPermissionRequest = {
        roles: [mappedPermission],
        grantedToIdentities: [
          {
            application: {
              id: registryAppId,
              displayName: displayName
            }
          }
        ]
      };

      // Add permission via Graph API
      await graphClient
        .api(`/sites/${siteId}/permissions`)
        .post(requestPayload);

      setSuccessMessage(`Successfully added ${permissionType} permission for app ${registryAppId}`);

      // Optionally refresh permissions list
      setTimeout(() => {
        void handleShowPermissions();
      }, 1000);
    } catch (error) {
      console.error('Error adding permission:', error);
      setErrorMessage(error instanceof Error ? error.message : 'Failed to add permission');
    } finally {
      setIsLoading(false);
    }
  }, [authType, siteUrl, appId, secret, registryAppId, displayName, permissionType, props.context, validateForm, handleShowPermissions]);

  // Revoke Permission Handler
  const handleRevokePermission = useCallback(async (permissionId: string) => {
    if (!permissionId) return;

    setErrorMessage('');
    setSuccessMessage('');
    setIsLoading(true);

    try {
      // Initialize Graph client
      let graphClient;
      if (authType === AuthMode.User) {
        graphClient = await GraphAuthService.getUserModeClient(props.context);
      } else {
        const tenantId = props.context.pageContext.aadInfo?.tenantId?.toString() || '';
        graphClient = await GraphAuthService.getAppModeClient(appId, secret, tenantId);
      }

      // Resolve site ID
      const siteId = await SiteUrlHelper.getSiteIdFromUrl(graphClient, siteUrl);

      // Revoke permission
      await graphClient
        .api(`/sites/${siteId}/permissions/${permissionId}`)
        .delete();

      setSuccessMessage('Permission revoked successfully');

      // Refresh permissions list
      setTimeout(() => {
        void handleShowPermissions();
      }, 500);
    } catch (error) {
      console.error('Error revoking permission:', error);
      setErrorMessage(error instanceof Error ? error.message : 'Failed to revoke permission');
    } finally {
      setIsLoading(false);
    }
  }, [authType, siteUrl, appId, secret, props.context, handleShowPermissions]);

  // DetailsList columns
  const columns: IColumn[] = [
    {
      key: 'appDisplayName',
      name: 'App Display Name',
      fieldName: 'appDisplayName',
      minWidth: 150,
      maxWidth: 250,
      isResizable: true
    },
    {
      key: 'appId',
      name: 'App ID',
      fieldName: 'appId',
      minWidth: 200,
      maxWidth: 300,
      isResizable: true
    },
    {
      key: 'roles',
      name: 'Granted Roles',
      fieldName: 'roles',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: IPermissionDisplay) => {
        return <span>{item.roles.join(', ')}</span>;
      }
    },
    {
      key: 'grantedDateTime',
      name: 'Granted Date/Time',
      fieldName: 'grantedDateTime',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true
    }
  ];

  // Command bar items for DetailsList
  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'revoke',
      text: 'Revoke Permission',
      iconProps: { iconName: 'Delete' },
      disabled: !selectedPermission,
      onClick: () => {
        if (selectedPermission) {
          handleRevokePermission(selectedPermission.id).catch(err => {
            console.error('Error revoking permission:', err);
          });
        }
      }
    },
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: () => {
        handleShowPermissions().catch(err => {
          console.error('Error refreshing permissions:', err);
        });
      }
    }
  ];

  return (
    <section className={`${styles.manageSiteSelected} ${props.hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
        <h2>SharePoint Site-Selected Permissions Manager</h2>
        <p>Manage app permissions for SharePoint sites using Microsoft Graph API</p>
      </div>

      <Stack tokens={{ childrenGap: 15 }}>
        {/* Error/Success Messages */}
        {errorMessage && (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
            onDismiss={() => setErrorMessage('')}
            dismissButtonAriaLabel="Close"
          >
            {errorMessage}
          </MessageBar>
        )}

        {successMessage && (
          <MessageBar
            messageBarType={MessageBarType.success}
            isMultiline={false}
            onDismiss={() => setSuccessMessage('')}
            dismissButtonAriaLabel="Close"
          >
            {successMessage}
          </MessageBar>
        )}

        {/* Site URL Input */}
        <TextField
          label="SharePoint Site URL"
          placeholder="https://contoso.sharepoint.com/sites/sitename"
          value={siteUrl}
          onChange={(_, newValue) => setSiteUrl(newValue || '')}
          required
        />

        {/* Authentication Type Dropdown */}
        <Dropdown
          label="Authentication Type"
          selectedKey={authType}
          options={authTypeOptions}
          onChange={(_, option) => setAuthType(option?.key as AuthMode)}
          required
        />

        {/* Conditional fields for App Mode */}
        {authType === AuthMode.App && (
          <>
            <TextField
              label="App ID (Client ID)"
              placeholder="Enter application ID"
              value={appId}
              onChange={(_, newValue) => setAppId(newValue || '')}
              required
            />
            <TextField
              label="Client Secret"
              placeholder="Enter client secret"
              type="password"
              value={secret}
              onChange={(_, newValue) => setSecret(newValue || '')}
              required
              canRevealPassword
            />
          </>
        )}

        {/* Registry App ID */}
        <TextField
          label="Registry App ID"
          placeholder="Enter the app ID to manage permissions for"
          value={registryAppId}
          onChange={(_, newValue) => setRegistryAppId(newValue || '')}
          required
        />

        {/* Loading Spinner */}
        {isLoading && (
          <Spinner size={SpinnerSize.large} label="Processing..." />
        )}

        <Pivot>
          <PivotItem headerText="Add Permission">
            <Stack tokens={{ childrenGap: 15 }}>
              {/* Registry App Display Name */}
              <TextField
                label="Registry App Display Name"
                placeholder="Enter the display name of the app"
                value={displayName}
                onChange={(_, newValue) => setDisplayName(newValue || '')}
                required
              />

              {/* Permission Type Dropdown */}
              <Dropdown
                label="Permission Type"
                selectedKey={permissionType}
                options={permissionTypeOptions}
                onChange={(_, option) => setPermissionType(option?.key as PermissionTypeLabel)}
                required
              />

              {/* Add Permission Button */}
              <DefaultButton
                text="Add Permission"
                onClick={handleAddPermission}
                disabled={isLoading}
              />
            </Stack>
          </PivotItem>

          <PivotItem headerText="List Permissions">
            <Stack tokens={{ childrenGap: 15 }}>
              {/* Show Permissions Button */}
              <PrimaryButton
                text="Show Permissions"
                onClick={handleShowPermissions}
                disabled={isLoading}
              />

              {/* Permissions List */}
              {permissions.length > 0 && (
                <Stack tokens={{ childrenGap: 10 }}>
                  <Label>Current Permissions</Label>
                  <CommandBar items={commandBarItems} />
                  <DetailsList
                    items={permissions}
                    columns={columns}
                    setKey="set"
                    layoutMode={DetailsListLayoutMode.justified}
                    selectionMode={SelectionMode.single}
                    onActiveItemChanged={(item) => setSelectedPermission(item)}
                  />
                </Stack>
              )}
            </Stack>
          </PivotItem>
        </Pivot>
      </Stack>
    </section>
  );
};

export default ManageSiteSelected;
