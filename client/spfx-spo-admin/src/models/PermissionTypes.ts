/**
 * Permission levels for site-selected permissions
 */
export enum PermissionLevel {
  Read = 'read',
  Write = 'write',
  FullControl = 'write.all'
}

/**
 * UI-friendly permission type labels
 */
export enum PermissionTypeLabel {
  Read = 'Read',
  Write = 'Write',
  FullControl = 'Full Control'
}

/**
 * Authentication mode options
 */
export enum AuthMode {
  User = 'User',
  App = 'App'
}

/**
 * Map UI labels to Graph API permission values
 */
export const PERMISSION_MAPPING: Record<PermissionTypeLabel, PermissionLevel> = {
  [PermissionTypeLabel.Read]: PermissionLevel.Read,
  [PermissionTypeLabel.Write]: PermissionLevel.Write,
  [PermissionTypeLabel.FullControl]: PermissionLevel.FullControl
};

/**
 * Interface for form state
 */
export interface IPermissionFormState {
  siteUrl: string;
  authType: AuthMode;
  appId: string;
  secret: string;
  tenantId: string;
  registryAppId: string;
  permissionType: PermissionTypeLabel;
}

/**
 * Interface for granted identity (application)
 */
export interface IGrantedIdentity {
  application?: {
    id: string;
    displayName?: string;
  };
}

/**
 * Interface for permission response from Graph API
 */
export interface ISitePermission {
  id: string;
  roles: string[];
  grantedToIdentities?: IGrantedIdentity[];
  grantedToIdentitiesV2?: IGrantedIdentity[];
  link?: {
    scope?: string;
    type?: string;
  };
  expirationDateTime?: string;
  hasPassword?: boolean;
}

/**
 * Interface for displaying permissions in the UI
 */
export interface IPermissionDisplay {
  id: string;
  appDisplayName: string;
  appId: string;
  roles: string[];
  grantedDateTime: string;
}

/**
 * Interface for adding new permission request
 */
export interface IAddPermissionRequest {
  roles: string[];
  grantedToIdentities: IGrantedIdentity[];
}

/**
 * Interface for error state
 */
export interface IErrorState {
  hasError: boolean;
  message: string;
}

/**
 * Interface for success state
 */
export interface ISuccessState {
  hasSuccess: boolean;
  message: string;
}
