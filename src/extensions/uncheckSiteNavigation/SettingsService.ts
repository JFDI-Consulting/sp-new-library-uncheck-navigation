import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';

import * as strings from 'UncheckSiteNavigationApplicationCustomizerStrings';

/** Component id from UncheckSiteNavigationApplicationCustomizer.manifest.json. */
export const COMPONENT_ID: string = 'd31c6f18-3a0d-462b-b677-c09314fbf3e6';
export const SETTINGS_CACHE_TTL_MS: number = 60 * 1000;
export const SETTINGS_CACHE_KEY_PREFIX: string = 'jfdi-unav:settings:v1:';

const LIST_PATH: string = 'Lists/JfdiUnavSettings';
const LIST_TITLE: string = 'JfdiUnavSettings';
const LIST_DESCRIPTION: string = `JFDI Uncheck Navigation settings store:${COMPONENT_ID}`;
const ITEM_KEY: string = 'configuration';
const MARKER_FIELD: string = 'JfdiUnavStoreId';

interface ISettingItem {
  Id: number;
  Enabled: boolean;
  JfdiUnavStoreId: string;
  '@odata.etag'?: string;
  'odata.etag'?: string;
}

interface ISettingState {
  exists: boolean;
  enabled: boolean;
  itemId?: number;
  etag?: string;
}

interface IBasePermissions {
  High: number | string;
  Low: number | string;
}

interface IListMetadata {
  Id: string;
  BaseTemplate: number;
  Description: string;
  Hidden: boolean;
  RootFolder?: { ServerRelativeUrl?: string };
}

interface IRoleAssignment {
  PrincipalId?: number;
  Member?: { Id?: number };
  RoleDefinitionBindings?: Array<{ Id?: number }>;
}

interface IFieldDefinition {
  InternalName: string;
  TypeAsString: string;
  Required: boolean;
  Indexed: boolean;
  EnforceUniqueValues: boolean;
}

/** Hidden-list-backed, site-local configuration for the customizer. */
export class SettingsService {
  private readonly _webUrl: string;
  private readonly _listServerRelativeUrl: string;
  private readonly _cacheKey: string;
  private _readPromise: Promise<boolean> | undefined;
  private _readIsForced: boolean = false;
  private _writePromise: Promise<void> | undefined;
  private _setting: ISettingState | undefined;
  private _cacheGeneration: number = 0;

  public constructor(
    private readonly _client: SPHttpClient,
    webUrl: string,
    userKey: string,
    private readonly _defaultEnabled: boolean = true
  ) {
    this._webUrl = webUrl.replace(/\/$/, '');
    const webPath: string = decodeURIComponent(new URL(this._webUrl).pathname).replace(/\/$/, '');
    this._listServerRelativeUrl = `${webPath}/${LIST_PATH}`;
    this._cacheKey = SETTINGS_CACHE_KEY_PREFIX +
      `${encodeURIComponent(this._webUrl.toLowerCase())}:${encodeURIComponent(userKey)}`;
  }

  /** Read once per instance, backed by a 60-second per-user/per-web cache. */
  public getEnabled(force: boolean = false): Promise<boolean> {
    if (this._writePromise) {
      return this._writePromise.then(
        () => this.getEnabled(force),
        () => this.getEnabled(true)
      );
    }
    if (this._readPromise) {
      if (force && !this._readIsForced) {
        return this._readPromise.then(() => this.getEnabled(true));
      }
      return this._readPromise;
    }
    if (!force) {
      const cached: boolean | undefined = this._readCache();
      if (cached !== undefined) {
        return Promise.resolve(cached);
      }
    } else {
      this._clearCache();
    }

    const generation: number = this._cacheGeneration;
    this._readIsForced = force;
    const read: Promise<boolean> = this._readRemote(force).then((state: ISettingState) => {
      if (generation !== this._cacheGeneration) {
        const write: Promise<void> | undefined = this._writePromise;
        return write ? write.then(
          () => this.getEnabled(force),
          () => this.getEnabled(true)
        ) : this.getEnabled(force);
      }
      this._setting = state;
      this._writeCache(state.enabled);
      return state.enabled;
    });
    this._readPromise = read.then(
      (enabled: boolean) => {
        if (generation === this._cacheGeneration) {
          this._readPromise = undefined;
          this._readIsForced = false;
        }
        return enabled;
      },
      (error: unknown) => {
        if (generation === this._cacheGeneration) {
          this._readPromise = undefined;
          this._readIsForced = false;
          this._clearCache();
        }
        throw error;
      }
    );
    return this._readPromise;
  }

  /** True only after an authoritative read confirmed that the list is absent. */
  public needsProvisioning(): boolean {
    return this._setting !== undefined && !this._setting.exists;
  }

  /** Check list edit rights, or the two permissions needed for first-Save setup. */
  public async canEdit(): Promise<boolean> {
    try {
      const listResponse: SPHttpClientResponse = await this._get(this._listUrl('/EffectiveBasePermissions'));
      if (listResponse.ok) {
        return this._hasPermission(await this._readPermissions(listResponse), SPPermission.editListItems);
      }
      if (listResponse.status !== 404) {
        return false;
      }
      const webResponse: SPHttpClientResponse = await this._get(`${this._webUrl}/_api/web/EffectiveBasePermissions`);
      if (!webResponse.ok) {
        return false;
      }
      const permissions: IBasePermissions = await this._readPermissions(webResponse);
      return this._hasPermission(permissions, SPPermission.manageLists) &&
        this._hasPermission(permissions, SPPermission.managePermissions);
    } catch {
      return false;
    }
  }

  /** Save with the ETag obtained by the latest uncached read. */
  public setEnabled(enabled: boolean): Promise<void> {
    this._cacheGeneration++;
    this._readPromise = undefined;
    this._readIsForced = false;
    this._clearCache();
    const previous: Promise<void> | undefined = this._writePromise;
    const operation: Promise<void> = (previous ? previous.catch(() => undefined) : Promise.resolve())
      .then(() => this._saveEnabled(enabled));
    this._writePromise = operation;
    operation.then(
      () => { if (this._writePromise === operation) { this._writePromise = undefined; } },
      () => { if (this._writePromise === operation) { this._writePromise = undefined; } }
    );
    return operation;
  }

  private async _saveEnabled(enabled: boolean): Promise<void> {
    try {
      if (!this._setting || (this._setting.exists && (!this._setting.itemId || !this._setting.etag))) {
        this._setting = await this._readRemote(true);
      }
      if (!this._setting.exists) {
        await this._provision(enabled);
        this._setting = await this._readRemote(true);
        this._writeCache(this._setting.enabled);
        return;
      }
      const response: SPHttpClientResponse = await this._client.post(
        this._listUrl(`/items(${this._setting.itemId})`), SPHttpClient.configurations.v1,
        {
          headers: this._headers({ 'X-HTTP-Method': 'MERGE', 'IF-MATCH': this._setting.etag as string }),
          body: JSON.stringify({ Enabled: enabled })
        }
      );
      if (!response.ok) {
        if (response.status === 412) {
          throw new Error(strings.SettingsConflict);
        }
        throw new Error(await this._describe(response));
      }
      this._setting = { exists: true, enabled, itemId: this._setting.itemId,
        etag: response.headers.get('ETag') || undefined };
      this._writeCache(enabled);
    } catch (error) {
      this._setting = undefined;
      this._clearCache();
      throw error;
    }
  }

  private async _readRemote(requireEtag: boolean = false): Promise<ISettingState> {
    const response: SPHttpClientResponse = await this._get(this._listUrl('/items',
      `$select=Id,Enabled,${MARKER_FIELD}&$filter=${encodeURIComponent(`Title eq '${ITEM_KEY}'`)}&$top=2`), requireEtag);
    if (response.status === 404) {
      return { exists: false, enabled: this._defaultEnabled };
    }
    if (!response.ok) {
      throw new Error(await this._describe(response));
    }
    let data: { value?: ISettingItem[] };
    try {
      data = await response.json();
    } catch {
      throw new Error(strings.SettingsInvalidResponse);
    }
    const items: ISettingItem[] = data.value || [];
    if (items.length !== 1 || typeof items[0].Id !== 'number' || typeof items[0].Enabled !== 'boolean' ||
      items[0].JfdiUnavStoreId !== COMPONENT_ID) {
      throw new Error(strings.SettingsIncomplete);
    }
    let resolvedEnabled: boolean = items[0].Enabled;
    let etag: string | undefined = items[0]['@odata.etag'] || items[0]['odata.etag'];
    if (requireEtag && !etag) {
      const itemResponse: SPHttpClientResponse = await this._get(this._listUrl(`/items(${items[0].Id})`,
        `$select=Id,Enabled,${MARKER_FIELD}`), true);
      if (!itemResponse.ok) {
        throw new Error(await this._describe(itemResponse));
      }
      const item: ISettingItem & { __metadata?: { etag?: string } } = await itemResponse.json();
      if (item.Id !== items[0].Id || typeof item.Enabled !== 'boolean' || item.JfdiUnavStoreId !== COMPONENT_ID) {
        throw new Error(strings.SettingsIncomplete);
      }
      resolvedEnabled = item.Enabled;
      etag = item['@odata.etag'] || item['odata.etag'] || (item.__metadata && item.__metadata.etag) ||
        itemResponse.headers.get('ETag') || undefined;
    }
    if (requireEtag && !etag) {
      throw new Error(strings.SettingsInvalidResponse);
    }
    return { exists: true, enabled: resolvedEnabled, itemId: items[0].Id, etag };
  }

  private async _provision(enabled: boolean): Promise<void> {
    if (!await this._hasBootstrapPermissions()) {
      throw new Error(strings.SettingsProvisionPermission);
    }
    const create: SPHttpClientResponse = await this._client.post(`${this._webUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1, {
        headers: this._headers(),
        body: JSON.stringify({ Title: LIST_TITLE, Description: LIST_DESCRIPTION, BaseTemplate: 100,
          Hidden: true, NoCrawl: true, OnQuickLaunch: false, EnableAttachments: false,
          EnableFolderCreation: false, EnableVersioning: true })
      });
    if (!create.ok) {
      if (create.status === 409) {
        await this._validateList();
        throw new Error(strings.SettingsIncomplete);
      }
      throw new Error(await this._describe(create));
    }
    await this._validateList();
    await this._secureNewList();
    await this._createSchema();
    await this._expectOk(await this._client.post(this._listUrl('/items'), SPHttpClient.configurations.v1, {
      headers: this._headers(),
      body: JSON.stringify({ Title: ITEM_KEY, Enabled: enabled, JfdiUnavStoreId: COMPONENT_ID })
    }));
  }

  private async _validateList(): Promise<void> {
    const response: SPHttpClientResponse = await this._get(this._listUrl('',
      '$select=Id,BaseTemplate,Description,Hidden,RootFolder/ServerRelativeUrl&$expand=RootFolder'));
    if (!response.ok) {
      throw new Error(await this._describe(response));
    }
    const list: IListMetadata = await response.json();
    const actualPath: string = list.RootFolder && list.RootFolder.ServerRelativeUrl || '';
    if (!list.Id || list.BaseTemplate !== 100 || list.Description !== LIST_DESCRIPTION || !list.Hidden ||
      actualPath.toLowerCase() !== this._listServerRelativeUrl.toLowerCase()) {
      throw new Error(strings.SettingsStoreCollision);
    }
  }

  /** Secure the newly-created empty list before adding its schema or value. */
  private async _secureNewList(): Promise<void> {
    await this._expectOk(await this._client.post(
      this._listUrl('/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)'),
      SPHttpClient.configurations.v1, { headers: this._headers() }));
    const response: SPHttpClientResponse = await this._get(`${this._webUrl}/_api/web?` +
      '$select=AssociatedOwnerGroup/Id,AssociatedMemberGroup/Id,AssociatedVisitorGroup/Id&' +
      '$expand=AssociatedOwnerGroup,AssociatedMemberGroup,AssociatedVisitorGroup');
    if (!response.ok) {
      throw new Error(await this._describe(response));
    }
    const groups: {
      AssociatedOwnerGroup?: { Id?: number };
      AssociatedMemberGroup?: { Id?: number };
      AssociatedVisitorGroup?: { Id?: number };
    } = await response.json();
    const ownerId: number | undefined = groups.AssociatedOwnerGroup && groups.AssociatedOwnerGroup.Id;
    const memberId: number | undefined = groups.AssociatedMemberGroup && groups.AssociatedMemberGroup.Id;
    const visitorId: number | undefined = groups.AssociatedVisitorGroup && groups.AssociatedVisitorGroup.Id;
    if (!ownerId || !memberId || !visitorId || ownerId === memberId || ownerId === visitorId || memberId === visitorId) {
      throw new Error(strings.SettingsAssociatedGroupsMissing);
    }
    const fullControlId: number = await this._getRoleDefinitionId(5);
    const readId: number = await this._getRoleDefinitionId(2);
    await this._addRole(ownerId, fullControlId);
    await this._addRole(memberId, readId);
    await this._addRole(visitorId, readId);

    const expected: number[] = [ownerId, memberId, visitorId];
    for (const assignment of await this._getRoleAssignments()) {
      const principalId: number | undefined = assignment.PrincipalId || (assignment.Member && assignment.Member.Id);
      if (principalId && expected.indexOf(principalId) === -1) {
        for (const binding of assignment.RoleDefinitionBindings || []) {
          if (binding.Id) {
            await this._expectOk(await this._client.post(this._listUrl(
              `/roleassignments/removeroleassignment(principalid=${principalId},roledefid=${binding.Id})`),
            SPHttpClient.configurations.v1, { headers: this._headers() }));
          }
        }
      }
    }
    const assignments: IRoleAssignment[] = await this._getRoleAssignments();
    if (assignments.length !== 3 || !this._hasRole(assignments, ownerId, fullControlId) ||
      !this._hasRole(assignments, memberId, readId) || !this._hasRole(assignments, visitorId, readId)) {
      throw new Error(strings.SettingsPermissionsIncomplete);
    }
  }

  private async _createSchema(): Promise<void> {
    await this._createField({ Title: 'Enabled', FieldTypeKind: 8, Required: true });
    await this._createField({ Title: MARKER_FIELD, FieldTypeKind: 2, Required: true });
    await this._expectOk(await this._client.post(this._listUrl("/fields/getbyinternalnameortitle('Title')"),
      SPHttpClient.configurations.v1, {
        headers: this._headers({ 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }),
        body: JSON.stringify({ Required: true, Indexed: true, EnforceUniqueValues: true })
      }));
    const response: SPHttpClientResponse = await this._get(this._listUrl('/fields',
      "$select=InternalName,TypeAsString,Required,Indexed,EnforceUniqueValues&$filter=InternalName eq 'Title' or InternalName eq 'Enabled' or InternalName eq 'JfdiUnavStoreId'"));
    if (!response.ok) {
      throw new Error(await this._describe(response));
    }
    const data: { value?: IFieldDefinition[] } = await response.json();
    const fields: IFieldDefinition[] = data.value || [];
    const title: IFieldDefinition | undefined = fields.filter((f: IFieldDefinition) => f.InternalName === 'Title')[0];
    const value: IFieldDefinition | undefined = fields.filter((f: IFieldDefinition) => f.InternalName === 'Enabled')[0];
    const marker: IFieldDefinition | undefined = fields.filter((f: IFieldDefinition) => f.InternalName === MARKER_FIELD)[0];
    if (!title || !title.Required || !title.Indexed || !title.EnforceUniqueValues ||
      !value || value.TypeAsString !== 'Boolean' || !value.Required ||
      !marker || marker.TypeAsString !== 'Text' || !marker.Required) {
      throw new Error(strings.SettingsIncomplete);
    }
  }

  private async _createField(definition: Record<string, unknown>): Promise<void> {
    await this._expectOk(await this._client.post(this._listUrl('/fields'), SPHttpClient.configurations.v1,
      { headers: this._headers(), body: JSON.stringify(definition) }));
  }

  private async _hasBootstrapPermissions(): Promise<boolean> {
    const response: SPHttpClientResponse = await this._get(`${this._webUrl}/_api/web/EffectiveBasePermissions`);
    if (!response.ok) {
      return false;
    }
    const permissions: IBasePermissions = await this._readPermissions(response);
    return this._hasPermission(permissions, SPPermission.manageLists) &&
      this._hasPermission(permissions, SPPermission.managePermissions);
  }

  private async _getRoleDefinitionId(type: number): Promise<number> {
    const response: SPHttpClientResponse = await this._get(
      `${this._webUrl}/_api/web/roledefinitions/getbytype(${type})?$select=Id`);
    if (!response.ok) {
      throw new Error(await this._describe(response));
    }
    const role: { Id?: number } = await response.json();
    if (!role.Id) {
      throw new Error(strings.SettingsInvalidResponse);
    }
    return role.Id;
  }

  private async _addRole(principalId: number, roleId: number): Promise<void> {
    await this._expectOk(await this._client.post(this._listUrl(
      `/roleassignments/addroleassignment(principalid=${principalId},roledefid=${roleId})`),
    SPHttpClient.configurations.v1, { headers: this._headers() }));
  }

  private async _getRoleAssignments(): Promise<IRoleAssignment[]> {
    const response: SPHttpClientResponse = await this._get(this._listUrl('/roleassignments',
      '$select=PrincipalId,Member/Id,RoleDefinitionBindings/Id&$expand=Member,RoleDefinitionBindings'));
    if (!response.ok) {
      throw new Error(await this._describe(response));
    }
    const data: { value?: IRoleAssignment[] } = await response.json();
    return data.value || [];
  }

  private _hasRole(assignments: IRoleAssignment[], principalId: number, roleId: number): boolean {
    return assignments.some((assignment: IRoleAssignment) => {
      const actual: number | undefined = assignment.PrincipalId || (assignment.Member && assignment.Member.Id);
      return actual === principalId && (assignment.RoleDefinitionBindings || [])
        .some((binding: { Id?: number }) => binding.Id === roleId);
    });
  }

  private async _readPermissions(response: SPHttpClientResponse): Promise<IBasePermissions> {
    const data: { High?: number | string; Low?: number | string } = await response.json();
    if (data.High === undefined || data.Low === undefined) {
      throw new Error(strings.SettingsInvalidResponse);
    }
    return { High: data.High, Low: data.Low };
  }

  private _hasPermission(value: IBasePermissions, permission: SPPermission): boolean {
    return new SPPermission({ High: Number(value.High), Low: Number(value.Low) }).hasPermission(permission);
  }

  private _listUrl(suffix: string = '', query: string = ''): string {
    const quoted: string = `'${this._listServerRelativeUrl.replace(/'/g, "''")}'`;
    return `${this._webUrl}/_api/web/GetList(@list)${suffix}?@list=${encodeURIComponent(quoted)}` +
      (query ? `&${query}` : '');
  }

  private _get(url: string, minimalMetadata: boolean = false): Promise<SPHttpClientResponse> {
    return this._client.get(url, SPHttpClient.configurations.v1,
      { headers: { 'Accept': minimalMetadata ? 'application/json;odata=minimalmetadata' : 'application/json;odata=nometadata' } });
  }

  private _headers(additional?: Record<string, string>): Record<string, string> {
    return { 'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata', ...additional };
  }

  private async _expectOk(response: SPHttpClientResponse): Promise<void> {
    if (!response.ok) {
      throw new Error(await this._describe(response));
    }
  }

  private _readCache(): boolean | undefined {
    try {
      const raw: string | null = window.sessionStorage.getItem(this._cacheKey);
      if (!raw) {
        return undefined;
      }
      const entry: { expiresAt?: number; enabled?: boolean } = JSON.parse(raw);
      const now: number = Date.now();
      if (typeof entry.enabled !== 'boolean' || typeof entry.expiresAt !== 'number' ||
        !Number.isFinite(entry.expiresAt) || entry.expiresAt <= now || entry.expiresAt > now + SETTINGS_CACHE_TTL_MS) {
        window.sessionStorage.removeItem(this._cacheKey);
        return undefined;
      }
      return entry.enabled;
    } catch {
      return undefined;
    }
  }

  private _writeCache(enabled: boolean): void {
    try {
      window.sessionStorage.setItem(this._cacheKey,
        JSON.stringify({ expiresAt: Date.now() + SETTINGS_CACHE_TTL_MS, enabled }));
    } catch { /* storage may be disabled */ }
  }

  private _clearCache(): void {
    try {
      window.sessionStorage.removeItem(this._cacheKey);
    } catch { /* storage may be disabled */ }
  }

  private async _describe(response: SPHttpClientResponse): Promise<string> {
    if (response.status === 403) {
      return strings.SettingsPermissionDenied;
    }
    let detail: string = '';
    try {
      const body: { 'odata.error'?: { message?: { value?: string } } } = await response.json();
      detail = body['odata.error'] && body['odata.error'].message && body['odata.error'].message.value || '';
    } catch { /* no JSON body */ }
    return strings.SettingsRequestFailed.replace('{0}', String(response.status))
      .replace('{1}', detail ? `: ${detail}` : '');
  }
}
