import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/** Component id from UncheckSiteNavigationApplicationCustomizer.manifest.json. */
export const COMPONENT_ID: string = 'd31c6f18-3a0d-462b-b677-c09314fbf3e6';

export interface ICustomActionProperties {
  labels?: string[];
  debug?: boolean;
  enabled?: boolean;
}

interface IUserCustomAction {
  Id: string;
  ClientSideComponentId: string;
  /** The REST API returns JSON null when no properties were ever set. */
  // eslint-disable-next-line @rushstack/no-new-null
  ClientSideComponentProperties: string | null;
}

interface ILocatedAction {
  action: IUserCustomAction;
  /** REST path segment the action lives under: web scope or site-collection scope. */
  scope: 'web' | 'site';
}

const GUID_RE: RegExp = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Reads and writes this customizer's own ClientSideComponentProperties on
 * the user custom action that hosts it (web scope first, then site
 * collection scope). Saving needs Manage Web, which site owners have;
 * nothing here requires custom script.
 */
export class SettingsService {
  public constructor(private readonly _client: SPHttpClient, private readonly _webUrl: string) { }

  /** Whether the customizer is switched on. Absent means on. */
  public static isEnabled(props: ICustomActionProperties | undefined): boolean {
    return !props || props.enabled !== false;
  }

  public async setEnabled(enabled: boolean): Promise<ICustomActionProperties> {
    const located: ILocatedAction = await this._getAction();
    const props: ICustomActionProperties = SettingsService._parse(located.action.ClientSideComponentProperties);
    props.enabled = enabled;

    const url: string = `${this._webUrl}/_api/${located.scope}/UserCustomActions('${encodeURIComponent(located.action.Id)}')`;
    const response: SPHttpClientResponse = await this._client.post(url, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*'
      },
      body: JSON.stringify({ ClientSideComponentProperties: JSON.stringify(props) })
    });
    if (!response.ok) {
      throw new Error(await SettingsService._describe(response));
    }
    return props;
  }

  private async _getAction(): Promise<ILocatedAction> {
    for (const scope of ['web', 'site'] as const) {
      const actions: IUserCustomAction[] = await this._listActions(scope);
      if (actions.length > 1) {
        throw new Error(`Found ${actions.length} copies of the customizer's custom action at ${scope} scope; remove the duplicates before changing this setting.`);
      }
      if (actions.length === 1) {
        if (!GUID_RE.test(actions[0].Id)) {
          throw new Error('The custom action id returned by SharePoint was not a GUID.');
        }
        return { action: actions[0], scope };
      }
    }
    throw new Error('The customizer\'s custom action was not found on this site or site collection. If the app is deployed tenant-wide, change the setting in the app catalog\'s Tenant Wide Extensions list instead.');
  }

  private async _listActions(scope: 'web' | 'site'): Promise<IUserCustomAction[]> {
    const url: string = `${this._webUrl}/_api/${scope}/UserCustomActions` +
      `?$select=Id,ClientSideComponentId,ClientSideComponentProperties` +
      `&$filter=ClientSideComponentId eq guid'${COMPONENT_ID}'`;
    const response: SPHttpClientResponse = await this._client.get(url, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=nometadata' }
    });
    if (!response.ok) {
      throw new Error(await SettingsService._describe(response));
    }
    const data: { value?: IUserCustomAction[] } = await response.json();
    return data.value || [];
  }

  /**
   * Absent/empty means "never configured". Anything else must be a JSON
   * object; refusing to guess here is what stops a toggle from silently
   * replacing a site's customised labels with `{"enabled":true}`.
   */
  private static _parse(raw: string | null): ICustomActionProperties {
    if (!raw || !raw.trim()) {
      return {};
    }
    let parsed: unknown;
    try {
      parsed = JSON.parse(raw);
    } catch {
      throw new Error('The custom action\'s existing properties are not valid JSON. Fix them with CLI for Microsoft 365 or PnP PowerShell, then try again.');
    }
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      throw new Error('The custom action\'s existing properties are not a JSON object. Fix them with CLI for Microsoft 365 or PnP PowerShell, then try again.');
    }
    return parsed as ICustomActionProperties;
  }

  private static async _describe(response: SPHttpClientResponse): Promise<string> {
    if (response.status === 403) {
      return 'You need to be a site owner (Manage Web permission) to change this setting.';
    }
    let detail: string = '';
    try {
      const body: { 'odata.error'?: { message?: { value?: string } } } = await response.json();
      detail = (body['odata.error'] && body['odata.error'].message && body['odata.error'].message.value) || '';
    } catch { /* no JSON body */ }
    return `Request failed (${response.status})${detail ? `: ${detail}` : ''}`;
  }
}
