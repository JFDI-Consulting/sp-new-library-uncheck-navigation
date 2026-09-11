jest.mock('@microsoft/sp-http', () => ({
  SPHttpClient: { configurations: { v1: {} } }
}));
jest.mock('@microsoft/sp-page-context', () => {
  class Permission {
    public static editListItems: Permission = new Permission({ High: 0, Low: 4 });
    public static manageLists: Permission = new Permission({ High: 0, Low: 2048 });
    public static managePermissions: Permission = new Permission({ High: 0, Low: 33554432 });
    public constructor(public readonly value: { High: number; Low: number }) { }
    public hasPermission(requested: Permission): boolean {
      return (this.value.High & requested.value.High) === requested.value.High &&
        (this.value.Low & requested.value.Low) === requested.value.Low;
    }
  }
  return { SPPermission: Permission };
});
jest.mock('UncheckSiteNavigationApplicationCustomizerStrings', () => ({
  SettingsPermissionDenied: 'permission denied',
  SettingsInvalidResponse: 'invalid response',
  SettingsIncomplete: 'incomplete',
  SettingsConflict: 'conflict',
  SettingsProvisionPermission: 'provision denied',
  SettingsStoreCollision: 'collision',
  SettingsAssociatedGroupsMissing: 'groups missing',
  SettingsPermissionsIncomplete: 'permissions incomplete',
  SettingsRequestFailed: 'request {0}{1}'
}), { virtual: true });

import { COMPONENT_ID, SETTINGS_CACHE_KEY_PREFIX, SETTINGS_CACHE_TTL_MS, SettingsService } from './SettingsService';

const WEB_URL: string = 'https://example.com/sites/a';
const USER_KEY: string = 'i:0#.f|membership|owner@example.com';

interface ITestResponse {
  ok: boolean;
  status: number;
  /** Mirrors the platform Headers.get contract. */
  // eslint-disable-next-line @rushstack/no-new-null
  headers: { get: (name: string) => string | null };
  json: () => Promise<unknown>;
}

function response(status: number, body: unknown, etag?: string): ITestResponse {
  return {
    ok: status >= 200 && status < 300,
    status,
    headers: { get: (name: string) => name.toLowerCase() === 'etag' ? etag || null : null },
    json: () => Promise.resolve(body)
  };
}

function item(enabled: boolean, etag?: string): object {
  return { value: [{ Id: 1, Enabled: enabled, JfdiUnavStoreId: COMPONENT_ID, 'odata.etag': etag }] };
}

function client(get: jest.Mock, post: jest.Mock = jest.fn()): never {
  return { get, post } as never;
}

function cacheKey(webUrl: string = WEB_URL, userKey: string = USER_KEY): string {
  return SETTINGS_CACHE_KEY_PREFIX + `${encodeURIComponent(webUrl.toLowerCase())}:${encodeURIComponent(userKey)}`;
}

describe('SettingsService', () => {
  beforeEach(() => {
    window.sessionStorage.clear();
    jest.restoreAllMocks();
  });

  it('deduplicates a cold read and lets another instance reuse its session cache', async () => {
    let finish: (value: ITestResponse) => void = () => { /* assigned by promise */ };
    const pending: Promise<ITestResponse> = new Promise((resolve) => { finish = resolve; });
    const get: jest.Mock = jest.fn().mockReturnValue(pending);
    const service: SettingsService = new SettingsService(client(get), WEB_URL, USER_KEY);

    const first: Promise<boolean> = service.getEnabled();
    const second: Promise<boolean> = service.getEnabled();
    expect(get).toHaveBeenCalledTimes(1);
    finish(response(200, item(false)));
    await expect(Promise.all([first, second])).resolves.toEqual([false, false]);

    const cachedClient: jest.Mock = jest.fn();
    await expect(new SettingsService(client(cachedClient), WEB_URL, USER_KEY).getEnabled()).resolves.toBe(false);
    expect(cachedClient).not.toHaveBeenCalled();
  });

  it('refreshes an expired cache entry', async () => {
    window.sessionStorage.setItem(cacheKey(), JSON.stringify({ enabled: false, expiresAt: Date.now() - 1 }));
    const get: jest.Mock = jest.fn().mockResolvedValue(response(200, item(true)));
    await expect(new SettingsService(client(get), WEB_URL, USER_KEY).getEnabled()).resolves.toBe(true);
    expect(get).toHaveBeenCalledTimes(1);
  });

  it('keeps caches separate for each web and user', async () => {
    const otherWeb: string = 'https://example.com/sites/b';
    const otherUser: string = 'other@example.com';
    await new SettingsService(client(jest.fn().mockResolvedValue(response(200, item(false)))),
      WEB_URL, USER_KEY).getEnabled();
    await new SettingsService(client(jest.fn().mockResolvedValue(response(200, item(true)))),
      WEB_URL, otherUser).getEnabled();
    await new SettingsService(client(jest.fn().mockResolvedValue(response(200, item(true)))),
      otherWeb, USER_KEY).getEnabled();

    expect(JSON.parse(window.sessionStorage.getItem(cacheKey()) as string).enabled).toBe(false);
    expect(JSON.parse(window.sessionStorage.getItem(cacheKey(WEB_URL, otherUser)) as string).enabled).toBe(true);
    expect(JSON.parse(window.sessionStorage.getItem(cacheKey(otherWeb, USER_KEY)) as string).enabled).toBe(true);
  });

  it('decodes the web path once before encoding the GetList parameter', async () => {
    const get: jest.Mock = jest.fn().mockResolvedValue(response(404, {}));
    await new SettingsService(client(get), 'https://example.com/sites/A%20B/%E2%9C%93', USER_KEY).getEnabled();

    const requestUrl: URL = new URL(get.mock.calls[0][0]);
    expect(requestUrl.searchParams.get('@list')).toBe("'/sites/A B/✓/Lists/JfdiUnavSettings'");
    expect(get.mock.calls[0][0]).not.toContain('%2520');
  });

  it('uses the default only for a confirmed missing list and rejects other failures', async () => {
    const missing: jest.Mock = jest.fn().mockResolvedValue(response(404, {}));
    const missingService: SettingsService = new SettingsService(client(missing), WEB_URL, USER_KEY, false);
    await expect(missingService.getEnabled()).resolves.toBe(false);
    expect(missingService.needsProvisioning()).toBe(true);

    window.sessionStorage.clear();
    const forbidden: jest.Mock = jest.fn().mockResolvedValue(response(403, {}));
    await expect(new SettingsService(client(forbidden), WEB_URL, USER_KEY).getEnabled())
      .rejects.toThrow('permission denied');
    expect(window.sessionStorage.getItem(cacheKey())).toBeNull();
  });

  it('waits for an ordinary in-flight read before making a forced ETag read', async () => {
    let finish: (value: ITestResponse) => void = () => { /* assigned by promise */ };
    const pending: Promise<ITestResponse> = new Promise((resolve) => { finish = resolve; });
    const get: jest.Mock = jest.fn()
      .mockReturnValueOnce(pending)
      .mockResolvedValueOnce(response(200, item(false, '"2"')));
    const service: SettingsService = new SettingsService(client(get), WEB_URL, USER_KEY);
    const ordinary: Promise<boolean> = service.getEnabled();
    const forced: Promise<boolean> = service.getEnabled(true);
    expect(get).toHaveBeenCalledTimes(1);
    finish(response(200, item(true)));
    await expect(ordinary).resolves.toBe(true);
    await expect(forced).resolves.toBe(false);
    expect(get).toHaveBeenCalledTimes(2);
  });

  it('saves with the last ETag and clears the cache after a conflict', async () => {
    const get: jest.Mock = jest.fn().mockResolvedValue(response(200, item(true, '"7"')));
    const post: jest.Mock = jest.fn().mockResolvedValue(response(412, {}));
    const service: SettingsService = new SettingsService(client(get, post), WEB_URL, USER_KEY);
    await service.getEnabled(true);

    await expect(service.setEnabled(false)).rejects.toThrow('conflict');
    expect(post.mock.calls[0][2].headers['IF-MATCH']).toBe('"7"');
    expect(window.sessionStorage.getItem(cacheKey())).toBeNull();
  });

  it('does not let a pre-save read overwrite the saved cache or resolve stale', async () => {
    let finishOldRead: (value: ITestResponse) => void = () => { /* assigned by promise */ };
    const oldRead: Promise<ITestResponse> = new Promise((resolve) => { finishOldRead = resolve; });
    const get: jest.Mock = jest.fn()
      .mockReturnValueOnce(oldRead)
      .mockResolvedValueOnce(response(200, item(true, '"9"')));
    const post: jest.Mock = jest.fn().mockResolvedValue(response(204, {}, '"10"'));
    const service: SettingsService = new SettingsService(client(get, post), WEB_URL, USER_KEY);

    const supersededRead: Promise<boolean> = service.getEnabled();
    const save: Promise<void> = service.setEnabled(false);
    await save;
    finishOldRead(response(200, item(true)));

    await expect(supersededRead).resolves.toBe(false);
    expect(JSON.parse(window.sessionStorage.getItem(cacheKey()) as string).enabled).toBe(false);
  });

  it('waits for a save before starting a navigation read', async () => {
    let finishSave: (value: ITestResponse) => void = () => { /* assigned by promise */ };
    const pendingSave: Promise<ITestResponse> = new Promise((resolve) => { finishSave = resolve; });
    const get: jest.Mock = jest.fn().mockResolvedValue(response(200, item(true, '"11"')));
    const post: jest.Mock = jest.fn().mockReturnValue(pendingSave);
    const service: SettingsService = new SettingsService(client(get, post), WEB_URL, USER_KEY);
    await service.getEnabled(true);
    const save: Promise<void> = service.setEnabled(false);
    const navigationRead: Promise<boolean> = service.getEnabled();
    expect(get).toHaveBeenCalledTimes(1);

    finishSave(response(204, {}, '"12"'));
    await save;
    await expect(navigationRead).resolves.toBe(false);
    expect(get).toHaveBeenCalledTimes(1);
  });

  it('does not mutate SharePoint when first-Save provisioning permissions are absent', async () => {
    const get: jest.Mock = jest.fn()
      .mockResolvedValueOnce(response(404, {}))
      .mockResolvedValueOnce(response(200, { High: 0, Low: 2048 }));
    const post: jest.Mock = jest.fn();
    const service: SettingsService = new SettingsService(client(get, post), WEB_URL, USER_KEY);
    await service.getEnabled();

    await expect(service.setEnabled(false)).rejects.toThrow('provision denied');
    expect(post).not.toHaveBeenCalled();
  });

  it('rejects cache entries with an implausibly long lifetime', async () => {
    window.sessionStorage.setItem(cacheKey(), JSON.stringify({
      enabled: false,
      expiresAt: Date.now() + SETTINGS_CACHE_TTL_MS + 1000
    }));
    const get: jest.Mock = jest.fn().mockResolvedValue(response(200, item(true)));
    await expect(new SettingsService(client(get), WEB_URL, USER_KEY).getEnabled()).resolves.toBe(true);
    expect(get).toHaveBeenCalledTimes(1);
  });
});
