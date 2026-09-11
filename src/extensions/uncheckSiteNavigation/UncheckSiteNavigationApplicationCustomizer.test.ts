jest.mock('@microsoft/sp-core-library', () => ({ Log: { info: jest.fn(), warn: jest.fn(), error: jest.fn() } }));
jest.mock('@microsoft/sp-application-base', () => ({
  BaseApplicationCustomizer: class { protected onDispose(): void { /* test base */ } },
  PlaceholderName: { Top: 1 }
}));
jest.mock('@microsoft/sp-page-context', () => ({ SPPermission: { manageWeb: 1 } }));
jest.mock('UncheckSiteNavigationApplicationCustomizerStrings', () => ({ Title: 'Test' }), { virtual: true });
jest.mock('./SettingsService', () => ({ SettingsService: jest.fn() }));

import Customizer from './UncheckSiteNavigationApplicationCustomizer';
import { SettingsService } from './SettingsService';


interface ITestCustomizer {
  context: object;
  properties: object;
  onInit: () => Promise<void>;
  onDispose: () => void;
  _refreshSettings: () => void;
  _scanDocument: (doc: Document) => void;
}

function setup(read: () => Promise<boolean>): ITestCustomizer {
  (SettingsService as jest.Mock).mockImplementation(() => ({ getEnabled: read }));
  const extension: ITestCustomizer = new Customizer() as unknown as ITestCustomizer;
  extension.properties = {};
  extension.context = {
    spHttpClient: {},
    pageContext: { web: { absoluteUrl: 'https://example.com/sites/a' }, user: { loginName: 'owner' } },
    application: { navigatedEvent: { add: jest.fn(), remove: jest.fn() } },
    placeholderProvider: { changedEvent: { add: jest.fn(), remove: jest.fn() } }
  };
  return extension;
}

function checkbox(): HTMLInputElement {
  document.body.innerHTML = '<input type="checkbox" checked aria-label="Show in site navigation">';
  return document.querySelector('input') as HTMLInputElement;
}

async function flush(): Promise<void> { await Promise.resolve(); await Promise.resolve(); }

describe('settings gate and navigation lifecycle', () => {
  afterEach(() => { jest.clearAllMocks(); });

  it('does not change a checkbox before the settings request resolves', async () => {
    let complete: (enabled: boolean) => void = () => { /* replaced below */ };
    const pending: Promise<boolean> = new Promise((resolve) => { complete = resolve; });
    const extension: ITestCustomizer = setup(() => pending);
    const control: HTMLInputElement = checkbox();
    await extension.onInit();
    extension._scanDocument(document);
    expect(control.checked).toBe(true);
    complete(true);
    await flush();
    extension._scanDocument(document);
    expect(control.checked).toBe(false);
    extension.onDispose();
  });

  it('keeps SharePoint defaults when reading settings fails', async () => {
    const extension: ITestCustomizer = setup(() => Promise.reject(new Error('Forbidden')));
    const control: HTMLInputElement = checkbox();
    await extension.onInit();
    await flush();
    extension._scanDocument(document);
    expect(control.checked).toBe(true);
    extension.onDispose();
  });

  it('ignores the previous navigation response after a newer disabled result', async () => {
    let complete: (enabled: boolean) => void = () => { /* replaced below */ };
    const first: Promise<boolean> = new Promise((resolve) => { complete = resolve; });
    const read: jest.Mock = jest.fn().mockReturnValueOnce(first).mockResolvedValue(false);
    const extension: ITestCustomizer = setup(read);
    const control: HTMLInputElement = checkbox();
    await extension.onInit();
    extension._refreshSettings();
    await flush();
    complete(true);
    await flush();
    extension._scanDocument(document);
    expect(control.checked).toBe(true);
    expect(SettingsService).toHaveBeenCalledTimes(1);
    extension.onDispose();
  });
});
