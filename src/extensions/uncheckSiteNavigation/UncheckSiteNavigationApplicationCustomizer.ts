import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import { SPPermission } from '@microsoft/sp-page-context';

import * as strings from 'UncheckSiteNavigationApplicationCustomizerStrings';
import { SettingsService } from './SettingsService';
import type { ISettingsUiHandle } from './SettingsUi';

const LOG_SOURCE: string = 'UncheckSiteNavigationApplicationCustomizer';

/** Query-string key that opens the settings panel on Site contents. */
export const SETTINGS_QUERY_KEY: string = 'jfdiUncheckNav';
const SETTINGS_QUERY_VALUE: string = 'settings';
const SITE_CONTENTS_PATH: RegExp = /\/_layouts\/15\/viewlsts\.aspx$/i;

/**
 * Properties supplied via ClientSideComponentProperties on the custom action.
 */
export interface IUncheckSiteNavigationApplicationCustomizerProperties {
  /**
   * Label texts (case-insensitive, whitespace-trimmed) that identify the
   * checkbox/toggle to default to unchecked. Add localised variants here
   * if your tenant UI language is not English.
   */
  labels?: string[];
  /** Emit verbose console logging. */
  debug?: boolean;
  /**
   * Switch the behaviour on or off for this site. Absent means on. Site
   * owners can change this from the settings panel on Site contents; it is
   * stored back into this same property on the custom action.
   */
  enabled?: boolean;
}

const DEFAULT_LABELS: string[] = [
  'Show in site navigation',
  'Show list in site navigation',
  'Show library in site navigation'
];

/** Delays between verification retries after a click, in milliseconds. */
const RETRY_DELAYS_MS: number[] = [150, 400, 800, 1500];

const CONTROL_SELECTOR: string = 'input[type="checkbox"], [role="switch"], [role="checkbox"]';

/**
 * Application Customizer that watches the page for the modern
 * "Create list" / "Create document library" experience and flips the
 * "Show in site navigation" checkbox off the first time it appears.
 *
 * SharePoint has no server-side setting for this default, so the customizer
 * works on the rendered DOM. The create-list experience is hosted in a
 * same-origin iframe (createlist.aspx) where SPFx extensions do not load, so
 * the customizer observes the top document AND every same-origin iframe it
 * can reach, re-hooking each iframe whenever it navigates.
 *
 * Each control element is only touched once, so a user who deliberately
 * re-ticks the box is not overridden.
 */
export default class UncheckSiteNavigationApplicationCustomizer
  extends BaseApplicationCustomizer<IUncheckSiteNavigationApplicationCustomizerProperties> {

  private _observers: MutationObserver[] = [];
  private _hookedDocs: WeakSet<Document> = new WeakSet<Document>();
  private _hookedFrames: WeakSet<HTMLIFrameElement> = new WeakSet<HTMLIFrameElement>();
  private _processed: WeakSet<Element> = new WeakSet<Element>();
  private _labels: string[] = [];
  private _scanScheduled: boolean = false;
  private _disposed: boolean = false;
  private _enabled: boolean = true;
  private _topPlaceholder: PlaceholderContent | undefined;
  private _settingsUi: ISettingsUiHandle | undefined;
  private _mounting: boolean = false;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    const configured: string[] | undefined = this.properties && this.properties.labels;
    const source: string[] = Array.isArray(configured) && configured.length > 0 ? configured : DEFAULT_LABELS;
    this._labels = source.map(this._normalise).filter((l: string) => l.length > 0);
    this._enabled = SettingsService.isEnabled(this.properties);
    this._log(`Enabled: ${this._enabled}`);

    if (typeof MutationObserver === 'undefined' || !document.body) {
      Log.warn(LOG_SOURCE, 'MutationObserver or document.body unavailable; customizer inactive.');
      return Promise.resolve();
    }

    this._hookDocument(document);
    this._scheduleScan();

    // Modern SharePoint navigates client-side without re-running onInit, so
    // the settings UI has to follow route and placeholder changes.
    this.context.application.navigatedEvent.add(this, this._syncSettingsUi);
    this.context.placeholderProvider.changedEvent.add(this, this._syncSettingsUi);
    this._syncSettingsUi();

    return Promise.resolve();
  }

  protected onDispose(): void {
    this._disposed = true;
    for (const o of this._observers) {
      o.disconnect();
    }
    this._observers = [];
    this.context.application.navigatedEvent.remove(this, this._syncSettingsUi);
    this.context.placeholderProvider.changedEvent.remove(this, this._syncSettingsUi);
    this._unmountSettingsUi();
    super.onDispose();
  }

  /**
   * Mount the settings UI on Site contents and remove it everywhere else.
   * Idempotent: safe to call on every navigation and placeholder change.
   *
   * On Site contents, users who can manage the web get a status bar in the
   * Top placeholder with a link to the settings panel. The panel also opens
   * when the page URL carries ?jfdiUncheckNav=settings, so it can be linked
   * from anywhere. Site Settings itself is a classic page where this
   * customizer cannot run, and NoScript sites refuse Site Settings links.
   */
  private _syncSettingsUi(): void {
    if (this._disposed) {
      return;
    }
    if (!SITE_CONTENTS_PATH.test(window.location.pathname)) {
      this._unmountSettingsUi();
      return;
    }
    if (this._settingsUi || this._mounting) {
      return;
    }
    const canManage: boolean = this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb);
    const requested: boolean = new URLSearchParams(window.location.search).get(SETTINGS_QUERY_KEY) === SETTINGS_QUERY_VALUE;
    if (!canManage && !requested) {
      return;
    }

    const placeholder: PlaceholderContent | undefined = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
    if (!placeholder) {
      this._log('Top placeholder not available yet; waiting for placeholderProvider.changedEvent.');
      return;
    }
    this._topPlaceholder = placeholder;
    this._mounting = true;

    const service: SettingsService = new SettingsService(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl);
    // React and Fluent UI live in a separate chunk so every other page load
    // pays nothing for a feature only site owners see on Site contents.
    import(/* webpackChunkName: 'uncheck-nav-settings' */ './SettingsUi')
      .then((mod) => {
        this._mounting = false;
        if (this._disposed || this._topPlaceholder !== placeholder) {
          return;
        }
        this._settingsUi = mod.renderSettingsUi(placeholder.domElement, {
          enabled: this._enabled,
          canEdit: canManage,
          openOnMount: requested,
          onSave: async (enabled: boolean): Promise<void> => {
            await service.setEnabled(enabled);
            this._enabled = enabled;
            this._log(`Setting saved: enabled=${enabled}`);
            this._scheduleScan();
          }
        });
        if (requested) {
          this._stripSettingsQuery();
        }
      })
      .catch((e: Error) => {
        this._mounting = false;
        Log.error(LOG_SOURCE, e);
      });
  }

  private _unmountSettingsUi(): void {
    if (this._settingsUi) {
      this._settingsUi.unmount();
      this._settingsUi = undefined;
    }
    if (this._topPlaceholder) {
      this._topPlaceholder.dispose();
      this._topPlaceholder = undefined;
    }
  }

  /** Remove only our own query parameter so a refresh does not reopen the panel. */
  private _stripSettingsQuery(): void {
    const search: string = window.location.search.replace(/^\?/, '');
    const kept: string[] = search.split('&').filter((p: string) => p && p.split('=')[0] !== SETTINGS_QUERY_KEY);
    const url: string = window.location.pathname + (kept.length ? `?${kept.join('&')}` : '') + window.location.hash;
    window.history.replaceState(window.history.state, '', url);
  }

  /** Observe a document (top or iframe) for DOM changes. Idempotent per document. */
  private _hookDocument(doc: Document): void {
    if (this._hookedDocs.has(doc) || !doc.body) {
      return;
    }
    this._hookedDocs.add(doc);
    const observer: MutationObserver = new MutationObserver(() => this._scheduleScan());
    observer.observe(doc.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['checked', 'aria-checked'] });
    this._observers.push(observer);
    this._log(`Observing document ${doc.location ? doc.location.pathname : '(unknown)'}`);
  }

  /** Find same-origin iframes in a document and hook their documents. */
  private _hookFrames(doc: Document): void {
    const frames: HTMLIFrameElement[] = Array.from(doc.querySelectorAll('iframe'));
    for (const frame of frames) {
      if (!this._hookedFrames.has(frame)) {
        this._hookedFrames.add(frame);
        // Re-hook after every navigation inside the frame (new document each time).
        frame.addEventListener('load', () => this._scheduleScan());
      }
      const inner: Document | null = this._frameDocument(frame);
      if (inner) {
        this._hookDocument(inner);
      }
    }
  }

  private _frameDocument(frame: HTMLIFrameElement): Document | null {
    try {
      const doc: Document | null = frame.contentDocument;
      return doc && doc.body ? doc : null; // cross-origin access throws or returns null
    } catch {
      return null;
    }
  }

  /** Coalesce bursts of mutations into a single scan per animation frame. */
  private _scheduleScan(): void {
    if (this._scanScheduled || this._disposed) {
      return;
    }
    this._scanScheduled = true;
    const run = (): void => {
      this._scanScheduled = false;
      try {
        this._scanDocument(document);
      } catch (e) {
        Log.error(LOG_SOURCE, e instanceof Error ? e : new Error(String(e)));
      }
    };
    if (typeof window.requestAnimationFrame === 'function') {
      window.requestAnimationFrame(run);
    } else {
      window.setTimeout(run, 50);
    }
  }

  private _scanDocument(doc: Document): void {
    this._hookFrames(doc);

    if (!this._enabled) {
      return; // switched off for this site; keep observing so a save takes effect live
    }

    const controls: Element[] = Array.from(doc.querySelectorAll(CONTROL_SELECTOR));
    for (const control of controls) {
      if (this._processed.has(control)) {
        continue;
      }
      const label: string = this._normalise(this._getAccessibleName(control));
      if (!label || this._labels.indexOf(label) === -1) {
        continue;
      }
      this._processed.add(control);
      this._uncheck(control, label);
    }

    // Recurse into hooked same-origin iframes.
    const frames: HTMLIFrameElement[] = Array.from(doc.querySelectorAll('iframe'));
    for (const frame of frames) {
      const inner: Document | null = this._frameDocument(frame);
      if (inner) {
        this._scanDocument(inner);
      }
    }
  }

  private _isChecked(control: Element): boolean {
    return control instanceof (control.ownerDocument.defaultView || window).HTMLInputElement
      ? (control as HTMLInputElement).checked
      : control.getAttribute('aria-checked') === 'true';
  }

  /**
   * Click the control to uncheck it, then verify. The create-list experience
   * can render the checkbox before React has attached its handlers (or reset
   * it on a later re-render), so a dropped click is retried with backoff.
   */
  private _uncheck(control: Element, label: string, attempt: number = 0): void {
    if (!this._isChecked(control)) {
      this._log(`"${label}" ${attempt === 0 ? 'already unchecked' : 'defaulted to unchecked'}.`);
      return;
    }

    // Use a real click so the React-controlled component updates its state.
    (control as HTMLElement).click();

    if (attempt >= RETRY_DELAYS_MS.length) {
      this._log(`"${label}" could not be unchecked after ${attempt + 1} attempts.`);
      return;
    }
    window.setTimeout(() => {
      if (this._disposed || !control.isConnected) {
        return;
      }
      this._uncheck(control, label, attempt + 1);
    }, RETRY_DELAYS_MS[attempt]);
  }

  /** Best-effort accessible name: aria-label, aria-labelledby, <label for>, wrapping <label>. */
  private _getAccessibleName(el: Element): string {
    const doc: Document = el.ownerDocument;

    const ariaLabel: string | null = el.getAttribute('aria-label');
    if (ariaLabel) {
      return ariaLabel;
    }

    const labelledBy: string | null = el.getAttribute('aria-labelledby');
    if (labelledBy) {
      const text: string = labelledBy
        .split(/\s+/)
        .map((id: string) => {
          const node: HTMLElement | null = doc.getElementById(id);
          return node ? node.textContent || '' : '';
        })
        .join(' ');
      if (text.trim()) {
        return text;
      }
    }

    if (el.id) {
      const forLabel: Element | null = doc.querySelector(`label[for="${CSS.escape(el.id)}"]`);
      if (forLabel && forLabel.textContent) {
        return forLabel.textContent;
      }
    }

    const wrapping: Element | null = el.closest('label');
    if (wrapping && wrapping.textContent) {
      return wrapping.textContent;
    }

    return '';
  }

  private _normalise(text: string): string {
    return (text || '').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  private _log(message: string): void {
    if (this.properties && this.properties.debug) {
      console.log(`[${LOG_SOURCE}] ${message}`);
    }
  }
}
