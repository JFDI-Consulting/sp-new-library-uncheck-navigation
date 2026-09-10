import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';

import * as strings from 'UncheckSiteNavigationApplicationCustomizerStrings';

const LOG_SOURCE: string = 'UncheckSiteNavigationApplicationCustomizer';

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
}

const DEFAULT_LABELS: string[] = [
  'Show in site navigation',
  'Show list in site navigation',
  'Show library in site navigation'
];

/**
 * Application Customizer that watches the page for the modern
 * "Create list" / "Create document library" panels and flips the
 * "Show in site navigation" checkbox off the first time it appears.
 *
 * SharePoint has no server-side setting for this default, so the customizer
 * works on the rendered DOM: it observes mutations, finds a checkbox (or
 * Fluent UI toggle) whose accessible label matches one of the configured
 * strings, and clicks it once if it is checked. Each control element is only
 * touched once, so a user who deliberately re-ticks the box is not overridden.
 */
export default class UncheckSiteNavigationApplicationCustomizer
  extends BaseApplicationCustomizer<IUncheckSiteNavigationApplicationCustomizerProperties> {

  private _observer: MutationObserver | undefined;
  private _processed: WeakSet<Element> = new WeakSet<Element>();
  private _labels: string[] = [];
  private _scanScheduled: boolean = false;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    const configured: string[] | undefined = this.properties && this.properties.labels;
    const source: string[] = Array.isArray(configured) && configured.length > 0 ? configured : DEFAULT_LABELS;
    this._labels = source.map(this._normalise).filter((l: string) => l.length > 0);

    if (typeof MutationObserver === 'undefined' || !document.body) {
      Log.warn(LOG_SOURCE, 'MutationObserver or document.body unavailable; customizer inactive.');
      return Promise.resolve();
    }

    this._observer = new MutationObserver(() => this._scheduleScan());
    this._observer.observe(document.body, { childList: true, subtree: true });

    // Cover anything already on the page when the customizer loads.
    this._scheduleScan();

    return Promise.resolve();
  }

  protected onDispose(): void {
    if (this._observer) {
      this._observer.disconnect();
      this._observer = undefined;
    }
    super.onDispose();
  }

  /** Coalesce bursts of mutations into a single scan per animation frame. */
  private _scheduleScan(): void {
    if (this._scanScheduled) {
      return;
    }
    this._scanScheduled = true;
    const run = (): void => {
      this._scanScheduled = false;
      try {
        this._scan();
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

  private _scan(): void {
    // Only look inside overlays (Fluent UI panels/dialogs render into layers).
    // Falls back to the whole document if no layer host exists.
    const roots: Element[] = Array.from(document.querySelectorAll('.ms-Layer, [role="dialog"]'));
    const scopes: ParentNode[] = roots.length > 0 ? roots : [document];

    for (const scope of scopes) {
      const controls: Element[] = Array.from(
        scope.querySelectorAll('input[type="checkbox"], [role="switch"], [role="checkbox"]')
      );
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
    }
  }

  private _uncheck(control: Element, label: string): void {
    const isChecked: boolean = control instanceof HTMLInputElement
      ? control.checked
      : control.getAttribute('aria-checked') === 'true';

    if (!isChecked) {
      this._log(`"${label}" already unchecked.`);
      return;
    }

    // Use a real click so the React-controlled component updates its state.
    (control as HTMLElement).click();

    const nowChecked: boolean = control instanceof HTMLInputElement
      ? control.checked
      : control.getAttribute('aria-checked') === 'true';
    this._log(`"${label}" ${nowChecked ? 'could not be unchecked' : 'defaulted to unchecked'}.`);
  }

  /** Best-effort accessible name: aria-label, aria-labelledby, <label for>, wrapping <label>. */
  private _getAccessibleName(el: Element): string {
    const ariaLabel: string | null = el.getAttribute('aria-label');
    if (ariaLabel) {
      return ariaLabel;
    }

    const labelledBy: string | null = el.getAttribute('aria-labelledby');
    if (labelledBy) {
      const text: string = labelledBy
        .split(/\s+/)
        .map((id: string) => {
          const node: HTMLElement | null = document.getElementById(id);
          return node ? node.textContent || '' : '';
        })
        .join(' ');
      if (text.trim()) {
        return text;
      }
    }

    if (el.id) {
      const forLabel: Element | null = document.querySelector(`label[for="${CSS.escape(el.id)}"]`);
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
      Log.info(LOG_SOURCE, message);
    }
  }
}
