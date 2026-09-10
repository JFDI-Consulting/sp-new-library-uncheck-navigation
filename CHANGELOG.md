# Changelog

All notable changes to sp-new-library-uncheck-navigation are documented in this file.

## [1.0.0.2] - 2026-09-10

First published release.

### Added

- **"Show in site navigation" now defaults to unchecked** in the modern Create list and Create
  document library panels on any site where the app is installed. A user who wants the new list or
  library in the navigation ticks the box as before; the customizer only acts once per control and
  never overrides a deliberate choice
- **Configurable label matching.** The custom action's `labels` property lists the checkbox label
  texts to match, for tenants whose UI language is not English; `debug` logs each match to the
  browser console

### Internal

- The create experience is hosted in a same-origin `createlist.aspx` iframe where SPFx extensions
  do not load, so the customizer observes every same-origin iframe it can reach and re-hooks them
  when they navigate
- The document-library panel renders its checkbox before React attaches its handlers, so a click
  is verified and retried with a short backoff
- Headed Playwright proof (`e2e/prove.js`) against a site with the app and a control site without
  it; screenshots from the passing run on g53.sharepoint.com are in `docs/proof/`
- Developer metadata set to the JFDI Consulting Ltd defaults (MPN ID 2339010)

---

*sp-new-library-uncheck-navigation — a SharePoint Framework application customizer by JFDI Consulting.*
