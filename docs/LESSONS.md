# Lessons learned

Things that cost time or surprised us while building and proving this on a
real tenant (g53.sharepoint.com, September 2026). Written so the next person
does not pay for them again.

## SharePoint platform

1. **The modern create panel is an iframe.** *New → List* and *New → Document
   library* load `/_layouts/15/createlist.aspx?dlg=true&mode=sites|sitesDoclib`
   into `<iframe name="createListFrame">`. SPFx extensions do not run inside it.
   It is same-origin, so a top-page customizer can hook `frame.contentDocument`,
   but every navigation inside the frame is a new document: listen for `load`
   and re-hook.

2. **The library panel drops the first click.** Its checkbox is Fluent UI 9,
   rendered before React attaches handlers. A single `click()` looked like it
   worked and did nothing. Verify after clicking and retry with backoff; four
   retries up to 1.5 s were enough in practice.

3. **Feature XML in an app package is a strict subset.** `Location="Microsoft.SharePoint.SiteSettings"`
   is not in the `CustomActionLocations` enumeration the app framework accepts;
   `UrlAction` URLs must start with `http:`, `https:`, `~appWebUrl` or
   `~remoteAppUrl` (`~site` is rejected). The catalog validates the *package*
   (visible as `IsValidAppPackage` / `AppPackageErrorMessage` on the "Apps for
   SharePoint" list item), but the *feature* is only validated at install, so a
   package can look fine in the catalog and then fail on every site.

4. **NoScript blocks more than scripts.** On default modern sites, a tenant admin
   gets 403 `Attempted to perform an unauthorized operation` adding a plain URL
   custom action, and `Access denied` writing the web property bag. SPFx
   (ClientSideComponentId) custom actions and their properties are allowed:
   MERGE on `/_api/web/UserCustomActions('<id>')` works for Manage Web users.

5. **Site Settings is classic.** Application customizers never load on
   `settings.aspx`, so even a working link there could not open an SPFx panel.

6. **Site contents keeps your query string.** `viewlsts.aspx?foo=bar` becomes
   `viewlsts.aspx?foo=bar&view=14`; `foo` survives, which makes it a usable
   deep-link target. Strip only your own key afterwards, not the whole query.

7. **Client-side navigation does not re-run `onInit`.** UI that depends on the
   route must be driven from `application.navigatedEvent` and
   `placeholderProvider.changedEvent`, and must unmount itself when the route
   no longer matches. A proof that only uses `page.goto` will never notice.

8. **`tryCreateContent` can legitimately return undefined** early in the page
   life. Treat it as "not yet", not "never", and wait for `changedEvent`.
   Do not destroy state (like the deep-link query string) before the UI is
   actually up.

9. **App upgrade resets custom action properties** to the package defaults, and
   re-uploading the *same* solution version then running `app upgrade` changes
   nothing on the site. For dev iterations, uninstall and install.

10. **A failed install is silent in the UI.** The app tile is greyed out, the
    instance sits at `AppStatus 6`, `AppMonitoringDetails.aspx` is retired, and
    `app instance list` shows no error. The reason is only available from
    `POST /_api/web/GetAppInstanceById('<instance-id>')/GetErrorDetails`.

11. **`app deploy` right after `app add --overwrite` can throw
    `ResourceNotFoundException`.** The catalog needs a few seconds. Retry.

## SPFx / build

12. **SPFx 1.23 needs Node 22.** Node 24 fails immediately and unhelpfully.
    `scripts/use-node.sh` exists because version managers can report success
    without changing `node` on PATH; it verifies `node --version`.

13. **`npm install --save x --save-dev y` puts everything in devDependencies.**
    The last flag wins for the whole command. Check `package.json` afterwards;
    runtime packages in devDependencies still bundle, so the mistake is
    invisible until someone audits the manifest.

14. **An Application Customizer is a tax on every page view.** Adding React and
    Fluent UI statically took the bundle from ~10 KB to ~235 KB for a feature
    seen by owners on one page. `import()` splits it into a chunk; measure
    `dist/` sizes after touching imports. This came out of code review, not
    testing; the tenant proof cannot see bundle size.

15. **The SPFx ESLint profile fails the production build on warnings** (for
    example `@rushstack/no-new-null`). Suppress with a reason when the API
    genuinely returns `null`.

16. **Fluent UI 8 for our UI, Fluent UI 9 in Microsoft's.** SPFx 1.23 pairs with
    `@fluentui/react` 8 and React 17. The create panel we manipulate uses
    Fluent 9 (`.fui-Checkbox__input`); do not confuse the two when reading DOM.

## Testing

17. **Headed Playwright with a persistent profile is the only credible test.**
    Sign in once; `~/.cache/pw-sp-profile` keeps the session. Behind an
    authenticating proxy, pass `{server, username, password}` from
    `HTTPS_PROXY` in launch options or Chromium fails with
    `ERR_INVALID_AUTH_CREDENTIALS`.

18. **Assert on both sides.** A control site without the app (checkbox must be
    *checked*) turned a "the box is unchecked" observation into a proof that
    *we* unchecked it. The same idea applies to the switch: prove off leaves it
    checked, then prove on unchecks it again.

19. **`getByRole('button', {name:'New'})` did not match SharePoint's New button;**
    `button:has-text("New")` did. Menu items needed a short wait and
    `click({force:true})` because an animation intercepts pointer events.

20. **Expect one spurious 30 s locator timeout in a long run.** It happened once
    in an otherwise passing sequence and the identical rerun passed. Rerun
    before investigating; do not loosen the assertions.

21. **Never edit test files with a naive `String.replace`** that contains `$`
    sequences; it corrupted the proof script once. Use a script that slices, or
    an editor.

## Process

22. **Prove platform assumptions before designing around them.** The original
    plan had a Site Settings link. Three cheap tenant experiments (CLI add,
    package validation, install) killed it in under an hour and produced the
    evidence now in `DECISIONS.md`. Designing the panel first would have wasted
    a day.

23. **Get an independent review before releasing.** The review caught the
    lifecycle gap (UI only mounted from `onInit`), the bundle size, the
    site-scope fallback and the malformed-JSON overwrite, none of which the
    passing tenant proof could reveal.

24. **Keep the release mechanism boring and shared.** `scripts/release`,
    `increment-version.sh` and `build.sh` are copied from sp-calendar-planner.
    Same conventions across JFDI SPFx projects: `X.Y.0.Z` solution versions,
    `vX.Y.0.Z` tags, `.sppkg` on the GitHub Release, never in git, CHANGELOG
    entry required before release.

25. **Sandboxed shells can make config files read-only.** `sed -i` on
    `config/package-solution.json` failed with `EROFS` in the sandbox used for
    development, twice, half-way through the version bump. The release script
    should be run outside such sandboxes, or the bump finished by hand across
    all three files.
