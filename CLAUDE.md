# Project notes

## Versioning
Bump the version in `package.json` (shown in the page title and top bar) every time a change is pushed:
- small fix or tweak → patch (`npm version patch --no-git-tag-version`, e.g. 1.1.0 → 1.1.1)
- new feature or visible behaviour change → minor (`npm version minor --no-git-tag-version`)
- breaking change to saved settings or the run-sheet format → major

Commit the bump (`package.json` and `package-lock.json`) with the change.

## Merge and deploy every time
The owner wants every change merged and deployed, without being asked:
1. Push the working branch, open a pull request into `main`, and merge it once CI passes.
2. Merging deploys GitHub Pages (`pages.yml`) and publishes the offline file as release `vX.Y.Z` (`release.yml` creates the tag from `package.json`; no tag push needed). The owner uses that file at the hospital, where Pages is blocked.
3. Check the Pages and release runs succeed, and report the release link.

Each pushed change gets its own version, so every merge makes a new release. A version that already has a release is skipped, so forgetting the bump means no new download.

## Offline file
Name the single-file build with its version when sharing it, e.g. `theatre-list-prepper-v1.1.1.html` (copy `dist/index.html`; Pages needs it to stay `index.html`). The release workflow does the same.

## Privacy
Never commit real run sheets, patient details, surgeon names or email addresses. Tests use invented data; `.gitignore` blocks CSV/Excel files and `*settings*.json`.

- **No real names anywhere**: code, tests, README, comments, commit messages, screenshots. Examples use invented names (e.g. "Smith"). Never copy a name from an uploaded sample into the repo.
- Before every push, scan the whole working tree and the new commit(s) for names taken from the user's uploaded files. Keep that name list outside the repo (session scratchpad only). Don't push if anything matches.
- CI also fails on spreadsheet/CSV files, real health-service email domains, and email addresses outside the placeholder domains (see `ci.yml`).
