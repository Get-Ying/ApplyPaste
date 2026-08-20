# ApplyPaste Project Structure

## Root Files

`manifest.json`, `popup.html`, and `sidepanel.html` are extension entry points. `README.md` is the project and user entry page. `package.json` defines development commands.

## src/

Browser extension runtime source and styles.

## icons/

Extension icons referenced by `manifest.json`.

## data/

Empty profile data bundled for first startup. It contains no user information.

## public-template/

Empty public data shape used by release and privacy checks.

## docs/

User, development, audit, and test documentation.

## docs/images/

Official ApplyPaste usage-guide images.

## tests/

Automated behavior, integration, release, and regression tests.

## tests/fixtures/

Public, synthetic test pages and files.

## tools/

Build, static checks, privacy scanning, and release utilities.

## release/

Current shareable Beta build and ZIP archive.

## archive/

Historical documentation, legacy tests, and legacy empty layout directories. Archive content does not participate in runtime or release builds.
