# Διακίνηση

Personal mobile-friendly web app for viewing and editing the five worksheets in a shared Excel workbook.

## How it works

- Microsoft sign-in uses the `consumers` authority and the delegated `Files.ReadWrite` scope.
- The user enters the OneDrive sharing link on their own device; it is not stored in this repository.
- The app edits only changed worksheet cells in the `.xlsx` package and uploads the resulting workbook to the original shared item.
- Before upload, the app compares the workbook eTag to detect changes made since it was loaded.
- The source workbook itself is never included in the app or repository.

## Microsoft app registration

Configure the app registration for personal Microsoft accounts and add this SPA redirect URI after the site URL is known:

`https://lgeor.github.io/diakinisi-app/auth.html`

The Application (client) ID is public and is configured in `app.js`. Do not add a client secret to this static app.

## Hosting

The site is a static PWA. It must be served over HTTPS for Microsoft sign-in and service worker installation.
