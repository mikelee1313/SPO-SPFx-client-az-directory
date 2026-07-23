# Client A-Z Directory

SPFx 1.20 React web part for hosting a searchable A-Z client directory in SharePoint Online.


<img width="1184" height="582" alt="image" src="https://github.com/user-attachments/assets/ab41ab97-e0c7-42a5-82c4-5b640de948ff" />

## Features

- Loads client records from a configurable SharePoint list.
- Searches client names and optional secondary text.
- Shows A-Z letter tiles with live counts for each starting letter.
- Expands individual letters to browse matching clients.
- Supports optional secondary text and optional client links.
- Paginates SharePoint REST results so lists with thousands of clients can be loaded.

## SharePoint list setup

Create or use a SharePoint list with at least one text column for the client name. The default configuration expects:

| Setting | Default | Notes |
| --- | --- | --- |
| Client list name | `Clients` | SharePoint list display name |
| Client name column internal name | `Title` | Text column used for search, sorting, A-Z grouping, and display |
| Secondary text column internal name | blank | Optional text shown below each client, such as region or account owner |
| Client link column internal name | blank | Optional Hyperlink column or text URL used to make each client clickable |

Use SharePoint column **internal names** in the web part property pane.

## Build and package

```powershell
npm install
npx gulp build --no-color
npx gulp bundle --ship --no-color
npx gulp package-solution --ship --no-color
```

The package is generated at:

```text
sharepoint\solution\client-az-directory.sppkg
```

## Deploy

1. Upload `sharepoint\solution\client-az-directory.sppkg` to the SharePoint tenant App Catalog.
2. Deploy the app when prompted.
3. Add **Client A-Z Directory** to a SharePoint page.
4. Edit the web part properties and confirm the list name and column internal names.

## Used SharePoint Framework Version

SPFx 1.20.0
