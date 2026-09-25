<p><img src="docs/assets/app-icon.svg" width="88" height="88" alt="puresheets icon"></p>

# puresheets

## What puresheets does

A spreadsheet editor built on Univer, with cells, formulas, formatting, worksheet tabs, data rules, comments, and charts. Work with editable workbooks and exchange supported spreadsheet formats.

## App layout

| Area | What you use it for |
| --- | --- |
| **Command bar and menus** | Choose editing, formatting, and workbook actions. |
| **Formula bar** | Inspect or edit the value or formula of the selected cell. |
| **Spreadsheet surface** | Select cells, enter data, and work across rows and columns. |
| **Sheet tabs and supporting panels** | Switch worksheets and work with comments, charts, data rules, and agent activity. |

The app also uses the shared [puredesktop](https://puredesktop.ai) shell and drawer agent. Panels can vary with the current view and selection.

## Getting started

1. Create or open a workbook and enter your own data.
2. Use cells, formulas, formatting, and worksheet tabs to organize the workbook.
3. Save as `.sheets` or portable `.sheets.html`; use the available import/export controls when exchanging files with other spreadsheet tools.

Read the [app guide](docs/app-guide.md) for development, loading, and source-layout details.

## Develop and customize

You can develop this app outside [puredesktop](https://puredesktop.ai), using your preferred editor, terminal, and coding tools, then load the module into [puredesktop](https://puredesktop.ai) to use and test it. You can also change your local version from **purefactory** or through **the app’s drawer agent**.

### Use your own development tools

1. Fork or clone this repository and work on a local copy in your editor.
2. Set up the app’s dependencies and run its development server or build. See the [app guide](docs/app-guide.md#development-and-loading) for this repository’s requirements and scripts.
3. Load the module into [puredesktop](https://puredesktop.ai). For a local web development server, the platform guide describes **File → Register App…**: register its URL, app name, and required permissions, then open it from **Browse Apps**. Keep the development server running while using that entry point.
4. Make changes in your editor, reload the app as needed, and test its file, account, and agent integrations inside the desktop. A distributable `.pureapp` package can be loaded through **File → Install App…**.

See the [app development and integration guide](https://puredesktop.ai/docs/apps/) for registration, the app manifest, the bridge, and packaging. Editing outside the desktop does not remove this module’s shared-dependency requirements.

### Use purefactory or the app’s drawer agent

Open your local app project in **purefactory** to develop it there, or open the app’s **drawer agent** and describe the change you want to make to your local version. Specify whether you want to change the app itself or work on the document or data currently open. Review the resulting source changes, run the relevant checks, and reload your local app to try them. You can keep the changes for yourself, develop a fork, or contribute them back with a pull request.

## Developer accounts and the marketplace

[Create a developer account on puredesktop.ai](https://puredesktop.ai/developers) to take part in the developer community and submit apps for review. We welcome contributions to this app, forks that take it in a different direction, and entirely new apps to offer on [puredesktop](https://puredesktop.ai).

We welcome **open-source and proprietary projects alike** to the [puredesktop](https://puredesktop.ai) marketplace. A marketplace with support for **paid apps is coming soon**, so developers will be able to charge for their apps if they choose. When distributing a fork, follow the licenses of the code and dependencies you use.

For more information about developer accounts, app submissions, or the upcoming marketplace, contact [info@puredesktop.ai](mailto:info@puredesktop.ai).

## Open source and contributions

A spreadsheet app based on the open-source [Univer](https://github.com/dream-num/univer) project (Apache-2.0), with [ExcelJS](https://github.com/exceljs/exceljs) for workbook interchange.

Anyone may use, study, modify, and share this software under the applicable licenses.
We welcome pull requests, bug reports, documentation improvements, and new ideas.
See [CONTRIBUTING.md](CONTRIBUTING.md) for how to contribute.

### License

Original code by pure.science inc is licensed under the [MIT License](LICENSE).
Copyright (c) 2026 pure.science inc. Third-party code, dependencies, and assets retain their own licenses and copyright notices.

### Major open-source projects

| Project / source | Homepage or documentation | Support the maintainers |
| --- | --- | --- |
| [dream-num/univer](https://github.com/dream-num/univer) | [Homepage / docs](https://docs.univer.ai) | [Open Collective](https://opencollective.com/univer) |
| [exceljs/exceljs](https://github.com/exceljs/exceljs) | [Project home](https://github.com/exceljs/exceljs) | — |
| [react/react](https://github.com/react/react) | [Homepage / docs](https://react.dev) | — |
| [styled-components/styled-components](https://github.com/styled-components/styled-components) | [Homepage / docs](https://styled-components.com) | [GitHub Sponsors](https://github.com/sponsors/quantizor) · [Open Collective](https://opencollective.com/styled-components) |

Thank you to these projects and their contributors. Additional direct dependencies,
upstream links, and asset notices are listed in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
