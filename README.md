<p><img src="docs/assets/app-icon.svg" width="88" height="88" alt="puresheets icon"></p>

# puresheets

**Work with data, formulas, and editable spreadsheets.** An app for [puredesktop](https://puredesktop.ai).

[Get started](#getting-started) · [App guide](docs/app-guide.md) · [Develop](docs/development.md) · [Developer account](https://puredesktop.ai/developers)

## What it does

A spreadsheet editor built on Univer, with cells, formulas, formatting, worksheet tabs, data rules, comments, and charts. Work with editable workbooks and exchange supported spreadsheet formats.

## Requirements

Use a compatible [puredesktop](https://puredesktop.ai) build for desktop integration, storage, and the app drawer. Developer setup is covered in the [development guide](docs/development.md).

Open or create a workbook. Agent assistance and external exchange-rate lookups use their configured host services.

## Getting started

1. Create or open a workbook and enter your own data.
2. Use cells, formulas, formatting, and worksheet tabs to organize the workbook.
3. Save as `.sheets` or portable `.sheets.html`; use the available import/export controls when exchanging files with other spreadsheet tools.

## App layout

| Area | What you use it for |
| --- | --- |
| **Command bar and menus** | Choose editing, formatting, and workbook actions. |
| **Formula bar** | Inspect or edit the value or formula of the selected cell. |
| **Spreadsheet surface** | Select cells, enter data, and work across rows and columns. |
| **Sheet tabs and supporting panels** | Switch worksheets and work with comments, charts, data rules, and agent activity. |

The app also uses the shared [puredesktop](https://puredesktop.ai) shell and drawer agent. Panels can vary with the current view and selection.

## Working with the agent

Open the app’s drawer in [puredesktop](https://puredesktop.ai) and describe what you want to do. For example:

> Read the totals in this range.
>
> Format the header row and add a chart.

The app exposes 15 tools, including `getSheetsContext`, `readRange`, `findInWorkbook`. See [agents.md](agents.md) for workflows and [plugin.json](plugin.json) for the complete tool schemas and approval flags. Some actions apply directly, while approval-marked actions ask first. Check the result in the app after a change.

## Files and data

Save editable `.sheets` workbooks or portable `.sheets.html` files. Spreadsheet import/export uses the supported formats exposed by the app.

## Develop and customize

We welcome **developers and vibecoders alike**. Fork puresheets, add a feature, or use what you learn to build a new app.

| Develop your way | Workflow |
| --- | --- |
| **Claude Code, Codex, or your editor** | Open the app’s source folder, read `README.md`, `plugin.json`, `package.json`, and `agents.md`, then make changes and run the app’s checks. Test inside [puredesktop](https://puredesktop.ai) with matching shared platform packages. |
| **purefactory** | Choose **Start building** for a new app, or select an available app project to extend it. Use **Open folder** for external tools and **Open app** to test. |
| **App drawer** | Request a local app change where app-development integration is available. Make clear whether you want to change the app itself or its current document. |

Use **Share** in purefactory to create a `.pureapp` package, then **Settings → System → Install an app → Choose package…** to load it in current builds. Source availability and integration vary by host build.

Follow the [development guide](docs/development.md) for Claude Code/Codex commands, app-specific setup and checks, and packaging. A standalone browser preview does not provide every desktop service.

## Documentation and limitations

| Guide | What it covers |
| --- | --- |
| [App guide](docs/app-guide.md) | App overview, source layout, and usage. |
| [Development guide](docs/development.md) | External coding tools, purefactory, checks, and installation. |
| [Agent guide](agents.md) | App-specific agent workflows and constraints. |
| [Agent contribution skill](.agents/skills/contribute-puresheets/SKILL.md) | Clone or fork, implement and check changes, open PRs, create issues, and comment. |

Verify formulas, formatting, and charts after exchanging files with other spreadsheet tools; not every feature has an exact equivalent.

## Contributing and marketplace

We welcome **developers and vibecoders alike**. Go to [puredesktop.ai](https://puredesktop.ai) and [create a developer account](https://puredesktop.ai/developers) to join the developer community and submit your app for review.

Bring improvements to this app, develop a fork, or build something entirely new. We welcome **open-source and proprietary projects alike** to the [puredesktop](https://puredesktop.ai) marketplace. Support for **paid apps is coming soon**, so you will be able to charge for your apps if you choose. Forks and redistributed dependencies must follow their applicable licenses.

For developer access, app submissions, or marketplace questions, contact [info@puredesktop.ai](mailto:info@puredesktop.ai).

Anyone may use, study, modify, and share this app under its applicable licenses. We welcome pull requests, bug reports, and documentation improvements. See [CONTRIBUTING.md](CONTRIBUTING.md).

### Contribute with a coding agent

Give your agent the [contribution skill](.agents/skills/contribute-puresheets/SKILL.md) and describe the change, issue, or comment you want it to make. Codex can discover it in `.agents/skills/contribute-puresheets/`; with Claude Code or another tool, ask it to read that `SKILL.md` explicitly. For example:

> Read `.agents/skills/contribute-puresheets/SKILL.md`, implement [describe the change], run the relevant checks, and open a pull request to `puredesktop/puresheets` from my fork.

The skill includes app-specific checks and workflows for PRs, issues, and comments. Anyone with a GitHub account can contribute; merging is reserved for `esetera` and `MRdevTagg`.

## Credits and license

A spreadsheet app based on the open-source [Univer](https://github.com/dream-num/univer) project (Apache-2.0), with [ExcelJS](https://github.com/exceljs/exceljs) for workbook interchange.

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
