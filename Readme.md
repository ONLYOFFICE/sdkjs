[![License](https://img.shields.io/badge/License-GNU%20AGPL%20V3-green.svg?style=flat)](https://www.gnu.org/licenses/agpl-3.0.en.html)

# 🧩 ONLYOFFICE JavaScript SDK (SDKJS)

SDKJS is the official JavaScript Software Development Kit (SDK) for ONLYOFFICE’s document editing components.

It is integrated into:

- [ONLYOFFICE Docs (Document Server)][2]
- [ONLYOFFICE Desktop Editors][4]

JavaScript SDK provides the necessary **client-side APIs** for integrating and customizing the ONLYOFFICE editors. It also includes an **implementation layer** for the [Office JavaScript APIs][5], enabling advanced document manipulation and integration operations.

## 🌐 Project Resources

- **Official website:** [ONLYOFFICE Homepage](https://www.onlyoffice.com?utm_source=github&utm_medium=cpc&utm_campaign=GitHubSdkjs)
- **Source repository:** [SDKJS on GitHub](https://github.com/ONLYOFFICE/sdkjs)
- **ONLYOFFICE Docs:** [Suite overview](https://www.onlyoffice.com/docs?utm_source=github&utm_medium=cpc&utm_campaign=GitHubSdkjs)

📖 **Developer documentation:** [ONLYOFFICE API Documentation](https://api.onlyoffice.com?utm_source=github&utm_medium=cpc&utm_campaign=GitHubSdkjs) — The essential
reference guide for working with ONLYOFFICE APIs and integration modules.

## 📁 Repository Structure Overview

The directory layout below helps developers quickly navigate and understand SDKJS directory purposes.

| Folder    | Description                                                                |
| :-------- | :------------------------------------------------------------------------- |
| `.github` | Contains GitHub workflows and issue/pr templates for CI/CD automation.     |
| `build`   | Scripts and configuration files used to build SDKJS bundles.               |
| `cell`    | Core functionality and UI logic for spreadsheet editor.                    |
| `common`  | Shared modules, utilities, and core logic used across all editor types.    |
| `configs` | Configuration files and constants used for environment and runtime setup.  |
| `pdf`     | Modules and UI components for viewing and annotating PDF files.            |
| `slide`   | Logic and rendering components for presentation editor.                    |
| `tests`   | Automated test suites and configs for validating SDKJS behavior.           |
| `tools`   | Helper utilities, build scripts, and developer tools.                      |
| `vendor`  | Third-party libraries and external dependencies used by SDKJS.             |
| `visio`   | Modules related to drawing and diagram editing (Visio-like functionality). |
| `word`    | Core logic and UI components for text document editor.                     |

## 💬 User Feedback and Support

We welcome community participation, technical insights, and feedback. For questions, integration issues, or troubleshooting related to [ONLYOFFICE Document Server][2], please explore these resources:

- **Report issues:** [GitHub Issues](https://github.com/ONLYOFFICE/DocumentServer/issues)
- **Forum:** [ONLYOFFICE Community][1]
- **Feedback platform:** [feedback.onlyoffice.com](https://feedback.onlyoffice.com/forums/966080-your-voice-matters)
- **Developer Q&A:** [Stack Overflow][3]

[1]: https://community.onlyoffice.com/
[2]: https://github.com/ONLYOFFICE/DocumentServer
[3]: https://stackoverflow.com/questions/tagged/onlyoffice
[4]: https://github.com/ONLYOFFICE/DesktopEditors
[5]: https://github.com/ONLYOFFICE/office-js-api

## 📜 License

**SDKJS** is licensed under the **GNU Affero General Public License (AGPL) v3.0**. For full details, refer to the [LICENSE](https://github.com/ONLYOFFICE/sdkjs/blob/master/LICENSE.txt) file.