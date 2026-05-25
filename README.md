# NetDTL Installer

Windows installer for deploying **NetDTL**.

NetDTL Installer automates the installation, deployment and initial configuration of the NetDTL environment on Windows systems.

version 1.0-0 25-may-2026 DTL didier.morandi@gmail.com

---

## Features

- Automated dependency installation
- NetDTL deployment
- Initial application configuration
- Patch panel tools
- Desktop shortcut creation
- Windows-friendly installer (Inno Setup)

---

## Project structure

```text
netdtl-installer/
│
├── NetDTL_Setup.iss
├── assets/
│   └── netdtl.ico
├── scripts/
│   ├── install_dependencies.ps1
│   ├── deploy_netdtl.ps1
│   └── configure_netdtl.ps1
├── tools/
│   ├── patch_panel_engine.py
│   ├── patch_panel_launcher.py
│   └── patch_panel_config.json.example
└── README.md
