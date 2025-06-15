# Policy Plus MSI Installer

This folder contains a WiX Toolset project that builds an MSI package for Policy Plus.

## Building
1. Install [WiX Toolset 3.x](https://wixtoolset.org/) and ensure `candle` and `light` are in your `PATH`.
2. Build the main application in `Release` configuration so that `PolicyPlus.exe` is available at `PolicyPlus\bin\Release`.
3. Run `msbuild PolicyPlusInstaller.wixproj` from this directory. The resulting `PolicyPlusInstaller.msi` will be placed in the `bin` output folder.

## Installer Features
- Verifies .NET Framework 4.5 or later is installed.
- Optional creation of desktop and Start Menu shortcuts controlled via properties `DESKTOP_SHORTCUT` and `STARTMENU_SHORTCUT`.
- Presents the project license for acceptance.
- Supports standard MSI silent installation options, for example:
  ```
  msiexec /i PolicyPlusInstaller.msi DESKTOP_SHORTCUT=0 /quiet
  ```
