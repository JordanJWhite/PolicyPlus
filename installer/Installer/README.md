# Policy Plus MSI Installer

This directory contains a WiX Toolset v3 project for building a standard MSI package for Policy Plus.

## Prerequisites

Before building the installer, you must have the following installed:

1.  **Visual Studio:** (e.g., 2019, 2022) with the "Desktop development with .NET" workload.
2.  **WiX Toolset v3:** You can download it from the [official releases page](https://wixtoolset.org/releases/).
3.  **WiX Toolset v3 Visual Studio Extension:** This is required to load and build the project within Visual Studio.
    *   [For Visual Studio 2022](https://marketplace.visualstudio.com/items?itemName=WixToolset.WixToolsetVisualStudio2022Extension)
    *   [For Visual Studio 2019](https://marketplace.visualstudio.com/items?itemName=WixToolset.WixToolsetVisualStudio2019Extension)

## Building the Installer

The `Installer.wixproj` is configured with a project reference to `PolicyPlus.vbproj`. This means Visual Studio will automatically build the main application first, then package it into the MSI.

1.  Ensure all the prerequisites listed above are installed.
2.  Open the `PolicyPlus.sln` solution file in Visual Studio.
3.  Select your desired build configuration from the dropdown menu (e.g., `Release`).
4.  In the **Solution Explorer**, right-click on the **Installer** project.
5.  Click **Build**.
6.  The compiled `PolicyPlus.msi` package will be available in the `Installer\bin\<Configuration>\` folder (e.g., `Installer\bin\Release\`).

## Installer Features

*   **System Validation:** Verifies that the target machine has .NET Framework 4.5 or a newer version installed.
*   **Shortcuts:** Creates a Start Menu shortcut for Policy Plus.
*   **Standard Installation:** Installs the application for all users (`perMachine`) into the `Program Files` directory.
*   **License Agreement:** Presents the project's license for user acceptance during installation.
*   **Proper Integration:** Adds details and an icon to the 'Apps & features' (Add/Remove Programs) list.
*   **Silent Installation:** Supports all standard `msiexec` command-line options for automated deployment. For example:
    ```shell
    msiexec /i PolicyPlus.msi /quiet
    ```