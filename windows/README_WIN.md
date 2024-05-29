### Installation and Configuration Guide for ConvertyPDF on Windows

This guide provides detailed instructions for installing and configuring ConvertyPDF on Windows, including generating a standalone executable, adding an option to the context menu, and creating an installer.

## 1. Installing PyInstaller and Generating the Executable

### Installing PyInstaller

1. **Open Command Prompt as Administrator**:

   - Search for "cmd" in the Start menu.
   - Right-click on "Command Prompt" and select "Run as administrator".

2. **Install PyInstaller**:

   ```sh
   pip install pyinstaller
   ```

### Generating the Executable

1. **Navigate to the Project Directory**:

   ```sh
   cd path\to\your\project
   ```

2. **Ensure All Dependencies are Installed**:

   ```sh
   pip install python-docx lxml pillow
   ```

3. **Use PyInstaller to Create the Executable**:

   ```sh
   pyinstaller --onefile --windowed --hidden-import=docx --hidden-import=lxml --hidden-import=PIL app_gui.py
   ```

   This command will generate a standalone executable in the `dist` directory.

## 2. Adding the Executable to the Windows Context Menu

### Creating a Registry File

1. **Create a Text File** and rename it with a `.reg` extension, for example, `add_to_context_menu.reg`.
2. **Open the File with a Text Editor** and add the following lines:

   ```reg
   Windows Registry Editor Version 5.00

   [HKEY_CLASSES_ROOT\*\shell\ConvertToPDF]
   @="Convert to PDF"

   [HKEY_CLASSES_ROOT\*\shell\ConvertToPDF\command]
   @="\"C:\\path\\to\\your\\app_gui.exe\" \"%1\""
   ```

   Make sure to replace `C:\\path\\to\\your\\app_gui.exe` with the actual path to the generated executable.

### Adding the Entry to the Windows Registry

1. **Double-click the `.reg` File** you created and confirm to add the entries to the Windows Registry.

## 3. Creating an Installer (Optional)

### Downloading and Installing Inno Setup

1. **Download Inno Setup** from the official website: [Inno Setup](https://jrsoftware.org/isinfo.php).
2. **Install Inno Setup** following the instructions on the website.

### Creating the Setup Script

1. **Create a Text File** and save it with a `.iss` extension, for example, `setup.iss`.
2. **Open the File with a Text Editor** and add the following lines:

   ```iss
   [Setup]
   AppName=ConvertifyPDF
   AppVersion=1.0
   DefaultDirName={pf}\ConvertifyPDF
   DefaultGroupName=ConvertifyPDF
   OutputBaseFilename=ConvertifyPDF_Installer
   Compression=lzma
   SolidCompression=yes

   [Files]
   Source: "dist\app_gui.exe"; DestDir: "{app}"; Flags: ignoreversion
   Source: "dist\app.exe"; DestDir: "{app}"; Flags: ignoreversion

   [Icons]
   Name: "{group}\ConvertifyPDF GUI"; Filename: "{app}\app_gui.exe"
   Name: "{group}\ConvertifyPDF CLI"; Filename: "{app}\app.exe"

   [Run]
   Filename: "{app}\app_gui.exe"; Description: "{cm:LaunchProgram,ConvertifyPDF}"; Flags: nowait postinstall skipifsilent
   ```

### Compiling the Setup Script

1. **Open Inno Setup** and load the `setup.iss` file.
2. **Compile the Script** following the instructions in the Inno Setup interface. This will generate an `.exe` file that can be used to install your application on other Windows computers.

## Conclusion

By following these steps, you will be able to install and configure ConvertyPDF on Windows, add the option to the context menu, and create an installer for easy distribution.
