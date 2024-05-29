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