; This needs to be compiled with the UNICODE version of the Inno Setup Compiler
; (version 5.5.9 or later)

[Setup]
AppName=ZoneEdit Dynamic DNS Update Client
AppVersion=2.8.0
LicenseFile=LICENSE.TXT
OutputBaseFilename=zeDyn-2.8-setup
DefaultDirName={pf32}\ZoneEdit Dynamic DNS Update Client
DefaultGroupName=ZoneEdit Dynamic DNS Update Client
OutputDir=.
DisableStartupPrompt=False
AlwaysShowGroupOnReadyPage=True
AlwaysShowDirOnReadyPage=True
UninstallDisplayIcon={app}\zeDyn.exe
PrivilegesRequired=none
AlwaysRestart=yes

[Files]
Source: "SysTrayDLL\SysTrayDll.dll"; DestDir: "{syswow64}"; Flags: ignoreversion regserver uninsneveruninstall overwritereadonly
Source: "English.lng"; DestDir: "{app}"; Flags: ignoreversion
Source: "zeDyn.exe"; DestDir: "{app}"; Flags: ignoreversion overwritereadonly; Permissions: everyone-full
Source: "عربي.lng"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\ZoneEdit Dynamic DNS Update Client\zeDyn.exe"; Filename: "{app}\zeDyn.exe"; WorkingDir: "{app}"
Name: "{commondesktop}\zeDyn.exe"; Filename: "{app}\zeDyn.exe"; WorkingDir: "{app}"
Name: "{commonstartup}\zeDyn.exe"; Filename: "{app}\zeDyn.exe"; WorkingDir: "{app}"
Name: "{group}\Uninstall"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\zeDyn.exe"; Description: "Launch application"; Flags: postinstall nowait skipifsilent