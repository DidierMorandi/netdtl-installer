; ============================================================
; NetDTL stand alone Installer
; v1.0-0 22-may-2026 DTL initial release
; v1.0-1 23-may-2026 DTL Flags: waituntilterminated runhidden 
; ============================================================

[Setup]
AppName=NetDTL
AppVersion=1.0.1
AppPublisher=Didier Morandi
AppPublisherURL=https://github.com/DidierMorandi/netdtl
DefaultDirName={autopf}\NetDTL
DefaultGroupName=NetDTL
OutputDir=Output
OutputBaseFilename=NetDTL_Installer_v1.0.1
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
DisableProgramGroupPage=yes
SetupLogging=yes
UninstallDisplayIcon={app}\assets\netdtl.ico

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: desktopicon; Description: "Créer un raccourci sur le bureau"; GroupDescription: "Raccourcis :"
Name: patchpanel; Description: "Installer le générateur de tableau de brassage"; Flags: checkedonce

[Files]
Source: "scripts\install_dependencies.ps1"; DestDir: "{app}\scripts"
Source: "scripts\deploy_netdtl.ps1"; DestDir: "{app}\scripts"
Source: "scripts\configure_netdtl.ps1"; DestDir: "{app}\scripts"

Source: "tools\patch_panel_engine.py"; DestDir: "{app}\tools"; Tasks: patchpanel
Source: "tools\patch_panel_launcher.py"; DestDir: "{app}\tools"; Tasks: patchpanel
Source: "tools\patch_panel_config.json.example"; DestDir: "{app}\tools"; Tasks: patchpanel

Source: "assets\*"; DestDir: "{app}\assets"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{group}\NetDTL"; Filename: "http://localhost/netdtl"

Name: "{commondesktop}\NetDTL"; \
Filename: "http://localhost/netdtl"; \
Tasks: desktopicon

Name: "{group}\Générateur tableau de brassage"; \
Filename: "{sys}\cmd.exe"; \
Parameters: "/c start """" python ""{app}\tools\patch_panel_launcher.py"""; \
Tasks: patchpanel

[Run]
Filename: "powershell.exe"; \
Parameters: "-ExecutionPolicy Bypass -File ""{app}\scripts\install_dependencies.ps1"""; \
Flags: waituntilterminated runhidden

Filename: "powershell.exe"; \
Parameters: "-ExecutionPolicy Bypass -File ""{app}\scripts\deploy_netdtl.ps1"""; \
Flags: waituntilterminated runhidden

Filename: "powershell.exe"; \
Parameters: "-ExecutionPolicy Bypass -File ""{app}\scripts\configure_netdtl.ps1"" -AdminUser ""{code:GetAdminUser}"" -AdminPass ""{code:GetAdminPass}"" -CIDR ""{code:GetCIDR}"""; \
Flags: waituntilterminated runhidden

Filename: "http://localhost/netdtl"; \
Description: "Lancer NetDTL"; \
Flags: shellexec postinstall nowait skipifsilent

[Code]
var
  ConfigPage: TInputQueryWizardPage;

function GetAdminUser(Param: String): String;
begin
  Result := ConfigPage.Values[0];
end;

function GetAdminPass(Param: String): String;
begin
  Result := ConfigPage.Values[1];
end;

function GetCIDR(Param: String): String;
begin
  Result := ConfigPage.Values[2];
end;

procedure InitializeWizard;
begin
  ConfigPage := CreateInputQueryPage(
    wpSelectTasks,
    'Configuration NetDTL',
    'Paramètres initiaux',
    'Saisissez les paramètres nécessaires à l''installation.'
  );

  ConfigPage.Add('Utilisateur administrateur NetDTL :', False);
  ConfigPage.Values[0] := 'admin';

  ConfigPage.Add('Mot de passe administrateur NetDTL :', True);
  ConfigPage.Values[1] := '';

  ConfigPage.Add('Réseau CIDR par défaut :', False);
  ConfigPage.Values[2] := '192.168.1.0/24';
end;