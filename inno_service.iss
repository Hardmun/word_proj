; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "Word split"
#define MyAppVersion "1.0"
#define MyAppPublisher "Nikulin Vitaly"
#define MyAppExeName "Wordsplit.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{9AD8A1F1-819A-45CC-B586-32E61C4F8B3B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName=C:\Program Files\Word split
DisableProgramGroupPage=yes
; Uncomment the following line to run in non administrative install mode (install for current user only.)
;PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=commandline
OutputBaseFilename=Wordsplit
SetupIconFile=C:\install\wordsplit\word.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "C:\install\wordsplit\dist\Wordsplit.exe"; DestDir: "{app}"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

;[Icons]
;Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"

[INI]
FileName: "{app}\settings.ini"; Section: "DEFAULT"; Key: "AllowSendErrors"; String: "{code:GetCheckError}"
FileName: "{app}\settings.ini"; Section: "DEFAULT"; Key: "path"; String: "{code:GetWordDir}"

[Dirs]
Name: "{app}"; Flags: uninsalwaysuninstall

[run]
Filename: {sys}\sc.exe; Parameters: "create Wordsplit start=auto binPath= ""{app}\Wordsplit.exe"" displayname=""Word split""" ; Flags: runhidden
Filename: {sys}\sc.exe; Parameters: "description Wordsplit ""Word split application""" ; Flags: runhidden
Filename: {sys}\sc.exe; Parameters: "start Wordsplit" ; Flags: runhidden

[UninstallRun]
Filename: {sys}\sc.exe; Parameters: "stop Wordsplit" ; Flags: runhidden
Filename: {sys}\sc.exe; Parameters: "delete Wordsplit" ; Flags: runhidden

[UninstallDelete]
Type: filesandordirs; Name: "{app}\unnecessary_files"

[Code]
var
  OtherInputDirPage: TInputDirWizardPage;
  PageCheckBox: TInputOptionWizardPage;

procedure InitializeWizard;
begin
  OtherInputDirPage :=
    CreateInputDirPage(wpSelectDir, 'Select the WORD directory to analyze', '', '', False, '');
  OtherInputDirPage.Add('');
  OtherInputDirPage.Values[0] := ExpandConstant('{userdesktop}') + '\WordToSplit'

  PageCheckBox :=
    CreateInputOptionPage(wpWelcome,
  'License Information', 'Are you a registered user?',
  'If you are a registered user, please check the box below, then click Next.',
  False, False);
  PageCheckBox.Add('Allow the programm to send errors to developers.');
  PageCheckBox.Values[0] := True;
end;

function GetWordDir(Param: String): String;
begin
  Result := OtherInputDirPage.Values[0];
end;

function GetCheckError(Param: String): String;
begin
  if PageCheckBox.Values[0] then begin
    Result := '1';
  end
  else begin
    Result := '0';
  end; 
  
end;

/////////////////////////////////////////////////////////////////////
function GetUninstallString(): String;
var
  sUnInstPath: String;
  sUnInstallString: String;
begin
  sUnInstPath := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{#emit SetupSetting("AppId")}_is1');
  sUnInstallString := '';
  if not RegQueryStringValue(HKLM, sUnInstPath, 'UninstallString', sUnInstallString) then
    RegQueryStringValue(HKCU, sUnInstPath, 'UninstallString', sUnInstallString);
  Result := sUnInstallString;
end;

/////////////////////////////////////////////////////////////////////
function IsUpgrade(): Boolean;
begin
  Result := (GetUninstallString() <> '');
end;

/////////////////////////////////////////////////////////////////////
function UnInstallOldVersion(): Integer;
var
  sUnInstallString: String;
  iResultCode: Integer;
begin
// Return Values:
// 1 - uninstall string is empty
// 2 - error executing the UnInstallString
// 3 - successfully executed the UnInstallString

  // default return value
  Result := 0;

  // get the uninstall string of the old app
  sUnInstallString := GetUninstallString();
  if sUnInstallString <> '' then begin
    sUnInstallString := RemoveQuotes(sUnInstallString);
    if Exec(sUnInstallString, '/SILENT /NORESTART /SUPPRESSMSGBOXES','', SW_HIDE, ewWaitUntilTerminated, iResultCode) then
      Result := 3
    else
      Result := 2;
  end else
    Result := 1;
end;

/////////////////////////////////////////////////////////////////////
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if (CurStep=ssInstall) then
  begin
    if (IsUpgrade()) then
    begin
      UnInstallOldVersion();
    end;
  end;
end;