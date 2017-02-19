unit U_Npp_PreviewHTML;

////////////////////////////////////////////////////////////////////////////////////////////////////
interface

uses
  SysUtils, Windows, IniFiles,
  NppPlugin, SciSupport,
  F_About, F_PreviewHTML;

type
  TNppPluginPreviewHTML = class(TNppPlugin)
  private
    FSettings: TIniFile;
    FUpdated: Boolean;
  public
    constructor Create;

    procedure SetInfo(NppData: TNppData); override;

    procedure CommandShowPreview;
    procedure CommandSetIEVersion(const BrowserEmulation: Integer);
    procedure CommandOpenFile(const Filename: nppString);
    procedure CommandCheckUpdates;
    procedure CommandShowAbout;

    procedure DoNppnToolbarModification; override;
    procedure DoNppnFileClosed(const BufferID: THandle); override;
    procedure DoNppnBufferActivated(const BufferID: THandle); override;
    procedure DoModified(const hwnd: HWND; const modificationType: Integer); override;

    function  GetSettings(const Name: string = 'Settings.ini'): TIniFile;
  end {TNppPluginPreviewHTML};

procedure _FuncCheckUpdate; cdecl;
procedure _FuncShowPreview; cdecl;
procedure _FuncOpenSettings; cdecl;
procedure _FuncOpenFilters; cdecl;
procedure _FuncShowAbout; cdecl;

procedure _FuncSetIE7; cdecl;
procedure _FuncSetIE8; cdecl;
procedure _FuncSetIE9; cdecl;
procedure _FuncSetIE10; cdecl;
procedure _FuncSetIE11; cdecl;
procedure _FuncSetIE12; cdecl;
procedure _FuncSetIE13; cdecl;


var
  Npp: TNppPluginPreviewHTML;

////////////////////////////////////////////////////////////////////////////////////////////////////
implementation
uses
  WebBrowser, Registry,
  U_AutoUpdate;

{ ------------------------------------------------------------------------------------------------ }
procedure _FuncCheckUpdate; cdecl;
begin
  Npp.CommandCheckUpdates;
end;
{ ------------------------------------------------------------------------------------------------ }
procedure _FuncOpenSettings; cdecl;
begin
  Npp.CommandOpenFile('Settings.ini');
end;
{ ------------------------------------------------------------------------------------------------ }
procedure _FuncOpenFilters; cdecl;
begin
  Npp.CommandOpenFile('Filters.ini');
end;
{ ------------------------------------------------------------------------------------------------ }
procedure _FuncShowAbout; cdecl;
begin
  Npp.CommandShowAbout;
end;
{ ------------------------------------------------------------------------------------------------ }
procedure _FuncShowPreview; cdecl;
begin
  Npp.CommandShowPreview;
end;
{ ------------------------------------------------------------------------------------------------ }
procedure _FuncSetIE7; cdecl;
begin
  Npp.CommandSetIEVersion(7000);
end;
{ ------------------------------------------------------------------------------------------------ }
procedure _FuncSetIE8; cdecl;
begin
  Npp.CommandSetIEVersion(8000);
end;
{ ------------------------------------------------------------------------------------------------ }
procedure _FuncSetIE9; cdecl;
begin
  Npp.CommandSetIEVersion(9000);
end;
{ ------------------------------------------------------------------------------------------------ }
procedure _FuncSetIE10; cdecl;
begin
  Npp.CommandSetIEVersion(10000);
end;
{ ------------------------------------------------------------------------------------------------ }
procedure _FuncSetIE11; cdecl;
begin
  Npp.CommandSetIEVersion(11000);
end;
{ ------------------------------------------------------------------------------------------------ }
procedure _FuncSetIE12; cdecl;
begin
  Npp.CommandSetIEVersion(12000);
end;
{ ------------------------------------------------------------------------------------------------ }
procedure _FuncSetIE13; cdecl;
begin
  Npp.CommandSetIEVersion(13000);
end;


{ ================================================================================================ }
{ TNppPluginPreviewHTML }

{ ------------------------------------------------------------------------------------------------ }
constructor TNppPluginPreviewHTML.Create;
begin
  inherited;
  self.PluginName := '&Preview HTML';
end {TNppPluginPreviewHTML.Create};

{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.SetInfo(NppData: TNppData);
var
  IEVersion: string;
  MajorIEVersion, Code, EmulatedVersion: Integer;
  sk: TShortcutKey;
begin
  inherited;

  sk.IsShift := True; sk.IsCtrl := true; sk.IsAlt := False;
  sk.Key := 'H'; // Ctrl-Shift-H
  self.AddFuncItem('&Preview HTML', _FuncShowPreview, sk);

  IEVersion := GetIEVersion;
  with GetSettings do begin
    IEVersion := ReadString('Emulation', 'Installed IE version', IEVersion);
    Free;
  end;
  Val(IEVersion, MajorIEVersion, Code);
  if Code <= 1 then
    MajorIEVersion := 0;

  EmulatedVersion := GetBrowserEmulation div 1000;

  if MajorIEVersion > 7 then begin
    self.AddFuncSeparator;
    self.AddFuncItem('View as IE&7', _FuncSetIE7, EmulatedVersion = 7);
  end;

  if MajorIEVersion >= 8 then
    self.AddFuncItem('View as IE&8', _FuncSetIE8, EmulatedVersion = 8);
  if MajorIEVersion >= 9 then
    self.AddFuncItem('View as IE&9', _FuncSetIE9, EmulatedVersion = 9);
  if MajorIEVersion >= 10 then
    self.AddFuncItem('View as IE1&0', _FuncSetIE10, EmulatedVersion = 10);
  if MajorIEVersion >= 11 then
    self.AddFuncItem('View as IE1&1', _FuncSetIE11, EmulatedVersion = 11);
  if MajorIEVersion >= 12 then
    self.AddFuncItem('View as IE1&2', _FuncSetIE12, EmulatedVersion = 12);
  if MajorIEVersion >= 13 then
    self.AddFuncItem('View as IE1&3', _FuncSetIE13, EmulatedVersion = 13);

  self.AddFuncSeparator;

  self.AddFuncItem('Edit &settings', _FuncOpenSettings);
  self.AddFuncItem('Edit &filter definitions', _FuncOpenFilters);

  self.AddFuncSeparator;

  self.AddFuncItem('Check for &updates', _FuncCheckUpdate);

  self.AddFuncSeparator;

  self.AddFuncItem('&About', _FuncShowAbout);
end {TNppPluginPreviewHTML.SetInfo};

{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.CommandCheckUpdates;
const
  scYesNo: array[Boolean] of string = ('No', 'Yes');
var
  Current, Latest, Notes: string;
  Update: TPluginUpdate;
  ZipPath: string;
begin
  try
    if FUpdated then begin
      MessageBox(Npp.NppData.NppHandle, 'Please restart Notepad++ first!', PChar(Caption), MB_ICONWARNING);
      Exit;
    end;

    Update := TPluginUpdate.Create;
    try
      Current := Update.CurrentVersion;
      if not Update.IsUpdateAvailable(Latest, Notes) then begin
        MessageBox(Npp.NppData.NppHandle,
                   PChar(Format('Nothing to update!'#10#10'Your current version of this plugin, v%s, is the most recent one.',
                                [Current])), PChar(Caption),
                   MB_ICONINFORMATION);
        Exit;
      end;

      if ID_YES = MessageBox(Npp.NppData.NppHandle,
                             PChar(Format('Update available!'#10#10
                                        + 'Your current version of this plugin, v%0:s, is out of date.'
                                        + ' We recommend you update to version %1:s.'#10#10
                                        + 'These are the changes since the current version:'#10#10
                                        + '%2:s'#10#10
                                        + 'Do you want to download and install %1:s?',
                                          [Current, Latest, Notes])),
                             PChar(Caption), MB_YESNO or MB_DEFBUTTON1 or MB_ICONQUESTION) then begin
        ZipPath := Update.DownloadUpdate;
        FUpdated := Update.ReplacePlugin(ZipPath);
        if FUpdated then
          MessageBox(Npp.NppData.NppHandle,
                     'Update completed. Please restart Notepad++ to use the updated plugin.',
                     PChar(Caption), MB_ICONINFORMATION);
      end;
    finally
      Update.Free;
    end;
  except
    ShowException(ExceptObject, ExceptAddr);
  end;
end {TNppPluginPreviewHTML.CommandCheckUpdate};

{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.CommandOpenFile(const Filename: nppString);
var
  FullPath: nppString;
begin
  try
    FullPath := Npp.ConfigDir + '\PreviewHTML\' + Filename;
    if not FileExists(FullPath) and FileExists(ChangeFileExt(FullPath, '.sample' + ExtractFileExt(FullPath))) then
      FullPath := ChangeFileExt(FullPath, '.sample' + ExtractFileExt(FullPath));
    if not DoOpen(FullPath) then
      MessageBox(Npp.NppData.NppHandle, PChar(Format('Unable to open "%s".', [FullPath])), PChar(Caption), MB_ICONWARNING);
  except
    ShowException(ExceptObject, ExceptAddr);
  end;
end {TNppPluginPreviewHTML.CommandOpenFilters};

{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.CommandSetIEVersion(const BrowserEmulation: Integer);
begin
  if GetBrowserEmulation <> BrowserEmulation then begin
    SetBrowserEmulation(BrowserEmulation);
    MessageBox(Npp.NppData.NppHandle,
                PChar(Format('The preview browser mode has been set to correspond to Internet Explorer version %d.'#13#10#13#10 +
                             'Please restart Notepad++ for the new browser mode to be taken into account.',
                             [BrowserEmulation div 1000])),
                PChar(Self.Caption), MB_ICONWARNING);
  end else begin
    MessageBox(Npp.NppData.NppHandle,
                PChar(Format('The preview browser mode was already set to Internet Explorer version %d.',
                             [BrowserEmulation div 1000])),
                PChar(Self.Caption), MB_ICONINFORMATION);
  end;
end {TNppPluginPreviewHTML.CommandSetIEVersion};

{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.CommandShowAbout;
begin
  with TAboutForm.Create(self) do begin
    ShowModal;
    Free;
  end;
end {TNppPluginPreviewHTML.CommandShowAbout};

{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.CommandShowPreview;
const
  ncDlgId = 0;
begin
  if (not Assigned(frmHTMLPreview)) then begin
    frmHTMLPreview := TfrmHTMLPreview.Create(self, ncDlgId);
    frmHTMLPreview.Show;
  end else begin
    if not frmHTMLPreview.Visible then
      frmHTMLPreview.Show
    else
      frmHTMLPreview.Hide
    ;
  end;
  if frmHTMLPreview.Visible then begin
    frmHTMLPreview.btnRefresh.Click;
  end;
end {TNppPluginPreviewHTML.CommandShowPreview};

{ ------------------------------------------------------------------------------------------------ }
function TNppPluginPreviewHTML.GetSettings(const Name: string): TIniFile;
begin
  ForceDirectories(ConfigDir + '\PreviewHTML');
  Result := TIniFile.Create(ConfigDir + '\PreviewHTML\' + Name);
end {TNppPluginPreviewHTML.GetSettings};


{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.DoNppnToolbarModification;
var
  tb: TToolbarIcons;
begin
  tb.ToolbarIcon := 0;
  tb.ToolbarBmp := LoadImage(Hinstance, 'TB_PREVIEW_HTML', IMAGE_BITMAP, 0, 0, (LR_DEFAULTSIZE or LR_LOADMAP3DCOLORS));
  SendMessage(self.NppData.NppHandle, NPPM_ADDTOOLBARICON, WPARAM(self.CmdIdFromDlgId(0)), LPARAM(@tb));

//  SendMessage(self.NppData.ScintillaMainHandle, SCI_SETMODEVENTMASK, SC_MOD_INSERTTEXT or SC_MOD_DELETETEXT, 0);
//  SendMessage(self.NppData.ScintillaSecondHandle, SCI_SETMODEVENTMASK, SC_MOD_INSERTTEXT or SC_MOD_DELETETEXT, 0);
end {TNppPluginPreviewHTML.DoNppnToolbarModification};

{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.DoNppnBufferActivated(const BufferID: THandle);
begin
  inherited;
  if Assigned(frmHTMLPreview) and frmHTMLPreview.Visible then begin
    frmHTMLPreview.btnRefresh.Click;
  end;
end {TNppPluginPreviewHTML.DoNppnBufferActivated};

{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.DoNppnFileClosed(const BufferID: THandle);
begin
  if Assigned(frmHTMLPreview) then begin
    frmHTMLPreview.ForgetBuffer(BufferID);
  end;
  inherited;
end {TNppPluginPreviewHTML.DoNppnFileClosed};

{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.DoModified(const hwnd: HWND; const modificationType: Integer);
begin
  if Assigned(frmHTMLPreview) and frmHTMLPreview.Visible and (modificationType and (SC_MOD_INSERTTEXT or SC_MOD_DELETETEXT) <> 0) then begin
    frmHTMLPreview.ResetTimer;
  end;
  inherited;
end {TNppPluginPreviewHTML.DoModified};



////////////////////////////////////////////////////////////////////////////////////////////////////
initialization
  try
    Npp := TNppPluginPreviewHTML.Create;
  except
    ShowException(ExceptObject, ExceptAddr);
  end;
end.
