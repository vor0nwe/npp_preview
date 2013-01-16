unit U_Npp_PreviewHTML;

////////////////////////////////////////////////////////////////////////////////////////////////////
interface

uses
  SysUtils, Windows,
  NppPlugin, SciSupport,
  F_About, F_PreviewHTML;

type
  TNppPluginPreviewHTML = class(TNppPlugin)
  public
    constructor Create;
    procedure CommandShowPreview;
    procedure CommandReplaceHelloWorld;
    procedure CommandShowAbout;
    procedure CommandSetIEVersion(const BrowserEmulation: Integer);
    procedure DoNppnToolbarModification; override;
    procedure DoNppnBufferActivated(const BufferID: Cardinal); override;
  end;

procedure _FuncShowPreview; cdecl;
procedure _FuncReplaceHelloWorld; cdecl;
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
  WebBrowser, Registry;

{ ================================================================================================ }
{ TNppPluginPreviewHTML }

{ ------------------------------------------------------------------------------------------------ }
constructor TNppPluginPreviewHTML.Create;
var
  IEVersion: string;
  MajorIEVersion, Code, EmulatedVersion: Integer;
  sk: TShortcutKey;
begin
  inherited;
  self.PluginName := '&Preview HTML';

  sk.IsShift := True; sk.IsCtrl := true; sk.IsAlt := False;
  sk.Key := 'H'; // Ctrl-Shift-H
  self.AddFuncItem('&Preview HTML', _FuncShowPreview, sk);

  IEVersion := GetIEVersion;
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

  self.AddFuncItem('&About', _FuncShowAbout);
end {TNppPluginPreviewHTML.Create};

{ ------------------------------------------------------------------------------------------------ }
procedure _FuncReplaceHelloWorld; cdecl;
begin
  Npp.CommandReplaceHelloWorld;
end;
procedure _FuncShowAbout; cdecl;
{ ------------------------------------------------------------------------------------------------ }
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


{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.CommandReplaceHelloWorld;
var
  s: UTF8String;
begin
  s := 'Hello World';
  SendMessage(self.NppData.ScintillaMainHandle, SCI_REPLACESEL, 0, LPARAM(PAnsiChar(s)));
end;

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
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.CommandShowAbout;
begin
  with TAboutForm.Create(self) do begin
    ShowModal;
    Free;
  end;
end;

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
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.DoNppnToolbarModification;
var
  tb: TToolbarIcons;
begin
  tb.ToolbarIcon := 0;
  tb.ToolbarBmp := LoadImage(Hinstance, 'TB_PREVIEW_HTML', IMAGE_BITMAP, 0, 0, (LR_DEFAULTSIZE or LR_LOADMAP3DCOLORS));
  SendMessage(self.NppData.NppHandle, NPPM_ADDTOOLBARICON, WPARAM(self.CmdIdFromDlgId(0)), LPARAM(@tb));
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.DoNppnBufferActivated(const BufferID: Cardinal);
begin
  inherited;
  if Assigned(frmHTMLPreview) and frmHTMLPreview.Visible then begin
{$MESSAGE HINT 'TODO: only refresh the preview if it’s configured to follow the active tab'}
    frmHTMLPreview.btnRefresh.Click;
  end;
end;


////////////////////////////////////////////////////////////////////////////////////////////////////
initialization
  Npp := TNppPluginPreviewHTML.Create;
end.
