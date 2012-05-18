unit U_Npp_PreviewHTML;

////////////////////////////////////////////////////////////////////////////////////////////////////
interface

uses
  NppPlugin, SysUtils, Windows, SciSupport, F_About, F_PreviewHTML;

type
  TNppPluginPreviewHTML = class(TNppPlugin)
  public
    constructor Create;
    procedure CommandShowPreview;
    procedure CommandReplaceHelloWorld;
    procedure CommandShowAbout;
    procedure DoNppnToolbarModification; override;
    procedure DoNppnBufferActivated(const BufferID: Cardinal); override;
  end;

procedure _FuncShowPreview; cdecl;
procedure _FuncReplaceHelloWorld; cdecl;
procedure _FuncShowAbout; cdecl;

var
  Npp: TNppPluginPreviewHTML;

////////////////////////////////////////////////////////////////////////////////////////////////////
implementation

{ ================================================================================================ }
{ TNppPluginPreviewHTML }

{ ------------------------------------------------------------------------------------------------ }
constructor TNppPluginPreviewHTML.Create;
var
  sk: TShortcutKey;
  i: Integer;
begin
  inherited;
  self.PluginName := '&Preview HTML';
  i := 0;

  self.AddFuncItem('&Show preview', _FuncShowPreview);

//  sk.IsCtrl := true; sk.IsAlt := true; sk.IsShift := false;
//  sk.Key := #118; // CTRL ALT SHIFT F7
//  self.AddFuncItem('Replace Hello World', _FuncHelloWorld, sk);

  self.AddFuncSeparator;

  self.AddFuncItem('About', _FuncShowAbout);
end;

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
procedure TNppPluginPreviewHTML.CommandReplaceHelloWorld;
var
  s: UTF8String;
begin
  s := 'Hello World';
  SendMessage(self.NppData.ScintillaMainHandle, SCI_REPLACESEL, 0, LPARAM(PAnsiChar(s)));
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
//  end else begin
//    SendMessage(Npp.NppData.NppHandle, NPPM_SETMENUITEMCHECK, WPARAM(self.CmdIdFromDlgId(ncDlgId)), 0);
//    FreeAndNil(frmHTMLPreview);
  end;
  frmHTMLPreview.Show;
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.DoNppnToolbarModification;
var
  tb: TToolbarIcons;
begin
  tb.ToolbarIcon := 0;
  tb.ToolbarBmp := LoadImage(Hinstance, 'IDB_TB_TEST', IMAGE_BITMAP, 0, 0, (LR_DEFAULTSIZE or LR_LOADMAP3DCOLORS));
  SendMessage(self.NppData.NppHandle, NPPM_ADDTOOLBARICON, WPARAM(self.CmdIdFromDlgId(0)), LPARAM(@tb));
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TNppPluginPreviewHTML.DoNppnBufferActivated(const BufferID: Cardinal);
begin
  inherited;
  if Assigned(frmHTMLPreview) then begin
{$MESSAGE HINT 'TODO: only refresh the preview if it’s visible (and if so configured)'}
    frmHTMLPreview.Button1.Click;
  end;
end;


////////////////////////////////////////////////////////////////////////////////////////////////////
initialization
  Npp := TNppPluginPreviewHTML.Create;
end.
