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

  self.AddFuncItem('&Preview HTML', _FuncShowPreview);

//  sk.IsCtrl := true; sk.IsAlt := true; sk.IsShift := false;
//  sk.Key := #118; // CTRL ALT SHIFT F7
//  self.AddFuncItem('Replace Hello World', _FuncHelloWorld, sk);

  self.AddFuncSeparator;

  self.AddFuncItem('&About', _FuncShowAbout);
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
