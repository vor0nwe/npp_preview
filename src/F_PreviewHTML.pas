unit F_PreviewHTML;

////////////////////////////////////////////////////////////////////////////////////////////////////
interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, NppDockingForms, StdCtrls, NppPlugin, Vcl.OleCtrls, SHDocVw, Vcl.ExtCtrls, Vcl.ComCtrls;

type
  TfrmHTMLPreview = class(TNppDockingForm)
    wbIE: TWebBrowser;
    pnlButtons: TPanel;
    Button1: TButton;
    Button2: TButton;
    sbrIE: TStatusBar;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure FormHide(Sender: TObject);
    procedure FormFloat(Sender: TObject);
    procedure FormDock(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure wbIETitleChange(ASender: TObject; const Text: WideString);
    procedure FormActivate(Sender: TObject);
    procedure wbIEBeforeNavigate2(ASender: TObject; const pDisp: IDispatch; const URL, Flags,
      TargetFrameName, PostData, Headers: OleVariant; var Cancel: WordBool);
    procedure wbIENewWindow3(ASender: TObject; var ppDisp: IDispatch; var Cancel: WordBool;
      dwFlags: Cardinal; const bstrUrlContext, bstrUrl: WideString);
    procedure wbIEStatusTextChange(ASender: TObject; const Text: WideString);
    procedure wbIEStatusBar(ASender: TObject; StatusBar: WordBool);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmHTMLPreview: TfrmHTMLPreview;

////////////////////////////////////////////////////////////////////////////////////////////////////
implementation
uses
  ShellAPI,
  WebBrowser, SciSupport;

{$R *.dfm}

{ ================================================================================================ }

procedure TfrmHTMLPreview.FormCreate(Sender: TObject);
begin
  self.NppDefaultDockingMask := DWS_DF_FLOATING; // whats the default docking position
  self.KeyPreview := true; // special hack for input forms
  self.OnFloat := self.FormFloat;
  self.OnDock := self.FormDock;
  inherited;
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.FormActivate(Sender: TObject);
begin
  inherited;
  Button1.Click;
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.Button1Click(Sender: TObject);
var
  View: Integer;
  BufferID: Integer;
  hScintilla: THandle;
  Size: Integer;
  Filename: nppString;
  Content: UTF8String;
  HTML: string;
  HeadStart: Integer;
begin
  inherited;

  SendMessage(Self.Npp.NppData.NppHandle, NPPM_GETCURRENTSCINTILLA, 0, LPARAM(@View));
  if View = 0 then begin
    hScintilla := Self.Npp.NppData.ScintillaMainHandle;
  end else begin
    hScintilla := Self.Npp.NppData.ScintillaSecondHandle;
  end;
  BufferID := SendMessage(Self.Npp.NppData.NppHandle, NPPM_GETCURRENTBUFFERID, 0, 0);

  Size := SendMessage(Self.Npp.NppData.NppHandle, NPPM_GETFULLPATHFROMBUFFERID, BufferID, LPARAM(nil));
  SetLength(Filename, Size);
  SetLength(Filename, SendMessage(Self.Npp.NppData.NppHandle, NPPM_GETFULLPATHFROMBUFFERID, BufferID, LPARAM(nppPChar(Filename))));
//  Filename := nppString(nppPChar(Filename));

//  Filename := StringOfChar(#0, MAX_PATH);
//  Size := SendMessage(Self.Npp.NppData.NppHandle, NPPM_GETFULLCURRENTPATH, Length(Filename), LPARAM(nppPChar(Filename)));
//  Filename := nppString(nppPChar(Filename));

  Size := SendMessage(hScintilla, SCI_GETTEXT, 0, 0);
  SetLength(Content, Size);
  SendMessage(hScintilla, SCI_GETTEXT, Size, NativeInt(PAnsiChar(Content)));
  Content := UTF8String(PAnsiChar(Content));

  HTML := string(Content);

  if SendMessage(Self.Npp.NppData.NppHandle, NPPM_GETBUFFERLANGTYPE, BufferID, 0) = NativeInt(L_HTML) then begin
    if Pos('<base ', HTML) = 0 then begin
      HeadStart := Pos('<head>', HTML);
      if HeadStart > 0 then
        Inc(HeadStart, 6)
      else
        HeadStart := 1;
      Insert('<base href="' + Filename + '" />', HTML, HeadStart);
    end;
  end;
  wbIE.LoadDocFromString(string(HTML));

  if wbIE.GetDocument <> nil then
    self.UpdateDisplayInfo(wbIE.GetDocument.title)
  else
    self.UpdateDisplayInfo('');
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.Button2Click(Sender: TObject);
begin
  inherited;
  self.Hide;
end;

{ ------------------------------------------------------------------------------------------------ }
// special hack for input forms
// This is the best possible hack I could came up for
// memo boxes that don't process enter keys for reasons
// too complicated... Has something to do with Dialog Messages
// I sends a Ctrl+Enter in place of Enter
procedure TfrmHTMLPreview.FormKeyPress(Sender: TObject;
  var Key: Char);
begin
  inherited;
//  if (Key = #13) and (self.Memo1.Focused) then self.Memo1.Perform(WM_CHAR, 10, 0);
end;

{ ------------------------------------------------------------------------------------------------ }
// Docking code calls this when the form is hidden by either "x" or self.Hide
procedure TfrmHTMLPreview.FormHide(Sender: TObject);
begin
  inherited;
  SendMessage(self.Npp.NppData.NppHandle, NPPM_SETMENUITEMCHECK, self.CmdID, 0);
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.FormDock(Sender: TObject);
begin
  SendMessage(self.Npp.NppData.NppHandle, NPPM_SETMENUITEMCHECK, self.CmdID, 1);
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.FormFloat(Sender: TObject);
begin
  SendMessage(self.Npp.NppData.NppHandle, NPPM_SETMENUITEMCHECK, self.CmdID, 1);
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.FormShow(Sender: TObject);
begin
  inherited;
  SendMessage(self.Npp.NppData.NppHandle, NPPM_SETMENUITEMCHECK, self.CmdID, 1);
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.wbIEBeforeNavigate2(ASender: TObject; const pDisp: IDispatch; const URL,
  Flags, TargetFrameName, PostData, Headers: OleVariant; var Cancel: WordBool);
begin
  inherited;
  if (URL <> 'about:blank') and Assigned(Npp) then begin
    ShellExecute(Npp.NppData.NppHandle, nil, PChar(VarToStr(URL)), nil, nil, SW_SHOWDEFAULT);
    Cancel := True;
  end;
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.wbIENewWindow3(ASender: TObject; var ppDisp: IDispatch;
  var Cancel: WordBool; dwFlags: Cardinal; const bstrUrlContext, bstrUrl: WideString);
begin
  inherited;

  if bstrUrl <> 'about:blank' then begin
    ShellExecute(Npp.NppData.NppHandle, nil, PChar(bstrUrl), nil, nil, SW_SHOWDEFAULT);
  end;
  Cancel := True;
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.wbIEStatusBar(ASender: TObject; StatusBar: WordBool);
begin
  sbrIE.Visible := StatusBar;
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.wbIEStatusTextChange(ASender: TObject; const Text: WideString);
begin
  sbrIE.SimpleText := Text;
  sbrIE.Visible := Length(Text) > 0;
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.wbIETitleChange(ASender: TObject; const Text: WideString);
begin
  inherited;
  self.UpdateDisplayInfo(StringReplace(Text, 'about:blank', '', [rfReplaceAll]));
end;

end.
