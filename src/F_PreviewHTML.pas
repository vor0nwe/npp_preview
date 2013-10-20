unit F_PreviewHTML;

////////////////////////////////////////////////////////////////////////////////////////////////////
interface

uses
  Windows, Messages, SysUtils, Classes, Variants, Graphics, Controls, Forms,
  Dialogs, StdCtrls, SHDocVw, OleCtrls, ComCtrls, ExtCtrls, IniFiles,
  NppPlugin, NppDockingForms;

type
  TfrmHTMLPreview = class(TNppDockingForm)
    wbIE: TWebBrowser;
    pnlButtons: TPanel;
    btnRefresh: TButton;
    btnClose: TButton;
    sbrIE: TStatusBar;
    pnlPreview: TPanel;
    pnlHTML: TPanel;
    btnAbout: TButton;
    tmrAutorefresh: TTimer;
    procedure btnRefreshClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure FormHide(Sender: TObject);
    procedure FormFloat(Sender: TObject);
    procedure FormDock(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure wbIETitleChange(ASender: TObject; const Text: WideString);
    procedure wbIEBeforeNavigate2(ASender: TObject; const pDisp: IDispatch; const URL, Flags,
      TargetFrameName, PostData, Headers: OleVariant; var Cancel: WordBool);
    procedure wbIENewWindow3(ASender: TObject; var ppDisp: IDispatch; var Cancel: WordBool;
      dwFlags: Cardinal; const bstrUrlContext, bstrUrl: WideString);
    procedure wbIEStatusTextChange(ASender: TObject; const Text: WideString);
    procedure wbIEStatusBar(ASender: TObject; StatusBar: WordBool);
    procedure btnCloseStatusbarClick(Sender: TObject);
    procedure btnAboutClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure tmrAutorefreshTimer(Sender: TObject);
  private
    { Private declarations }
    FBufferID: NativeInt;
    FScrollPositions: TStringList;

    function  GetSettings(const Name: string = 'Settings.ini'): TIniFile;

    procedure SaveScrollPos;
    procedure RestoreScrollPos(const BufferID: NativeInt);

    function TransformXMLToHTML(const XML: WideString): string;
  public
    { Public declarations }
    procedure ResetTimer;
    procedure ForgetBuffer(const BufferID: NativeInt);
  end;

var
  frmHTMLPreview: TfrmHTMLPreview;

////////////////////////////////////////////////////////////////////////////////////////////////////
implementation
uses
  ShellAPI, ComObj, StrUtils, IOUtils,
  RegExpr,
  WebBrowser, SciSupport, U_Npp_PreviewHTML, MSHTML;

{$R *.dfm}

{ ================================================================================================ }

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.FormCreate(Sender: TObject);
begin
  FScrollPositions := TStringList.Create;
  self.NppDefaultDockingMask := DWS_DF_FLOATING; // whats the default docking position
  //self.KeyPreview := true; // special hack for input forms
  self.OnFloat := self.FormFloat;
  self.OnDock := self.FormDock;
  inherited;
  FBufferID := -1;
  with GetSettings() do begin
    tmrAutorefresh.Interval := ReadInteger('Autorefresh', 'Interval', tmrAutorefresh.Interval);
  end;
end {TfrmHTMLPreview.FormCreate};
{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.FormDestroy(Sender: TObject);
begin
  FreeAndNil(FScrollPositions);
  inherited;
end {TfrmHTMLPreview.FormDestroy};


{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.btnCloseStatusbarClick(Sender: TObject);
begin
  sbrIE.Visible := False;
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.tmrAutorefreshTimer(Sender: TObject);
begin
  tmrAutorefresh.Enabled := False;
  btnRefresh.Click;
end {TfrmHTMLPreview.tmrAutorefreshTimer};

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.btnRefreshClick(Sender: TObject);
var
  View: Integer;
  BufferID: Integer;
  hScintilla: THandle;
  Lexer: NativeInt;
  IsHTML, IsXML, IsCustom: Boolean;
  Size: WPARAM;
  Filename: nppString;
  Content: UTF8String;
  HTML: string;
  HeadStart: Integer;
  FilterName: string;
  CodePage: NativeInt;
begin
  try
    SaveScrollPos;

    SendMessage(Self.Npp.NppData.NppHandle, NPPM_GETCURRENTSCINTILLA, 0, LPARAM(@View));
    if View = 0 then begin
      hScintilla := Self.Npp.NppData.ScintillaMainHandle;
    end else begin
      hScintilla := Self.Npp.NppData.ScintillaSecondHandle;
    end;
    BufferID := SendMessage(Self.Npp.NppData.NppHandle, NPPM_GETCURRENTBUFFERID, 0, 0);

    Lexer := SendMessage(hScintilla, SCI_GETLEXER, 0, 0);
    IsHTML := (Lexer = SCLEX_HTML);
    IsXML := (Lexer = SCLEX_XML);

    {--- MCO 22-01-2013: determine whether the current document matches a custom filter ---}
    FilterName := ''; // DetermineCustomFilter;
    IsCustom := Length(FilterName) > 0;

    {$MESSAGE HINT 'TODO: Find a way to communicate why there is no preview, depending on the situation — MCO 22-01-2013'}

    if IsXML or IsHTML or IsCustom then begin
      CodePage := SendMessage(hScintilla, SCI_GETCODEPAGE, 0, 0);
      Size := SendMessage(hScintilla, SCI_GETTEXT, 0, 0);
      SetLength(Content, Size);
      SendMessage(hScintilla, SCI_GETTEXT, Size, LPARAM(PAnsiChar(Content)));
      if CodePage = CP_ACP then begin
        HTML := string(PAnsiChar(Content));
      end else begin
        SetLength(HTML, Size);
        if Size > 0 then begin
          SetLength(HTML, MultiByteToWideChar(CodePage, MB_PRECOMPOSED, PAnsiChar(Content), Size, PWideChar(HTML), Length(HTML)));
          if Length(HTML) = 0 then
            RaiseLastOSError;
        end;
      end;
    end;

    if IsCustom then begin
      HTML := ''; // ExecuteCustomFilter(FilterName, HTML);
      IsHTML := Length(HTML) > 0;
    end else if IsXML then begin
      HTML := TransformXMLToHTML(HTML);
      IsHTML := Length(HTML) > 0;
    end;

    pnlHTML.Visible := IsHTML;
    sbrIE.Visible := IsHTML and (Length(sbrIE.SimpleText) > 0);
    if IsHTML then begin
      Size := SendMessage(Self.Npp.NppData.NppHandle, NPPM_GETFULLPATHFROMBUFFERID, BufferID, LPARAM(nil));
      SetLength(Filename, Size);
      SetLength(Filename, SendMessage(Self.Npp.NppData.NppHandle, NPPM_GETFULLPATHFROMBUFFERID, BufferID, LPARAM(nppPChar(Filename))));

      if (Pos('<base ', HTML) = 0) and FileExists(Filename) then begin
        HeadStart := Pos('<head>', HTML);
        if HeadStart > 0 then
          Inc(HeadStart, 6)
        else
          HeadStart := 1;
        Insert('<base href="' + Filename + '" />', HTML, HeadStart);
      end;
      wbIE.LoadDocFromString(HTML);

      if wbIE.GetDocument <> nil then
        self.UpdateDisplayInfo(wbIE.GetDocument.title)
      else
        self.UpdateDisplayInfo('');

      {--- 2013-01-26 Martijn: the WebBrowser control has a tendency to steal the focus. We'll let
                                the editor take it back. ---}
      SendMessage(hScintilla, SCI_GRABFOCUS, 0, 0);
    end else begin
      self.UpdateDisplayInfo('');
    end;

    RestoreScrollPos(BufferID);
  except
    on E: Exception do begin
      sbrIE.SimpleText := E.Message;
      sbrIE.Visible := True;
    end;
  end;
end {TfrmHTMLPreview.btnRefreshClick};

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.SaveScrollPos;
var
  Index, ScrollTop, ScrollLeft: Integer;
  docEl: IHTMLElement2;
begin
  if FBufferID = -1 then
    Exit;

  if Assigned(wbIE.Document) and Assigned((wbIE.Document as IHTMLDocument3).documentElement) then begin
    docEl := (wbIE.Document as IHTMLDocument3).documentElement AS IHTMLElement2;
    ScrollTop := docEl.scrollTop;
    ScrollLeft := docEl.scrollLeft;
  end else begin
    ScrollTop := -1;
    ScrollLeft := -1;
  end;
  Index := FScrollPositions.IndexOfObject(TObject(FBufferID));
  if Index = -1 then
    FScrollPositions.AddObject(IntToStr(ScrollTop) + '=' + IntToStr(ScrollLeft), TObject(FBufferID))
  else
    FScrollPositions[Index] := IntToStr(ScrollTop) + '=' + IntToStr(ScrollLeft);
end {TfrmHTMLPreview.SaveScrollPos};

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.RestoreScrollPos(const BufferID: NativeInt);
var
  Index, ScrollTop, ScrollLeft: Integer;
  docEl: IHTMLElement2;
begin
  {--- MCO 22-01-2013: Look up this buffer's scroll position; if we know one, wait for the page
                          to finish loading, then restore the scroll position. ---}
  Index := FScrollPositions.IndexOfObject(TObject(BufferID));
  if Index > -1 then begin
    ScrollTop := StrToInt(FScrollPositions.Names[Index]);
    ScrollLeft := StrToInt(FScrollPositions.ValueFromIndex[Index]);
    if ScrollTop <> -1 then begin
      {$MESSAGE HINT 'TODO: This would be better if done in the browsercontrol's DocumentComplete event,
                            so as to prevent blocking Notepad++ — MCO 22-01-2013'}
      while not wbIE.ReadyState in [READYSTATE_INTERACTIVE, READYSTATE_COMPLETE] do begin
        Forms.Application.ProcessMessages;
        Sleep(0);
      end;
      if Assigned(wbIE.Document) and Assigned((wbIE.Document as IHTMLDocument3).documentElement) then begin
        docEl := (wbIE.Document as IHTMLDocument3).documentElement as IHTMLElement2;
        docEl.scrollTop := ScrollTop;
        docEl.scrollLeft := ScrollLeft;
      end;
    end;
  end;
  FBufferID := BufferID;
end {TfrmHTMLPreview.RestoreScrollPos};

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.ForgetBuffer(const BufferID: NativeInt);
var
  Index: Integer;
begin
  if FBufferID = BufferID then
    FBufferID := -1;
  if Assigned(FScrollPositions) then begin
    Index := FScrollPositions.IndexOfObject(TObject(BufferID));
    if Index > -1 then
      FScrollPositions.Delete(Index);
  end;
end {TfrmHTMLPreview.ForgetBuffer};

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.ResetTimer;
begin
  tmrAutorefresh.Enabled := False;
  tmrAutorefresh.Enabled := True;
end {TfrmHTMLPreview.ResetTimer};

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.btnAboutClick(Sender: TObject);
begin
  (npp as TNppPluginPreviewHTML).CommandShowAbout;
end {TfrmHTMLPreview.btnAboutClick};

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.btnCloseClick(Sender: TObject);
begin
  self.Hide;
end {TfrmHTMLPreview.btnCloseClick};

{ ------------------------------------------------------------------------------------------------ }
// special hack for input forms
// This is the best possible hack I could came up for
// memo boxes that don't process enter keys for reasons
// too complicated... Has something to do with Dialog Messages
// I sends a Ctrl+Enter in place of Enter
procedure TfrmHTMLPreview.FormKeyPress(Sender: TObject;
  var Key: Char);
begin
//  if (Key = #13) and (self.Memo1.Focused) then self.Memo1.Perform(WM_CHAR, 10, 0);
end;

{ ------------------------------------------------------------------------------------------------ }
// Docking code calls this when the form is hidden by either "x" or self.Hide
procedure TfrmHTMLPreview.FormHide(Sender: TObject);
begin
  SaveScrollPos;
  SendMessage(self.Npp.NppData.NppHandle, NPPM_SETMENUITEMCHECK, self.CmdID, 0);
  self.Visible := False;
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
function TfrmHTMLPreview.GetSettings(const Name: string): TIniFile;
begin
  ForceDirectories(Npp.ConfigDir + '\PreviewHTML');
  Result := TIniFile.Create(Npp.ConfigDir + '\PreviewHTML\' + Name);
end {TfrmHTMLPreview.GetSettings};

{ ------------------------------------------------------------------------------------------------ }
function TfrmHTMLPreview.TransformXMLToHTML(const XML: WideString): string;
  { - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }
  function CreateDOMDocument: OleVariant;
  var
    nVersion: Integer;
  begin
    VarClear(Result);
    for nVersion := 7 downto 4 do begin
      try
        Result := CreateOleObject(Format('MSXML2.DOMDocument.%d.0', [nVersion]));
        if not VarIsClear(Result) then begin
          if nVersion >= 4 then begin
            Result.setProperty('NewParser', True);
          end;
          if nVersion >= 6 then begin
            Result.setProperty('AllowDocumentFunction', True);
            Result.setProperty('AllowXsltScript', True);
            Result.setProperty('ResolveExternals', True);
            Result.setProperty('UseInlineSchema', True);
            Result.setProperty('ValidateOnParse', False);
          end;
          Break;
        end;
      except
        VarClear(Result);
      end;
    end{for};
  end {CreateDOMDocument};
  { - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }
var
  bMethodHTML: Boolean;
  xDoc, xPI, xStylesheet, xOutput: OleVariant;
  rexHref: TRegExpr;
begin
  Result := '';
  try
    try
      {--- MCO 30-05-2012: Check to see if there's an xml-stylesheet to convert the XML to HTML. ---}
      xDoc := CreateDOMDocument;
      if VarIsClear(xDoc) then Exit;
      if not xDoc.LoadXML(XML) then Exit;

      xPI := xDoc.selectSingleNode('//processing-instruction("xml-stylesheet")');
      if VarIsClear(xPI) then Exit;

      rexHref := TRegExpr.Create;
      try
        rexHref.ModifierI := False;
        rexHref.Expression := '(^|\s+)href=["'']([^"'']*?)["'']';
        if not rexHref.Exec(xPI.nodeValue) then Exit;

        xStylesheet := CreateDOMDocument;
        if not xStylesheet.Load(rexHref.Match[2]) then Exit;
      finally
        rexHref.Free;
      end;

      bMethodHTML := SameText(xDoc.documentElement.nodeName, 'html');
      if not bMethodHTML then begin
        xStylesheet.setProperty('SelectionNamespaces', 'xmlns:xsl="http://www.w3.org/1999/XSL/Transform"');
        xOutput := xStylesheet.selectSingleNode('/*/xsl:output');
        if VarIsClear(xOutput) then
          Exit;

        bMethodHTML := SameStr(VarToStrDef(xOutput.getAttribute('method'), 'xml'), 'html');
      end;
      if not bMethodHTML then Exit;

      Result := xDoc.transformNode(xStylesheet.documentElement);
    except
      on E: Exception do begin
        {--- MCO 30-05-2012: Ignore any errors; we weren't able to perform the transformation ---}
        Result := '<html><title>Error transforming XML to HTML</title><body><pre style="color: red">' + StringReplace(E.Message, '<', '&lt;', [rfReplaceAll]) + '</pre></body></html>';
      end;
    end;
  finally
    VarClear(xOutput);
    VarClear(xStylesheet);
    VarClear(xPI);
    VarClear(xDoc);
  end;
end {TfrmHTMLPreview.TransformXMLToHTML};

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.wbIEBeforeNavigate2(ASender: TObject; const pDisp: IDispatch; const URL,
  Flags, TargetFrameName, PostData, Headers: OleVariant; var Cancel: WordBool);
var
  Handle: HWND;
begin
  inherited;
  if not SameText(URL, 'about:blank') and not StartsText('javascript:', URL) then begin
    if Assigned(Npp) then
      Handle := Npp.NppData.NppHandle
    else
      Handle := 0;
    ShellExecute(Handle, nil, PChar(VarToStr(URL)), nil, nil, SW_SHOWDEFAULT);
    Cancel := True;
  end;
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.wbIENewWindow3(ASender: TObject; var ppDisp: IDispatch;
  var Cancel: WordBool; dwFlags: Cardinal; const bstrUrlContext, bstrUrl: WideString);
var
  Handle: HWND;
begin
  if not SameText(bstrUrl, 'about:blank') and not StartsText('javascript:', bstrURL) then begin
    if Assigned(Npp)  then
      Handle := Npp.NppData.NppHandle
    else
      Handle := 0;
    ShellExecute(Handle, nil, PChar(bstrUrl), nil, nil, SW_SHOWDEFAULT)
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
