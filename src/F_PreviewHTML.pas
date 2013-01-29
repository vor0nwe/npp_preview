unit F_PreviewHTML;

////////////////////////////////////////////////////////////////////////////////////////////////////
interface

uses
  Windows, Messages, SysUtils, Classes, Variants, Graphics, Controls, Forms,
  Dialogs, StdCtrls, SHDocVw, OleCtrls, ComCtrls, ExtCtrls, IniFiles,
  NppPlugin, NppDockingForms,
  U_CustomFilter;

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
    FFilterThread: TCustomFilterThread;
    function  GetSettings(const Name: string = 'Settings.ini'): TIniFile;

    procedure SaveScrollPos;
    procedure RestoreScrollPos(const BufferID: NativeInt);

    function  DetermineCustomFilter: string;
    function  ExecuteCustomFilter(const FilterName, HTML: string; const BufferID: NativeInt): Boolean;
    function  TransformXMLToHTML(const XML: WideString): string;

    procedure FilterThreadTerminate(Sender: TObject);
  public
    { Public declarations }
    procedure ResetTimer;
    procedure ForgetBuffer(const BufferID: NativeInt);
    procedure DisplayPreview(HTML: string; const BufferID: NativeInt);
  end;

var
  frmHTMLPreview: TfrmHTMLPreview;

procedure ODS(const DebugOutput: string); overload;
procedure ODS(const DebugOutput: string; const Args: array of const); overload;

////////////////////////////////////////////////////////////////////////////////////////////////////
implementation
uses
  ShellAPI, ComObj, StrUtils, IOUtils, MSHTML,
  RegExpr,
  WebBrowser, SciSupport, U_Npp_PreviewHTML;

{$R *.dfm}

{ ------------------------------------------------------------------------------------------------ }
procedure ODS(const DebugOutput: string); overload;
begin
  OutputDebugString(PChar('PreviewHTML['+IntToHex(GetCurrentThreadId, 4)+']: ' + DebugOutput));
end {ODS};
{ ------------------------------------------------------------------------------------------------ }
procedure ODS(const DebugOutput: string; const Args: array of const); overload;
begin
  ODS(Format(DebugOutput, Args));
end{ODS};


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
  FreeAndNil(FFilterThread);
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
  Content: UTF8String;
  HTML: string;
  FilterName: string;
begin
  try
    SaveScrollPos;

    if Assigned(FFilterThread) then begin
ODS('FreeAndNil(FFilterThread);');
      FreeAndNil(FFilterThread);
    end;

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
    FilterName := DetermineCustomFilter;
    IsCustom := Length(FilterName) > 0;

    {$MESSAGE HINT 'TODO: Find a way to communicate why there is no preview, depending on the situation — MCO 22-01-2013'}

    if IsXML or IsHTML or IsCustom then begin
      Size := SendMessage(hScintilla, SCI_GETTEXT, 0, 0);
      SetLength(Content, Size);
      SendMessage(hScintilla, SCI_GETTEXT, Size, LPARAM(PAnsiChar(Content)));
      Content := UTF8String(PAnsiChar(Content));
      HTML := string(Content);
    end;

    if IsCustom then begin
//MessageBox(Npp.NppData.NppHandle, PChar(Format('FilterName: %s', [FilterName])), 'PreviewHTML', MB_ICONINFORMATION);
      if ExecuteCustomFilter(FilterName, HTML, BufferID) then begin
        Exit;
      end else begin
        HTML := '<pre style="color: darkred">ExecuteCustomFilter returned False</pre>';
      end;
    end else if IsXML then begin
      HTML := TransformXMLToHTML(HTML);
    end;

    DisplayPreview(HTML, BufferID);
  except
    on E: Exception do begin
ODS('btnRefreshClick ### %s: %s', [E.ClassName, StringReplace(E.Message, sLineBreak, '', [rfReplaceAll])]);
      sbrIE.SimpleText := E.Message;
      sbrIE.Visible := True;
    end;
  end;
end {TfrmHTMLPreview.btnRefreshClick};

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.DisplayPreview(HTML: string; const BufferID: NativeInt);
var
  IsHTML: Boolean;
  HeadStart: Integer;
  Size: WPARAM;
  Filename: nppString;
  View: Integer;
  hScintilla: THandle;
begin
ODS('DisplayPreview(HTML: "%s"(%d); BufferID: %x)', [StringReplace(Copy(HTML, 1, 10), #13#10, '', [rfReplaceAll]), Length(HTML), BufferID]);
  try
    IsHTML := Length(HTML) > 0;
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
      SendMessage(Self.Npp.NppData.NppHandle, NPPM_GETCURRENTSCINTILLA, 0, LPARAM(@View));
      if View = 0 then begin
        hScintilla := Self.Npp.NppData.ScintillaMainHandle;
      end else begin
        hScintilla := Self.Npp.NppData.ScintillaSecondHandle;
      end;
      SendMessage(hScintilla, SCI_GRABFOCUS, 0, 0);
    end else begin
      self.UpdateDisplayInfo('');
    end;

    RestoreScrollPos(BufferID);
  except
    on E: Exception do begin
ODS('DisplayPreview ### %s: %s', [E.ClassName, StringReplace(E.Message, sLineBreak, '', [rfReplaceAll])]);
      sbrIE.SimpleText := E.Message;
      sbrIE.Visible := True;
    end;
  end;
end {TfrmHTMLPreview.DisplayPreview};

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
function TfrmHTMLPreview.DetermineCustomFilter: string;
var
  DocFileName: nppString;
  Filters: TIniFile;
  Names: TStringList;
  i: Integer;
  Match: Boolean;
  Extension, Language, DocLanguage: string;
  DocLangType, LangType: Integer;
begin
  DocFileName := StringOfChar(#0, MAX_PATH);
  SendMessage(Npp.NppData.NppHandle, NPPM_GETFILENAME, WPARAM(Length(DocFileName)), LPARAM(nppPChar(DocFileName)));
  DocFileName := nppString(nppPChar(DocFileName));

  ForceDirectories(Npp.ConfigDir + '\PreviewHTML');
  Filters := TIniFile.Create(Npp.ConfigDir + '\PreviewHTML\Filters.ini');
  Names := TStringList.Create;
  try
    Filters.ReadSections(Names);
    for i := 0 to Names.Count - 1 do begin
      Match := False;

      {$MESSAGE HINT 'TODO: Test entire file name — MCO 22-01-2013'}

      {--- MCO 22-01-2013: Test extension ---}
      Extension := Filters.ReadString(Names[i], 'Extension', '');
      if (Extension <> '') and SameFileName(Extension, ExtractFileExt(DocFileName)) then begin
        Match := True;
      end;

      {--- MCO 22-01-2013: Test highlighter language ---}
      Language := Filters.ReadString(Names[i], 'Language', '');
      if Language <> '' then begin
        DocLangType := -1;
        SendMessage(Npp.NppData.NppHandle, NPPM_GETCURRENTLANGTYPE, WPARAM(0), LPARAM(@DocLangType));
        if DocLangType > -1 then begin
          if TryStrToInt(Language, LangType) and (LangType = DocLangType) then begin
            Match := True;
          end else begin
            SetLength(DocLanguage, SendMessage(Npp.NppData.NppHandle, NPPM_GETLANGUAGENAME, WPARAM(LangType), LPARAM(nil)));
            SetLength(DocLanguage, SendMessage(Npp.NppData.NppHandle, NPPM_GETLANGUAGENAME, WPARAM(LangType), LPARAM(PChar(DocLanguage))));
            if SameText(Language, DocLanguage) then begin
              Match := True;
            end else begin
              SetLength(DocLanguage, SendMessage(Npp.NppData.NppHandle, NPPM_GETLANGUAGEDESC, WPARAM(LangType), LPARAM(nil)));
              SetLength(DocLanguage, SendMessage(Npp.NppData.NppHandle, NPPM_GETLANGUAGEDESC, WPARAM(LangType), LPARAM(PChar(DocLanguage))));
              if SameText(Language, DocLanguage) then
                Match := True;
            end;
          end;
        end;
      end;

      {$MESSAGE HINT 'TODO: Test lexer — MCO 22-01-2013'}

      if Match then
        Exit(Names[i]);
    end;
  finally
    Names.Free;
    Filters.Free;
  end;
end {TfrmHTMLPreview.DetermineCustomFilter};

{ ------------------------------------------------------------------------------------------------ }
function TfrmHTMLPreview.ExecuteCustomFilter(const FilterName, HTML: string; const BufferID: NativeInt): Boolean;
var
  FilterData: TFilterData;
  DocFile: TFileName;
  View: Integer;
  hScintilla: THandle;
  Filters: TIniFile;
  BufferEncoding: NativeInt;
begin
  FilterData.Name := FilterName;
  FilterData.BufferID := BufferID;

  DocFile := StringOfChar(#0, MAX_PATH);
  SendMessage(Npp.NppData.NppHandle, NPPM_GETFULLCURRENTPATH, WPARAM(Length(DocFile)), LPARAM(PChar(DocFile)));
  DocFile := string(PChar(DocFile));
  FilterData.DocFile := DocFile;
  FilterData.Contents := HTML;

  SendMessage(Self.Npp.NppData.NppHandle, NPPM_GETCURRENTSCINTILLA, 0, LPARAM(@View));
  if View = 0 then begin
    hScintilla := Self.Npp.NppData.ScintillaMainHandle;
  end else begin
    hScintilla := Self.Npp.NppData.ScintillaSecondHandle;
  end;
  BufferEncoding := SendMessage(Npp.NppData.NppHandle, NPPM_GETBUFFERENCODING, BufferID, 0);
  case BufferEncoding of
    1, 4: FilterData.Encoding := TEncoding.UTF8;
    2, 6: FilterData.Encoding := TEncoding.BigEndianUnicode;
    3, 7: FilterData.Encoding := TEncoding.Unicode;
    5:    FilterData.Encoding := TEncoding.UTF7;
    else  FilterData.Encoding := TEncoding.ANSI;
  end;
  FilterData.UseBOM := BufferEncoding in [1, 2, 3];
  FilterData.Modified := SendMessage(hScintilla, SCI_GETMODIFY, 0, 0) <> 0;

  Filters := GetSettings('Filters.ini');
  try
    FilterData.FilterInfo := TStringList.Create;
    Filters.ReadSectionValues(FilterName, FilterData.FilterInfo);
  finally
    Filters.Free;
  end;

  FilterData.OnTerminate := FilterThreadTerminate;

  {--- 2013-01-26 Martijn: Create a new TCustomFilterThread ---}
  FFilterThread := TCustomFilterThread.Create(FilterData);
  Result := Assigned(FFilterThread);
end {TfrmHTMLPreview.ExecuteCustomFilter};

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.FilterThreadTerminate(Sender: TObject);
begin
ODS('FilterThreadTerminate');
if (Sender as TThread).FatalException is Exception then
  ODS('Fatal %s: "%s"', [((Sender as TThread).FatalException as Exception).ClassName, ((Sender as TThread).FatalException as Exception).Message]);

  FFilterThread := nil;
end {TfrmHTMLPreview.FilterThreadTerminate};


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
