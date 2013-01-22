unit F_PreviewHTML;

////////////////////////////////////////////////////////////////////////////////////////////////////
interface

uses
  Windows, Messages, SysUtils, Classes, Variants, Graphics, Controls, Forms,
  Dialogs, StdCtrls, SHDocVw, Vcl.OleCtrls, Vcl.ComCtrls, Vcl.ExtCtrls,
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
  private
    { Private declarations }
    function TransformXMLToHTML(const XML: WideString): string;
    function DetermineCustomFilter: string;
    function ExecuteCustomFilter(const FilterName, HTML: string): string;
  public
    { Public declarations }
  end;

var
  frmHTMLPreview: TfrmHTMLPreview;

////////////////////////////////////////////////////////////////////////////////////////////////////
implementation
uses
  ShellAPI, ComObj, StrUtils, IniFiles,
  RegExpr,
  WebBrowser, SciSupport, U_Npp_PreviewHTML;

{$R *.dfm}

{ ================================================================================================ }

procedure TfrmHTMLPreview.FormCreate(Sender: TObject);
begin
  self.NppDefaultDockingMask := DWS_DF_FLOATING; // whats the default docking position
  //self.KeyPreview := true; // special hack for input forms
  self.OnFloat := self.FormFloat;
  self.OnDock := self.FormDock;
  inherited;
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.btnCloseStatusbarClick(Sender: TObject);
begin
  sbrIE.Visible := False;
end;

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
begin
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
    {--- MCO 22-01-2013: execute the custom filter, and assign the output to the HTML variable ---}
    HTML := ExecuteCustomFilter(FilterName, HTML);
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

    if Pos('<base ', HTML) = 0 then begin
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
  end else begin
    self.UpdateDisplayInfo('');
  end;
end;

{ ------------------------------------------------------------------------------------------------ }
function TfrmHTMLPreview.DetermineCustomFilter: string;
var
  DocFileName, ConfigDir: TFileName;
  Filters: TIniFile;
  Names: TStringList;
  i: Integer;
  Match: Boolean;
  Extension, Language, DocLanguage: string;
  DocLangType, LangType: Integer;
begin
  ConfigDir := StringOfChar(#0, MAX_PATH);
  SendMessage(Npp.NppData.NppHandle, NPPM_GETPLUGINSCONFIGDIR, WPARAM(Length(ConfigDir)), LPARAM(PChar(ConfigDir)));
  DocFileName := StringOfChar(#0, MAX_PATH);
  SendMessage(Npp.NppData.NppHandle, NPPM_GETFILENAME, WPARAM(Length(DocFileName)), LPARAM(PChar(DocFileName)));

  ForceDirectories(ConfigDir + '\PreviewHTML');
  Filters := TIniFile.Create(ConfigDir + '\PreviewHTML\Filters.ini');
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
function TfrmHTMLPreview.ExecuteCustomFilter(const FilterName, HTML: string): string;
var
  ConfigDir: string;
  Filters: TIniFile;
  i: Integer;
begin
  ConfigDir := StringOfChar(#0, MAX_PATH);
  SendMessage(Npp.NppData.NppHandle, NPPM_GETPLUGINSCONFIGDIR, WPARAM(Length(ConfigDir)), LPARAM(PChar(ConfigDir)));

  ForceDirectories(ConfigDir + '\PreviewHTML');
  Filters := TIniFile.Create(ConfigDir + '\PreviewHTML\Filters.ini');
  try
    Filters.ReadString(FilterName, 'Command', '');
    {$MESSAGE HINT 'TODO: Output the HTML to temp file, substitute the temp file name, execute the command with pipe, and read the output — MCO 22-01-2013'}
    {--- MCO 22-01-2013: TODO: ways of transferring the HTML to the filter exe: on the input stream, write to file and pass the file name ---}
    {--- MCO 22-01-2013: TODO: ways of transferring the result from the filter exe: read from the output stream, read the changed input file, or read a (separate) output file ---}
  finally
    Filters.Free;
  end;
end {TfrmHTMLPreview.ExecuteCustomFilter};

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.btnAboutClick(Sender: TObject);
begin
  (npp as TNppPluginPreviewHTML).CommandShowAbout;
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TfrmHTMLPreview.btnCloseClick(Sender: TObject);
begin
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
  if not SameText(bstrUrl, 'about:blank') and not StartsText('javascript:', bstrURL) then begin
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
