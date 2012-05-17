unit WebBrowser;
(* (c) DGMR raadgevende ingenieurs bv
 *
 * Created by: MCo On: 2009-01-15 16:06:04
 *
 * $Header:   J:/PVCSPROJ/EP/archives/EPGCheck/src/WebBrowser.pas-arc   1.0   24 Feb 2012 19:18:52   MCo  $
 *
 * Description:
 *
 * Exported Classes:
 * Exported Interfaces:
 * Exported Types:
 * Exported Functions:
 * Exported Variables:
 * Exported Constants:
 *
 * $Log:   J:/PVCSPROJ/EP/archives/EPGCheck/src/WebBrowser.pas-arc  $
 * 
 *    Rev 1.0   24 Feb 2012 19:18:52   MCo
 * Initial revision.
 * 
 *    Rev 1.0   30 Jul 2009 16:57:26   MCo
 * Initial revision.
 * 
 *    Rev 1.1   10 Jul 2009 17:30:24   MCo
 * * Headers toegevoegd
 *)

interface
uses
  Classes, SHDocVw, MSHTML;

type

  TWBHelper = class helper for TWebBrowser
    public
      procedure CheckDocReady();
      procedure LoadBlankDoc();
      procedure LoadDocFromStream(Stream: TStream);
      procedure LoadDocFromString(const HTMLString: string);
      function  GetDocument(): IHTMLDocument2;
      function  GetRootElement(): IHTMLElement;
      procedure ExecJS(const Code: string);
  end;

implementation
uses
  SysUtils, Forms, ActiveX, Variants;

{ ------------------------------------------------------------------------------------------------ }
procedure TWBHelper.ExecJS(const Code: string);
begin
  GetDocument.parentWindow.execScript(Code, 'JScript');
end;

{ ------------------------------------------------------------------------------------------------ }
function TWBHelper.GetDocument: IHTMLDocument2;
begin
  Result := Self.Document as IHTMLDocument2;
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TWBHelper.LoadBlankDoc();
begin
  Self.Navigate('about:blank');
  while not Self.ReadyState in [READYSTATE_INTERACTIVE, READYSTATE_COMPLETE] do begin
    Forms.Application.ProcessMessages;
    Sleep(0);
  end;
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TWBHelper.CheckDocReady();
begin
  if not Assigned(Self.Document) then
    LoadBlankDoc();
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TWBHelper.LoadDocFromStream(Stream: TStream);
begin
  CheckDocReady();
  (Self.Document as IPersistStreamInit).Load(TStreamAdapter.Create(Stream));
end;

{ ------------------------------------------------------------------------------------------------ }
procedure TWBHelper.LoadDocFromString(const HTMLString: string);
var
  v: OleVariant;
  HTMLDocument: IHTMLDocument2;
begin
  CheckDocReady();
  HTMLDocument := Self.Document as IHTMLDocument2;
  v := VarArrayCreate([0, 0], varVariant);
  v[0] := HTMLString;
  HTMLDocument.Write(PSafeArray(TVarData(v).VArray));
  HTMLDocument.Close;
end;

{ ------------------------------------------------------------------------------------------------ }
function TWBHelper.GetRootElement: IHTMLElement;
begin
  Result := Self.GetDocument.body;
  while Assigned(Result) and Assigned(Result.parentElement) do begin
    Result := Result.parentElement;
  end;
end;

end.
