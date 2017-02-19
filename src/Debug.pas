unit Debug;

interface

procedure ODS(const DebugOutput: string); overload;
procedure ODS(const DebugOutput: string; const Args: array of const); overload;

implementation
uses
  Classes, SysUtils,
  Windows,
  L_SpecialFolders;

var
  OutputLog: TStreamWriter;

{ ------------------------------------------------------------------------------------------------ }
procedure ODS(const DebugOutput: string); overload;
begin
  OutputDebugString(PChar('PreviewHTML['+IntToHex(GetCurrentThreadId, 4)+']: ' + DebugOutput));
  {$IFDEF DEBUG}
  if OutputLog = nil then begin
    OutputLog := TStreamWriter.Create(TFileStream.Create(ChangeFileExt(TSpecialFolders.DLLFullName, '.log'), fmCreate or fmShareDenyWrite), TEncoding.UTF8);
    OutputLog.OwnStream;
    OutputLog.BaseStream.Seek(0, soFromEnd);
  end;
  OutputLog.Write(FormatDateTime('yyyy-MM-dd hh:nn:ss.zzz: ', Now));
  OutputLog.WriteLine(DebugOutput.Replace(#10, #10 + StringOfChar(' ', 25)));
  {$ENDIF}
end {ODS};
{ ------------------------------------------------------------------------------------------------ }
procedure ODS(const DebugOutput: string; const Args: array of const); overload;
begin
  ODS(Format(DebugOutput, Args));
end{ODS};


initialization

finalization
  OutputLog.Free;

end.
