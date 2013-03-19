unit U_AutoUpdate;

interface

uses
  SysUtils;

type
  TPluginUpdate = class
  private
    FURL: string;
//    FOnProgress: TProgressEvent; // TODO: see L_HttpClient

    function GetCurrentVersion: string;
    function GetLatestVersion: string;
  public
    constructor Create{(const AURL: string)};

    property URL: string            read FURL;

    property CurrentVersion: string read GetCurrentVersion;
    property LatestVersion: string  read GetLatestVersion;
  public
    class function CompareVersions(const VersionA, VersionB: string): Integer;
  end;

  EUpdateError = class(Exception);

implementation
uses
  RegularExpressions, Windows,
  L_HttpClient, L_VersionInfoW, L_SpecialFolders;

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


{ ------------------------------------------------------------------------------------------------ }
{ TPluginUpdate }

{ ------------------------------------------------------------------------------------------------ }
constructor TPluginUpdate.Create;
begin
  //FURL := AURL;
end {TPluginUpdate.Create};

{ ------------------------------------------------------------------------------------------------ }
function TPluginUpdate.GetCurrentVersion: string;
begin
  with TFileVersionInfo.Create(TSpecialFolders.DLLFullName) do begin
    Result := FileVersion;
    Free;
  end;
end {TPluginUpdate.GetCurrentVersion};

{ ------------------------------------------------------------------------------------------------ }
function TPluginUpdate.GetLatestVersion: string;
var
  Http: THttpClient;
  RxVersion: TRegEx;
  Match: TMatch;
  i: Integer;
begin
  Result := '';

  // Download the current release notes
ODS('Create HttpClient');
  Http := THttpClient.Create('http://fossil.2of4.net/npp_preview/doc/publish/ReleaseNotes.txt');
  try
ODS('Http.Get');
    if Http.Get() = 200 then begin
ODS('%d %s', [Http.StatusCode, Http.StatusText]);
for i := 0 to Http.ResponseHeaders.Count - 1 do
  ODS(Http.ResponseHeaders[i]);
ODS('Response: "%s"', [Copy(StringReplace(StringReplace(Http.ResponseString, #10, '◙', [rfReplaceAll]), #13, '♪', [rfReplaceAll]), 1, 250)]);
      // get the response stream, and look for the latest version in there
      Match := TRegEx.Match(Http.ResponseString, 'v[0-9]+(\.[0-9]+){3}');
ODS('Match: Success=%s; Value="%s"', [BoolToStr(Match.Success, True), Match.Value]);
      if Match.Success then begin
        Result := Match.Value;
      end;
    end else begin
      raise EUpdateError.CreateFmt('%d %s', [Http.StatusCode, Http.StatusText]);
    end;
  finally
    Http.Free;
  end;
end {TPluginUpdate.GetLatestVersion};

{ ------------------------------------------------------------------------------------------------ }
class function TPluginUpdate.CompareVersions(const VersionA, VersionB: string): Integer;
var
  va, vb: string;
  ia, ib: Integer;
  na, nb: Integer;
  Code: Integer;
begin
  Result := 0;
  if (VersionA = '') or (VersionB = '') then
    Exit;

  if VersionA[1] = 'v' then va := Copy(VersionA, 2) else va := VersionA;
  if VersionB[1] = 'v' then vb := Copy(VersionB, 2) else vb := VersionB;

  repeat
    na := 0;
    if va <> '' then begin
//ODS('va: "%s"; na: %d', [va, na]);
      Val(va, na, Code);
//ODS('va: "%s"; na: %d; Code: %d', [va, na, Code]);
      if Code in [0, 1] then
        va := ''
      else
        va := Copy(va, Code + 1);
    end;

    nb := 0;
    if vb <> '' then begin
//ODS('vb: "%s"; nb: %d', [vb, nb]);
      Val(vb, nb, Code);
//ODS('vb: "%s"; nb: %d; Code: %d', [vb, nb, Code]);
      if Code in [0, 1] then
        vb := ''
      else
        vb := Copy(vb, Code + 1);
    end;

    Result := na - nb;
//ODS('Result = %d := %d - %d', [Result, nb, na]);
    if Result <> 0 then
      Exit;
  until (va = '') and (vb = '');
end {TPluginUpdate.CompareVersions};

end.
