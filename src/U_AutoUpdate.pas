unit U_AutoUpdate;

interface

uses
  SysUtils;

type
  TPluginUpdate = class
  private
    FURL: string;
//    FOnProgress: TProgressEvent; // TODO: see L_HttpClient
    FLatestVersion: string;

    function GetCurrentVersion: string;
    function GetLatestVersion: string;
  public
    constructor Create{(const AURL: string)};

    function IsUpdateAvailable(out NewVersion, Changes: string): Boolean;
    function DownloadUpdate: string;
    function ReplacePlugin(const PathDownloaded: string): Boolean;

    property URL: string            read FURL;

    property CurrentVersion: string read GetCurrentVersion;
    property LatestVersion: string  read GetLatestVersion;
  public
    class function CompareVersions(const VersionA, VersionB: string): Integer;
  end;

  EUpdateError = class(Exception);

implementation
uses
  Classes, RegularExpressions, Windows, NetEncoding, Zip,
  L_HttpClient, L_VersionInfoW, L_SpecialFolders,
  Debug;

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
begin
  // TODO
end {TPluginUpdate.GetLatestVersion};

{ ------------------------------------------------------------------------------------------------ }
function TPluginUpdate.IsUpdateAvailable(out NewVersion, Changes: string): Boolean;
var
  Http: THttpClient;
  Notes: string;
  Matches: TMatchCollection;
  Match: TMatch;
  s: string;
  FirstPos, LastPos: Integer;
begin
  // Download the current release notes
  ODS('Create HttpClient');
  Http := THttpClient.Create('http://fossil.2of4.net/npp_preview/doc/publish/ReleaseNotes.txt');
  try
    ODS('Http.Get');
    if Http.Get() = 200 then begin
      ODS('%d %s', [Http.StatusCode, Http.StatusText]);
      for s in Http.ResponseHeaders do
        ODS(s);
      // get the response stream, and look for the latest version in there
      Notes := Http.ResponseString;
      ODS('Response: "%s"', [Copy(StringReplace(StringReplace(Notes, #10, '·', [rfReplaceAll]), #13, '·', [rfReplaceAll]), 1, 250)]);

      NewVersion := '';
      FirstPos := -1;
      LastPos := -1;

      Matches := TRegEx.Matches(Notes, 'v[0-9]+(\.[0-9]+){3}');
      ODS('Matches: %d', [Matches.Count]);
      for Match in Matches do begin
        ODS('Match: Success=%s; Value="%s"; Index=%d', [BoolToStr(Match.Success, True), Match.Value, Match.Index]);
        if NewVersion = '' then
          NewVersion := Match.Value;
        if FirstPos = -1 then
          FirstPos := Match.Index;
        if CompareVersions(CurrentVersion, Match.Value) <= 0 then begin
          // This version is equal or older than the current version
          LastPos := Match.Index;
          Break;
        end;
      end {for};

      if FirstPos = -1 then
        FirstPos := Notes.IndexOf('<pre>') + 5;
      if (LastPos = -1) or (LastPos = FirstPos) then
        LastPos := Notes.IndexOf('</pre>');

      Changes := Notes.Substring(FirstPos - 1, LastPos - FirstPos);
      ODS('NewVersion: "%s"; Changes: "%s"', [NewVersion, Changes]);

      Changes := TNetEncoding.HTML.Decode(Changes.Trim);
    end else begin
      // TODO: show message, open project's main page?
      raise EUpdateError.CreateFmt('%d %s', [Http.StatusCode, Http.StatusText]);
    end;
  finally
    Http.Free;
  end;

  Result := CompareVersions(CurrentVersion, NewVersion) > 0;

  FLatestVersion := NewVersion;
end {TPluginUpdate.IsUpdateAvailable};

{ ------------------------------------------------------------------------------------------------ }
function TPluginUpdate.DownloadUpdate: string;
var
  Http: THttpClient;
  FS: TFileStream;
begin
  // TODO: Download http://fossil.2of4.net/npp_preview/zip/Preview_Plugin.zip?uuid=publish&name=plugins to a temp dir,
  //  extract it to a custom temp folder, and return that folder's path
  Http := THttpClient.Create('http://fossil.2of4.net/npp_preview/zip/Preview_Plugin.zip?uuid=publish&name=plugins');
  try
    if Http.Get() = 200 then begin
      Result := TSpecialFolders.TempDll + 'Preview_Plugin.zip';
      FS := TFileStream.Create(Result, fmCreate or fmShareDenyWrite);
      try
        FS.CopyFrom(Http.ResponseStream, 0);
      finally
        FS.Free;
      end;
    end else begin
      // TODO: show message, open project's main page?
      raise EUpdateError.CreateFmt('%d %s', [Http.StatusCode, Http.StatusText]);
    end;
  finally
    Http.Free;
  end;
end {TPluginUpdate.DownloadUpdate};

{ ------------------------------------------------------------------------------------------------ }
function TPluginUpdate.ReplacePlugin(const PathDownloaded: string): Boolean;
var
  NppDir: string;
  Zip: TZipFile;
  i: Integer;
  Info: TZipHeader;
  Name: string;
  Backup: string;
begin
  NppDir := TSpecialFolders.DLL + '..\';

  Zip := TZipFile.Create;
  try
    ODS(PathDownloaded);
    Zip.Open(PathDownloaded, zmRead);
    for i := 0 to Zip.FileCount - 1 do begin
      Info := Zip.FileInfo[i];
      Name := Zip.FileName[i].Replace('/', '\');
      ODS(Name);

      if Name.EndsWith('\') then Continue; // Skip directories

      if SameFileName(ExtractFileName(Name), 'ReleaseNotes.txt') then begin
        Name := Name.Replace('ReleaseNotes.txt', 'Doc\PreviewHTML\ReleaseNotes.txt', [rfIgnoreCase]);
        ODS('=> ' + Name);
      end;

      if FileExists(NppDir + Name) then begin
        Backup := NppDir + ChangeFileExt(Name, '-' + CurrentVersion + ExtractFileExt(Name));
        if SameFileName(ExtractFileExt(Name), '.dll') then begin
          Backup := ChangeFileExt(Backup, '.~' + ExtractFileExt(Backup).Substring(1));
          ODS('Backup: ' + Backup);
          Win32Check(RenameFile(NppDir + Name, Backup));
        end else begin
          ODS('Backup: ' + Backup);
          Win32Check(CopyFile(PChar(NppDir + Name), PChar(Backup), False));
        end;
      end;

      ODS('Extracting file to: ' + NppDir + Name);
      ForceDirectories(ExtractFileDir(NppDir + Name));
      Zip.Extract(i, ExtractFileDir(NppDir + Name), False);
    end {for};
  finally
    Zip.Free;
  end;

  // TODO: Rename the current DLL to ChangeFileExt(DllName, '-' + OwnVersion + '.~dll')
  // HardlinkOrCopy all files in extract location to path relative to plugins folder. ./Config should
  //  be translated to PluginsConfigFolder. ReleaseNotes.txt gets special treatment: it's in the
  //  root folder, but should be moved to ./Doc/PreviewHTML.
  Result := True;
end {TPluginUpdate.ReplacePlugin};

{ ------------------------------------------------------------------------------------------------ }
class function TPluginUpdate.CompareVersions(const VersionA, VersionB: string): Integer;
var
  va, vb: string;
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

    Result := nb - na;
//ODS('Result = %d := %d - %d', [Result, nb, na]);
    if Result <> 0 then
      Exit;
  until (va = '') and (vb = '');
end {TPluginUpdate.CompareVersions};

end.
