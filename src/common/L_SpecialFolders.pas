unit L_SpecialFolders;

////////////////////////////////////////////////////////////////////////////////////////////////////
interface
{$REGION 'PVCS-header'}
(* © DGMR raadgevende ingenieurs BV
 *
 * Created by: MCO on 31-01-2012
 *
 * $Header: $
 *
 * Description: TSpecialFolders is een class waar het pad van alle 'speciale' mappen
 *               mee opgevraagd kan worden.  Het zijn allemaal class properties, dus er
 *               hoeft niet eerst een instantie aangemaakt te worden.
 *              Deze class is betrouwbaarder dan EnvironmentVariables, Application.ExeName
 *               of ParamStr(0), want EnvironmentVariables zijn niet altijd allemaal gedefinieerd,
 *               en ParamStr(0) is niet beschikbaar vanuit een Dll.
 *              Wanneer de Dll-properties vanuit een EXE worden aangeroepen, wordt het pad van
 *               de EXE gebruikt.
 *              Als met de functies TempExe, TempDll, en TempExeDll een nieuwe directory wordt
 *               gemaakt, dan wordt deze geleegd bij het afsluiten van het programma.
 *              Met de functie Settings kan het pad opgevraagd worden waar instellingen opgeslagen
 *               zouden moeten worden, rekening houdend met de volgende factoren:
 *               - sfUserSpecific: of het pad gebruikersgebonden moet zijn, of voor alle gebruikers moet gelden;
 *               - sfMachineSpecific: of het pad alleen op de huidige machine moet gelden (local),
 *                  of ook met het gebruikersprofiel op het netwerk gekopieerd mag worden (roaming)
 *               - sfExeSpecific: of de naam van de Exe aan het pad toegevoegd moet worden;
 *               - sfDllSpecific: of de naam van de Dll (of van de Exe als we niet in een Dll zitten)
 *                  aan het pad toegevoegd moet worden.
 *               De functie Settings zonder parameters haalt [sfUserSpecific, sfDllSpecific] op.
 * Examples:
 *    TSpecialFolders.Settings;                 // C:\Users\MCo\AppData\Roaming\ENORM\
 *    TSpecialFolders.Settings([sfMachineSpecific, sfDllSpecific]); // C:\ProgramData\ENORM\
 *    TSpecialFolders.Settings([sfUserSpecific, sfMachineSpecific, sfExeSpecific, sfDllSpecific]); // C:\Users\MCo\AppData\Local\ISL2\SLC_A3\
 *    TSpecialFolders.Exe;                      // C:\Program Files (x86)\ISL\ISL2 V4.01\
 *    TSpecialFolders.ExeFullName;              // C:\Program Files (x86)\ISL\ISL2 V4.01\ISL2.EXE
 *                                              // C:\Program Files (x86)\ENORM V0.70\ENORM.EXE
 *    TSpecialFolders.DLL;                      // C:\Program Files (x86)\ISL\ISL2 V4.01\
 *    TSpecialFolders.DLLFullName;              // C:\Program Files (x86)\ISL\ISL2 V4.01\SLC_A3.dll
 *                                              // C:\Program Files (x86)\ENORM V0.70\ENORM.EXE
 *    TSpecialFolders.Temp;                     // C:\Users\MCo\AppData\Local\Temp\
 *    TSpecialFolders.TempExe;                  // C:\Users\MCo\AppData\Local\Temp\ISL2\
 *                                              // C:\Users\MCo\AppData\Local\Temp\ENORM\
 *    TSpecialFolders.TempDll;                  // C:\Users\MCo\AppData\Local\Temp\SLC_A3\
 *                                              // C:\Users\MCo\AppData\Local\Temp\ENORM\
 *    TSpecialFolders.TempExeDll;               // C:\Users\MCo\AppData\Local\Temp\ISL2\SLC_A3\
 *                                              // C:\Users\MCo\AppData\Local\Temp\ENORM\
 *    TSpecialFolders.Windows;                  // C:\Windows\
 *    TSpecialFolders.System;                   // C:\Windows\system32\
 *    TSpecialFolders.AppData;                  // C:\Users\MCo\AppData\Roaming\
 *    TSpecialFolders.CommonAppData;            // C:\ProgramData\
 *                                              // C:\Documents and Settings\All Users\Application Data\
 *    TSpecialFolders.AppDataLocal;             // C:\Users\MCo\AppData\Local\
 *                                              // C:\Documents and Settings\MCo\Local Settings\Application Data\
 *    TSpecialFolders.ProgramFiles;             // C:\Program Files\          <= 64-bits EXE on 64-bits OS or 32-bits EXE on 32-bits OS
 *                                              // C:\Program Files (x86)\    <= 32-bits EXE on 64-bits OS
 *    TSpecialFolders.CommonProgramFiles;       // C:\Program Files (x86)\Common Files\
 *    TSpecialFolders.ProgramFilesX86;          // C:\Program Files (x86)\
 *    TSpecialFolders.CommonProgramFilesX86;    // C:\Program Files (x86)\Common Files\
 *    TSpecialFolders.Desktop;                  // C:\Users\MCo\Desktop\
 *    TSpecialFolders.CommonDesktop;            // C:\Users\Public\Desktop\
 *    TSpecialFolders.Documents;                // C:\Users\MCo\Documents\
 *    TSpecialFolders.CommonDocuments;          // C:\Users\Public\Documents\
 *    TSpecialFolders.Pictures;                 // C:\Users\MCo\Pictures\
 *    TSpecialFolders.CommonPictures;           // C:\Users\Public\Pictures\
 *    TSpecialFolders.Music;                    // C:\Users\MCo\Music\
 *    TSpecialFolders.CommonMusic;              // C:\Users\Public\Music\
 *    TSpecialFolders.Video;                    // C:\Users\MCo\Videos\
 *    TSpecialFolders.CommonVideo;              // C:\Users\Public\Videos\
 *    TSpecialFolders.AdminTools;               // C:\Users\MCo\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Administrative Tools\
 *    TSpecialFolders.CommonAdminTools;         // C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Administrative Tools\
 *    TSpecialFolders.Favorites;                // C:\Users\MCo\Favorites\
 *    TSpecialFolders.CommonFavorites;          // C:\Users\MCo\Favorites\
 *    TSpecialFolders.StartMenu;                // C:\Users\MCo\AppData\Roaming\Microsoft\Windows\Start Menu\
 *    TSpecialFolders.CommonStartMenu;          // C:\ProgramData\Microsoft\Windows\Start Menu\
 *    TSpecialFolders.StartMenuPrograms;        // C:\Users\MCo\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\
 *    TSpecialFolders.CommonStartMenuPrograms;  // C:\ProgramData\Microsoft\Windows\Start Menu\Programs\
 *    TSpecialFolders.StartUp;                  // C:\Users\MCo\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\
 *    TSpecialFolders.CommonStartUp;            // C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\
 *    TSpecialFolders.Templates;                // C:\Users\MCo\AppData\Roaming\Microsoft\Windows\Templates\
 *    TSpecialFolders.CommonTemplates;          // C:\ProgramData\Microsoft\Windows\Templates\
 *    TSpecialFolders.Cookies;                  // C:\Users\MCo\AppData\Roaming\Microsoft\Windows\Cookies\
 *    TSpecialFolders.DiscBurnArea;             // C:\Users\MCo\AppData\Local\Microsoft\Windows\Burn\Burn\
 *    TSpecialFolders.History;                  // C:\Users\MCo\AppData\Local\Microsoft\Windows\History\
 *    TSpecialFolders.NetHood;                  // C:\Users\MCo\AppData\Roaming\Microsoft\Windows\Network Shortcuts\
 *    TSpecialFolders.PrintHood;                // C:\Users\MCo\AppData\Roaming\Microsoft\Windows\Printer Shortcuts\
 *    TSpecialFolders.Recent;                   // C:\Users\MCo\AppData\Roaming\Microsoft\Windows\Recent\
 *    TSpecialFolders.SendTo;                   // C:\Users\MCo\AppData\Roaming\Microsoft\Windows\SendTo\
 *
 *
 * $Log: $
 *)
{$ENDREGION}
uses
  ShlObj, SHFolder;

type
  TSettingsFlag = (sfUserSpecific, sfMachineSpecific, sfExeSpecific, sfDllSpecific);
  TSettingsFlags = set of TSettingsFlag;

  TSpecialFolders = class
  private
    class procedure AddTempFolder(const TempDir: string);
    class procedure DeleteTempFolders;
    class function GetModulePathName(const hInst: Cardinal): string;

    class function GetExe: string; static; inline;
    class function GetExeDir: string; static; inline;
    class function GetExeBaseName: string; static; inline;
    class function GetDll: string; static; inline;
    class function GetDllDir: string; static; inline;
    class function GetDllBaseName: string; static; inline;
    class function GetWindowsDir: string; static;
    class function GetWindowsSysDir: string; static;
    class function GetTempDir(const Index: Integer): string; static;
    class function GetCSIDLDir(const CSIDL: Integer): string; static; // Constant Special Item ID List
  protected
    class function GetHToken: Cardinal; virtual;
  public
    class function Settings(Flags: TSettingsFlags = [sfUserSpecific, sfDllSpecific]): string;

    class property Exe: string                    read GetExeDir;
    class property ExeFullName: string            read GetExe;
    class property DLL: string                    read GetDllDir;
    class property DLLFullName: string            read GetDll;
    class property Temp: string           index 0 read GetTempDir;
    class property TempExe: string        index 1 read GetTempDir;
    class property TempDll: string        index 2 read GetTempDir;
    class property TempExeDll: string     index 3 read GetTempDir;

    class property Windows: string                read GetWindowsDir;
    class property System: string                 read GetWindowsSysDir;

    class property AppData: string                  index CSIDL_APPDATA                 read GetCSIDLDir;
    class property CommonAppData: string            index CSIDL_COMMON_APPDATA          read GetCSIDLDir;
    class property AppDataLocal: string             index CSIDL_LOCAL_APPDATA           read GetCSIDLDir;
    class property ProgramFiles: string             index CSIDL_PROGRAM_FILES           read GetCSIDLDir;
    class property ProgramFilesCommon: string       index CSIDL_PROGRAM_FILES_COMMON    read GetCSIDLDir;
    class property ProgramFilesX86: string          index CSIDL_PROGRAM_FILESX86        read GetCSIDLDir;
    class property ProgramFilesX86Common: string    index CSIDL_PROGRAM_FILES_COMMONX86 read GetCSIDLDir;

    class property Desktop: string                  index CSIDL_DESKTOPDIRECTORY        read GetCSIDLDir;
    class property CommonDesktop: string            index CSIDL_COMMON_DESKTOPDIRECTORY read GetCSIDLDir;
    class property Documents: string                index CSIDL_PERSONAL                read GetCSIDLDir;
    class property CommonDocuments: string          index CSIDL_COMMON_DOCUMENTS        read GetCSIDLDir;
    class property Pictures: string                 index CSIDL_MYPICTURES              read GetCSIDLDir;
    class property CommonPictures: string           index CSIDL_COMMON_PICTURES         read GetCSIDLDir;
    class property Music: string                    index CSIDL_MYMUSIC                 read GetCSIDLDir;
    class property CommonMusic: string              index CSIDL_COMMON_MUSIC            read GetCSIDLDir;
    class property Video: string                    index CSIDL_MYVIDEO                 read GetCSIDLDir;
    class property CommonVideo: string              index CSIDL_COMMON_VIDEO            read GetCSIDLDir;

    class property AdminTools: string               index CSIDL_ADMINTOOLS              read GetCSIDLDir;
    class property CommonAdminTools: string         index CSIDL_COMMON_ADMINTOOLS       read GetCSIDLDir;
    class property Favorites: string                index CSIDL_FAVORITES               read GetCSIDLDir;
    class property CommonFavorites: string          index CSIDL_COMMON_FAVORITES        read GetCSIDLDir;
    class property StartMenu: string                index CSIDL_STARTMENU               read GetCSIDLDir;
    class property CommonStartMenu: string          index CSIDL_COMMON_STARTMENU        read GetCSIDLDir;
    class property StartMenuPrograms: string        index CSIDL_PROGRAMS                read GetCSIDLDir;
    class property CommonStartMenuPrograms: string  index CSIDL_COMMON_PROGRAMS         read GetCSIDLDir;
    class property StartUp: string                  index CSIDL_STARTUP                 read GetCSIDLDir;
    class property CommonStartUp: string            index CSIDL_COMMON_STARTUP          read GetCSIDLDir;
    class property Templates: string                index CSIDL_TEMPLATES               read GetCSIDLDir;
    class property CommonTemplates: string          index CSIDL_COMMON_TEMPLATES        read GetCSIDLDir;

    class property Cookies: string                  index CSIDL_COOKIES                 read GetCSIDLDir;
    class property DiscBurnArea: string             index CSIDL_CDBURN_AREA             read GetCSIDLDir;
    class property History: string                  index CSIDL_HISTORY                 read GetCSIDLDir;
    class property NetHood: string                  index CSIDL_NETHOOD                 read GetCSIDLDir;
    class property PrintHood: string                index CSIDL_PRINTHOOD               read GetCSIDLDir;
    class property Recent: string                   index CSIDL_RECENT                  read GetCSIDLDir;
    class property SendTo: string                   index CSIDL_SENDTO                  read GetCSIDLDir;

    class property ByCSIDL[const CSIDL: Integer]: string                                read GetCSIDLDir;
  end;


////////////////////////////////////////////////////////////////////////////////////////////////////
implementation
uses
  Windows, SysUtils, ComObj, Classes, IOUtils;

var
  TempFolders: TStringList;

type
  TStringAPICallback = reference to function(lpBuffer: PChar; nSize: Cardinal): Cardinal;

{ ------------------------------------------------------------------------------------------------ }
function CallAPIStringFunction(const Callback: TStringAPICallback;
                               const InitialSize: integer = MAX_PATH): string;
var
  iSize, iResult, iError: integer;
begin
  iSize := InitialSize;
  repeat
    SetLength(Result, iSize);
    iResult := Callback(PChar(Result), iSize);
    iError := GetLastError;
    if iResult = 0 then begin
      if iError = ERROR_SUCCESS then begin
        Result := '';
        Exit;
      end else begin
        RaiseLastOSError;
      end;
    end else if iResult >= iSize then begin
      iSize := iResult + 1;
    end else begin
      SetLength(Result, iResult);
      Break;
    end;
  until iResult < iSize;
end {GetStringFromAPI};


{ ================================================================================================ }
{ TSpecialFolders }

{ ------------------------------------------------------------------------------------------------ }
class procedure TSpecialFolders.AddTempFolder(const TempDir: string);
begin
  if not DirectoryExists(TempDir) then begin
    if not Assigned(TempFolders) then
      TempFolders := TStringList.Create;
    TempFolders.Add(TempDir);
  end;
end {TSpecialFolders.AddTempFolder};

{ ------------------------------------------------------------------------------------------------ }
class procedure TSpecialFolders.DeleteTempFolders;
var
  i: Integer;
begin
  if Assigned(TempFolders) then begin
    for i := TempFolders.Count - 1 downto 0 do begin
      // remove all files first
      TDirectory.Delete(TempFolders[i], True);
      RemoveDir(TempFolders[i]);
    end;
    FreeAndNil(TempFolders);
  end;
end {TSpecialFolders.DeleteTempFolders};

{ ------------------------------------------------------------------------------------------------ }
class function TSpecialFolders.GetCSIDLDir(const CSIDL: Integer): string;
var
  Buffer: array[0..MAX_PATH] of Char;
  PBuffer: PChar;
begin
  PBuffer := PChar(@Buffer[0]);
  OleCheck(SHGetFolderPath(0, CSIDL or CSIDL_FLAG_CREATE, GetHToken, SHGFP_TYPE_CURRENT, PBuffer));
  Result := IncludeTrailingPathDelimiter(string(PBuffer));
end {TSpecialFolders.GetCSIDLDir};

{ ------------------------------------------------------------------------------------------------ }
class function TSpecialFolders.GetDll: string;
begin
  Result := GetModulePathName(HInstance);
end {TSpecialFolders.GetDll};

{ ------------------------------------------------------------------------------------------------ }
class function TSpecialFolders.GetDllDir: string;
begin
  Result := ExtractFilePath(GetDll);
end {TSpecialFolders.GetDllDir};

{ ------------------------------------------------------------------------------------------------ }
class function TSpecialFolders.GetDllBaseName: string;
begin
  Result := ChangeFileExt(ExtractFileName(GetDll), '');
end {TSpecialFolders.GetDllName};

{ ------------------------------------------------------------------------------------------------ }
class function TSpecialFolders.GetExe: string;
begin
  Result := GetModulePathName(0);
end {TSpecialFolders.GetExe};

{ ------------------------------------------------------------------------------------------------ }
class function TSpecialFolders.GetExeDir: string;
begin
  Result := ExtractFilePath(GetExe);
end {TSpecialFolders.GetExeDir};

{ ------------------------------------------------------------------------------------------------ }
class function TSpecialFolders.GetHToken: Cardinal;
begin
  Result := 0; // use `High(Cardinal)` instead of `0` for default user's paths
end {TSpecialFolders.GetHToken};

{ ------------------------------------------------------------------------------------------------ }
class function TSpecialFolders.GetExeBaseName: string;
begin
  Result := ChangeFileExt(ExtractFileName(GetExe), '');
end {TSpecialFolders.GetExeName};

{ ------------------------------------------------------------------------------------------------ }
class function TSpecialFolders.GetModulePathName(const hInst: Cardinal): string;
begin
  Result := CallAPIStringFunction(
              function(lpBuffer: PChar; nSize: Cardinal): Cardinal
              begin
                Result := GetModuleFileName(hInst, lpBuffer, nSize);
              end);
  if SameStr(Copy(Result, 1, 4), '\\?\') then
    Result := Copy(Result, 5);
end {TSpecialFolders.GetModulePathName};

{ ------------------------------------------------------------------------------------------------ }
class function TSpecialFolders.GetTempDir(const Index: Integer): string;
begin
  Result := CallAPIStringFunction(
              function(lpBuffer: PChar; nSize: Cardinal): Cardinal
              begin
                Result := GetTempPath(nSize, lpBuffer);
              end);
  if Index in [1, 3] then begin
    Result := Result + IncludeTrailingPathDelimiter(GetExeBaseName);
    AddTempFolder(Result);
  end;
  if (Index = 2) or ((Index = 3) and not SameFileName(GetExe, GetDll)) then begin
    Result := Result + IncludeTrailingPathDelimiter(GetDllBaseName);
    AddTempFolder(Result);
  end;
  ForceDirectories(Result);
end {TSpecialFolders.GetTempDir};

{ ------------------------------------------------------------------------------------------------ }
class function TSpecialFolders.GetWindowsDir: string;
begin
  Result := CallAPIStringFunction(
              function(lpBuffer: PChar; nSize: Cardinal): Cardinal
              begin
                Result := GetWindowsDirectory(lpBuffer, nSize);
              end);
  Result := IncludeTrailingPathDelimiter(Result);
end {TSpecialFolders.GetWindowsDir};

{ ------------------------------------------------------------------------------------------------ }
class function TSpecialFolders.GetWindowsSysDir: string;
begin
  Result := CallAPIStringFunction(
              function(lpBuffer: PChar; nSize: Cardinal): Cardinal
              begin
                Result := GetSystemDirectory(lpBuffer, nSize);
              end);
  Result := IncludeTrailingPathDelimiter(Result);
end {TSpecialFolders.GetWindowSysDir};

{ ------------------------------------------------------------------------------------------------ }
class function TSpecialFolders.Settings(Flags: TSettingsFlags): string;
var
  CSIDL: Integer;
begin
  if not (sfUserSpecific in Flags) then
    CSIDL := CSIDL_COMMON_APPDATA
  else if sfMachineSpecific in Flags then
    CSIDL := CSIDL_LOCAL_APPDATA
  else
    CSIDL := CSIDL_APPDATA
  ;
  Result := GetCSIDLDir(CSIDL);

  if (sfExeSpecific in Flags) then begin
    Result := Result + IncludeTrailingPathDelimiter(GetExeBaseName);
  end;
  if (sfDllSpecific in Flags) and not ((sfExeSpecific in Flags) and SameFilename(GetExe, GetDll)) then begin
    Result := Result + IncludeTrailingPathDelimiter(GetDllBaseName);
  end;
  ForceDirectories(Result);
end {TSpecialFolders.SettingsDir};


////////////////////////////////////////////////////////////////////////////////////////////////////
initialization

finalization
  TSpecialFolders.DeleteTempFolders;

end.
