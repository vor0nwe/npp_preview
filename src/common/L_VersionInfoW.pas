unit L_VersionInfoW;

interface

uses
  Windows, SysUtils;

type
  TFileVersionInfo = class
  private
    { Private declarations }
    FFilename         : WideString;
    FHasVersionInfo   : boolean;

    FCompanyName      : WideString;
    FFileDescription  : WideString;
    FFileVersion      : WideString;
    FInternalname     : WideString;
    FLegalCopyright   : WideString;
    FLegalTradeMarks  : WideString;
    FOriginalFilename : WideString;
    FProductName      : WideString;
    FProductVersion   : WideString;
    FComments         : WideString;
    FMajorVersion     : Word;
    FMinorVersion     : Word;
    FRevision          : Word;
    FBuild            : Word;

    procedure SetFileName(AFileName: WideString);
  protected
    { Protected declarations }
  public
    { Public declarations }
    constructor Create(AFileName: WideString);
    destructor  Destroy; override;

    property FileName         : WideString read FFileName           write SetFileName;
  public
    { Published declarations }
    property CompanyName      : WideString  read FCompanyName;
    property FileDescription  : WideString  read FFileDescription;
    property FileVersion      : WideString  read FFileVersion;
    property Internalname     : WideString  read FInternalname;
    property LegalCopyright   : WideString  read FLegalCopyright;
    property LegalTradeMarks  : WideString  read FLegalTradeMarks;
    property OriginalFilename : WideString  read FOriginalFilename;
    property ProductName      : WideString  read FProductName;
    property ProductVersion   : WideString  read FProductVersion;
    property Comments         : WideString  read FComments;
    property MajorVersion     : Word        read FMajorVersion;
    property MinorVersion     : Word        read FMinorVersion;
    property Revision         : Word        read FRevision;
    property Build            : Word        read FBuild;
  end;

implementation

type
  TLangAndCP = record
    wLanguage : word;
    wCodePage : word;
  end;
  PLangAndCP = ^TLangAndCP;

constructor TFileVersionInfo.Create(AFileName: WideString);
begin
  inherited Create;
  SetFileName(AFileName);
end;

destructor TFileVersionInfo.Destroy;
begin
  inherited Destroy;
end;

procedure TFileVersionInfo.SetFileName(AFileName: WideString);
var
  Dummy     : cardinal;
  BufferSize: integer;
  Buffer    : Pointer;
  Lang      : PLangAndCP;
  SubBlock  : WideString;
  InfoBlock : VS_FIXEDFILEINFO;
  InfoPtr   : Pointer;
  function QueryValue(AName: WideString): WideString;
  var
    Value   : PWChar;
  begin
    SubBlock := WideFormat('\\StringFileInfo\\%.4x%.4x\\%s', [Lang.wLanguage, Lang.wCodePage, AName]);
    VerQueryValueW(Buffer, PWChar(SubBlock), Pointer(Value), Dummy);
    Result := WideString(Value);
  end;
begin
  FFilename := AFileName;

  BufferSize := GetFileVersionInfoSizeW(PWChar(AFileName), Dummy);
  FHasVersionInfo := (Buffersize > 0);
  if BufferSize > 0 then begin
    Buffer := AllocMem(BufferSize);
    try
      GetFileVersionInfoW(PWChar(AFileName),0,BufferSize,Buffer);

      SubBlock := '\\VarFileInfo\\Translation';
      VerQueryValueW(Buffer, PWChar(SubBlock), Pointer(Lang), Dummy);

      FCompanyName      := QueryValue('CompanyName');
      FFileDescription  := QueryValue('FileDescription');
      FFileVersion      := QueryValue('FileVersion');
      FInternalName     := QueryValue('InternalName');
      FLegalCopyright   := QueryValue('LegalCopyright');
      FLegalTradeMarks  := QueryValue('LegalTradeMarks');
      FOriginalFilename := QueryValue('OriginalFilename');
      FProductName      := QueryValue('ProductName');
      FProductVersion   := QueryValue('ProductVersion');
      FComments         := QueryValue('Comments');

      VerQueryValue(Buffer, '\', InfoPtr, Dummy);
      Move(InfoPtr^, InfoBlock, SizeOf(VS_FIXEDFILEINFO));
      FMajorVersion := InfoBlock.dwFileVersionMS shr 16;
      FMinorVersion := InfoBlock.dwFileVersionMS and 65535;
      FRevision     := InfoBlock.dwFileVersionLS shr 16;
      FBuild        := InfoBlock.dwFileVersionLS and 65535;
    finally
      FreeMem(Buffer,BufferSize);
    end;
  end
  else begin
    FCompanyname      := '';
    FFileDescription  := '';
    FFileVersion      := '';
    FInternalname     := '';
    FLegalCopyright   := '';
    FLegalTradeMarks  := '';
    FOriginalFilename := '';
    FProductName      := '';
    FProductVersion   := '';
    FComments         := '';
    FMajorVersion     := 0;
    FMinorVersion     := 0;
    FRevision         := 0;
    FBuild            := 0;
  end;
end;


end.

