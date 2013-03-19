unit L_HttpClient;

interface
uses
  SysUtils, Classes, WinInet;

type
  THttpConnectEnum = (
    hcDirect          = INTERNET_OPEN_TYPE_DIRECT,
    hcPreconfig       = INTERNET_OPEN_TYPE_PRECONFIG,
    hcPreconfigNoAuto = INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY,
    hcProxy           = INTERNET_OPEN_TYPE_PROXY
  );

  THttpClient = class;
  TProgressNotify = procedure(Sender: THttpClient; Msg: string; BytesRead, TotalBytes: Cardinal; out Cancel: Boolean) of object;

  THttpClient = class
    private
      FhReq: HINTERNET;
      FOnProgress: TProgressNotify;
      FURL: string;

      function  SendRequestAndGetResponse(const Verb, Resource: string;
                                          const Data: Pointer; const DataSize: Cardinal;
                                          Flags: Cardinal = 0): integer;
      procedure ReadResponseHeaders;
    protected
      FPath: string;
      FRequestHeaders: TStringList;
      FResponseStatus: cardinal;
      FResponseStatusText: string;
      FResponseHeaders: TStringList;
      FResponseStream: TMemoryStream;

      function QueryInfoString(const InfoFlag: cardinal): string; overload; inline;
      function QueryInfoString(const InfoFlag: cardinal; out Value: string): boolean; overload;
      function QueryInfoInteger(const InfoFlag: cardinal): cardinal; overload; inline;
      function QueryInfoInteger(const InfoFlag: cardinal; out Value: cardinal): Boolean; overload; inline;

      function GetResponseString: string; // RawString
    public
      UserAgent: string;
      Hostname: string;
      Port: integer;
      UseProxy: THttpConnectEnum;
      Proxy, ProxyByPass: string;
      UserName: string;
      Password: string;
      ConnectTimeout, SendTimeout, ReceiveTimeout: Cardinal;

      constructor Create(const URL: string); overload;
      constructor Create(const Hostname, Path: string; const Port: integer = INTERNET_DEFAULT_HTTP_PORT); overload;
      destructor  Destroy; override;

      function Head(Path: string = ''; const BypassCache: Boolean = False): integer;
      function Get(Path: string = ''; const BypassCache: Boolean = False): integer;
      function Post(Data: string): integer; overload;
      function Post(Data: TStream): integer; overload;
      function Post(Path: string; Data: string): integer; overload;
      function Post(Path: string; Data: TStream): integer; overload;

      property Path: string                   read FPath;
      property RequestHeaders: TStringList    read FRequestHeaders;
      property StatusCode: cardinal           read FResponseStatus;
      property StatusText: string             read FResponseStatusText;
      property ResponseHeaders: TStringList   read FResponseHeaders;
      property ResponseStream: TMemoryStream  read FResponseStream;
      property ResponseString: string         read GetResponseString; // RawString
      property URL: string                    read FURL;

      property OnProgress: TProgressNotify    read FOnProgress  write FOnProgress;
  end;

type
  EHttpError = class(Exception);


implementation
uses
  Windows, Math;

const
  READBUFFERSIZE = 4096;
  ERROR_INSUFFICIENT_BUFFER = 122;   { dderror }

resourcestring
  scConnecting =  'Connecting';
  scUploading  =  'Sending';
  scDownloading = 'Receiving';
  scFinished    = 'Done';

{ ================================================================================================ }
{ THttpClient }

{ ------------------------------------------------------------------------------------------------ }
constructor THttpClient.Create(const URL: string);
var
  UC: URL_COMPONENTS;
begin
  UC.dwStructSize := sizeof(UC);
  UC.lpszScheme := nil;
  UC.dwSchemeLength := 0;
  UC.nScheme := INTERNET_SCHEME_UNKNOWN;
  UC.lpszHostName := PChar(StringOfChar(#0, Length(URL) + 1));
  UC.dwHostNameLength := Length(URL) + 1;
  UC.nPort := INTERNET_INVALID_PORT_NUMBER;
  UC.lpszUserName := PChar(StringOfChar(#0, Length(URL) + 1));
  UC.dwUserNameLength := Length(URL) + 1;
  UC.lpszPassword := PChar(StringOfChar(#0, Length(URL) + 1));
  UC.dwPasswordLength := Length(URL) + 1;
  UC.lpszUrlPath := PChar(StringOfChar(#0, Length(URL) + 1));
  UC.dwUrlPathLength := Length(URL) + 1;
  UC.lpszExtraInfo := nil;
  UC.dwExtraInfoLength := 0;

  Win32Check(InternetCrackUrl(PChar(URL), 0, 0, UC));
  if not UC.nScheme in [INTERNET_SCHEME_HTTP, INTERNET_SCHEME_HTTPS, INTERNET_SCHEME_FILE] then begin
    raise EHttpError.CreateFmt('Unsupported scheme used in URL "%s". Only http:// and https:// are supported', [URL]);
  end;
  Create(UC.lpszHostName, UC.lpszUrlPath, UC.nPort);
  Self.UserName := UC.lpszUserName;
  Self.Password := UC.lpszPassword;
end {THttpClient.Create};
{ ------------------------------------------------------------------------------------------------ }
constructor THttpClient.Create(const Hostname, Path: string; const Port: integer);
var
  UC: URL_COMPONENTS;
  Buffer: string;
  Size: Cardinal;
begin
  FPath := Path;
  Self.UserAgent := ChangeFileExt(ExtractFileName(ParamStr(0)), '');
  Self.Hostname := Hostname;
  Self.Port := Port;
  Self.UseProxy := hcPreconfig;
  FRequestHeaders := TStringList.Create;
  FRequestHeaders.NameValueSeparator := ':';
  ConnectTimeout := 0;
  SendTimeout := 0;
  ReceiveTimeout := 0;

  Size := INTERNET_MAX_URL_LENGTH;
  SetLength(Buffer, Size);
  if InternetCreateUrl(UC, ICU_ESCAPE, PChar(Buffer), Size) then begin
    FURL := Copy(Buffer, 1, Size);
  end else if GetLastError = ERROR_INSUFFICIENT_BUFFER then begin
    SetLength(Buffer, Size);
    if InternetCreateUrl(UC, ICU_ESCAPE, PChar(Buffer), Size) then begin
      FURL := Copy(Buffer, 1, Size);
    end;
  end;
end {THttpClient.Create};
{ ------------------------------------------------------------------------------------------------ }
destructor THttpClient.Destroy;
begin
  if Assigned(FResponseStream) then FreeAndNil(FResponseStream);
  if Assigned(FResponseHeaders) then FreeAndNil(FResponseHeaders);
  FRequestHeaders.Free;
  inherited;
end {THttpClient.Destroy};

{ ------------------------------------------------------------------------------------------------ }
function THttpClient.Head(Path: string = ''; const BypassCache: Boolean = False): integer;
begin
  if Path = '' then begin
    Path := FPath;
  end;
  if BypassCache then begin
    Result := SendRequestAndGetResponse('HEAD', Path, nil, 0, INTERNET_FLAG_RELOAD);
  end else begin
    Result := SendRequestAndGetResponse('HEAD', Path, nil, 0);
  end;
end {THttpClient.Head};

{ ------------------------------------------------------------------------------------------------ }
function THttpClient.Get(Path: string = ''; const BypassCache: Boolean = False): integer;
begin
  if Path = '' then begin
    Path := FPath;
  end;
  if BypassCache then begin
    Result := SendRequestAndGetResponse('GET', Path, nil, 0, INTERNET_FLAG_RELOAD);
  end else begin
    Result := SendRequestAndGetResponse('GET', Path, nil, 0);
  end;
end {THttpClient.Get};

{ ------------------------------------------------------------------------------------------------ }
function THttpClient.Post(Data: TStream): integer;
begin
  Result := Self.Post(FPath, Data);
end {THttpClient.Post};
{ ------------------------------------------------------------------------------------------------ }
function THttpClient.Post(Path: string; Data: TStream): integer;
var
  MemData: TMemoryStream;
  CleanUp: boolean;
begin
  if Data is TMemoryStream then begin
    MemData := TMemoryStream(Data);
    CleanUp := False;
  end else begin
    Data.Position := 0;
    MemData := TMemoryStream.Create;
    MemData.CopyFrom(Data, Data.Size);
    CleanUp := True;
  end;
  try
    Result := SendRequestAndGetResponse('POST', Path, MemData.Memory, MemData.Size);
  finally
    if CleanUp then begin
      MemData.Free;
    end;
  end;
end {THttpClient.Post};
{ ------------------------------------------------------------------------------------------------ }
function THttpClient.Post(Data: string): integer;
begin
  Result := Self.Post(FPath, Data);
end {THttpClient.Post};
{ ------------------------------------------------------------------------------------------------ }
function THttpClient.Post(Path: string; Data: string): integer;
begin
  Result := SendRequestAndGetResponse('POST', Path, PChar(Data), ByteLength(Data));
end {THttpClient.Post};

{ ------------------------------------------------------------------------------------------------ }
function THttpClient.SendRequestAndGetResponse(const Verb, Resource: string;
                                               const Data: Pointer;
                                               const DataSize: Cardinal;
                                               Flags: Cardinal = 0): integer;
var
  hInt, hConn: HINTERNET;
  PProxy, PProxyBypass, PUserName, PPassword: PChar;
  BytesRead, TotalBytesToRead, TotalBytesRead: Cardinal;
  Buffer: TMemoryStream;
  Size: Cardinal;
  Cancel: Boolean;
begin
  Result := 0; // means the request has been canceled

  FPath := Resource;

  { Initialize the response data }
  FResponseStatus := 0;
  FResponseStatusText := '';
  if Assigned(FResponseHeaders) then FreeAndNil(FResponseHeaders);
  if Assigned(FResponseStream) then FreeAndNil(FResponseStream);

  { Set the proxy parameters }
  if UseProxy = hcProxy then begin
    if Length(Self.Proxy) = 0 then begin
      PProxy := nil;
    end else begin
      PProxy := PChar(Self.Proxy);
    end;
    if Length(Self.ProxyByPass) = 0 then begin
      PProxyByPass := nil;
    end else begin
      PProxyByPass := PChar(Self.ProxyByPass);
    end;
  end else begin
    PProxy := nil;
    PProxyBypass := nil;
  end;

  Cancel := False;
  if Assigned(FOnProgress) then begin
    FOnProgress(Self, scConnecting, 0, 0, Cancel);
    if Cancel then
      Exit;
  end;

  { Open an internet handle and set the request parameters }
  hInt := InternetOpen(PChar(UserAgent), Ord(UseProxy), PProxy, PProxyBypass, 0);
  Win32Check(Assigned(hInt));
  try
    if ConnectTimeout > 0 then InternetSetOption(hInt, INTERNET_OPTION_CONNECT_TIMEOUT, @ConnectTimeout, sizeof(ConnectTimeout));
    if SendTimeout > 0    then InternetSetOption(hInt, INTERNET_OPTION_SEND_TIMEOUT,    @SendTimeout,    sizeof(SendTimeout));
    if ReceiveTimeout > 0 then InternetSetOption(hInt, INTERNET_OPTION_RECEIVE_TIMEOUT, @ReceiveTimeout, sizeof(ReceiveTimeout));

    if Length(Self.UserName) = 0 then begin
      PUserName := nil;
    end else begin
      PUserName := PChar(Self.UserName);
    end;
    if Length(Self.Password) = 0 then begin
      PPassword := nil;
    end else begin
      PPassword := PChar(Self.Password);
    end;

    { Try to connect, then send the request }
    hConn := InternetConnect(hInt, PChar(Hostname), Self.Port, PUserName, PPassword, INTERNET_SERVICE_HTTP, 0, 1);
    Win32Check(Assigned(hConn));
    try
      if Self.Port = INTERNET_DEFAULT_HTTPS_PORT then begin
        Flags := Flags or INTERNET_FLAG_SECURE or INTERNET_FLAG_IGNORE_CERT_CN_INVALID or INTERNET_FLAG_IGNORE_CERT_DATE_INVALID;
      end;
      FhReq := HttpOpenRequest(hConn, PChar(Verb), PChar(Resource), nil, nil, nil, Flags, 1);
      Win32Check(Assigned(FhReq));
      try
        InternetQueryOption(FhReq, INTERNET_OPTION_URL, nil, Size);
        FURL := StringOfChar(#0, Size);
        Win32Check(InternetQueryOption(FhReq, INTERNET_OPTION_URL, PChar(FURL), Size));
        SetLength(FURL, Size);

        if Assigned(FOnProgress) then begin
          FOnProgress(Self, scUploading, 0, 0, Cancel);
          if Cancel then
            Exit;
        end;

        Win32Check(HttpSendRequest(FhReq,
                                   PChar(FRequestHeaders.Text), Length(FRequestHeaders.Text),
                                   Data, DataSize));

        { Try to get the content-length }
        if not QueryInfoInteger(HTTP_QUERY_CONTENT_LENGTH, TotalBytesToRead) then
          TotalBytesToRead := 0;

        FResponseStream := TMemoryStream.Create;
        Buffer := TMemoryStream.Create;
        try
          TotalBytesRead := 0;
          Buffer.SetSize(READBUFFERSIZE);
          repeat
            Sleep(0);
            if not InternetReadFile(FhReq, Buffer.Memory, READBUFFERSIZE, BytesRead) then begin
              Break;
            end;
            if BytesRead > 0 then begin
              Buffer.Position := 0;
              FResponseStream.CopyFrom(Buffer, Min(BytesRead, READBUFFERSIZE));
              Inc(TotalBytesRead, BytesRead);
              if Assigned(FOnProgress) then begin
                FOnProgress(Self, scDownloading, TotalBytesRead, TotalBytesToRead, Cancel);
                if Cancel then Exit;
              end;
            end;
          until (BytesRead = 0);
        finally
          Buffer.Free;
          FResponseStream.Position := 0;
        end;

        { Read the response status and headers }
        FResponseStatus := QueryInfoInteger(HTTP_QUERY_STATUS_CODE);
        FResponseStatusText := QueryInfoString(HTTP_QUERY_STATUS_TEXT);
        ReadResponseHeaders;

      finally
        InternetCloseHandle(FhReq);
      end;
    finally
      InternetCloseHandle(hConn);
    end;
  finally
    InternetCloseHandle(hInt);
  end;

  if Assigned(FOnProgress) then
    FOnProgress(Self, scFinished, TotalBytesRead, TotalBytesToRead, Cancel);

  Result := FResponseStatus;
end {THttpClient.SendRequestAndGetResponse};


{ ------------------------------------------------------------------------------------------------ }
function THttpClient.QueryInfoInteger(const InfoFlag: cardinal): cardinal;
begin
  Win32Check(QueryInfoInteger(InfoFlag, Result));
end {THttpClient.QueryInfoInteger};
{ ------------------------------------------------------------------------------------------------ }
function THttpClient.QueryInfoInteger(const InfoFlag: cardinal; out Value: cardinal): boolean;
var
  Index, BufSize: Cardinal;
begin
  Index := 0;
  bufsize := sizeOf(Value);
  Result := HttpQueryInfo(FhReq, InfoFlag or HTTP_QUERY_FLAG_NUMBER, @Value, BufSize, Index);
end {THttpClient.QueryInfoInteger};

{ ------------------------------------------------------------------------------------------------ }
function THttpClient.QueryInfoString(const InfoFlag: cardinal): string;
begin
  Win32Check(QueryInfoString(InfoFlag, Result));
end {THttpClient.QueryInfoString};
{ ------------------------------------------------------------------------------------------------ }
function THttpClient.QueryInfoString(const InfoFlag: cardinal; out Value: string): Boolean;
var
  BufSize, Index: cardinal;
  i: byte;
begin
  Result := True;
  Index := 0;
  bufsize := READBUFFERSIZE;
  for i := 0 to 1 do begin
    SetLength(Value, bufsize div SizeOf(Char));
    if not HttpQueryInfo(FhReq, InfoFlag, PChar(Value), bufsize, Index) then begin
      case GetLastError of
        ERROR_INSUFFICIENT_BUFFER: begin
          // bufsize now contains the required size; just go to the next iteration.
        end;
        ERROR_HTTP_HEADER_NOT_FOUND: begin
          // no more headers; just return the current result
          SetLength(Value, bufsize div SizeOf(Char));
          Break;
        end;
        else begin
          Result := False;
          Break;
        end;
      end;
    end else begin
      // HttpQueryInfo was successful
      SetLength(Value, bufsize div SizeOf(Char));
      Break;
    end;
  end {for};
end {THttpClient.QueryInfoString};

{ ------------------------------------------------------------------------------------------------ }
procedure THttpClient.ReadResponseHeaders;
var
  Headers, Header: string;
  i, Offset: integer;
begin
  Headers := QueryInfoString(HTTP_QUERY_RAW_HEADERS);
  Header := StringOfChar(#0, Length(Headers));
  FResponseHeaders := TStringList.Create;
  FResponseHeaders.NameValueSeparator := ':';
  FResponseHeaders.CaseSensitive := False;

  Offset := 0;
  for i := 1 to Length(Headers) - 1 do begin
    if Headers[i] = #0 then begin
      SetLength(Header, i - Offset - 1);
      if Length(Header) = 0 then begin
        Break;
      end;
      FResponseHeaders.Add(Header);
      Header := StringOfChar(#0, Length(Headers) - i);
      Offset := i;
    end else begin
      Header[i - Offset] := Headers[i];
    end;
  end;
  if (Length(Header) > 0) and (Header[1] <> #0) then begin
    Offset := Pos(#0, Header);
    if Offset > 0 then begin
      FResponseHeaders.Add(Copy(Header, 1, Offset - 1));
    end else begin
      FResponseHeaders.Add(Header);
    end;
  end;
end {THttpClient.ReadResponseHeaders};

{ ------------------------------------------------------------------------------------------------ }
function THttpClient.GetResponseString: string; // RawString
var
  ContentType: string;
  CharPos: Integer;
  Encoding: TEncoding;
  SS: TStringStream;
begin
  if Assigned(FResponseStream) and (FResponseStream.Size > 0) then begin
//    SetLength(Result, FResponseStream.Size div StringElementSize(Result));
//    CopyMemory(@Result[1], FResponseStream.Memory, FResponseStream.Size);
    ContentType := ResponseHeaders.Values['Content-Type'];
    CharPos := Pos('charset=', ContentType);
    if CharPos > 0 then begin
      Encoding := TEncoding.GetEncoding(Copy(ContentType, CharPos + 8));
    end else begin
      Encoding := TEncoding.Default;
    end;
    SS := TStringStream.Create('', Encoding, True);
    try
      SS.CopyFrom(FResponseStream, 0);
      Result := SS.DataString;
    finally
      SS.Free;
    end;
  end else begin
    Result := '';
  end;
end {THttpClient.GetResponseString};

end.
