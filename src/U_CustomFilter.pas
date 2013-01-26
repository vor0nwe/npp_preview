unit U_CustomFilter;

////////////////////////////////////////////////////////////////////////////////////////////////////
interface
uses
  Classes, SysUtils;

type
  {--- 2013-01-26 Martijn: TFilterData contains all the information needed for the filter to do its
                            job (in a different thread, so the filter processing becomes thread-safe). ---}
  TFilterData = record
    Name: string;
    DocFile: TFileName;
    Contents: string;
    CodePage: Integer;
    FilterInfo: TStringList; // the contents of the filter
  end;

type
  TCustomFilterThread = class(TThread)
  private
    FData: TFilterData;
    FHTML: string;
  public
    constructor Create(const Data: TFilterData); reintroduce;
    procedure Execute; override;
    property HTML: string read FHTML;
  end;


////////////////////////////////////////////////////////////////////////////////////////////////////
implementation

{ TCustomFilterThread }

{ ------------------------------------------------------------------------------------------------ }
constructor TCustomFilterThread.Create(const Data: TFilterData);
begin
  inherited Create(True);
  FData := Data;
end {TCustomFilter.Create};

{ ------------------------------------------------------------------------------------------------ }
procedure TCustomFilterThread.Execute;
type
  TContentInputType = (citStandardInput, citFileParameter);
  TContentOutputType = (cotStandardOutput, cotInputFile, cotOutputFile);
var
  Command: string;
  DocFile, InFile, OutFile, Ext: TFileName;
  SS: TStringStream;
  Start: TDateTime;
  ContentInput: TContentInputType;
  ContentOutput: TContentOutputType;
begin
  inherited;

  Command := FData.FilterInfo.Values['Command'];
  if Command = '' then begin
    FHTML := FData.Contents;
    Exit;
  end;

    (*
    UseTempFile := (SCI_GETMODIFY <> 0) or (ContentOutput = cotInputFile);

    if ContentInput = citStandardInput then begin

      // TODO: run the command, and pass the contents to the standard input stream of the new process.

      if ContentOutput = cotInputFile then begin
        // TODO: send out a warning to the user that the thing's been misconfigured
        ContentOutput := cotStandardOutput;
      end;
      if ContentOutput = cotStandardOutput then begin
        // TODO: read the output from the process's output stream
      end else if ContentOutput = cotOutputFile then begin
        // TODO: read the output from the designated output file
      end else begin
        // TODO: emit warning to the user that this filter's been misconfigured
      end;
    end else begin // ContentInput = citFile
      if (SCI_GETMODIFY <> 0) or (ContentOutput = cotInputFile) then begin
        // TODO: save the contents to a temp input file
      end else begin
        // TODO: use the original file as input file
      end;
      if (ContentOutput = cotOutputFile) then begin
        // TODO: prepare a temp file as output file
      end;

      // TODO: Run the command

      if (ContentOutput = cotStandardOutput) then begin
        // TODO: read the output from the process's standard output stream
      end else begin
        // TODO: read the output from the output file
        // TODO: if not SameFile(DocFile, OutputFile) then DeleteFile(OutputFile);
      end;
    end;
    *)


(*    if SendMessage(hScintilla, SCI_GETMODIFY, 0, 0) <> 0 then begin
      // The document is modified, so we’ll need to save it to a temp file, and pass that along
      InFile := TPath.GetTempFileName;
      Ext := ExtractFileExt(DocFile);
      if Ext <> '' then begin
        TFile.Move(InFile, ChangeFileExt(InFile, Ext));
        InFile := ChangeFileExt(InFile, Ext);
      end;
      SS := TStringStream.Create(HTML, SendMessage(hScintilla, SCI_GETCODEPAGE, 0, 0));
      try
        SS.SaveToFile(InFile);
      finally
        SS.Free;
      end;
    end else begin
      // The document is unmodified, so we can just reference the original file
      InFile := DocFile;
    end;
    try
      {$MESSAGE HINT 'TODO: substitute the temp file name, execute the command with pipe, and read the output — MCO 22-01-2013'}
      {$MESSAGE HINT 'TODO: perhaps we could wait for the result in a different thread, so as not to block Notepad++ — MCO 22-01-2013'}

      {$MESSAGE WARN 'TODO: REPLACE THIS TEMPORARY CODE!!! — MCO 22-01-2013'}
      Command := StringReplace(Command, '%1', '"' + InFile + '"', [rfReplaceAll]);
      OutFile := TPath.GetTempFileName;
      TFile.Delete(OutFile);
//ShowMessage(Format('InFile: "%s"; Command: "%s", OutFile: "%s"', [InFile, Command, OutFile]));
      ShellExecute(Npp.NppData.NppHandle, nil, 'cmd.exe', PChar('/c ' + Command + ' > "' + OutFile + '"'), PChar(ExtractFilePath(OutFile)), SW_HIDE);

      Start := Now;
      repeat
        Sleep(500);
      until FileExists(OutFile) or (Now - Start > 10 * (1 / 86400));
      {$MESSAGE WARN 'TODO: /REPLACE THIS TEMPORARY CODE!!! — MCO 22-01-2013'}

      SS := TStringStream.Create('', SendMessage(hScintilla, SCI_GETCODEPAGE, 0, 0));
      try
        if FileExists(OutFile) then begin
          SS.LoadFromFile(OutFile);
          try TFile.Delete(OutFile); except end;
        end;
        Result := SS.DataString;
      finally
        SS.Free;
      end;
      {--- MCO 22-01-2013: TODO: ways of transferring the HTML to the filter exe: on the input stream; write to file and pass the file name ---}
      {--- MCO 22-01-2013: TODO: ways of transferring the result from the filter exe: read from the output stream, read the changed input file, or read a (separate) output file ---}
    finally
      if InFile <> DocFile then begin
        try TFile.Delete(InFile); except end;
      end;
    end;
  finally
    Filters.Free;
  end;
*)


  {--- 2012-01-26 Martijn: When finished, we need to populate the webbrowser component, but this
                  has to happen in the main thead, so that will have to be done in the OnTerminate. ---}
end {TCustomFilter.Execute};

end.
