unit F_About;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, NppForms, StdCtrls;

type
  TAboutForm = class(TNppForm)
    btnOK: TButton;
    Label1: TLabel;
    Label2: TLabel;
    lblPlugin: TLabel;
    lblAuthor: TLabel;
    txtAuthor: TStaticText;
    lblVersion: TLabel;
    txtURL: TStaticText;
    procedure txtAuthorClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure txtURLClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AboutForm: TAboutForm;

implementation
uses
  ShellAPI,
  IdURI,
  L_VersionInfoW, L_SpecialFolders,
  NppPlugin;

{$R *.dfm}

{ ------------------------------------------------------------------------------------------------ }
procedure TAboutForm.FormCreate(Sender: TObject);
var
  WR, BR: TRect;
begin
  // Position the form centered on Notepad++
  if GetWindowRect(Npp.NppData.NppHandle, WR) then begin
    BR.Left := WR.Left + (WR.Width div 2) - (Self.Width div 2);
    BR.Width := Self.Width;
    BR.Top := WR.Top + (WR.Height div 2) - (Self.Height div 2);
    BR.Height := Self.Height;
    Self.BoundsRect := BR;
  end;

  with TFileVersionInfo.Create(TSpecialFolders.DLLFullName) do begin
    lblVersion.Caption := Format('v%s (%d.%d.%d.%d)', [FileVersion, MajorVersion, MinorVersion, Revision, Build]);
    Free;
  end;

  txtAuthor.Font.Color := clHotLight;
  txtAuthor.Font.Style := txtAuthor.Font.Style + [fsUnderline];
  txtURL.Font.Color := clHotLight;
  txtURL.Font.Style := txtAuthor.Font.Style + [fsUnderline];
end {TAboutForm.FormCreate};

{ ------------------------------------------------------------------------------------------------ }
procedure TAboutForm.txtAuthorClick(Sender: TObject);
var
  Subject: string;
begin
  with TFileVersionInfo.Create(TSpecialFolders.DLLFullName) do begin
    Subject := Self.Caption + Format(' v%s (%d.%d.%d.%d)', [FileVersion, MajorVersion, MinorVersion, Revision, Build]);
    Free;
  end;
  Subject := TIdURI.ParamsEncode(Subject);
  ShellExecute(Self.Handle, nil, PChar('mailto:vor0nwe@users.sf.net?subject=' + Subject), nil, nil, SW_SHOWDEFAULT);
end {TAboutForm.txtAuthorClick};

{ ------------------------------------------------------------------------------------------------ }
procedure TAboutForm.txtURLClick(Sender: TObject);
var
  Params: string;
begin
  with TFileVersionInfo.Create(TSpecialFolders.DLLFullName) do begin
    Params := '?client=' + TIdURI.ParamsEncode(Format('%s/%s', [Internalname, FileVersion])) +
                '&version=' + TIdURI.ParamsEncode(Format('%d.%d.%d.%d', [MajorVersion, MinorVersion, Revision, Build]));
    Free;
  end;
  ShellExecute(Self.Handle, nil, PChar(TStaticText(Sender).Caption + Params), nil, nil, SW_SHOWDEFAULT);
end {TAboutForm.txtURLClick};

end.
