unit F_About;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, NppForms, StdCtrls;

type
  TAboutForm = class(TNppForm)
    Button1: TButton;
    Label1: TLabel;
    Label2: TLabel;
    lblPlugin: TLabel;
    lblAuthor: TLabel;
    txtAuthor: TStaticText;
    lblVersion: TLabel;
    procedure txtAuthorClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
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
  L_VersionInfoW, L_SpecialFolders;

{$R *.dfm}

{ ------------------------------------------------------------------------------------------------ }
procedure TAboutForm.FormCreate(Sender: TObject);
begin
  with TFileVersionInfo.Create(TSpecialFolders.DLLFullName) do begin
    lblVersion.Caption := Format('v%s (%d.%d.%d.%d)', [FileVersion, MajorVersion, MinorVersion, Revision, Build]);
    Free;
  end;
  txtAuthor.Font.Color := clHotLight;
  txtAuthor.Font.Style := txtAuthor.Font.Style + [fsUnderline];
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

end.
