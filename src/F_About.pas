unit F_About;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, NppForms, StdCtrls, ExtCtrls;

type
  TAboutForm = class(TNppForm)
    btnOK: TButton;
    lblBasedOn: TLabel;
    lblTribute: TLinkLabel;
    lblPlugin: TLabel;
    lblAuthor: TLinkLabel;
    lblVersion: TLabel;
    lblURL: TLinkLabel;
    lblIEVersion: TLinkLabel;
    procedure FormCreate(Sender: TObject);
    procedure lblLinkClick(Sender: TObject; const Link: string; LinkType: TSysLinkType);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AboutForm: TAboutForm;

implementation
uses
  ShellAPI, StrUtils,
  IdURI,
  L_VersionInfoW, L_SpecialFolders, WebBrowser,
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

  lblIEVersion.Caption := Format(lblIEVersion.Caption, [GetIEVersion]);
end {TAboutForm.FormCreate};

{ ------------------------------------------------------------------------------------------------ }
procedure TAboutForm.lblLinkClick(Sender: TObject; const Link: string; LinkType: TSysLinkType);
var
  URL, Subject: string;
begin
  if LinkType = sltURL then begin
    URL := Link;
    if StartsText('http', URL) then begin
      with TFileVersionInfo.Create(TSpecialFolders.DLLFullName) do begin
        URL := URL + '?client=' + TIdURI.ParamsEncode(Format('%s/%s', [Internalname, FileVersion])) +
                      '&version=' + TIdURI.ParamsEncode(Format('%d.%d.%d.%d', [MajorVersion, MinorVersion, Revision, Build]));
        Free;
      end;
    end else if StartsText('mailto:', URL) then begin
      with TFileVersionInfo.Create(TSpecialFolders.DLLFullName) do begin
        Subject := Self.Caption + Format(' v%s (%d.%d.%d.%d)', [FileVersion, MajorVersion, MinorVersion, Revision, Build]);
        Free;
      end;
      URL := URL + '?subject=' + TIdURI.ParamsEncode(Subject);
    end;
    ShellExecute(Self.Handle, nil, PChar(URL), nil, nil, SW_SHOWDEFAULT);
  end;
end {TAboutForm.lblLinkClick};

end.
