inherited AboutForm: TAboutForm
  BorderIcons = []
  BorderStyle = bsSizeable
  Caption = 'About Preview HTML'
  ClientHeight = 256
  ParentFont = True
  Position = poDesigned
  OnCreate = FormCreate
  ExplicitWidth = 351
  ExplicitHeight = 294
  PixelsPerInch = 96
  TextHeight = 13
  object lblBasedOn: TLabel
    Left = 8
    Top = 96
    Width = 152
    Height = 13
    Caption = 'Based on the example plugin by'
  end
  object lblPlugin: TLabel
    Left = 8
    Top = 8
    Width = 175
    Height = 13
    Caption = 'HTML Preview plugin for Notepad++'
    ShowAccelChar = False
  end
  object lblVersion: TLabel
    Left = 189
    Top = 8
    Width = 42
    Height = 13
    Caption = 'v0.0.0.0'
  end
  object lblAuthor: TLinkLabel
    Left = 8
    Top = 27
    Width = 223
    Height = 17
    Caption = 
      'by Martijn Coppoolse, <a href="mailto:vor0nwe@users.sf.net">vor0' +
      'nwe@users.sf.net</a>'
    TabOrder = 1
    OnLinkClick = lblLinkClick
  end
  object lblTribute: TLinkLabel
    Left = 8
    Top = 115
    Width = 203
    Height = 17
    Caption = 
      'Damjan Zobo Cvetko, <a href="mailto:zobo@users.sf.net">zobo@user' +
      's.sf.net</a>'
    TabOrder = 3
    OnLinkClick = lblLinkClick
  end
  object btnOK: TButton
    Left = 136
    Top = 220
    Width = 75
    Height = 25
    Anchors = [akLeft, akRight, akBottom]
    Cancel = True
    Caption = '&OK'
    Default = True
    ModalResult = 1
    TabOrder = 0
  end
  object lblURL: TLinkLabel
    Left = 8
    Top = 46
    Width = 172
    Height = 17
    Cursor = crHandPoint
    Caption = 
      '<a href="http://fossil.2of4.net/npp_preview">http://fossil.2of4.' +
      'net/npp_preview</a>'
    TabOrder = 2
    OnLinkClick = lblLinkClick
  end
end
