inherited AboutForm: TAboutForm
  BorderIcons = []
  BorderStyle = bsSizeable
  Caption = 'About Preview HTML'
  ClientHeight = 256
  ParentFont = True
  OnCreate = FormCreate
  ExplicitWidth = 351
  ExplicitHeight = 294
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 96
    Width = 152
    Height = 13
    Caption = 'Based on the example plugin by'
  end
  object Label2: TLabel
    Left = 8
    Top = 115
    Width = 199
    Height = 13
    Caption = 'Damjan Zobo Cvetko, zobo@users.sf.net'
  end
  object lblPlugin: TLabel
    Left = 8
    Top = 8
    Width = 175
    Height = 13
    Caption = 'HTML Preview plugin for Notepad++'
    ShowAccelChar = False
  end
  object lblAuthor: TLabel
    Left = 8
    Top = 27
    Width = 105
    Height = 13
    Caption = 'by Martijn Coppoolse,'
    ShowAccelChar = False
  end
  object lblVersion: TLabel
    Left = 189
    Top = 8
    Width = 42
    Height = 13
    Caption = 'v0.0.0.0'
  end
  object Button1: TButton
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
  object txtAuthor: TStaticText
    Left = 119
    Top = 27
    Width = 115
    Height = 17
    Cursor = crHandPoint
    Caption = 'vor0nwe@users.sf.net'
    ShowAccelChar = False
    TabOrder = 1
    TabStop = True
    Transparent = False
    OnClick = txtAuthorClick
  end
end
