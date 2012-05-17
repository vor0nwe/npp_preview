inherited AboutForm: TAboutForm
  BorderIcons = []
  BorderStyle = bsSizeable
  Caption = 'About Preview HTML'
  ClientHeight = 256
  ParentFont = True
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 24
    Top = 24
    Width = 152
    Height = 13
    Caption = 'Based on the example plugin by'
  end
  object Label2: TLabel
    Left = 24
    Top = 48
    Width = 199
    Height = 13
    Caption = 'Damjan Zobo Cvetko, zobo@users.sf.net'
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
    ExplicitTop = 224
  end
end
