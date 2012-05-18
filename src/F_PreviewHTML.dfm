inherited frmHTMLPreview: TfrmHTMLPreview
  BorderStyle = bsSizeToolWin
  Caption = 'HTML preview'
  ClientHeight = 420
  ClientWidth = 504
  ParentFont = True
  OnCreate = FormCreate
  OnHide = FormHide
  OnKeyPress = FormKeyPress
  OnShow = FormShow
  ExplicitWidth = 520
  ExplicitHeight = 454
  PixelsPerInch = 96
  TextHeight = 13
  object pnlButtons: TPanel
    Left = 0
    Top = 379
    Width = 504
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      504
      41)
    object btnRefresh: TButton
      Left = 8
      Top = 6
      Width = 75
      Height = 25
      Anchors = [akLeft, akBottom]
      Caption = 'Refresh'
      TabOrder = 0
      OnClick = btnRefreshClick
    end
    object btnClose: TButton
      Left = 421
      Top = 6
      Width = 75
      Height = 25
      Anchors = [akRight, akBottom]
      Caption = 'Close'
      TabOrder = 1
      OnClick = btnCloseClick
    end
    object sbrIE: TStatusBar
      Left = 89
      Top = 10
      Width = 326
      Height = 19
      Align = alNone
      Anchors = [akLeft, akRight, akBottom]
      Panels = <>
      SimplePanel = True
    end
  end
  object pnlPreview: TPanel
    Left = 0
    Top = 0
    Width = 504
    Height = 379
    Align = alClient
    BevelOuter = bvNone
    Caption = '(no preview available)'
    TabOrder = 1
    object pnlHTML: TPanel
      Left = 0
      Top = 0
      Width = 504
      Height = 379
      Align = alClient
      BevelOuter = bvNone
      Caption = 'pnlHTML'
      TabOrder = 0
      ExplicitWidth = 496
      ExplicitHeight = 369
      object wbIE: TWebBrowser
        Left = 0
        Top = 0
        Width = 504
        Height = 379
        TabStop = False
        Align = alClient
        TabOrder = 0
        OnStatusTextChange = wbIEStatusTextChange
        OnTitleChange = wbIETitleChange
        OnBeforeNavigate2 = wbIEBeforeNavigate2
        OnStatusBar = wbIEStatusBar
        OnNewWindow3 = wbIENewWindow3
        ExplicitLeft = 8
        ExplicitTop = 8
        ExplicitWidth = 288
        ExplicitHeight = 159
        ControlData = {
          4C000000173400002C2700000000000000000000000000000000000000000000
          000000004C000000000000000000000001000000E0D057007335CF11AE690800
          2B2E12620B000000000000004C0000000114020000000000C000000000000046
          8000000000000000000000000000000000000000000000000000000000000000
          00000000000000000100000000000000000000000000000000000000}
      end
    end
  end
end
