object frmAjuda: TfrmAjuda
  Left = 0
  Top = 0
  Caption = 'Ajuda'
  ClientHeight = 302
  ClientWidth = 506
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object pnlButton: TPanel
    Left = 0
    Top = 261
    Width = 506
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      506
      41)
    object btnFechar: TButton
      Left = 411
      Top = 6
      Width = 83
      Height = 25
      Anchors = [akTop, akRight, akBottom]
      Caption = '&Fechar'
      TabOrder = 0
      OnClick = btnFecharClick
    end
  end
  object pnlDados: TPanel
    Left = 0
    Top = 0
    Width = 506
    Height = 261
    Align = alClient
    BevelOuter = bvNone
    Caption = 'pnlDados'
    TabOrder = 1
    object mmoAjuda: TMemo
      Left = 0
      Top = 0
      Width = 506
      Height = 261
      Align = alClient
      ParentShowHint = False
      ReadOnly = True
      ScrollBars = ssVertical
      ShowHint = True
      TabOrder = 0
    end
  end
end
