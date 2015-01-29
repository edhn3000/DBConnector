object F_ResultBox: TF_ResultBox
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = 'F_ResultBox'
  ClientHeight = 234
  ClientWidth = 340
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
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 340
    Height = 190
    Align = alClient
    TabOrder = 0
    ExplicitLeft = 72
    ExplicitTop = 40
    ExplicitWidth = 185
    ExplicitHeight = 41
  end
  object Panel2: TPanel
    Left = 0
    Top = 190
    Width = 340
    Height = 44
    Align = alBottom
    TabOrder = 1
    object Button1: TButton
      Left = 48
      Top = 6
      Width = 75
      Height = 25
      Caption = 'Button1'
      TabOrder = 0
    end
    object Button2: TButton
      Left = 129
      Top = 6
      Width = 75
      Height = 25
      Caption = 'Button1'
      TabOrder = 1
    end
    object Button3: TButton
      Left = 210
      Top = 6
      Width = 75
      Height = 25
      Caption = 'Button1'
      TabOrder = 2
    end
  end
end
