object F_About: TF_About
  Left = 410
  Top = 291
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = #20851#20110
  ClientHeight = 215
  ClientWidth = 330
  Color = clBtnFace
  TransparentColorValue = clGray
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #23435#20307
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 12
  object lblWinVersion: TLabel
    Left = 80
    Top = 32
    Width = 78
    Height = 12
    Caption = 'lblWinVersion'
  end
  object lblSoftVersion: TLabel
    Left = 80
    Top = 88
    Width = 84
    Height = 12
    Caption = 'lblSoftVersion'
  end
  object bvl1: TBevel
    Left = 16
    Top = 162
    Width = 297
    Height = 2
    Shape = bsFrame
  end
  object img1: TImage
    Left = 14
    Top = 32
    Width = 50
    Height = 65
    AutoSize = True
  end
  object lblDescription: TLabel
    Left = 16
    Top = 110
    Width = 84
    Height = 12
    Caption = 'lblDescription'
  end
  object lblURL: TLabel
    Left = 16
    Top = 142
    Width = 36
    Height = 12
    Caption = 'lblURL'
  end
  object btnOk: TButton
    Left = 240
    Top = 172
    Width = 75
    Height = 25
    Caption = #20851#38381
    TabOrder = 0
    OnClick = btnOkClick
  end
  object tmr1: TTimer
    Enabled = False
    Interval = 100
    OnTimer = tmr1Timer
    Left = 32
    Top = 170
  end
  object tmr2: TTimer
    Enabled = False
    Interval = 50
    OnTimer = tmr2Timer
    Left = 80
    Top = 170
  end
end
