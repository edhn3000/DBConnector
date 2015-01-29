object F_ExecScript: TF_ExecScript
  Left = 363
  Top = 199
  Width = 463
  Height = 350
  BorderIcons = [biSystemMenu, biMinimize]
  Caption = #25191#34892#33050#26412
  Color = clBtnFace
  Font.Charset = GB2312_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #23435#20307
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 12
  object lblCurrLine: TLabel
    Left = 16
    Top = 80
    Width = 66
    Height = 12
    Caption = 'lblCurrLine'
    Visible = False
  end
  object lbledtScript: TLabeledEdit
    Left = 16
    Top = 40
    Width = 377
    Height = 20
    EditLabel.Width = 60
    EditLabel.Height = 12
    EditLabel.Caption = #33050#26412#25991#20214#65306
    TabOrder = 0
  end
  object btnOk: TButton
    Left = 356
    Top = 72
    Width = 73
    Height = 23
    Caption = #25191#34892
    TabOrder = 2
    OnClick = btnOkClick
  end
  object btnChoose: TButton
    Left = 402
    Top = 40
    Width = 27
    Height = 22
    Caption = '...'
    TabOrder = 1
    OnClick = btnChooseClick
  end
  object Memo1: TMemo
    Left = 0
    Top = 112
    Width = 447
    Height = 200
    Align = alBottom
    ScrollBars = ssVertical
    TabOrder = 3
    Visible = False
  end
end
