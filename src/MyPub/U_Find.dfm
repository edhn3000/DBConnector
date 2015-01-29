object F_Find: TF_Find
  Left = 192
  Top = 145
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = #26597#25214
  ClientHeight = 130
  ClientWidth = 337
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #23435#20307
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 12
  object lbl1: TLabel
    Left = 24
    Top = 29
    Width = 48
    Height = 12
    Caption = #20851#38190#23383#65306
  end
  object bvl1: TBevel
    Left = 0
    Top = 88
    Width = 321
    Height = 2
    Shape = bsFrame
  end
  object edtKeyWord: TEdit
    Left = 72
    Top = 25
    Width = 249
    Height = 20
    TabOrder = 0
    OnKeyPress = edtKeyWordKeyPress
  end
  object chkCaseSensitive: TCheckBox
    Left = 72
    Top = 61
    Width = 97
    Height = 17
    Caption = #21306#20998#22823#23567#20889
    TabOrder = 1
  end
  object chkWholeWordMatch: TCheckBox
    Left = 224
    Top = 61
    Width = 97
    Height = 17
    Caption = #20840#23383#21305#37197
    TabOrder = 2
  end
  object btnFind: TButton
    Left = 128
    Top = 98
    Width = 75
    Height = 23
    Caption = #26597#25214
    TabOrder = 3
    OnClick = btnFindClick
  end
  object btnFindNext: TButton
    Left = 232
    Top = 98
    Width = 75
    Height = 23
    Caption = #19979#19968#20010
    TabOrder = 4
    OnClick = btnFindNextClick
  end
end
