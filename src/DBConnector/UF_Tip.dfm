object F_Tip: TF_Tip
  Left = 270
  Top = 219
  Width = 459
  Height = 464
  Caption = #20869#23481#25552#31034
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
  OnShow = FormShow
  DesignSize = (
    451
    432)
  PixelsPerInch = 96
  TextHeight = 12
  object bvl1: TBevel
    Left = 24
    Top = 383
    Width = 391
    Height = 3
    Anchors = [akLeft, akRight, akBottom]
    Shape = bsTopLine
  end
  object pnl1: TPanel
    Left = 24
    Top = 24
    Width = 405
    Height = 350
    Anchors = [akLeft, akTop, akRight, akBottom]
    BevelInner = bvLowered
    BevelOuter = bvLowered
    TabOrder = 0
    object mmoTip: TMemo
      Left = 2
      Top = 2
      Width = 401
      Height = 346
      Align = alClient
      BevelInner = bvNone
      BorderStyle = bsNone
      Color = clBtnFace
      ReadOnly = True
      ScrollBars = ssBoth
      TabOrder = 0
    end
  end
  object btn1: TButton
    Left = 356
    Top = 393
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = #20851#38381
    TabOrder = 1
    OnClick = btn1Click
  end
end
