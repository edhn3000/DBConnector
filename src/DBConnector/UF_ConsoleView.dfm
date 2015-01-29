object F_ConsoleView: TF_ConsoleView
  Left = 284
  Top = 226
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsToolWindow
  Caption = 'ConsoleView'
  ClientHeight = 318
  ClientWidth = 480
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #23435#20307
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 12
  object mmoConsole: TMemo
    Left = 0
    Top = 0
    Width = 480
    Height = 318
    Align = alClient
    ReadOnly = True
    ScrollBars = ssBoth
    TabOrder = 0
    OnKeyDown = mmoConsoleKeyDown
  end
  object tmrClose: TTimer
    Enabled = False
    OnTimer = tmrCloseTimer
    Left = 67
    Top = 50
  end
end
