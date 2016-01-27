object F_ExportResult: TF_ExportResult
  Left = 262
  Top = 262
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = #23548#20986#32467#26524
  ClientHeight = 138
  ClientWidth = 356
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object lbl1: TLabel
    Left = 24
    Top = 29
    Width = 60
    Height = 13
    Caption = #25991#20214#20445#23384#22312
  end
  object bvl1: TBevel
    Left = 24
    Top = 87
    Width = 313
    Height = 30
    Shape = bsTopLine
  end
  object btnOpenWidhEditor: TButton
    Left = 132
    Top = 96
    Width = 91
    Height = 25
    Caption = #22312#32534#36753#22120#25171#24320
    TabOrder = 2
    OnClick = btnOpenWidhEditorClick
  end
  object edtExportFile: TEdit
    Left = 24
    Top = 48
    Width = 313
    Height = 21
    ReadOnly = True
    TabOrder = 0
  end
  object btnOpen: TButton
    Left = 24
    Top = 96
    Width = 89
    Height = 25
    Caption = #20851#32852#31243#24207#25171#24320
    TabOrder = 1
    OnClick = btnOpenClick
  end
  object btnOpenDir: TButton
    Left = 239
    Top = 96
    Width = 91
    Height = 25
    Caption = #25171#24320#25991#20214#22841
    TabOrder = 3
    OnClick = btnOpenDirClick
  end
end
