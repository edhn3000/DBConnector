object F_ExportTables: TF_ExportTables
  Left = 253
  Top = 115
  Caption = #23548#20986#25968#25454#24211
  ClientHeight = 454
  ClientWidth = 583
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object pnlBottom: TPanel
    Left = 0
    Top = 264
    Width = 583
    Height = 171
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 1
    DesignSize = (
      583
      171)
    object Label1: TLabel
      Left = 16
      Top = 108
      Width = 60
      Height = 13
      Anchors = [akLeft, akBottom]
      Caption = #23548#20986#30446#24405#65306
    end
    object Label2: TLabel
      Left = 24
      Top = 128
      Width = 300
      Height = 13
      Anchors = [akLeft, akBottom]
      Caption = #25552#31034#65306#27599#20010#34920#37117#20250#23548#20986#24314#34920#33050#26412#12289#25554#20837#33050#26412#31561#22810#20010#25991#20214#12290
    end
    object edtExportFile: TEdit
      Left = 72
      Top = 104
      Width = 457
      Height = 21
      Anchors = [akLeft, akBottom]
      TabOrder = 2
    end
    object btnSelect: TButton
      Left = 538
      Top = 103
      Width = 28
      Height = 22
      Anchors = [akLeft, akBottom]
      Caption = '...'
      TabOrder = 1
      OnClick = btnSelectClick
    end
    object btnOk: TButton
      Left = 472
      Top = 140
      Width = 94
      Height = 24
      Anchors = [akLeft, akBottom]
      Caption = #23548#20986
      TabOrder = 4
      OnClick = btnOkClick
    end
    object grpOption: TGroupBox
      Left = 13
      Top = 3
      Width = 558
      Height = 97
      Anchors = [akLeft, akTop, akRight]
      Caption = #36873#39033
      TabOrder = 0
      object chkDropTables: TCheckBox
        Left = 35
        Top = 12
        Width = 121
        Height = 17
        Caption = #23548#20986#21024#34920#35821#21477
        TabOrder = 0
        OnClick = chkDropTablesClick
      end
      object chkCreateTables: TCheckBox
        Left = 35
        Top = 32
        Width = 113
        Height = 17
        Caption = #23548#20986#24314#34920#35821#21477
        TabOrder = 2
      end
      object chkDisableTriggers: TCheckBox
        Left = 184
        Top = 64
        Width = 169
        Height = 17
        Hint = #25554#20837#25968#25454#21069#23631#34109#35302#21457#22120#65292#25554#20837#21518#24674#22797
        Caption = #23631#34109#35302#21457#22120'('#20165'Oracle)'
        Checked = True
        ParentShowHint = False
        ShowHint = True
        State = cbChecked
        TabOrder = 5
      end
      object chkDeleteTables: TCheckBox
        Left = 35
        Top = 52
        Width = 129
        Height = 17
        Caption = #23548#20986#28165#31354#25968#25454#35821#21477
        Checked = True
        State = cbChecked
        TabOrder = 4
      end
      object chkRemoveBreakLine: TCheckBox
        Left = 184
        Top = 16
        Width = 169
        Height = 17
        Hint = #25442#34892#21487#33021#23548#33268#25554#20837#25968#25454#26102#20986#38169
        Caption = #21435#38500#20869#23481#20013#25442#34892
        Checked = True
        ParentShowHint = False
        ShowHint = True
        State = cbChecked
        TabOrder = 1
      end
      object chkClobInFile: TCheckBox
        Left = 184
        Top = 40
        Width = 169
        Height = 17
        Hint = #36873#20013#23558#27599#26465#35760#24405#30340#22823#25991#26412#23383#27573#23548#20986#20026#25991#20214#65292#19981#36873#20013#21017#23548#20986#21040#35821#21477#20013
        Caption = #22823#25991#26412#23383#27573#23548#20986#20026#25991#20214
        Checked = True
        ParentShowHint = False
        ShowHint = True
        State = cbChecked
        TabOrder = 3
      end
      object chkExportData: TCheckBox
        Left = 35
        Top = 72
        Width = 129
        Height = 17
        Caption = #23548#20986#25968#25454
        Checked = True
        State = cbChecked
        TabOrder = 6
      end
    end
    object btnOpenDir: TButton
      Left = 379
      Top = 140
      Width = 82
      Height = 24
      Anchors = [akLeft, akBottom]
      Caption = #25171#24320#30446#24405
      TabOrder = 3
      OnClick = btnOpenDirClick
    end
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 435
    Width = 583
    Height = 19
    Panels = <
      item
        Width = 240
      end
      item
        Width = 50
      end>
  end
  object pnlClient: TPanel
    Left = 0
    Top = 0
    Width = 583
    Height = 264
    Align = alClient
    Caption = 'pnlClient'
    TabOrder = 0
    object tvwObjects: TCnCheckTreeView
      Left = 1
      Top = 14
      Width = 581
      Height = 249
      Align = alBottom
      AutoExpand = True
      Indent = 19
      ReadOnly = True
      TabOrder = 0
      Items.NodeData = {
        03020000002A000000000000000000000001000000FFFFFFFF00000000000000
        000100000001065400610062006C006500730024000000000000000000000001
        000000FFFFFFFF00000000000000000000000001033100310031002400000000
        0000000000000001000000FFFFFFFF0000000000000000010000000103320032
        00320024000000000000000000000001000000FFFFFFFF000000000000000000
        0000000103330033003300}
      CanDisableNode = False
    end
  end
end
