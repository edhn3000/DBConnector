object FM_DBOperate: TFM_DBOperate
  Left = 0
  Top = 0
  Width = 621
  Height = 467
  Font.Charset = GB2312_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = #23435#20307
  Font.Style = []
  ParentFont = False
  TabOrder = 0
  object stat1: TStatusBar
    Left = 0
    Top = 448
    Width = 621
    Height = 19
    Panels = <
      item
        Width = 120
      end
      item
        Width = 100
      end
      item
        Width = 100
      end
      item
        Width = 100
      end>
  end
  object pnl3: TPanel
    Left = 0
    Top = 0
    Width = 621
    Height = 448
    Align = alClient
    TabOrder = 0
    object splTopClient: TSplitter
      Left = 1
      Top = 101
      Width = 619
      Height = 4
      Cursor = crVSplit
      Align = alTop
    end
    object pnlTop: TPanel
      Left = 1
      Top = 1
      Width = 619
      Height = 100
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 0
      object pnl2: TPanel
        Left = 583
        Top = 0
        Width = 36
        Height = 100
        Align = alRight
        BevelOuter = bvNone
        TabOrder = 1
        object tlb2: TToolBar
          Left = -4
          Top = 0
          Width = 40
          Height = 100
          Align = alRight
          AutoSize = True
          Caption = 'tlb2'
          Images = DataModule1.ilBtnIcons
          TabOrder = 0
          object btn2: TToolButton
            Left = 0
            Top = 0
            Caption = 'btn2'
            DropdownMenu = pmSqlHistory
            ImageIndex = 20
            Wrap = True
            Style = tbsDropDown
            OnClick = btn2Click
          end
        end
      end
      object syndtSql: TSynEdit
        Left = 0
        Top = 0
        Width = 583
        Height = 100
        Align = alClient
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Courier New'
        Font.Style = []
        TabOrder = 0
        OnKeyDown = syndtSqlKeyDown
        Gutter.Font.Charset = DEFAULT_CHARSET
        Gutter.Font.Color = clWindowText
        Gutter.Font.Height = -11
        Gutter.Font.Name = 'Courier New'
        Gutter.Font.Style = []
        Highlighter = SynSQLSyn1
        OnChange = syndtSqlChange
        FontSmoothing = fsmNone
      end
    end
    object pnlClient: TPanel
      Left = 1
      Top = 105
      Width = 619
      Height = 342
      Align = alClient
      TabOrder = 1
      object pgcClient: TPageControl
        Left = 1
        Top = 1
        Width = 617
        Height = 340
        ActivePage = tsData
        Align = alClient
        TabOrder = 0
        object tsData: TTabSheet
          Caption = 'Data'
          object pnlGrid: TPanel
            Left = 0
            Top = 0
            Width = 609
            Height = 311
            Align = alClient
            BevelOuter = bvNone
            TabOrder = 0
            object splFieldDetail: TSplitter
              Left = 0
              Top = 196
              Width = 609
              Height = 4
              Cursor = crVSplit
              Align = alBottom
              Visible = False
              ExplicitTop = 197
            end
            object pnl6: TPanel
              Left = 0
              Top = 0
              Width = 609
              Height = 25
              Align = alTop
              AutoSize = True
              BevelOuter = bvNone
              TabOrder = 0
              object tlbtbr1: TToolBar
                Left = 0
                Top = 0
                Width = 316
                Height = 25
                Align = alLeft
                Caption = 'tbrMain'
                EdgeInner = esNone
                EdgeOuter = esNone
                Images = DataModule1.ilBtnIcons
                ParentShowHint = False
                ShowHint = True
                TabOrder = 0
                object btnFindAtList: TToolButton
                  Left = 0
                  Top = 0
                  Hint = #34920#26684#20013#26597#25214
                  ImageIndex = 15
                  ParentShowHint = False
                  ShowHint = True
                  OnClick = btnFindAtListClick
                end
                object btnExport: TToolButton
                  Left = 23
                  Top = 0
                  Hint = #23548#20986
                  Action = actExport
                  ImageIndex = 8
                  ParentShowHint = False
                  ShowHint = True
                end
                object dbnvgr1: TDBNavigator
                  Left = 46
                  Top = 0
                  Width = 180
                  Height = 22
                  DataSource = dsData
                  VisibleButtons = [nbInsert, nbDelete, nbEdit, nbPost, nbCancel, nbRefresh]
                  Flat = True
                  Hints.Strings = (
                    #31532#19968#20010
                    #21069#19968#20010
                    #19979#19968#20010
                    #26368#21518#19968#20010
                    #25554#20837
                    #21024#38500
                    #32534#36753
                    #25552#20132
                    #21462#28040
                    #21047#26032)
                  ParentShowHint = False
                  ShowHint = True
                  TabOrder = 0
                end
                object btntbtn1: TToolButton
                  Left = 226
                  Top = 0
                  Width = 8
                  Caption = 'btntbtn1'
                  ImageIndex = 9
                  Style = tbsSeparator
                end
                object btnPrint: TToolButton
                  Left = 234
                  Top = 0
                  Hint = #25171#21360
                  Caption = 'btnPrint'
                  ImageIndex = 19
                  ParentShowHint = False
                  ShowHint = True
                  OnClick = btnPrintClick
                end
                object btnAdjust: TToolButton
                  Left = 257
                  Top = 0
                  Hint = #35843#25972#23485#24230#26041#24335
                  Caption = 'btnAdjust'
                  DropdownMenu = pmAdjust
                  ImageIndex = 11
                  ParentShowHint = False
                  ShowHint = True
                  Style = tbsDropDown
                  OnClick = btnAdjustClick
                end
                object btnSortOnColumnClick: TToolButton
                  Left = 297
                  Top = 0
                  Hint = #28857#20987#21015#22836#25490#24207
                  ImageIndex = 21
                  Style = tbsCheck
                  OnClick = btnSortOnColumnClickClick
                end
              end
              object pnl1: TPanel
                Left = 589
                Top = 0
                Width = 20
                Height = 25
                Align = alRight
                AutoSize = True
                BevelOuter = bvNone
                TabOrder = 1
                object btn1: TSpeedButton
                  Left = 0
                  Top = 1
                  Width = 20
                  Height = 20
                  Caption = #215
                  Flat = True
                  OnClick = btn1Click
                end
              end
            end
            object dbgrdData: TfDBGrid
              Left = 0
              Top = 25
              Width = 609
              Height = 171
              AutoFitColumnWidth = True
              FitColumnMode = fcmBoth
              MaxFitColCount = 20
              BrushColorOdd = clWindow
              BrushColorEven = 12250296
              SortOnColumnClick = False
              Align = alClient
              DataSource = dsData
              PopupMenu = pmMenu
              TabOrder = 1
              TitleFont.Charset = GB2312_CHARSET
              TitleFont.Color = clWindowText
              TitleFont.Height = -13
              TitleFont.Name = #23435#20307
              TitleFont.Style = []
            end
            object pnlMsg: TPanel
              Left = 0
              Top = 200
              Width = 609
              Height = 111
              Align = alBottom
              BevelOuter = bvNone
              TabOrder = 2
              Visible = False
              OnResize = pnlMsgResize
              object pnl15: TPanel
                Left = 0
                Top = 0
                Width = 16
                Height = 111
                Align = alLeft
                AutoSize = True
                BevelOuter = bvNone
                TabOrder = 0
                object bvlBottom2: TBevel
                  Left = 9
                  Top = 20
                  Width = 2
                  Height = 100
                  Style = bsRaised
                end
                object bvlBottom1: TBevel
                  Left = 3
                  Top = 20
                  Width = 2
                  Height = 100
                  Style = bsRaised
                end
                object pnlBottomContranCloseBtn: TPanel
                  Left = 0
                  Top = 0
                  Width = 16
                  Height = 16
                  AutoSize = True
                  BevelOuter = bvNone
                  TabOrder = 0
                  object btnCloseMsg: TSpeedButton
                    Left = 0
                    Top = 0
                    Width = 16
                    Height = 16
                    Caption = #215
                    Flat = True
                    OnClick = btnCloseMsgClick
                  end
                end
              end
              object pnl5: TPanel
                Left = 16
                Top = 0
                Width = 593
                Height = 111
                Align = alClient
                BevelOuter = bvNone
                Caption = 'pnl5'
                TabOrder = 1
                object pnl4: TPanel
                  Left = 0
                  Top = 22
                  Width = 593
                  Height = 89
                  Align = alClient
                  BevelOuter = bvNone
                  TabOrder = 1
                  object tlbField: TToolBar
                    Left = 0
                    Top = 0
                    Width = 593
                    Height = 22
                    ButtonHeight = 21
                    ButtonWidth = 49
                    Caption = 'tlbField'
                    ShowCaptions = True
                    TabOrder = 0
                    object btnExportField: TToolButton
                      Left = 0
                      Top = 0
                      Action = actExportField
                    end
                    object btnImportField: TToolButton
                      Left = 49
                      Top = 0
                      Action = actImportField
                    end
                  end
                  object dbmmoShowFieldDetail: TDBMemo
                    Left = 0
                    Top = 22
                    Width = 593
                    Height = 67
                    Hint = #25913#21160#21487#20445#23384#21040#25968#25454#24211
                    Align = alClient
                    DataSource = dsData
                    ParentShowHint = False
                    ScrollBars = ssBoth
                    ShowHint = True
                    TabOrder = 1
                  end
                end
                object Panel1: TPanel
                  Left = 0
                  Top = 0
                  Width = 593
                  Height = 22
                  Align = alTop
                  BevelOuter = bvNone
                  TabOrder = 0
                  object lblFieldName: TLabel
                    Left = 12
                    Top = 4
                    Width = 87
                    Height = 13
                    Alignment = taCenter
                    BiDiMode = bdLeftToRight
                    Caption = #23383#27573':  '#31867#22411': '
                    ParentBiDiMode = False
                  end
                end
              end
            end
          end
        end
        object tsOutPut: TTabSheet
          Caption = 'OutPut'
          ImageIndex = 1
          ExplicitLeft = 0
          ExplicitTop = 0
          ExplicitWidth = 0
          ExplicitHeight = 0
          object pnl7: TPanel
            Left = 0
            Top = 0
            Width = 609
            Height = 311
            Align = alClient
            BevelOuter = bvNone
            Caption = 'pnl7'
            TabOrder = 0
            object mmoLog: TMemo
              Left = 0
              Top = 24
              Width = 609
              Height = 287
              Align = alClient
              ReadOnly = True
              ScrollBars = ssBoth
              TabOrder = 1
              OnChange = mmoLogChange
            end
            object pnl8: TPanel
              Left = 0
              Top = 0
              Width = 609
              Height = 24
              Align = alTop
              BevelOuter = bvNone
              TabOrder = 0
              object tlb1: TToolBar
                Left = 0
                Top = 0
                Width = 97
                Height = 24
                Align = alLeft
                Caption = 'tbrTree'
                EdgeInner = esNone
                EdgeOuter = esNone
                TabOrder = 0
                object btnManClearMsg: TToolButton
                  Left = 0
                  Top = 0
                  Hint = #28165#31354
                  ImageIndex = 5
                  ParentShowHint = False
                  Wrap = True
                  ShowHint = True
                  OnClick = btnManClearMsgClick
                end
              end
              object pnl9: TPanel
                Left = 589
                Top = 0
                Width = 20
                Height = 24
                Align = alRight
                AutoSize = True
                BevelOuter = bvNone
                TabOrder = 1
                object btn3: TSpeedButton
                  Left = 0
                  Top = 1
                  Width = 20
                  Height = 20
                  Caption = #215
                  Flat = True
                  OnClick = btn1Click
                end
              end
            end
          end
        end
      end
    end
  end
  object dsData: TDataSource
    Left = 24
    Top = 192
  end
  object pmAdjust: TPopupMenu
    Left = 64
    Top = 192
    object mniAdjustByTitleContent: TMenuItem
      AutoHotkeys = maManual
      Caption = #25353#29031#26631#39064#21644#20869#23481#20915#23450#21015#23485
      OnClick = mniAdjustByTitleContentClick
    end
    object mniAdjustByTitle: TMenuItem
      Caption = #25353#29031#26631#39064#20915#23450#21015#23485
      OnClick = mniAdjustByTitleClick
    end
    object mniAdjustByContent: TMenuItem
      Caption = #25353#29031#20869#23481#20915#23450#21015#23485
      OnClick = mniAdjustByContentClick
    end
  end
  object actlstMain: TActionList
    Left = 64
    Top = 232
    object actExecSelSql: TAction
      Category = #32534#36753
      Caption = #25191#34892'(F8)'
      ShortCut = 119
    end
    object actExecAllSqls: TAction
      Category = #32534#36753
      Caption = #20840#37096#25191#34892
    end
    object actLastError: TAction
      Category = #32534#36753
      Caption = #19978#27425#38169#35823
    end
    object actEditData: TAction
      Category = #32534#36753
      Caption = #32534#36753#25968#25454
    end
    object actExecScript: TAction
      Category = #32534#36753
      Caption = #25191#34892#33050#26412
    end
    object actExport: TAction
      Category = #32534#36753
      Caption = #23548#20986
      OnExecute = actExportExecute
    end
    object actFieldDetail: TAction
      Category = 'PopMenu'
      Caption = #23383#27573#35814#32454#20449#24687
      OnExecute = actFieldDetailExecute
    end
    object actStopExec: TAction
      Category = #32534#36753
      Caption = #20572#27490#25191#34892
      Enabled = False
    end
    object actExportField: TAction
      Category = #32534#36753
      Caption = 'Export'
      OnExecute = actExportFieldExecute
    end
    object actImportField: TAction
      Category = #32534#36753
      Caption = 'Import'
      OnExecute = actImportFieldExecute
    end
  end
  object SynSQLSyn1: TSynSQLSyn
    Options.AutoDetectEnabled = False
    Options.AutoDetectLineLimit = 0
    Options.Visible = False
    CommentAttri.Foreground = clGreen
    KeyAttri.Foreground = 10485760
    Left = 72
    Top = 40
  end
  object pmMenu: TPopupMenu
    Left = 24
    Top = 232
    object mniFieldDetail: TMenuItem
      Action = actFieldDetail
    end
    object mniCopyToClip: TMenuItem
      Caption = #22797#21046#21040#21098#36148#26495
      Visible = False
    end
  end
  object pmSqlHistory: TPopupMenu
    OnPopup = pmSqlHistoryPopup
    Left = 496
    Top = 32
  end
end
