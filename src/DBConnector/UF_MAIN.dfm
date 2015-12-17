object F_MAIN: TF_MAIN
  Left = 317
  Top = 102
  Caption = 'DBConnector'
  ClientHeight = 560
  ClientWidth = 785
  Color = clBtnFace
  Font.Charset = GB2312_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #23435#20307
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  Visible = True
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 12
  object stbMain: TStatusBar
    Left = 0
    Top = 541
    Width = 785
    Height = 19
    Panels = <
      item
        Width = 185
      end
      item
        Width = 150
      end>
  end
  object pnl13: TPanel
    Left = 0
    Top = 55
    Width = 785
    Height = 486
    Align = alClient
    TabOrder = 1
    object splLeftAndMain: TSplitter
      Left = 186
      Top = 1
      Width = 4
      Height = 484
    end
    object pnl14: TPanel
      Left = 190
      Top = 1
      Width = 594
      Height = 484
      Align = alClient
      TabOrder = 1
      object pnl16: TPanel
        Left = 1
        Top = 1
        Width = 592
        Height = 24
        Align = alTop
        AutoSize = True
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        object tbr2: TToolBar
          Left = 1
          Top = 1
          Width = 590
          Height = 22
          AutoSize = True
          Caption = 'tbr2'
          Images = DataModule1.ilBtnIcons
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
          OnMouseDown = BarMouseDown
          object tbtnAddPage: TToolButton
            Left = 0
            Top = 0
            Hint = #22686#21152#19968#39029
            Action = actAddPage
            ImageIndex = 3
            ParentShowHint = False
            ShowHint = True
          end
          object tbtnRemovePage: TToolButton
            Left = 23
            Top = 0
            Hint = #20851#38381#19968#39029
            Action = actRemovePage
            ImageIndex = 0
            ParentShowHint = False
            ShowHint = True
          end
          object tbtn2: TToolButton
            Left = 46
            Top = 0
            Width = 8
            Caption = 'tbtn2'
            ImageIndex = 16
            Style = tbsSeparator
          end
          object tbtnOpenWithEditor: TToolButton
            Left = 54
            Top = 0
            Hint = #25171#24320#25991#20214
            Action = actOpenWithEditor
            ImageIndex = 23
            ParentShowHint = False
            ShowHint = True
          end
          object tbtnSaveEditorText: TToolButton
            Left = 77
            Top = 0
            Hint = #20445#23384'Ctrl+S'
            Action = actSaveEditorText
            ImageIndex = 17
            ParentShowHint = False
            ShowHint = True
          end
          object tbtnSaveEditorTextAs: TToolButton
            Left = 100
            Top = 0
            Hint = #21478#23384#20026
            Action = actSaveEditorTextAs
            ImageIndex = 18
            ParentShowHint = False
            ShowHint = True
          end
          object btn3: TToolButton
            Left = 123
            Top = 0
            Width = 8
            Caption = 'btn3'
            ImageIndex = 20
            Style = tbsSeparator
          end
          object btnWrap: TToolButton
            Left = 131
            Top = 0
            Hint = #33258#21160#25442#34892'Ctrl+W'
            Caption = 'btnWrap'
            ImageIndex = 24
            OnClick = btnWrapClick
          end
          object btnSearchText: TToolButton
            Left = 154
            Top = 0
            Hint = #26597#25214'Ctrl+F'
            Caption = 'btnSearchText'
            ImageIndex = 15
            OnClick = btnSearchTextClick
          end
        end
      end
      object pnl18: TPanel
        Left = 1
        Top = 25
        Width = 592
        Height = 458
        Align = alClient
        TabOrder = 1
        object pgcMain: TPageControl
          Left = 1
          Top = 1
          Width = 590
          Height = 456
          ActivePage = ts3
          Align = alClient
          TabOrder = 0
          OnChange = pgcMainChange
          OnMouseDown = pgcMainMouseDown
          object ts3: TTabSheet
            Caption = 'SQL'
            ImageIndex = 2
            inline frmDBOperateMain: TFM_DBOperate
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
              inherited pnl3: TPanel
                inherited pnlClient: TPanel
                  inherited pgcClient: TPageControl
                    ActivePage = frmDBOperateMain.tsOutPut
                    inherited tsOutPut: TTabSheet
                      inherited pnl7: TPanel
                        inherited pnl8: TPanel
                          inherited tlb1: TToolBar
                            Images = DataModule1.ilBtnIcons
                          end
                        end
                      end
                    end
                  end
                end
              end
            end
          end
        end
      end
    end
    object pnlLeft: TPanel
      Left = 1
      Top = 1
      Width = 185
      Height = 484
      Align = alLeft
      TabOrder = 0
      OnResize = pnlLeftResize
      object pnl1: TPanel
        Left = 1
        Top = 1
        Width = 183
        Height = 16
        Align = alTop
        AutoSize = True
        BevelOuter = bvNone
        TabOrder = 0
        object bvlLeft1: TBevel
          Left = 0
          Top = 4
          Width = 167
          Height = 2
          Style = bsRaised
        end
        object bvlLeft2: TBevel
          Left = 0
          Top = 8
          Width = 167
          Height = 2
          Style = bsRaised
        end
        object pnlLeftContranCloseBtn: TPanel
          Left = 167
          Top = 0
          Width = 16
          Height = 16
          Align = alRight
          AutoSize = True
          BevelOuter = bvNone
          TabOrder = 0
          object btnClosePnlLeft: TButton
            Left = 0
            Top = 0
            Width = 16
            Height = 16
            Caption = #215
            Font.Charset = GB2312_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = #23435#20307
            Font.Style = []
            ParentFont = False
            TabOrder = 0
            TabStop = False
            OnClick = btnClosePnlLeftClick
          end
        end
      end
      object pnl17: TPanel
        Left = 1
        Top = 17
        Width = 183
        Height = 466
        Align = alClient
        BevelOuter = bvNone
        TabOrder = 1
        object pnlLeftTop: TPanel
          Left = 0
          Top = 0
          Width = 183
          Height = 44
          Align = alTop
          BevelOuter = bvNone
          TabOrder = 0
          object pnlObjSelect: TPanel
            Left = 0
            Top = 22
            Width = 183
            Height = 22
            Align = alBottom
            AutoSize = True
            Caption = 'pnlObjSelect'
            TabOrder = 1
            OnResize = pnlObjSelectResize
            DesignSize = (
              183
              22)
            object cbbTableFilter: TComboBox
              Left = 1
              Top = 1
              Width = 180
              Height = 20
              Style = csDropDownList
              Anchors = [akLeft, akTop, akRight, akBottom]
              ItemIndex = 0
              TabOrder = 0
              Text = #29992#25143#23545#35937
              OnChange = cbbTableFilterChange
              Items.Strings = (
                #29992#25143#23545#35937
                #20840#37096#23545#35937)
            end
          end
          object pnlSchemaSelect: TPanel
            Left = 0
            Top = 0
            Width = 183
            Height = 22
            Align = alTop
            TabOrder = 0
            Visible = False
            OnResize = pnlSchemaSelectResize
            DesignSize = (
              183
              22)
            object cbbSchemaSelect: TComboBox
              Left = 1
              Top = 1
              Width = 180
              Height = 20
              Style = csDropDownList
              Anchors = [akLeft, akTop, akRight, akBottom]
              TabOrder = 0
              OnChange = cbbSchemaSelectChange
            end
          end
        end
        object pnl19: TPanel
          Left = 0
          Top = 44
          Width = 183
          Height = 422
          Align = alClient
          BevelOuter = bvNone
          TabOrder = 1
          object tbrTree: TToolBar
            Left = 0
            Top = 0
            Width = 183
            Height = 22
            AutoSize = True
            Caption = 'tbrTree'
            Images = DataModule1.ilBtnIcons
            ParentShowHint = False
            ShowHint = True
            TabOrder = 0
            object tbtnRefresh: TToolButton
              Left = 0
              Top = 0
              Hint = #21047#26032#26641
              Action = actManRefreshTree
              ParentShowHint = False
              ShowHint = True
            end
            object tbtnSearch: TToolButton
              Left = 23
              Top = 0
              Hint = #26597#25214#25968#25454#24211#23545#35937
              Action = actFindAtTree
              ImageIndex = 15
              ParentShowHint = False
              ShowHint = True
            end
            object tbtnEditData: TToolButton
              Left = 46
              Top = 0
              Action = actEditData
              ImageIndex = 4
              ParentShowHint = False
              ShowHint = True
            end
            object btnSortNode: TToolButton
              Left = 69
              Top = 0
              Hint = #33410#28857#33258#21160#25490#24207
              Caption = 'btnSortNode'
              ImageIndex = 21
              Style = tbsCheck
              OnClick = btnSortNodeClick
            end
          end
          object pnlQuickSearch: TPanel
            Left = 0
            Top = 386
            Width = 183
            Height = 36
            Align = alBottom
            BevelOuter = bvNone
            TabOrder = 2
            Visible = False
            object lbl1: TLabel
              Left = 8
              Top = 12
              Width = 24
              Height = 12
              Caption = #26597#25214
            end
            object edtQuickSearch: TEdit
              Left = 46
              Top = 8
              Width = 121
              Height = 20
              TabOrder = 0
              OnChange = edtQuickSearchChange
              OnKeyDown = edtQuickSearchKeyDown
            end
          end
          object tvwObjects: TfHintTree
            Left = 0
            Top = 22
            Width = 183
            Height = 364
            Align = alClient
            Images = DataModule1.ilTree
            Indent = 19
            MultiSelect = True
            MultiSelectStyle = [msControlSelect, msShiftSelect]
            ParentShowHint = False
            ReadOnly = True
            RightClickSelect = True
            ShowHint = True
            TabOrder = 1
            OnChange = tvwObjectsChange
            OnCollapsed = tvwObjectsCollapsed
            OnCustomDrawItem = tvwObjectsCustomDrawItem
            OnExpanding = tvwObjectsExpanding
            OnExpanded = tvwObjectsExpanded
            OnKeyDown = tvwObjectsKeyDown
            OnKeyPress = tvwObjectsKeyPress
            OnMouseDown = tvwObjectsMouseDown
            HintBgColor = clInfoBk
            HintRemainTime = 5000
            OnShowHint = tvwObjectsShowHint
          end
        end
      end
    end
  end
  object clbr1: TCoolBar
    Left = 0
    Top = 0
    Width = 785
    Height = 55
    AutoSize = True
    Bands = <
      item
        Break = False
        Control = actmmb1
        ImageIndex = -1
        MinHeight = 23
        Width = 779
      end
      item
        Control = acttb1
        ImageIndex = -1
        MinHeight = 26
        Width = 779
      end>
    object acttb1: TActionToolBar
      Left = 11
      Top = 25
      Width = 770
      Height = 26
      ActionManager = actmgr1
      Caption = 'toolbar'
      ColorMap.HighlightColor = clWhite
      ColorMap.BtnSelectedColor = clBtnFace
      ColorMap.UnusedColor = clWhite
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      Spacing = 0
      OnMouseDown = BarMouseDown
    end
    object actmmb1: TActionMainMenuBar
      Left = 11
      Top = 0
      Width = 770
      Height = 23
      UseSystemFont = False
      ActionManager = actmgr1
      AllowHiding = True
      AnimationStyle = asFade
      Caption = 'menu'
      ColorMap.HighlightColor = clWhite
      ColorMap.BtnSelectedColor = clBtnFace
      ColorMap.UnusedColor = clWhite
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = #24494#36719#38597#40657
      Font.Style = []
      ParentShowHint = False
      ShowHint = False
      Spacing = 0
      OnMouseDown = BarMouseDown
    end
  end
  object actlstMain: TActionList
    Left = 56
    Top = 160
    object actExit: TAction
      Category = #31995#32479
      Caption = #36864#20986
      ImageIndex = 5
      OnExecute = actExitExecute
    end
    object actExecSelSql: TAction
      Category = #32534#36753
      Caption = #25191#34892'(F8)'
      ShortCut = 119
      OnExecute = actExecSelSqlExecute
    end
    object actRefreshTree: TAction
      Category = #26597#30475
      Caption = #21047#26032#26641
      OnExecute = actRefreshTreeExecute
    end
    object actConnect: TAction
      Category = #31995#32479
      Caption = #36830#25509
      OnExecute = actConnectExecute
    end
    object actExecAllSqls: TAction
      Category = #32534#36753
      Caption = #20840#37096#25191#34892
      OnExecute = actExecAllSqlsExecute
    end
    object actEditData: TAction
      Category = #32534#36753
      Caption = #32534#36753#25968#25454
      OnExecute = actEditDataExecute
    end
    object actExecScript: TAction
      Category = #32534#36753
      Caption = #25191#34892#33050#26412
      OnExecute = actExecScriptExecute
    end
    object actFindAtTree: TAction
      Category = #26597#30475
      Caption = #26641#20013#26597#25214
      OnExecute = actFindAtTreeExecute
    end
    object actFieldDetail: TAction
      Category = 'PopMenu'
      Caption = #23383#27573#35814#32454#20449#24687
    end
    object actStopExec: TAction
      Category = #32534#36753
      Caption = #20572#27490#25191#34892
      Enabled = False
      OnExecute = actStopExecExecute
    end
    object actOpenWithEditor: TAction
      Category = #31995#32479
      Caption = #22312#32534#36753#22120#20013#25171#24320'...'
      OnExecute = actOpenWithEditorExecute
    end
    object actSaveEditorText: TAction
      Category = #31995#32479
      Caption = #20445#23384#32534#36753#22120#20869#23481
      Enabled = False
      OnExecute = actSaveEditorTextExecute
    end
    object actAddPage: TAction
      Category = 'PgcMainControl'
      Caption = #22686#21152#39029
      OnExecute = actAddPageExecute
    end
    object actRemovePage: TAction
      Category = 'PgcMainControl'
      Caption = #20851#38381#39029
      Enabled = False
      OnExecute = actRemovePageExecute
    end
    object actTip: TAction
      Category = #24110#21161
      Caption = #25552#31034
      OnExecute = actTipExecute
    end
    object actAbout: TAction
      Category = #24110#21161
      Caption = #20851#20110'(&A)'
      OnExecute = actAboutExecute
    end
    object actDBParams: TAction
      Category = #35774#32622
      Caption = #25968#25454#24211#21442#25968
      ShortCut = 16463
      OnExecute = actDBParamsExecute
    end
    object actResetIconCache: TAction
      Category = #24110#21161
      Caption = #37325#24314#22270#26631#32531#23384
      OnExecute = actResetIconCacheExecute
    end
    object actSaveEditorTextAs: TAction
      Category = #31995#32479
      Caption = #21478#23384#32534#36753#22120#20869#23481#20026'...'
      OnExecute = actSaveEditorTextAsExecute
    end
    object actExport: TAction
      Category = #32534#36753
      Caption = #23548#20986#25968#25454
      OnExecute = actExportExecute
    end
    object actPrint: TAction
      Category = #31995#32479
      Caption = #25171#21360
      OnExecute = actPrintExecute
    end
    object actExportTableCreateSQL: TAction
      Category = #32534#36753
      Caption = #29983#25104#24314#34920#33050#26412
      OnExecute = actExportTableCreateSQLExecute
    end
    object actDropTable: TAction
      Category = #34920#25805#20316
      Caption = #21024#38500#34920
      OnExecute = actDropTableExecute
    end
    object actDeleteField: TAction
      Category = #23383#27573#25805#20316
      Caption = #21024#38500#23383#27573
      OnExecute = actDeleteFieldExecute
    end
    object actDelData: TAction
      Category = #32534#36753
      Caption = #21024#38500#25968#25454
      OnExecute = actDelDataExecute
    end
    object actAddField: TAction
      Category = #23383#27573#25805#20316
      Caption = 'actAddField'
      OnExecute = actAddFieldExecute
    end
    object actModifyField: TAction
      Category = #23383#27573#25805#20316
      Caption = 'actModifyField'
      OnExecute = actModifyFieldExecute
    end
    object actExportDB: TAction
      Category = #32534#36753
      Caption = #23548#20986#25968#25454
      OnExecute = actExportDBExecute
    end
    object actCommandHelp: TAction
      Category = #39640#32423
      Caption = #21629#20196#24110#21161
      OnExecute = actCommandHelpExecute
    end
    object actShowSettings: TAction
      Category = #39640#32423
      Caption = #26174#31034#24403#21069#35774#32622
      OnExecute = actShowSettingsExecute
    end
    object actEditObj: TAction
      Category = #32534#36753
      Caption = #32534#36753
      OnExecute = actEditObjExecute
    end
  end
  object pmTvw: TPopupMenu
    AutoHotkeys = maManual
    OnPopup = pmTvwPopup
    Left = 56
    Top = 200
    object mniFindAtTree: TMenuItem
      Action = actFindAtTree
    end
    object mniCopyName: TMenuItem
      AutoHotkeys = maManual
      Caption = #22797#21046#21517#31216' Ctrl+C'
      OnClick = mniCopyNameClick
    end
    object N3: TMenuItem
      Caption = '-'
    end
    object N7: TMenuItem
      Action = actEditObj
    end
    object N1: TMenuItem
      Action = actDropTable
      Caption = #38144#27585
    end
    object N8: TMenuItem
      Caption = '-'
    end
    object mniN4: TMenuItem
      Action = actEditData
    end
    object N6: TMenuItem
      Action = actDelData
    end
    object N5: TMenuItem
      Caption = '-'
    end
    object N2: TMenuItem
      Action = actDeleteField
      Caption = #21435#38500#23383#27573
    end
    object N4: TMenuItem
      Caption = '-'
    end
    object mniGenSql: TMenuItem
      Caption = #29983#25104'SQL'
      object mniGenCreateSQL: TMenuItem
        Action = actExportTableCreateSQL
        Caption = #29983#25104'Create'
      end
      object mniGenSelect: TMenuItem
        Caption = #29983#25104'Select'
        OnClick = mniGenSelectClick
      end
      object mniGenInsert: TMenuItem
        Caption = #29983#25104'Insert'
        OnClick = mniGenInsertClick
      end
      object mniGenUpdate: TMenuItem
        Caption = #29983#25104'Update'
        OnClick = mniGenUpdateClick
      end
    end
  end
  object actmgr1: TActionManager
    ActionBars.Customizable = False
    ActionBars = <
      item
        ContextItems.AutoHotKeys = False
        ContextItems = <>
        Items.AutoHotKeys = False
        Items.HideUnused = False
        Items = <
          item
            Items.HideUnused = False
            Items = <
              item
                Items.HideUnused = False
                Items = <>
                Action = actManConnect
                Caption = #36830#25509'(&W)'
                ImageIndex = 6
              end
              item
                Items.HideUnused = False
                Items = <>
                Caption = '-'
              end
              item
                Items.HideUnused = False
                Items = <>
                Action = actManOpenWithEditor
                Caption = #22312#32534#36753#22120#20013#25171#24320'(&O)...'
              end
              item
                Items.HideUnused = False
                Items = <>
                Action = actSaveEditorText
                Caption = #20445#23384#32534#36753#22120#20869#23481'(&X)'
              end
              item
                Items.HideUnused = False
                Items = <>
                Action = actSaveEditorTextAs
                Caption = #21478#23384#32534#36753#22120#20869#23481#20026'(&Y)...'
              end
              item
                Items.HideUnused = False
                Items = <>
                Caption = '-'
              end
              item
                Items.HideUnused = False
                Items = <>
                Action = actPrint
                Caption = #25171#21360'(&P)'
              end
              item
                Items.HideUnused = False
                Items = <>
                Caption = '-'
              end
              item
                Items.HideUnused = False
                Items = <>
                Action = actExit
                Caption = #36864#20986'(&Z)'
                ImageIndex = 5
              end>
            Caption = #31995#32479'(&S)'
          end
          item
            Items.HideUnused = False
            Items = <
              item
                Items.HideUnused = False
                Items = <>
                Action = actManExecSql
                Caption = #25191#34892'(&W)'
                ImageIndex = 1
                ShortCut = 119
              end
              item
                Items.HideUnused = False
                Items = <>
                Action = actManStopExec
                Caption = #20572#27490'(&S)'
                ImageIndex = 2
              end
              item
                Items.HideUnused = False
                Items = <>
                Action = actManEditData
                Caption = #32534#36753#25968#25454'(&Z)'
                ShortCut = 118
              end
              item
                Items.HideUnused = False
                Items = <>
                Caption = '-'
              end
              item
                Items.HideUnused = False
                Items = <>
                Action = actExecScript
                Caption = #25191#34892#33050#26412'(&X)...'
              end
              item
                Items.HideUnused = False
                Items = <>
                Action = actExportDB
                Caption = #23548#20986#25968#25454'(&Y)'
              end>
            Caption = #32534#36753'(&E)'
          end
          item
            Items.HideUnused = False
            Items = <
              item
                Items.HideUnused = False
                Items = <>
                Action = actManShowTree
                Caption = #26174#31034#27983#35272#26641'(&Z)'
              end
              item
                Items.HideUnused = False
                Items = <>
                Action = actManAutoShowError
                Caption = #33258#21160#26174#31034#38169#35823'(&W)'
              end
              item
                Items.HideUnused = False
                Items = <>
                Action = actManOnTop
                Caption = #24635#22312#26368#19978'(&V)'
              end
              item
                Items.HideUnused = False
                Items = <>
                Caption = '-'
              end
              item
                Items.HideUnused = False
                Items = <>
                Action = actManRefreshTree
                Caption = #21047#26032#26641'(&T)'
                ImageIndex = 4
              end>
            Caption = #26597#30475'(&V)'
          end
          item
            Items.HideUnused = False
            Items = <
              item
                Items.HideUnused = False
                Items = <>
                Action = actManDBParams
                Caption = #25968#25454#24211#21442#25968'(&O)...'
                ImageIndex = 3
                ShortCut = 16463
              end>
            Caption = #35774#32622'(&O)'
          end
          item
            Items.HideUnused = False
            Items = <
              item
                Items.HideUnused = False
                Items = <>
                Caption = '&actPlugin'
              end>
            Caption = #24037#20855'(&T)'
          end
          item
            Items.HideUnused = False
            Items = <
              item
                Items.HideUnused = False
                Items = <>
                Action = actCommandHelp
                Caption = #21629#20196#24110#21161'(&Y)'
              end
              item
                Items.HideUnused = False
                Items = <>
                Action = actShowSettings
                Caption = #26174#31034#24403#21069#35774#32622'(&Z)'
              end
              item
                Items.HideUnused = False
                Items = <>
                Caption = '-'
              end
              item
                Items.HideUnused = False
                Items = <
                  item
                    Items.HideUnused = False
                    Items = <>
                    Caption = '&ActionClientItem0'
                  end>
                Caption = 'DBAHelp(&H)'
              end>
            Caption = #39640#32423'(&A)'
          end
          item
            Items.HideUnused = False
            Items = <
              item
                Items.HideUnused = False
                Items = <>
                Action = actTip
                Caption = #25552#31034'(&Y)...'
              end
              item
                Items.HideUnused = False
                Items = <>
                Caption = '-'
              end
              item
                Items.HideUnused = False
                Items = <>
                Action = actManAbout
                Caption = #20851#20110'(&A)'
              end>
            Caption = #24110#21161'(&H)'
          end>
        ActionBar = actmmb1
      end
      item
        Items = <
          item
            Action = actManConnect
            Caption = #36830#25509'(&U)'
            ImageIndex = 6
            ShowShortCut = False
          end
          item
            Action = actManDBParams
            Caption = #21442#25968'(&V)'
            ImageIndex = 3
          end
          item
            Action = actManRefreshTree
            Caption = #21047#26032#26641'(&W)'
            ImageIndex = 26
            ShowShortCut = False
          end
          item
            Action = actManExecSql
            Caption = #25191#34892'(&X)'
            ImageIndex = 1
            ShowShortCut = False
          end
          item
            Action = actManStopExec
            Caption = #20572#27490'(&Y)'
            ImageIndex = 2
            ShowShortCut = False
          end
          item
            Action = actManExit
            Caption = #36864#20986'(&Z)'
            ImageIndex = 5
          end>
        ActionBar = acttb1
      end>
    LinkedActionLists = <
      item
        ActionList = actlstMain
        Caption = 'actlstMain'
      end>
    Images = DataModule1.ilBtnIcons
    Left = 96
    Top = 200
    StyleName = 'XP Style'
    object actManConnect: TAction
      Category = #31995#32479
      Caption = #36830#25509
      ImageIndex = 6
      OnExecute = actConnectExecute
    end
    object actManOpenWithEditor: TAction
      Category = #31995#32479
      Caption = #22312#32534#36753#22120#20013#25171#24320'...'
      OnExecute = actOpenWithEditorExecute
    end
    object actManSaveEditorText: TAction
      Category = #31995#32479
      Caption = #20445#23384#32534#36753#22120#20869#23481
    end
    object actManAbout: TAction
      Category = #24110#21161
      Caption = #20851#20110
      OnExecute = actAboutExecute
    end
    object actManEditData: TAction
      Category = #32534#36753
      Caption = #32534#36753#25968#25454
      ImageIndex = 4
      ShortCut = 118
      OnExecute = actEditDataExecute
    end
    object actManShowTree: TAction
      Category = #26597#30475
      Caption = #26174#31034#27983#35272#26641
      Checked = True
      OnExecute = actManShowTreeExecute
    end
    object actManAutoShowError: TAction
      Category = #26597#30475
      Caption = #33258#21160#26174#31034#38169#35823
      Checked = True
      OnExecute = actManAutoShowErrorExecute
    end
    object actManOnTop: TAction
      Category = #26597#30475
      Caption = #24635#22312#26368#19978
      OnExecute = actManOnTopExecute
    end
    object actManRefreshTree: TAction
      Category = #26597#30475
      Caption = #21047#26032#26641
      ImageIndex = 26
      OnExecute = actRefreshTreeExecute
    end
    object actManExecSql: TAction
      Category = #32534#36753
      Caption = #25191#34892
      ImageIndex = 1
      OnExecute = actExecSelSqlExecute
    end
    object actManStopExec: TAction
      Category = #32534#36753
      Caption = #20572#27490
      Enabled = False
      ImageIndex = 2
      OnExecute = actStopExecExecute
    end
    object actManSaveEditorTextAs: TAction
      Category = #31995#32479
      Caption = #21478#23384#32534#36753#22120#20869#23481#20026'...'
      OnExecute = actSaveEditorTextAsExecute
    end
    object actManDBParams: TAction
      Category = #35774#32622
      Caption = #21442#25968
      ImageIndex = 12
      OnExecute = actDBParamsExecute
    end
    object actManExecScript: TAction
      Category = #32534#36753
      Caption = #25191#34892#33050#26412
      OnExecute = actExecScriptExecute
    end
    object actManExit: TAction
      Category = #31995#32479
      Caption = #36864#20986
      ImageIndex = 5
      OnExecute = actExitExecute
    end
  end
  object pmBars: TPopupMenu
    Left = 96
    Top = 160
    object mniShowMenu: TMenuItem
      Caption = #26174#31034#33756#21333
      OnClick = mniShowMenuClick
    end
    object mniShowBar: TMenuItem
      Caption = #26174#31034#24037#20855#26639
      OnClick = mniShowBarClick
    end
    object mniShowTree: TMenuItem
      Caption = #26174#31034#27983#35272#26641
      OnClick = mniShowTreeClick
    end
  end
  object con1: TADOConnection
    Left = 48
    Top = 336
  end
  object Database1: TDatabase
    SessionName = 'Default'
    Left = 48
    Top = 368
  end
  object pmTab: TPopupMenu
    AutoHotkeys = maManual
    OnPopup = pmTabPopup
    Left = 272
    Top = 88
    object pmiCloseTab: TMenuItem
      Caption = #20851#38381
      OnClick = pmiCloseTabClick
    end
    object pmiRename: TMenuItem
      Caption = #37325#21629#21517
      OnClick = pmiRenameClick
    end
  end
  object OpenDialog1: TOpenDialog
    Left = 104
    Top = 368
  end
end
