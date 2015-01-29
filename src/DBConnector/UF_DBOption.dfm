inherited F_DBOption: TF_DBOption
  Left = 325
  Top = 127
  Caption = 'F_DBOption'
  ClientHeight = 513
  ClientWidth = 530
  OldCreateOrder = True
  OnCreate = FormCreate
  ExplicitWidth = 536
  ExplicitHeight = 541
  PixelsPerInch = 96
  TextHeight = 12
  inherited panelTitleArea: TPanel
    Width = 530
    ExplicitWidth = 530
    inherited Bevel1: TBevel
      Width = 530
      ExplicitWidth = 530
    end
    inherited Image1: TImage
      Left = 482
      Height = 57
      ExplicitLeft = 482
      ExplicitHeight = 48
    end
  end
  inherited panelCtlArea: TPanel
    Top = 466
    Width = 530
    ExplicitTop = 466
    ExplicitWidth = 530
    DesignSize = (
      530
      47)
    inherited Bevel2: TBevel
      Width = 530
      ExplicitWidth = 530
    end
    inherited labVer: TLabel
      Left = 506
      ExplicitLeft = 506
    end
    inherited btnCloseMe: TButton
      Left = 455
      TabOrder = 3
      ExplicitLeft = 455
    end
    inherited btnOK: TButton
      Left = 376
      Default = False
      TabOrder = 2
      OnClick = btnOKClick
      ExplicitLeft = 376
    end
    object btnApply: TButton
      Left = 296
      Top = 20
      Width = 70
      Height = 23
      Anchors = [akRight, akBottom]
      Caption = #24212#29992
      TabOrder = 1
      OnClick = btnApplyClick
    end
    object btnTest: TButton
      Left = 209
      Top = 20
      Width = 70
      Height = 23
      Anchors = [akRight, akBottom]
      Caption = #27979#35797#36830#25509
      TabOrder = 0
      Visible = False
      OnClick = btnTestClick
    end
  end
  inherited panelMainArea: TPanel
    Width = 530
    Height = 407
    ExplicitWidth = 530
    ExplicitHeight = 407
    object pgcOptions: TPageControl
      Left = 4
      Top = 201
      Width = 522
      Height = 202
      ActivePage = tsSqlite
      Align = alClient
      MultiLine = True
      TabOrder = 1
      OnChange = pgcOptionsChange
      object ts1: TTabSheet
        Caption = 'Access'
        ExplicitLeft = 0
        ExplicitTop = 0
        ExplicitWidth = 0
        ExplicitHeight = 0
        object lbl1: TLabel
          Left = 16
          Top = 20
          Width = 48
          Height = 12
          Caption = #25968#25454#28304#65306
        end
        object lbl6: TLabel
          Left = 16
          Top = 50
          Width = 48
          Height = 12
          Caption = #24037#20316#32452#65306
        end
        object edtAcsSource: TEdit
          Left = 76
          Top = 20
          Width = 350
          Height = 20
          TabOrder = 1
        end
        object edtAcsSec: TEdit
          Left = 76
          Top = 50
          Width = 350
          Height = 20
          TabOrder = 3
        end
        object btnChooseDB: TButton
          Left = 439
          Top = 18
          Width = 27
          Height = 25
          Caption = '...'
          TabOrder = 0
          OnClick = btnChooseDBClick
        end
        object btnChooseSec: TButton
          Left = 439
          Top = 48
          Width = 27
          Height = 25
          Caption = '...'
          TabOrder = 2
          OnClick = btnChooseSecClick
        end
        object chkAcs2007: TCheckBox
          Left = 16
          Top = 96
          Width = 97
          Height = 17
          Caption = 'Access2007'
          TabOrder = 4
        end
      end
      object ts2: TTabSheet
        Caption = 'Oracle'
        ImageIndex = 1
        ExplicitLeft = 0
        ExplicitTop = 0
        ExplicitWidth = 0
        ExplicitHeight = 0
        object lblOraSID: TLabel
          Left = 16
          Top = 20
          Width = 48
          Height = 12
          Caption = #26381#21153#21517#65306
        end
        object lblOraIP: TLabel
          Left = 16
          Top = 66
          Width = 36
          Height = 12
          Caption = #20027#26426#65306
        end
        object lblOraPort: TLabel
          Left = 292
          Top = 96
          Width = 36
          Height = 12
          Caption = #31471#21475#65306
        end
        object lblOraPro: TLabel
          Left = 16
          Top = 92
          Width = 36
          Height = 12
          Caption = #21327#35758#65306
        end
        object lbl10: TLabel
          Left = 16
          Top = 128
          Width = 96
          Height = 12
          Caption = #38656#35201'oracle'#23458#25143#31471
        end
        object Label4: TLabel
          Left = 380
          Top = 96
          Width = 48
          Height = 12
          Caption = #40664#35748'1521'
        end
        object edtOraIP: TEdit
          Left = 70
          Top = 66
          Width = 361
          Height = 20
          TabOrder = 2
        end
        object edtOraPort: TEdit
          Left = 330
          Top = 94
          Width = 47
          Height = 20
          TabOrder = 4
          Text = '1521'
        end
        object cbbOraPro: TComboBox
          Left = 70
          Top = 90
          Width = 80
          Height = 20
          Style = csDropDownList
          ItemIndex = 1
          TabOrder = 3
          Text = 'TCP'
          Items.Strings = (
            ''
            'TCP'
            'TCPS'
            'IPC')
        end
        object chkIsLocalTNS: TCheckBox
          Left = 70
          Top = 45
          Width = 110
          Height = 17
          Caption = #26159#21542#26412#22320#26381#21153#21517
          TabOrder = 1
          OnClick = chkIsLocalTNSClick
        end
        object cbbOraSID: TComboBox
          Left = 70
          Top = 19
          Width = 361
          Height = 20
          TabOrder = 0
          OnClick = cbbOraSIDClick
        end
      end
      object tsSybase: TTabSheet
        Caption = 'Sybase'
        ImageIndex = 2
        ExplicitLeft = 0
        ExplicitTop = 0
        ExplicitWidth = 0
        ExplicitHeight = 0
        object lbl9: TLabel
          Left = 16
          Top = 128
          Width = 84
          Height = 12
          Caption = #38656#35201'Sybase Dll'
        end
        object Label5: TLabel
          Left = 16
          Top = 49
          Width = 36
          Height = 12
          Caption = #31471#21475#65306
        end
        object Label6: TLabel
          Left = 16
          Top = 24
          Width = 60
          Height = 12
          Caption = #26381#21153#22320#22336#65306
        end
        object Label7: TLabel
          Left = 16
          Top = 76
          Width = 72
          Height = 12
          Caption = #40664#35748#25968#25454#24211#65306
        end
        object edtSybServerName: TEdit
          Left = 88
          Top = 20
          Width = 350
          Height = 20
          TabOrder = 0
        end
        object edtSybPort: TEdit
          Left = 88
          Top = 47
          Width = 145
          Height = 20
          TabOrder = 1
        end
        object cbbSybDataBase: TComboBox
          Left = 88
          Top = 72
          Width = 352
          Height = 20
          TabOrder = 2
        end
      end
      object tsMySql: TTabSheet
        Caption = 'MySql'
        ImageIndex = 3
        ExplicitLeft = 0
        ExplicitTop = 0
        ExplicitWidth = 0
        ExplicitHeight = 0
        object Label1: TLabel
          Left = 16
          Top = 20
          Width = 72
          Height = 12
          Caption = 'ODBC'#25968#25454#28304#65306
        end
        object Label2: TLabel
          Left = 17
          Top = 49
          Width = 48
          Height = 12
          Caption = #25968#25454#24211#65306
        end
        object Label3: TLabel
          Left = 16
          Top = 128
          Width = 300
          Height = 12
          Caption = #38656#35201#23433#35013'MySqlDDBCConnector'#65292#24182#22312'odbc'#20013#37197#32622#20102#25968#25454#28304
        end
        object Label8: TLabel
          Left = 16
          Top = 80
          Width = 48
          Height = 12
          Caption = #23383#31526#38598#65306
        end
        object edtmslHost: TComboBox
          Left = 88
          Top = 19
          Width = 335
          Height = 20
          TabOrder = 0
        end
        object edtmslDataBase: TComboBox
          Left = 88
          Top = 47
          Width = 335
          Height = 20
          TabOrder = 1
          OnDropDown = edtmslDataBaseDropDown
        end
        object edtMysqlCharset: TEdit
          Left = 88
          Top = 72
          Width = 121
          Height = 20
          TabOrder = 2
          Text = 'GBK'
        end
      end
      object tsSqlite: TTabSheet
        Caption = 'SQLite'
        ImageIndex = 4
        ExplicitLeft = 0
        ExplicitTop = 0
        ExplicitWidth = 0
        ExplicitHeight = 0
        object Label9: TLabel
          Left = 16
          Top = 20
          Width = 36
          Height = 12
          Caption = #25991#20214#65306
        end
        object edtSqliteDBFile: TEdit
          Left = 76
          Top = 20
          Width = 350
          Height = 20
          TabOrder = 0
        end
        object btnChooseSqliteDB: TButton
          Left = 439
          Top = 18
          Width = 27
          Height = 25
          Caption = '...'
          TabOrder = 1
          OnClick = btnChooseDBClick
        end
      end
    end
    object pnl1: TPanel
      Left = 4
      Top = 4
      Width = 522
      Height = 197
      Align = alTop
      BevelOuter = bvNone
      Font.Charset = GB2312_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = #23435#20307
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      object btnReadConfig: TButton
        Left = 12
        Top = 2
        Width = 82
        Height = 24
        Caption = #35835#21462#37197#32622'...'
        TabOrder = 0
        OnClick = btnReadConfigClick
      end
      object grp1: TGroupBox
        Left = 8
        Top = 84
        Width = 250
        Height = 100
        Caption = #25968#25454#24211
        TabOrder = 2
        object lbl2: TLabel
          Left = 8
          Top = 30
          Width = 60
          Height = 12
          Caption = #25968#25454#24211#31867#22411
        end
        object lbl5: TLabel
          Left = 18
          Top = 60
          Width = 48
          Height = 12
          Caption = #24341#25806#31867#22411
        end
        object cbbDBType: TComboBox
          Left = 81
          Top = 30
          Width = 150
          Height = 20
          Style = csDropDownList
          TabOrder = 0
          OnChange = cbbDBTypeChange
          Items.Strings = (
            ''
            'Access'
            'Oracle'
            'Sybase'
            'MySQL')
        end
        object cbbEngineType: TComboBox
          Left = 81
          Top = 60
          Width = 150
          Height = 20
          Style = csDropDownList
          TabOrder = 1
          OnChange = cbbEngineTypeChange
          Items.Strings = (
            #33258#21160#36873#25321
            'ADO'#39537#21160
            'BDE'#39537#21160
            'DBExpress'#39537#21160
            'SQLite'#39537#21160)
        end
      end
      object grp2: TGroupBox
        Left = 264
        Top = 84
        Width = 250
        Height = 100
        Caption = #29992#25143
        TabOrder = 3
        object lbl7: TLabel
          Left = 29
          Top = 30
          Width = 36
          Height = 12
          Caption = #29992#25143#21517
        end
        object lbl8: TLabel
          Left = 41
          Top = 60
          Width = 24
          Height = 12
          Caption = #23494#30721
        end
        object cbbUser: TComboBox
          Left = 75
          Top = 30
          Width = 150
          Height = 20
          TabOrder = 0
        end
        object edtPwd: TEdit
          Left = 75
          Top = 60
          Width = 150
          Height = 20
          PasswordChar = '*'
          TabOrder = 1
        end
      end
      object GroupBox1: TGroupBox
        Left = 8
        Top = 32
        Width = 506
        Height = 49
        Caption = #35774#32622
        TabOrder = 1
        object lbl11: TLabel
          Left = 16
          Top = 23
          Width = 96
          Height = 12
          Caption = #26368#36817#30340#25968#25454#24211#20010#25968
        end
        object edtRecentDBCount: TEdit
          Left = 125
          Top = 18
          Width = 102
          Height = 20
          TabOrder = 1
          OnKeyPress = edtRecentDBCountKeyPress
        end
        object chkSingle: TCheckBox
          Left = 280
          Top = 16
          Width = 177
          Height = 17
          Caption = #21482#20801#35768#21551#21160#19968#20010'DBConnector'
          TabOrder = 0
        end
      end
    end
  end
  object pmConfig: TPopupMenu
    AutoHotkeys = maManual
    Left = 112
    Top = 62
  end
  object ilTree: TImageList
    Height = 24
    Width = 24
    Left = 464
    Top = 256
    Bitmap = {
      494C010101000500040018001800FFFFFFFFFF10FFFFFFFFFFFFFFFF424D3600
      0000000000003600000028000000600000001800000001002000000000000024
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000FEFEFE000000000000000000F8F8F800F1F1F100EFEFEF00F8F8F800FBFB
      FB00FAFAFA00F9F9F900FAFAFA00000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000E8E8E800E8E8E800E9E9E900484847006A6A6900A9A9A800B4B4B300AFAE
      AE00A4A4A300A3A3A200DADADA00F4F4F4000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000009393
      920041403F0066666500585856007A7975005C5B5800605F5B00A2A19D00D7D6
      D400E3E3E300C8C7C50086868400DFDFDF0000000000FEFEFE00FEFEFE00FBFB
      FB00FAFAFA00FAFAFA0000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000F9F9F900FCFCFC008B8A
      87002A292900312E2D0048474600B5B3AF00C6C4C2005F5D5900C9C8C400918F
      8C00E2E3E200CBCAC90052514C00FFFFFF00FAF7F700F6F4F000F4EEE900F5ED
      E200F4E7DB00F4EDE700FDFDFD00000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FEFEFE00CECECE00CFD0D1008280
      7C002A2927002A272500C9C8C700ACABA900706E6B005D5C5700B6B4B2008B89
      8600E3E2E200D3D2CF0056555300FFE6C800F7E6D100F6DEC800F6D7BD00F9E4
      CD00F9E7CE00F5E8DC00FCFCFC00000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FEFEFE00FFFDF600FFECD5007C7B
      77002928270028272500EAE9E90054524E004B4945004F4E4B0059585400B5B4
      B200E4E4E400D5D3D20049494700FFEDD400F7DCC000F9EBD500F9E7D000FAEC
      D500F8E3C700F5E7D600FCFCFC00000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000FEFEFE00FFFFFC00FFF3DE007C7B
      77002A2927007C7B7A00EBEBEB00BBBAB9007978750081807D00DEDEDE00ECEC
      EC00EFEFEF00E1E1E00043434200FFE3BC00F6D1A400F6D7B200F6DBB800F1C1
      9100F2BD8A00F2CDAA00F9F9F900000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000FFFFFF00FED7B1005051
      4E0023232100C3C3C200CDCDCC004C4A48006D6D6A004B494900DADADA00B3B2
      B000F7F7F700C0BFBC004B4A4500FCD3A100CFC5AF00F5C49000F1BF8C00F0B6
      8100F0B17C00F0C29B00F8F8F800000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000FFFFFF00F2B98700D9C5
      AD00C5C4C300C1C1BF00DADAD800BAB9B700C4C3C300D2D2D100C5C5C400C7C7
      C600BFC1BF00BCBDBD00625F5B00FDCB90009E878B00DDB18600EFB67F00E9AC
      7400E9A66F00E7B18400F6F6F600000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000FCFCFE00F9E2CD00FDFC
      F000D4B79600CAB59B00C6A48000D7B18300F2B77A00EFB77900EEBA7B00F0BD
      7F00F5C08000E0AF8000515A8D002A6EA5004F73B300666A7F001A2D4B00E3A2
      6000DC985A00DA955F00F2F2F200000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000E9E9EA00FDFBF000FCFA
      EE00E6A26300E4A15F00E4A46200E4A76500E3AA6600E3AC6600EBB56800D3A0
      62005C6B820048BBE5009AAE9400EAB56C0032D5FF00000014000A378900DF9F
      5900E0994E00D68C4C00F0F0F000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000D1D1D100FFFEF300F9E5
      C800D9904F00D8925000D9965100DCA15800DDA45C00E1AB5C006F93AA001A8E
      D000307B9C00E7B36100E7B96900F0C36F0048DAFF0002000500072D8A00D091
      4E009A7B7200BE7E5100ECECEC00000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000C5C6C700EEC5A500F5D2
      AB00CD7D3900CD7F3B00D1864100D6934800D8994B00DC9E4A00E0A75000E1B1
      5900E0B35B00DFB15900E1B45B00BF924D002E72AF008A7A5600A2876400D599
      4E00D5954E00CC813800DBDBDB00000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000BBBCBC00FFFBED00FBF8
      E400F0C49B00C8763100CA7B3600D0863B00D48E3E00D5944200DA9C4500DBA1
      4B00DFA95000E0AD5300E3AE52009B6D30000A074900E5D2B900FAE3C200E8BB
      7E00F0CAA000E7AF7500D1D1D100000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000C1C1C100FFFFFE00FCF8
      EC00C9783000C46F2B00C6742E00CC7E3200D1833600D4883700D7903A00D994
      3C00DB983E00E0A54900E3A64600A29676008883AA00FFFFFF00FDFDFB00EDC6
      8D00F9DFCB00F2C39700C8C8C800000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000CCCCCC00FFFFFF00FFFF
      FF00F4D8B500CC7D3000C4692500CA732900CF762A00D27C2A00D7832C00DB86
      2D00DD892E00DE8B2F00E28F3200BC742800B5ABA300FDF7DE00E4BA6800CC71
      2800C86A2400C1621B00C3C3C400FEFEFE000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000DADADA00FFFFFF00D07F
      4000E6AA6A00C46E2900C2652200CA6C2400CD722500D2742800D97B2800DC7D
      2A00DE802900E1852C00E1862C00E5832400CD812400D9852900D3712200CA69
      2300C7652000BF5B1900C0C1C100FDFDFD000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000E7E7E700FFFFFF00AD45
      0E00BC5F1F00BD5C1F00C05F1F00C8682200CE6C2400D2702400D8762500DD79
      2400DE7A2400E27C2400E27C2200E3791F00DF721A00D5691600D0631300C659
      1000C35C1A00C2612700BCBDBE00FDFDFD000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000F3F3F300FCFEFF00D49C
      8600DEAC9100E6BAA200EDCAB800FDEADE00FFF9F300FFFFFF00FFFFFF00FDFF
      FF00EDF1F400CCCFD100C8CACB00BDBEBF00C3C3C300C4C5C500CDCDCD00E4E4
      E400EEEEEE00F1F1F100FCFCFC00000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000F9F9F900CCCCCC00CDCF
      CF00BBBCBD00BBBCBC00BFC0C000C7C7C700D1D2D200DEDEDE00EFEFEF00F2F2
      F200F5F5F500F9F9F900FBFBFB00FCFCFC00FEFEFE00FEFEFE00000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000FDFDFD00FAFA
      FA00FDFDFD00FDFDFD00FEFEFE00000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000424D3E000000000000003E000000
      2800000060000000180000000100010000000000200100000000000000000000
      000000000000000000000000FFFFFF00FFFFFF000000000000000000FFFFFF00
      0000000000000000F601FF000000000000000000F000FF000000000000000000
      E000830000000000000000008000010000000000000000000000010000000000
      0000000000000100000000000000000000000100000000000000000080000100
      0000000000000000800001000000000000000000800001000000000000000000
      8000010000000000000000008000010000000000000000008000010000000000
      0000000080000100000000000000000080000100000000000000000080000000
      0000000000000000800000000000000000000000800000000000000000000000
      80000100000000000000000080003F000000000000000000C1FFFF0000000000
      00000000FFFFFF00000000000000000000000000000000000000000000000000
      000000000000}
  end
end
