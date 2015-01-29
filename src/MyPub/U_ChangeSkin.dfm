object F_ChangeSkin: TF_ChangeSkin
  Left = 317
  Top = 233
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsDialog
  Caption = #36873#25321#30382#32932
  ClientHeight = 275
  ClientWidth = 400
  Color = clBtnFace
  Font.Charset = GB2312_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #23435#20307
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 12
  object Label1: TLabel
    Left = 49
    Top = 16
    Width = 65
    Height = 13
    Caption = #24403#21069#30382#32932#65306
    Font.Charset = GB2312_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #23435#20307
    Font.Style = []
    ParentFont = False
  end
  object lblCurrSkin: TLabel
    Left = 120
    Top = 16
    Width = 66
    Height = 12
    Caption = 'lblCurrSkin'
  end
  object btnOk: TBitBtn
    Left = 160
    Top = 240
    Width = 57
    Height = 23
    Caption = #30830#23450
    TabOrder = 3
    OnClick = btnOkClick
  end
  object btnCancel: TBitBtn
    Left = 320
    Top = 240
    Width = 57
    Height = 23
    Caption = #21462#28040
    ModalResult = 2
    TabOrder = 5
    OnClick = btnCancelClick
  end
  object btnPreview: TBitBtn
    Left = 240
    Top = 240
    Width = 57
    Height = 23
    Caption = #39044#35272
    TabOrder = 4
    OnClick = btnPreviewClick
  end
  object lvwSkins: TListView
    Left = 16
    Top = 40
    Width = 89
    Height = 169
    Columns = <
      item
        AutoSize = True
        Caption = #30382#32932#21517#31216
      end>
    ReadOnly = True
    RowSelect = True
    TabOrder = 0
    ViewStyle = vsReport
    OnSelectItem = lvwSkinsSelectItem
  end
  object pnl1: TPanel
    Left = 128
    Top = 40
    Width = 249
    Height = 169
    BevelOuter = bvLowered
    TabOrder = 2
    object imgPreView: TImage
      Left = 1
      Top = 1
      Width = 247
      Height = 167
      Align = alClient
      Stretch = True
    end
  end
  object pnl2: TPanel
    Left = 16
    Top = 216
    Width = 361
    Height = 17
    BevelOuter = bvNone
    TabOrder = 1
    object bvl1: TBevel
      Left = 96
      Top = 10
      Width = 265
      Height = 2
      Shape = bsFrame
    end
    object lbl1: TLabel
      Left = 16
      Top = 6
      Width = 36
      Height = 12
      Caption = 'DBTool'
    end
  end
end
