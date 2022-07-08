{**
 * <p>Title: UFM_DBOperate  </p>
 * <p>Copyright: Copyright (c) 2007/9/1</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 数据库操作frame</p>  
 * <p>Description: 通过创建多个此frame实现多个互相独立的数据库操作页面</p>
 *}
unit UFM_DBOperate;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, fDBGrid, StdCtrls, ExtCtrls, DBCtrls, ComCtrls,
  ToolWin, SynEdit, Menus, DB, ActnList, ADODB, Clipbrd,
  SynEditHighlighter, SynHighlighterSQL, Buttons, UF_ExportResult,
  U_DBConnnectManager, U_FieldInfo, U_ThreadUtil,
  U_DBEngineInterface, U_DBConnectInterface, U_DataMoudle;


type
  TEditorState = (esNone, esOpened, esChaged, esSaved);

  TEditStateChangedEvent = procedure (oldState, newState: TEditorState) of object;

  TExecSqlThread = class(TThread)
  public 
  end;

  TFM_DBOperate = class(TFrame)
    pnlTop: TPanel;
    pnlGrid: TPanel;
    pnl6: TPanel;
    tlbtbr1: TToolBar;
    btnFindAtList: TToolButton;
    btnExport: TToolButton;
    dbnvgr1: TDBNavigator;
    btntbtn1: TToolButton;
    btnPrint: TToolButton;
    btnAdjust: TToolButton;
    btnSortOnColumnClick: TToolButton;
    pnlMsg: TPanel;
    tlb1: TToolBar;
    btnManClearMsg: TToolButton;
    dbmmoShowFieldDetail: TDBMemo;
    pnl4: TPanel;
    pnl15: TPanel;
    bvlBottom2: TBevel;
    bvlBottom1: TBevel;
    pnlBottomContranCloseBtn: TPanel;
    dbgrdData: TfDBGrid;
    splTopClient: TSplitter;
    dsData: TDataSource;
    pmAdjust: TPopupMenu;
    mniAdjustByTitleContent: TMenuItem;
    mniAdjustByTitle: TMenuItem;
    mniAdjustByContent: TMenuItem;
    actlstMain: TActionList;
    actExecSelSql: TAction;
    actExecAllSqls: TAction;
    actLastError: TAction;
    actEditData: TAction;
    actExecScript: TAction;
    actExport: TAction;
    actFieldDetail: TAction;
    actStopExec: TAction;
    SynSQLSyn1: TSynSQLSyn;
    pmMenu: TPopupMenu;
    mniFieldDetail: TMenuItem;
    mniCopyToClip: TMenuItem;
    btnCloseMsg: TSpeedButton;
    pnl1: TPanel;
    btn1: TSpeedButton;
    stat1: TStatusBar;
    pnl2: TPanel;
    tlb2: TToolBar;
    btn2: TToolButton;
    pnl3: TPanel;
    pnlClient: TPanel;
    mmoLog: TMemo;
    syndtSql: TSynEdit;
    pgcClient: TPageControl;
    tsData: TTabSheet;
    tsOutPut: TTabSheet;
    pnl5: TPanel;
    pnl7: TPanel;
    splFieldDetail: TSplitter;
    pmSqlHistory: TPopupMenu;
    pnl8: TPanel;
    pnl9: TPanel;
    btn3: TSpeedButton;
    actExportField: TAction;
    actImportField: TAction;
    tlbField: TToolBar;
    btnExportField: TToolButton;
    btnImportField: TToolButton;
    Panel1: TPanel;
    lblFieldName: TLabel;
    procedure btnFindAtListClick(Sender: TObject);
    procedure actExportExecute(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure btnAdjustClick(Sender: TObject);
    procedure btnSortOnColumnClickClick(Sender: TObject);
    procedure btnManClearMsgClick(Sender: TObject);
    procedure btnCloseMsgClick(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure pnlMsgResize(Sender: TObject);
    procedure mniAdjustByTitleContentClick(Sender: TObject);
    procedure mniAdjustByTitleClick(Sender: TObject);
    procedure mniAdjustByContentClick(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure actFieldDetailExecute(Sender: TObject);
    procedure mmoLogChange(Sender: TObject);
    procedure pmSqlHistoryPopup(Sender: TObject);
    procedure actExportFieldExecute(Sender: TObject);
    procedure actImportFieldExecute(Sender: TObject);
    procedure SelectedField(Sender: TObject);
    procedure syndtSqlChange(Sender: TObject);
    procedure syndtSqlKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
//    FdLastExecTime: TDateTime;                     // sql或命令最后执行时间
    FnPnlTopLastHeight: Integer;                   // 上部面板最后top位置，即sql编辑器面板
    FSQLHistory: TStrings;                         // sql执行历史
    FEditorState: TEditorState;                    // 编辑器状态，见TEditorState
    FEditStateChanged: TEditStateChangedEvent;     // 编辑器状态改变事件
    FExecing: Boolean;
    FExecThread: TCustomThread;

    // 将qry中数据展示在dbgrid中
    procedure ShowQryData(qry: TDataSet);
    // 设置提示面板是否显示
    procedure ShowHintPanel(bShow: Boolean);
    // 将一个popmenu中所有设置为未选状态
    procedure UnCheckSubMnis(pm: TPopupMenu);
    // 将执行时间显示在状态栏
    procedure ShowExecTime;
    // 将记录条数显示在状态栏
    procedure ShowRecordCount;

    // 获取当前记录条数
    function getRecordCount: Integer;
    // 最后执行时间 获取其字符串形式
    function getLastExecTimeStr: string;
    // 获取最后执行的sql
    function GetLastSQL: string;
    // 设置最后执行的sql
    procedure SetLastSQL(value: string);
    // 获取编辑器控件对象
    function GetEditor: TSynEdit;
    // 设置编辑器状态，必须通过此方法，其中包含了触发状态改变事件
    procedure SetEditorState(value: TEditorState);

    // 增加sql历史
    procedure AddHistory(sSql: string);
    // sql历史菜单选中事件（将sql复制到编辑器中）
    procedure SQLHistoryMenuClick(Sender: TObject);
    // 释放本frame保存的dataset对象
    procedure FreeAndNilDataSet;

    function GetCanExport: Boolean;
    procedure DoExecuteLastSql;
    procedure OnExecuteSql(Sender: TObject);
  public
//    CurrPage: Integer;                        // 当前页号，分页功能相关
    DBConnect: IDBConnect;                    // 数据库对象，本frame专用
    FileName: string;                         // 编辑的文件名，在编辑文件模式中有效
    property ExecThread: TCustomThread read FExecThread;
    property LastExecTimeStr: string read getLastExecTimeStr;
    property RecordCount: Integer read getRecordCount;
    property LastSQL: string read GetLastSQL write SetLastSQL;
    property Editor: TSynEdit read GetEditor;
    property EditorState: TEditorState read FEditorState write SetEditorState;
    property EditStateChanged: TEditStateChangedEvent read FEditStateChanged
      write FEditStateChanged;
    property CanExport: Boolean read GetCanExport;
    property Execing: Boolean read FExecing;
  public
    constructor Create(AOwner: TComponent);override;
    destructor Destroy; override;

    // 将query中数据导出
    procedure ExportQuery(AQry: TADOQuery; AFileName: string;
      sSpaceBetween: string = ' ');

    // 执行sql
    function ExecuteSql(sSqlText: string): Integer;
    procedure StopExec();
//    function ExecuteSqlList(sqlList: TStrings):Integer;
    // 在hintpanel上显示信息
    procedure ShowMsg(s: string);

    // 决定需要展示的page  主要决定显示结果集页还是消息页
    procedure CheckShouldShowPage();
    // 是否显示中间的grid部分，不显示时sql编辑器空间会大些，便于编辑
    procedure ShowClient(bShow: Boolean);
    // 显示hintpanel，并显示其data页
    procedure ShowData(bShow: Boolean);

    //////// 事件
    // DBConnect产生log时触发
    procedure OnDBLog(sLog: string);
    // 连接成功后
    procedure OnConnected(Sender: TObject);
    // 断开连接后
    procedure OnDisConnected(Sender: TObject);

    // 重新创建该frame保存的dbconnector对象
    procedure ReCreateDBConnect(dbt: TDBType);
    procedure ShareDBConnect(db: IDBConnect);

    // 刷新UI各控件状态
    procedure RefreshUI;
    // 光标定位到sql编辑器
    procedure ForcusEdit;

    //////// editor operator
    // 将文件加载到编辑器，通常为sql
    procedure LoadFromFile(AFileName: string);
    // 将编辑器内容保存到文件
    procedure SaveToFile(AFileName: string);
  end;

const
  C_nPGCHINT_PAGE_ERROR = 0;
  C_nPGCHINT_PAGE_FIELD = 1;

  C_nPANELTOP_HEIGHT = 100;

var
  SearchTextMethod: TNotifyEvent;

implementation

{$R *.dfm}

uses
  UF_Find, U_CommonFunc, U_UIUtil, U_Pub, U_Const,
  U_fStrUtil, U_ini, U_SqlUtils, U_ExportUtil,
  U_WaitingForm;

const
  C_nSQLHistory_MaxShowCount = 15;
  C_PGCMAIN_PAGE_DATA = 0;
  C_PGCMAIN_PAGE_OUTPUT = 1;

constructor TFM_DBOperate.Create(AOwner: TComponent);
begin
  inherited;
  ShowClient(false);
//  CurrPage := 1;
  FnPnlTopLastHeight := 0;
  FSQLHistory := TStringList.Create;
  RefreshUI;
  dbgrdData.OnSelectedField := SelectedField;
  EditorState := esNone;
  g_Global.RegisterFrame(Self);
  FExecThread := TCustomThread.Create(TCustomThreadTask.Create(DoExecuteLastSql), False);
end;

procedure TFM_DBOperate.ReCreateDBConnect(dbt: TDBType);
begin
  if Assigned(DBConnect) then begin
    DBConnect._Release;
    DBConnect := nil;
  end;
  DBConnect := TDBConnectManager.CreateDBConnect(dbt);
  DBConnect.OnConnected := OnConnected;
  DBConnect.OnDisConnected := OnDisConnected;
  DBConnect.OnLog := OnDBLog;
  DBConnect.OnExecuted := Self.OnExecuteSql;
end;

procedure TFM_DBOperate.ShareDBConnect(db: IDBConnect);
begin
  DBConnect.ShareEngine(db.DBEngine);
end;

destructor TFM_DBOperate.Destroy;
begin
  g_Global.UnRegisterFrame(Self);
  FreeAndNilDataSet;
  if Assigned(FSQLHistory) then
    FreeAndNil(FSQLHistory);
  if Assigned(DBConnect) then begin
    DBConnect._Release;
    DBConnect := nil;
  end;
  if Assigned(FExecThread) then
    FreeAndNil(FExecThread);

  inherited;
end;

procedure TFM_DBOperate.btnFindAtListClick(Sender: TObject);
begin
  if not Assigned(gFindForm) then
    gFindForm := TF_Find.Create(Self);
  with gFindForm do
  begin
    InitControl(TWinControl(dbgrdData), '使用字段内容查找');
    FormStyle := fsStayOnTop;
    Show;
  end;
end;

procedure TFM_DBOperate.actExportExecute(Sender: TObject);
const
  C_sFilterStr = 'sql脚本*.sql|*.sql|文本文件*.txt|*.txt|Excel文件*.xls|*.xls|CSV文件*.csv|*.csv';
  function FitFileExt(sFileName: string; exts: array of String): String;
  var
    i: Integer;
    bMatch: Boolean;
  begin
    Result := sFileName;
    bMatch := False;
    for i := 0 to High(exts) do begin
      if SameText(exts[i], ExtractFileExt(Result)) then begin
        bMatch := True;
        Break;
      end;
    end;
    if not bMatch then
      Result := Result + exts[0];
  end;
  function FitFileName(sFileName: string; dlgExport: TSaveDialog):string;
  begin
    Result := sFileName;
    case dlgExport.FilterIndex of
      1:
      begin
        Result := FitFileExt(Result, ['.sql']);
      end;
      2:
      begin
        Result := FitFileExt(Result, ['.txt']);
      end;
      3:
      begin
        Result := FitFileExt(Result, ['.xls', 'xlsx']);
      end;
      4:
      begin
        Result := FitFileExt(Result, ['.csv']);
      end;
    end;
  end;
var
  sFileName: String;
  qryOld : TADOQuery;
  dlgExport: TSaveDialog;
  waitForm: TWaitingForm;
begin
  // 合法性判断
  if not Assigned(dbgrdData.DataSource.DataSet) then
    Exit;
  with dbgrdData.DataSource.DataSet do begin
    if not Active then
      Exit;

    First;
    if Fields.Count = 0 then
      Exit;
  end;
                                
  actExport.Enabled := False;
  dbgrdData.Columns.BeginUpdate;      
  dlgExport := TSaveDialog.Create(Self);
  Application.ProcessMessages;
  try
    dlgExport.Filter := C_sFilterStr;
    // TODO: 导出的内容不光是表   前缀要分析一下  
    sFileName := TSqlUtils.getTableNameBySQL(LastSQL);  
    if sFileName = '' then
      sFileName := '查询结果'
    else    
      dlgExport.FileName := C_FilePrefix_InsertTable + sFileName;

    dlgExport.InitialDir := GetAppRootPath;
    if not dlgExport.Execute then
      Exit;

    sFileName := FitFileName(dlgExport.FileName, dlgExport);
    qryOld := TADOQuery(dbgrdData.DataSource.DataSet);
    dbgrdData.DataSource.DataSet := nil;                    // DBGrid的DataSet置空，这样导出的速度快

    waitForm := TWaitingForm.Create(Self);
    try
      waitForm.SetMesasge('正在导出...');
      waitForm.Show;
      ExportQuery(qryOld, sFileName);   // 做导出操作的方法
      waitForm.Close;
    finally
      waitForm.Free;
    end;

    with TF_ExportResult.Create(Self) do begin
      ShowExportFile(sFileName);
      Show;
    end;
    
    dbgrdData.DataSource.DataSet := qryOld;                 // 还原DBGrid的DataSet
  finally
    dlgExport.Free;
    dbgrdData.Columns.EndUpdate;
    actExport.Enabled := True;
  end;
end;

procedure TFM_DBOperate.ExportQuery(AQry: TADOQuery; AFileName,
  sSpaceBetween: string);
const
  C_sSeparatorLine = '-';
var
  slstExport: TStringList;
begin
  slstExport := TStringList.Create;
  try
    if SameText(ExtractFileExt(AFileName), '.csv') then begin
      TExportUtil.ExportToCSV(AQry, AFileName);
    end else if SameText(ExtractFileExt(AFileName), '.xls') then begin
      TExportUtil.ExportToExcel(AQry, AFileName);
    end else if SameText(ExtractFileExt(AFileName), '.sql') then begin
      DBConnect.ExportQuery(AQry, AFileName);
    end else begin  // txt
      with AQry do
      begin
//        aiFieldLengths := UIUtil.GetFieldLengths(AQry);
        TExportUtil.FillExportQueryList(AQry, slstExport, '  ', '-');
        slstExport.SaveToFile( AFileName );
      end;
    end;
  finally
    slstExport.Free;
  end;
end;

procedure TFM_DBOperate.ForcusEdit;
begin
  syndtSql.SetFocus;
end;

procedure TFM_DBOperate.FreeAndNilDataSet;
begin
  if Assigned(dsData.DataSet) then
  begin
    dsData.DataSet.Free;
    dsData.DataSet := nil;
  end;
end;

function TFM_DBOperate.GetCanExport: Boolean;
begin
  Result := Assigned(dsData.DataSet) and not dsData.DataSet.Eof;
end;

function TFM_DBOperate.GetEditor: TSynEdit;
begin
  Result := syndtSql;
end;

procedure TFM_DBOperate.btnPrintClick(Sender: TObject);
begin
  if IDYES = FormatMsgBox('是否要打印列表中所有数据？', mbsQuestion, Handle) then
    UIUtil.PrintDbGrid(TDBGrid(dbgrdData), 'title');
end;

procedure TFM_DBOperate.btnAdjustClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to pmAdjust.Items.Count - 1 do
  begin
    if pmAdjust.Items[i].Checked then
    begin
      pmAdjust.Items[i].Click;
      Break;
    end;
  end;
end;

procedure TFM_DBOperate.btnSortOnColumnClickClick(Sender: TObject);
begin
  dbgrdData.SortOnColumnClick := btnSortOnColumnClick.Down;
  RefreshUI;
end;

procedure TFM_DBOperate.btnManClearMsgClick(Sender: TObject);
begin
  mmoLog.Clear;
  mmoLogChange(mmoLog);
  DBConnect.ClearLog;
end;

procedure TFM_DBOperate.DoExecuteLastSql();
begin
  dbmmoShowFieldDetail.DataField := '';

  FExecing := True;
  try
//    FdLastExecTime := Now;
  //  dsData.DataSet.DisableControls;
    Self.DBConnect.ExecSqlText(LastSQL, gC_Delimiter, dsData.DataSet);
  //  dsData.DataSet.EnableControls;
  finally
    FExecing := False;
  end;
end;

procedure TFM_DBOperate.OnExecuteSql(Sender: TObject);
begin
//  SendMessage(Self.Handle, WM_FRAME_SHOW_QUERY, 0, 0);
  if Self.DBConnect.Connected and Self.DBConnect.LastOperSucc then
  begin
    ShowQryData(dsData.DataSet);
//    FdLastExecTime := Now - FdLastExecTime;
    ShowExecTime;
    ShowRecordCount;
  end;
end;

function TFM_DBOperate.ExecuteSql(sSqlText: string): Integer;
begin
  Result := 0;
  if Trim(sSqlText) = '' then
    Exit;

//  if LastSQL <> sSqlText then
//  begin
    LastSQL := sSqlText;
//    CurrPage := 1;
//  end;
  ExecThread.Resume;
  Result := 1;
end;

procedure TFM_DBOperate.StopExec;
begin
  if Assigned(Self.DBConnect) then
    Self.DBConnect.StopExec();
  if Self.Execing then
    ExecThread.Abort();
end; 

procedure TFM_DBOperate.ShowQryData(qry: TDataSet);
begin
  if not Assigned(qry) then Exit;
  if dsData.DataSet <> qry then
    dsData.DataSet := qry;
  
  try
    dbgrdData.FitColumuWidth;
  except
    on ex: Exception do
    ShowMsg(ex.Message);
  end;
  
  if not pnlClient.Visible then
  begin
    ShowClient(True);
  end
  else 
    CheckShouldShowPage;
end;

//procedure TFM_DBOperate.CheckExecError(bShowNoError: Boolean);
//var
//  i: Integer;
//begin
//  if DBConnect.LastError = '' then
//  begin
//   if bShowNoError then
//    ShowMsg('')
//  end
//  else
//  begin
//    mmoLog.Lines.Clear;
//    for i := 0 to DBConnect.Log.Count - 1 do
//      ShowMsg(DBConnect.Log[i]);
//    ShowClient(True);
//    if (dsData.DataSet = nil) or (dsData.DataSet.Eof) then
//      pgcClient.ActivePageIndex := 1;
//  end;
//end;

procedure TFM_DBOperate.OnDBLog(sLog: string);
begin
  if GlobalParams.AutoShowError then
    ShowMsg(sLog);
end; 

procedure TFM_DBOperate.OnConnected(Sender: TObject);
begin  
  FreeAndNilDataSet;
  dsData.DataSet := Self.DBConnect.GetNewQuery;
  g_Global.ShareDB(Self.DBConnect);
  g_Global.RefreshMainUIOnConected;
end;

procedure TFM_DBOperate.OnDisConnected(Sender: TObject);
begin
  FreeAndNilDataSet;
  g_Global.ShareDB(Self.DBConnect);  
  g_Global.RefreshMainUIOnDisConected;
end;

function TFM_DBOperate.getRecordCount: Integer;
begin
  if not Assigned(dbgrdData.DataSource.DataSet)
     or not dbgrdData.DataSource.DataSet.Active then
    Result := 0
  else
    Result := dbgrdData.DataSource.DataSet.RecordCount;
end;

procedure TFM_DBOperate.LoadFromFile(AFileName: string);
begin 
  Editor.Lines.LoadFromFile(AFileName);   
  FileName := AFileName;
  EditorState  := esOpened;
end;  

procedure TFM_DBOperate.SaveToFile(AFileName: string);
begin
  FileName := AFileName;
  Editor.Lines.SaveToFile(AFileName);
  EditorState  := esSaved;
end;

procedure TFM_DBOperate.ShowMsg(s: string);
//var
//  i, nMaxWidth: Integer;
begin  
//  ShowHintPanel(True, C_nPGCHINT_PAGE_ERROR);
  if Trim(s) <> '' then
  begin
  // FormatDateTime('yyyy-mm-dd hh:mm:ss',Now) + ' '
    mmoLog.Lines.Add(s);  

    if not pnlClient.Visible then
    begin
      ShowClient(True);
    end;
  end;

//  nMaxWidth := 0;
//  for i := 0 to lstMsg.Items.Count-1 do
//    if nMaxWidth <
//       lstMsg.Canvas.TextWidth(lstMsg.Items.Strings[i]) then
//      nMaxWidth := lstMsg.Canvas.TextWidth(lstMsg.Items.Strings[i]);
//
//  SendMessage(lstMsg.Handle, LB_SETHORIZONTALEXTENT,
//              Trunc(nMaxWidth * 1.2), 0 );
end;

procedure TFM_DBOperate.ShowHintPanel(bShow: Boolean);
begin
  if bShow then
  begin
    splFieldDetail.Visible := False;

    pnlMsg.Visible := True;
    pnlMsg.Align := alBottom;
    splFieldDetail.Visible := True;
    splFieldDetail.Align := alBottom;
    SelectedField(dbgrdData);
  end
  else
  begin           
    splFieldDetail.Visible := False; 
    pnlMsg.Visible := False;
  end;
end;

procedure TFM_DBOperate.SelectedField(Sender: TObject);
var
  selectedField: TField;
begin
  selectedField := dbgrdData.SelectedField;
  tlbField.Visible := selectedField.DataType in [ftOraBlob, ftBlob];
  actExportField.Enabled := (selectedField.DataType in [ftOraBlob, ftBlob])
    and not selectedField.IsNull;
  actImportField.Enabled := selectedField.DataType in [ftOraBlob, ftBlob];
  if pnlMsg.Visible then
  begin
    lblFieldName.Caption := Format('字段: %s  类型: %s',[selectedField.FieldName,
      FieldTypeToStr(selectedField.DataType, False)]);
    dbmmoShowFieldDetail.DataField := selectedField.FieldName;
  end;
end;

function TFM_DBOperate.getLastExecTimeStr: string;
var
  d: Double;
begin
  d := DBConnect.GetLastElapsedMilis;
  Result := commonFunc.PrecisionNumStr(FloatToStr(d), 3) + '毫秒';
end;

procedure TFM_DBOperate.btnCloseMsgClick(Sender: TObject);
begin
  ShowHintPanel(False);
end;

procedure TFM_DBOperate.btn1Click(Sender: TObject);
begin
  FnPnlTopLastHeight := pnlTop.Height;
  ShowClient(False);
end;

procedure TFM_DBOperate.pnlMsgResize(Sender: TObject);
  function GetNoLessThanZero(n: Integer): Integer;
  begin
    if n > 0 then  Result := n
    else Result := 0;
  end;
begin
  bvlBottom1.Top := GetNoLessThanZero(pnlBottomContranCloseBtn.Height + 5);
  bvlBottom1.Height := GetNoLessThanZero(pnlMsg.Height - pnlBottomContranCloseBtn.Height);
  bvlBottom2.Top := bvlBottom1.Top;
  bvlBottom2.Height := bvlBottom1.Height;
end;

procedure TFM_DBOperate.RefreshUI;
begin
  if btnSortOnColumnClick.Down then
    btnSortOnColumnClick.Hint := '启用了点击列头排序，按住Shift实现多列排序'
  else
    btnSortOnColumnClick.Hint := '未启用点击列头排序';
  UnCheckSubMnis(pmAdjust);  
  case dbgrdData.FitColumnMode of
    fcmTitle: mniAdjustByTitle.Checked := True;
    fcmContent: mniAdjustByContent.Checked := True;
    fcmBoth: mniAdjustByTitleContent.Checked := True;
  end;
end;

procedure TFM_DBOperate.mniAdjustByTitleContentClick(Sender: TObject);
begin
  dbgrdData.FitColumnMode := fcmBoth;
  dbgrdData.FitColumuWidth;

  UnCheckSubMnis(pmAdjust);  
  TMenuItem(Sender).Checked := True;
end;

procedure TFM_DBOperate.mniAdjustByTitleClick(Sender: TObject);
begin
  dbgrdData.FitColumnMode := fcmTitle;
  dbgrdData.FitColumuWidth;

  UnCheckSubMnis(pmAdjust);  
  TMenuItem(Sender).Checked := True;
end;

procedure TFM_DBOperate.mmoLogChange(Sender: TObject);
begin
  if mmoLog.Lines.Count <> 0 then
    pgcClient.Pages[1].Caption := Format('OutPut(%d)', [mmoLog.Lines.Count])
  else
    pgcClient.Pages[1].Caption := 'OutPut';
end;

procedure TFM_DBOperate.mniAdjustByContentClick(Sender: TObject);
begin
  dbgrdData.FitColumnMode := fcmContent;
  dbgrdData.FitColumuWidth;
  
  UnCheckSubMnis(pmAdjust);   
  TMenuItem(Sender).Checked := True;
end;

procedure TFM_DBOperate.UnCheckSubMnis(pm: TPopupMenu);
var
  i: Integer;
begin
  for i := 0 to pm.Items.Count - 1 do
    pm.Items[i].Checked := False;
end;

procedure TFM_DBOperate.ShowExecTime;
begin
  stat1.Panels[0].Text := getLastExecTimeStr;
end;

procedure TFM_DBOperate.ShowRecordCount;
var
  nRecCount: Integer;
  sFormat: string;
begin
  if not Assigned(dsData.DataSet) then Exit;

  sFormat := '记录数：%d';
  if dsData.DataSet.Active then
  begin
    nRecCount := dsData.DataSet.RecordCount;
    if (nRecCount > 0) and (nRecCount = GlobalParams.MaxRecord) then
      sFormat := '记录数：≥%d';  
    stat1.Panels[2].Text := Format('字段数：%d',[dsData.DataSet.FieldCount]);
  end
  else
    nRecCount := 0;
  stat1.Panels[1].Text := Format(sFormat, [nRecCount]);
  if stat1.Panels[1].Width < stat1.Canvas.TextWidth(stat1.Panels[1].Text) then
     stat1.Panels[1].Width := stat1.Canvas.TextWidth(stat1.Panels[1].Text);
end;

procedure TFM_DBOperate.btn2Click(Sender: TObject);
begin
  if LastSQL <> '' then
  begin
    if (syndtSql.Lines.Count = 0) or (syndtSql.Lines[syndtSql.Lines.Count-1] <> '') then
      syndtSql.Lines.Add(LastSQL)
    else
      syndtSql.Lines[syndtSql.Lines.Count-1] := LastSQL;
  end;
end;

procedure TFM_DBOperate.actExportFieldExecute(Sender: TObject);
var
  blobfield: TBlobField;
  stream: TMemoryStream;
begin
  with TSaveDialog.Create(Self) do
  try
    FileName := dbgrdData.SelectedField.FieldName;
    if Execute() then
    begin
      blobfield := dbgrdData.SelectedField as TBlobField;
      stream := TMemoryStream.Create;
      try
        blobfield.SaveToStream(stream);
        stream.SaveToFile(FileName);
      finally
        stream.Free;
      end;
    end;
  finally
    Free;
  end;
end;     

procedure TFM_DBOperate.actImportFieldExecute(Sender: TObject);
var
  blobfield: TBlobField;
  stream: TMemoryStream;
begin
  with TOpenDialog.Create(Self) do
  try
    if Execute() then
    begin
      dsData.DataSet.Edit;
      blobfield := dbgrdData.SelectedField as TBlobField;
      stream := TMemoryStream.Create;
      try
        stream.LoadFromFile(FileName);
        blobfield.LoadFromStream(stream);
      finally
        stream.Free;
      end;
    end;
  finally
    Free;
  end;
end;

procedure TFM_DBOperate.actFieldDetailExecute(Sender: TObject);
begin
  ShowHintPanel(True);
end;

procedure TFM_DBOperate.CheckShouldShowPage();
begin     
  // sql执行错误 始终显示输出结果页
  // sql执行成功 根据情况看显示哪个页面
  if not Assigned(DBConnect) then
    Exit;
  if not DBConnect.LastOperSucc then
  begin
    pgcClient.ActivePageIndex := C_PGCMAIN_PAGE_OUTPUT
  end
  else
  begin
    if DBConnect.LastSQLType in [dbcstCommand, dbcstUpdate] then
      pgcClient.ActivePageIndex := C_PGCMAIN_PAGE_OUTPUT
    else if DBConnect.LastSQLType = dbcstQuery then
      pgcClient.ActivePageIndex := C_PGCMAIN_PAGE_DATA;
  end;
end;  

procedure TFM_DBOperate.ShowClient(bShow: Boolean);
begin
  if bShow = pnlClient.Visible then
    Exit;
  if bShow then
  begin
    pnlTop.Align := alTop;
    if FnPnlTopLastHeight <> 0 then
      pnlTop.Height := FnPnlTopLastHeight
    else
      pnlTop.Height := C_nPANELTOP_HEIGHT;
    splTopClient.Visible := True;
    splTopClient.Align := alTop;
    pnlClient.Visible := True;
    pnlClient.Align := alClient;
  end
  else
  begin
    pnlClient.Visible := False;
    splTopClient.Visible := False;
    pnlTop.Align := alClient;
  end;
  CheckShouldShowPage;
end;

procedure TFM_DBOperate.ShowData(bShow: Boolean);
begin
  ShowClient(True);
  pgcClient.ActivePageIndex := C_PGCMAIN_PAGE_DATA;
end;

procedure TFM_DBOperate.AddHistory(sSql: string);
var
  nIndex: Integer;
begin
  nIndex := FSQLHistory.IndexOf(sSql);
  if nIndex = -1 then
    FSQLHistory.Add(sSql)
  else
  begin
    FSQLHistory.Add(sSql);
    FSQLHistory.Delete(nIndex);
  end;
  while FSQLHistory.Count > C_nSQLHistory_MaxShowCount do
  begin
    FSQLHistory.Delete(0);
  end;
end;

function TFM_DBOperate.GetLastSQL: string;
begin
  if FSQLHistory.Count > 0 then
  Result := FSQLHistory[FSQLHistory.Count-1];
end;

procedure TFM_DBOperate.SetEditorState(value: TEditorState);
var
  oldStatus: TEditorState;
begin
  if FEditorState <> value then
  begin
    oldStatus := FEditorState;
    FEditorState := value;
    if Assigned(EditStateChanged) then
      EditStateChanged(oldStatus, FEditorState);
  end;  
end;

procedure TFM_DBOperate.SetLastSQL(value: string);
begin
  AddHistory(value);
end;

procedure TFM_DBOperate.SQLHistoryMenuClick(Sender: TObject);
var
  mni: TMenuItem;
  s: string;
begin
  mni := Sender as TMenuItem;
  s := FSQLHistory[mni.Tag];
  if (syndtSql.Lines.Count = 0) or (syndtSql.Lines[syndtSql.Lines.Count-1] <> '') then
    syndtSql.Lines.Add(s)
  else
    syndtSql.Lines[syndtSql.Lines.Count-1] := s;
end;

procedure TFM_DBOperate.syndtSqlChange(Sender: TObject);
begin
  EditorState := esChaged;
end;

procedure TFM_DBOperate.syndtSqlKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if ssCtrl in Shift then
  begin
    // 保存编辑器内容
    if (Key = Ord('s')) or (Key = Ord('S')) then
    begin
      if EditorState in [esChaged] then
        SendMessage(Application.MainForm.Handle, WMUSER_DBCONNECTOR_EDITOR_SAVE, 0, 0);
//        SaveToFile(Self.FileName);
    end;
    // 切换自动换行
    if (Key = Ord('w')) or (Key = Ord('W')) then
    begin
      syndtSql.WordWrap := not syndtSql.WordWrap;
    end;
    // 查找
    if (Key = Ord('f')) or (Key = Ord('F')) then
    begin
      if Assigned(SearchTextMethod) then
        SearchTextMethod(nil);
    end;
  end;
end;

procedure TFM_DBOperate.pmSqlHistoryPopup(Sender: TObject);
var   
  mni: TMenuItem;
  i: Integer;
begin
  pmSqlHistory.Items.Clear;
  for i := FSQLHistory.Count-1 downto 0 do
  begin
    mni := TMenuItem.Create(pmSqlHistory);
    mni.Tag := i;
    if Length(FSQLHistory[i]) < 50 then
      mni.Caption := FSQLHistory[i]
    else
      mni.Caption := Copy(FSQLHistory[i], 1, 47) + '...'; 
    mni.OnClick := SQLHistoryMenuClick;
    pmSqlHistory.Items.Add(mni);
  end;  
end;

end.
