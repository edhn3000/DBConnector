{**
 * <p>Title: UFM_DBOperate  </p>
 * <p>Copyright: Copyright (c) 2007/9/1</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: ���ݿ����frame</p>  
 * <p>Description: ͨ�����������frameʵ�ֶ��������������ݿ����ҳ��</p>
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
//    FdLastExecTime: TDateTime;                     // sql���������ִ��ʱ��
    FnPnlTopLastHeight: Integer;                   // �ϲ�������topλ�ã���sql�༭�����
    FSQLHistory: TStrings;                         // sqlִ����ʷ
    FEditorState: TEditorState;                    // �༭��״̬����TEditorState
    FEditStateChanged: TEditStateChangedEvent;     // �༭��״̬�ı��¼�
    FExecing: Boolean;
    FExecThread: TCustomThread;

    // ��qry������չʾ��dbgrid��
    procedure ShowQryData(qry: TDataSet);
    // ������ʾ����Ƿ���ʾ
    procedure ShowHintPanel(bShow: Boolean);
    // ��һ��popmenu����������Ϊδѡ״̬
    procedure UnCheckSubMnis(pm: TPopupMenu);
    // ��ִ��ʱ����ʾ��״̬��
    procedure ShowExecTime;
    // ����¼������ʾ��״̬��
    procedure ShowRecordCount;

    // ��ȡ��ǰ��¼����
    function getRecordCount: Integer;
    // ���ִ��ʱ�� ��ȡ���ַ�����ʽ
    function getLastExecTimeStr: string;
    // ��ȡ���ִ�е�sql
    function GetLastSQL: string;
    // �������ִ�е�sql
    procedure SetLastSQL(value: string);
    // ��ȡ�༭���ؼ�����
    function GetEditor: TSynEdit;
    // ���ñ༭��״̬������ͨ���˷��������а����˴���״̬�ı��¼�
    procedure SetEditorState(value: TEditorState);

    // ����sql��ʷ
    procedure AddHistory(sSql: string);
    // sql��ʷ�˵�ѡ���¼�����sql���Ƶ��༭���У�
    procedure SQLHistoryMenuClick(Sender: TObject);
    // �ͷű�frame�����dataset����
    procedure FreeAndNilDataSet;

    function GetCanExport: Boolean;
    procedure DoExecuteLastSql;
    procedure OnExecuteSql(Sender: TObject);
  public
//    CurrPage: Integer;                        // ��ǰҳ�ţ���ҳ�������
    DBConnect: IDBConnect;                    // ���ݿ���󣬱�frameר��
    FileName: string;                         // �༭���ļ������ڱ༭�ļ�ģʽ����Ч
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

    // ��query�����ݵ���
    procedure ExportQuery(AQry: TADOQuery; AFileName: string;
      sSpaceBetween: string = ' ');

    // ִ��sql
    function ExecuteSql(sSqlText: string): Integer;
    procedure StopExec();
//    function ExecuteSqlList(sqlList: TStrings):Integer;
    // ��hintpanel����ʾ��Ϣ
    procedure ShowMsg(s: string);

    // ������Ҫչʾ��page  ��Ҫ������ʾ�����ҳ������Ϣҳ
    procedure CheckShouldShowPage();
    // �Ƿ���ʾ�м��grid���֣�����ʾʱsql�༭���ռ���Щ�����ڱ༭
    procedure ShowClient(bShow: Boolean);
    // ��ʾhintpanel������ʾ��dataҳ
    procedure ShowData(bShow: Boolean);

    //////// �¼�
    // DBConnect����logʱ����
    procedure OnDBLog(sLog: string);
    // ���ӳɹ���
    procedure OnConnected(Sender: TObject);
    // �Ͽ����Ӻ�
    procedure OnDisConnected(Sender: TObject);

    // ���´�����frame�����dbconnector����
    procedure ReCreateDBConnect(dbt: TDBType);
    procedure ShareDBConnect(db: IDBConnect);

    // ˢ��UI���ؼ�״̬
    procedure RefreshUI;
    // ��궨λ��sql�༭��
    procedure ForcusEdit;

    //////// editor operator
    // ���ļ����ص��༭����ͨ��Ϊsql
    procedure LoadFromFile(AFileName: string);
    // ���༭�����ݱ��浽�ļ�
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
    InitControl(TWinControl(dbgrdData), 'ʹ���ֶ����ݲ���');
    FormStyle := fsStayOnTop;
    Show;
  end;
end;

procedure TFM_DBOperate.actExportExecute(Sender: TObject);
const
  C_sFilterStr = 'sql�ű�*.sql|*.sql|�ı��ļ�*.txt|*.txt|Excel�ļ�*.xls|*.xls|CSV�ļ�*.csv|*.csv';
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
  // �Ϸ����ж�
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
    // TODO: ���������ݲ����Ǳ�   ǰ׺Ҫ����һ��  
    sFileName := TSqlUtils.getTableNameBySQL(LastSQL);  
    if sFileName = '' then
      sFileName := '��ѯ���'
    else    
      dlgExport.FileName := C_FilePrefix_InsertTable + sFileName;

    dlgExport.InitialDir := GetAppRootPath;
    if not dlgExport.Execute then
      Exit;

    sFileName := FitFileName(dlgExport.FileName, dlgExport);
    qryOld := TADOQuery(dbgrdData.DataSource.DataSet);
    dbgrdData.DataSource.DataSet := nil;                    // DBGrid��DataSet�ÿգ������������ٶȿ�

    waitForm := TWaitingForm.Create(Self);
    try
      waitForm.SetMesasge('���ڵ���...');
      waitForm.Show;
      ExportQuery(qryOld, sFileName);   // �����������ķ���
      waitForm.Close;
    finally
      waitForm.Free;
    end;

    with TF_ExportResult.Create(Self) do begin
      ShowExportFile(sFileName);
      Show;
    end;
    
    dbgrdData.DataSource.DataSet := qryOld;                 // ��ԭDBGrid��DataSet
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
  if IDYES = FormatMsgBox('�Ƿ�Ҫ��ӡ�б����������ݣ�', mbsQuestion, Handle) then
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
    lblFieldName.Caption := Format('�ֶ�: %s  ����: %s',[selectedField.FieldName,
      FieldTypeToStr(selectedField.DataType, False)]);
    dbmmoShowFieldDetail.DataField := selectedField.FieldName;
  end;
end;

function TFM_DBOperate.getLastExecTimeStr: string;
var
  d: Double;
begin
  d := DBConnect.GetLastElapsedMilis;
  Result := commonFunc.PrecisionNumStr(FloatToStr(d), 3) + '����';
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
    btnSortOnColumnClick.Hint := '�����˵����ͷ���򣬰�סShiftʵ�ֶ�������'
  else
    btnSortOnColumnClick.Hint := 'δ���õ����ͷ����';
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

  sFormat := '��¼����%d';
  if dsData.DataSet.Active then
  begin
    nRecCount := dsData.DataSet.RecordCount;
    if (nRecCount > 0) and (nRecCount = GlobalParams.MaxRecord) then
      sFormat := '��¼������%d';  
    stat1.Panels[2].Text := Format('�ֶ�����%d',[dsData.DataSet.FieldCount]);
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
  // sqlִ�д��� ʼ����ʾ������ҳ
  // sqlִ�гɹ� �����������ʾ�ĸ�ҳ��
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
    // ����༭������
    if (Key = Ord('s')) or (Key = Ord('S')) then
    begin
      if EditorState in [esChaged] then
        SendMessage(Application.MainForm.Handle, WMUSER_DBCONNECTOR_EDITOR_SAVE, 0, 0);
//        SaveToFile(Self.FileName);
    end;
    // �л��Զ�����
    if (Key = Ord('w')) or (Key = Ord('W')) then
    begin
      syndtSql.WordWrap := not syndtSql.WordWrap;
    end;
    // ����
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
