{**
 * <p>Title: UF_MAIN  </p>
 * <p>Copyright: Copyright (c) 2006/7/24</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: ������</p>
 *}

unit UF_MAIN;

interface

//{$DEFINE DEBUG}
{$WARN UNIT_PLATFORM OFF} 
{$WARN SYMBOL_PLATFORM OFF}

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, ToolWin, Menus, ActnList, ExtCtrls, 
  ADODB, ActnMan, ActnMenus, ActnCtrls, UFM_DBOperate, DBTables,
  XPStyleActnCtrls, U_Pub, DB, SynEdit, SynEditHighlighter,
  SynHighlighterSQL, DBCtrls, UF_Find, U_ThreadUtil, StdStyleActnCtrls,
  U_DataMoudle, fDBGrid, UF_DBOption,U_Const, OleServer,
  UF_Tip, U_Plugin, U_DBAHelp, U_DBConfig,
  U_DBEngineInterface, U_DBConnectInterface,
  UF_Rename, U_PubFunc, UF_ExportTables, U_fHintTree, U_DBTreeFunc;

//  TNodeData = class
//  public
//    Tag: TNodeTag;
//  end;

type
  TF_MAIN = class(TForm)
    actlstMain: TActionList;
    actExit: TAction;
    actExecSelSql: TAction;
    actRefreshTree: TAction;
    splLeftAndMain: TSplitter;
    actConnect: TAction;
    actExecAllSqls: TAction;
    pmTvw: TPopupMenu;
    mniN4: TMenuItem;
    actEditData: TAction;
    actExecScript: TAction;
    stbMain: TStatusBar;
    mniFindAtTree: TMenuItem;
    actFieldDetail: TAction;
    pnl14: TPanel;
    actStopExec: TAction;
    actOpenWithEditor: TAction;
    actSaveEditorText: TAction;
    pnlLeft: TPanel;
    pnl1: TPanel;
    pnlLeftContranCloseBtn: TPanel;
    btnClosePnlLeft: TButton;
    pnl17: TPanel;
    pnlLeftTop: TPanel;
    cbbTableFilter: TComboBox;
    bvlLeft1: TBevel;
    bvlLeft2: TBevel;
    pnl16: TPanel;
    tbr2: TToolBar;
    tbtnAddPage: TToolButton;
    pnl18: TPanel;
    pgcMain: TPageControl;
    actAddPage: TAction;
    actRemovePage: TAction;
    tbtnRemovePage: TToolButton;
    actmgr1: TActionManager;
    actmmb1: TActionMainMenuBar;
    actManOpenWithEditor: TAction;
    actManSaveEditorText: TAction;
    actManAbout: TAction;
    actManEditData: TAction;
    actManShowTree: TAction;
    actManAutoShowError: TAction;
    actManOnTop: TAction;
    acttb1: TActionToolBar;
    actManConnect: TAction;
    actManRefreshTree: TAction;
    actManExecSql: TAction;
    actManStopExec: TAction;
    tbtn2: TToolButton;
    actTip: TAction;
    actAbout: TAction;
    actManDBParams: TAction;
    actDBParams: TAction;
    pnl13: TPanel;
    actManExecScript: TAction;
    pnl19: TPanel;
    tbrTree: TToolBar;
    tbtnRefresh: TToolButton;
    tbtnSearch: TToolButton;
    tbtnEditData: TToolButton;
    actResetIconCache: TAction;
    tbtnOpenWithEditor: TToolButton;
    tbtnSaveEditorText: TToolButton;
    actSaveEditorTextAs: TAction;
    actManSaveEditorTextAs: TAction;
    tbtnSaveEditorTextAs: TToolButton;
    clbr1: TCoolBar;
    pmBars: TPopupMenu;
    mniShowMenu: TMenuItem;
    mniShowBar: TMenuItem;
    mniShowTree: TMenuItem;
    btnSearchText: TToolButton;
    btn3: TToolButton;
    pnlQuickSearch: TPanel;
    edtQuickSearch: TEdit;
    lbl1: TLabel;
    btnSortNode: TToolButton;
    ts3: TTabSheet;
    mniGenCreateSQL: TMenuItem;
    frmDBOperateMain: TFM_DBOperate;
    actExport: TAction;
    actPrint: TAction;
    actManExit: TAction;
    con1: TADOConnection;
    actExportTableCreateSQL: TAction;
    actDropTable: TAction;
    actDeleteField: TAction;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    Database1: TDatabase;
    pmTab: TPopupMenu;
    pmiCloseTab: TMenuItem;
    pmiRename: TMenuItem;
    mniCopyName: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    actDelData: TAction;
    actAddField: TAction;
    actModifyField: TAction;
    actExportDB: TAction;
    mniGenSql: TMenuItem;
    mniGenSelect: TMenuItem;
    mniGenInsert: TMenuItem;
    mniGenUpdate: TMenuItem;
    btnWrap: TToolButton;
    actCommandHelp: TAction;
    actShowSettings: TAction;
    tvwObjects: TfHintTree;
    OpenDialog1: TOpenDialog;
    actEditObj: TAction;
    N7: TMenuItem;
    N8: TMenuItem;
    pnlObjSelect: TPanel;
    pnlSchemaSelect: TPanel;
    cbbSchemaSelect: TComboBox;
    procedure actExitExecute(Sender: TObject);
    procedure actExecSelSqlExecute(Sender: TObject);
    procedure lvwDataClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure actConnectExecute(Sender: TObject);
    procedure actExecAllSqlsExecute(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure cbbTableFilterChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure actRefreshTreeExecute(Sender: TObject);
    procedure actEditDataExecute(Sender: TObject);
    procedure tvwObjectsMouseDown(Sender: TObject;Button: TMouseButton;
      Shift: TShiftState;X, Y: Integer);
    procedure actExecScriptExecute(Sender: TObject);
    procedure mniAutoShowErrClick(Sender: TObject);
    procedure mniTvwGoToClick(Sender: TObject);
    procedure pnlLeftResize(Sender: TObject);
    procedure cbbHighLightChange(Sender: TObject);
    procedure mniOnTopClick(Sender: TObject);
    procedure actFindAtTreeExecute(Sender: TObject);
    procedure actStopExecExecute(Sender: TObject);
    procedure actOpenWithEditorExecute(Sender: TObject);
    procedure btnClosePnlLeftClick(Sender: TObject);
    procedure mniShowTreeClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure actAddPageExecute(Sender: TObject);
    procedure actRemovePageExecute(Sender: TObject);
    procedure actManShowTreeExecute(Sender: TObject);
    procedure actManAutoShowErrorExecute(Sender: TObject);
    procedure actManOnTopExecute(Sender: TObject);
    procedure actManRefreshTreeExecute(Sender: TObject);
    procedure act15Execute(Sender: TObject);
    procedure actAboutExecute(Sender: TObject);
    procedure tvwObjectsExpanding(Sender: TObject; Node: TTreeNode;
      var AllowExpansion: Boolean);
    procedure actDBParamsExecute(Sender: TObject);
    procedure actResetIconCacheExecute(Sender: TObject);
    procedure lbxMsgMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure actSaveEditorTextAsExecute(Sender: TObject);
    procedure actSaveEditorTextExecute(Sender: TObject);
    procedure BarMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure mniShowMenuClick(Sender: TObject);
    procedure mniShowBarClick(Sender: TObject);
    procedure btnSearchTextClick(Sender: TObject);
    procedure tvwObjectsKeyPress(Sender: TObject; var Key: Char);
    procedure edtQuickSearchChange(Sender: TObject);
    procedure actTipExecute(Sender: TObject);
    procedure btnSortNodeClick(Sender: TObject);
    procedure pgcMainChange(Sender: TObject);
    procedure pgcMainMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure tvwObjectsExpanded(Sender: TObject; Node: TTreeNode);
    procedure tvwObjectsCollapsed(Sender: TObject; Node: TTreeNode);
    procedure actExportExecute(Sender: TObject);
    procedure actPrintExecute(Sender: TObject);
    procedure tvwObjectsCustomDrawItem(Sender: TCustomTreeView;
      Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure actExportTableCreateSQLExecute(Sender: TObject);
    procedure actDropTableExecute(Sender: TObject);
    procedure actDeleteFieldExecute(Sender: TObject);
    procedure pmiCloseTabClick(Sender: TObject);
    procedure pmiRenameClick(Sender: TObject);
    procedure pmTabPopup(Sender: TObject);
    procedure mniCopyNameClick(Sender: TObject);
    procedure tvwObjectsKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure edtQuickSearchKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure tvwObjectsChange(Sender: TObject; Node: TTreeNode);
    procedure actDelDataExecute(Sender: TObject);
    procedure actAddFieldExecute(Sender: TObject);
    procedure actModifyFieldExecute(Sender: TObject);
    procedure mniGenSelectClick(Sender: TObject);
    procedure mniGenInsertClick(Sender: TObject);
    procedure mniGenUpdateClick(Sender: TObject);
    procedure btnWrapClick(Sender: TObject);
    procedure actCommandHelpExecute(Sender: TObject);
    procedure actShowSettingsExecute(Sender: TObject);
    procedure tvwObjectsShowHint(Sender: TObject; Node: TTreeNode);
    procedure actExportDBExecute(Sender: TObject);
    procedure actEditObjExecute(Sender: TObject);
    procedure pmTvwPopup(Sender: TObject);
    procedure pnlObjSelectResize(Sender: TObject);
    procedure pnlSchemaSelectResize(Sender: TObject);
    procedure cbbSchemaSelectChange(Sender: TObject);
  private
    { Private declarations }
    FslstNodes: TStringList;
    FbSortNode: Boolean;
    FaryTabs: array of TRect;
    FNodeHitted: TTreeNode;
    FbInReconnect: Boolean;

    FThreadEditData: TCustomThread;

    FPluginList: TPluginList;
    FDBAHelpList: TDBAHelpList;
                                                  
    function DoConnectDB(DBConnect: IDBConnect; DBConfig: TDBConfig;
      bConnect: Boolean): Boolean;
    function ConnectDB(bConnect: Boolean): Boolean;

    // �жϵ�ǰʹ�õ�dbframe
    function GetActiveDBFrame: TFM_DBOperate;
    procedure AdjustActiveDBFrame;
    function GetDBFrameInPage(pageIndex: Integer): TFM_DBOperate;

    // ˢ��pageControl������tab��rect��������˫��tabʱ���˫�����ĸ�tab���Ӷ��ر�
    procedure RefreshTabArray;

    // �жϵ�����ĸ�tab
    function getTabIndexByPoint(X, Y: Integer): Integer;

    // ��״̬����ʾ��ʾ
    procedure ShowStatus(s: string);     

    procedure ShowTreePanel(bShow: Boolean);

    // ˢ������� ���������¼��ظ����ڵ�
    procedure RefreshTree;
    // ���ý���״̬
    procedure RefreshUIStatus; 
    procedure LoadDataBaseNames;
    procedure LoadPlugin;
    procedure LoadDBAHelp;
    procedure actOnePluginExecute(Sender: TObject);    

    function IsSystemObjects: Boolean;

    // ִ��sql���
    function ExecuteSql(sSqlText: string):Boolean;
    procedure SetHighLight(ASQLDialect: TSQLDialect);
    procedure EditorGoToLine(nLine, nIndex, nCount: Integer); 
    function GetActiveEditor: TSynEdit;

    // ʹ��Global����������ݿ����ӹ����˷����������ӳɹ�����Ϣ
    procedure WMUSERDBConnectorOnConnectSucc(var msg: TMsg); message
      WMUSER_DBCONNECTOR_ONCONNECT_SUCC; 
    procedure WMUSERDBConnectorOnConnectFail(var msg: TMsg); message
      WMUSER_DBCONNECTOR_ONCONNECT_FAIL;
    procedure WMUserDBConnectorEditorSave(var msg: TMsg);message
      WMUSER_DBCONNECTOR_EDITOR_SAVE;
    procedure WMUserDBConnectorMainRefreshOnConnect(var msg: TMsg);message
      WMUSER_DBCONNECTOR_MAIN_REFRESHONCONNECT; 
    procedure WMUserDBConnectorMainRefreshOnDisConnect(var msg: TMsg);message
      WMUSER_DBCONNECTOR_MAIN_REFRESHONDISCONNECT;


    function AddPage(sPageName: string = ''; active: Boolean = True): Integer;

    procedure EditorStateChanged(oldState, newState: TEditorState);

    procedure ShowQuickSearch(sKey: string);
    procedure CloseQuickSearch;
    function GenSelectSqlByNode(node: TTreeNode; alias: string = 't'): string;  
    function GenInsertSqlByNode(node: TTreeNode): string;
    function GetFieldsByTableNode(tableNode: TTreeNode; schema: string = ''): string;
    function GetFieldsByFieldNode(fieldNode: TTreeNode; schema: string = ''): string;
    function GetFieldValuesByTableNode(tableNode: TTreeNode): string;
    function GetFieldValuesByFieldNode(fieldNode: TTreeNode): string;
    function GetFieldUpdatesByFieldNode(fieldNode: TTreeNode): string;
    function GetFieldUpdatesByTableNode(tableNode: TTreeNode): string;
    function GenUpdateSqlByNode(node: TTreeNode): string;
    procedure RunDBAHelp(item: TDBAHelpItem);
    procedure actOneDBAHelpExecute(Sender: TObject);
    procedure ReConnectDB;
    procedure DoExecSelSql;
    function DBType2Dialect(dbt: TDBType): TSQLDialect;
    procedure DoAfterExecSqlThread;
  public
    { Public declarations }
    procedure OpenWithEditor(AFileName: string);
  end;

var
  F_MAIN: TF_MAIN;

implementation

uses
  U_Ini, U_CommonFunc, UF_About, Math, U_UIUtil, Clipbrd
  {$IFDEF USECNDEBUG},CnDebug{$ENDIF}, U_fStrUtil, U_DBUtil;

const
  C_nPgcMainPageSQL = 0;
//  C_nPgcMainPageEditor = 1;

{$R *.dfm} 

function TF_MAIN.DBType2Dialect(dbt: TDBType): TSQLDialect;
var
  ConnectedSQLDialect: TSQLDialect;
begin
  case dbt of
  dbtAccess, dbtAccess2007:
  begin
    ConnectedSQLDialect := sqlMSSQL2K;
  end;
  dbtOracle:      
  begin
    ConnectedSQLDialect := sqlOracle;
  end;
  dbtSybase, dbtSybaseASA:
  begin
    ConnectedSQLDialect := sqlSybase;
  end;
  dbtMySQL:
  begin
    ConnectedSQLDialect := sqlMySQL;
  end;   
  else
    ConnectedSQLDialect := sqlStandard;
  end;
  Result := ConnectedSQLDialect;
end; 

function TF_MAIN.DoConnectDB(DBConnect: IDBConnect; DBConfig: TDBConfig;
  bConnect: Boolean): Boolean;
var
  sDataSource, sUser, sPwd: string;
  eDBType: TDBType;
  eDBET: TDBEngineType;

  function GetValidFullPath(Str: string):string;
  begin
    if (Trim(Str) = '') or (Pos(':', Str) <> 0)
       or (Pos('\\', Str) = 1) then
      Result := Str
    else
      Result := GetAppRootPath+Str;
  end;
begin 
  if not bConnect then
    Result := g_Global.CloseDB
  else
  begin
    eDBType := DBconfig.DBType;
    eDBET := DBconfig.DBEngineType;
    case eDBType of
    dbtAccess, dbtAccess2007:
    begin
      sDataSource := BuildAccessDataSource(GetValidFullPath(DBconfig.AcsDataSource),
        GetValidFullPath(DBconfig.AcsSecuredDB));
//      if FileExists(DBconfig.AcsDataSource) then
//        FileSetAttr(sDataSource,
//          FileGetAttr(sDataSource) and (not SysUtils.faReadOnly));
      sUser := DBconfig.AcsUser;
      sPwd := DBconfig.AcsPwd;
    end;
    dbtOracle:
    begin
      sDataSource := BuildOracleDataSource(DBconfig.SID, DBconfig.Host,
        DBconfig.Port, DBconfig.Protocol,
        DBconfig.IsLocalTns or (DBconfig.Host = ''));
      sUser := DBconfig.OraUser;
      sPwd := DBconfig.OraPwd;
    end;
    dbtSyBase, dbtSybaseASA:
    begin
      sDataSource := BuildSybaseDataSource(DBconfig.SybServerName,
        DBconfig.SybPort, DBconfig.SybDatabaseName);
      sUser := DBconfig.SybUser;
      sPwd := DBconfig.SybPwd;
    end;
    dbtMySQL:
    begin
      sDataSource := BuildMySqlDataSource(DBconfig.mslHost,
        DBconfig.mslDataBase, DBconfig.mslCharset);
      sUser := DBconfig.mslUser;
      sPwd := DBconfig.mslPwd;
    end;
    dbtSqlite:
    begin
      sDataSource := DBConfig.sltDB;
    end;
    else
      sDataSource := DBConfig.DataSource;
    end;

    if Trim(sDataSource) = '' then
    begin
      raise Exception.Create('��Ч����Դ������������á�');
    end;
    
    if eDBType <> dbtUnKnown then
    begin
      Result := g_Global.OpenDB(sDataSource, sUser, sPwd, eDBType, eDBET);
    end
    else
    begin
      Result := g_Global.DBConnect.OpenDB(GetAppRootPath + C_sFILENAME_INI,
        'UnKnownDB');
      g_DBTreeFunc.SetDBConnect(g_Global.DBConnect);
    end;
  end;
end;

function TF_MAIN.ConnectDB(bConnect: Boolean): Boolean;
begin
  if not bConnect then
  begin        
    Result := DoConnectDB(g_Global.DBConnect, GlobalParams.DBconfig, bConnect);
    if Result then
    begin
      RefreshTree;
      SetHighLight(sqlStandard);
      ShowStatus('�Ѿ��Ͽ����ӣ�');
    end
    else
      ShowStatus('�Ͽ�����ʧ�ܣ�');
  end
  else
  begin     
    ShowStatus('�����������ݿ�...');
    Result := DoConnectDB(g_Global.DBConnect, GlobalParams.DBconfig, bConnect);
    if Result then
    begin
      RefreshTree;
      U_Pub.SaveConfig;   
      SetHighLight(DBType2Dialect(g_Global.DBConnect.DBType));
      LoadDBAHelp;
      ShowStatus('���ӳɹ������ݿ�����' + DBTypeToStr(
        g_Global.DBConnect.DBType, False));
    end
    else
    begin
      ShowStatus('�������ݿ�ʧ�ܣ�');
    end;
  end;  
end;

procedure TF_MAIN.ReConnectDB();
var
  oldDb: string;
begin
  FbInReconnect := True;
  try
    if g_Global.Connected then
      ConnectDB(False);
    case g_Global.DBConnect.DBType of
    dbtMySql:
    begin
      oldDb := GlobalParams.DBconfig.mslDataBase;
      GlobalParams.DBconfig.mslDataBase := cbbSchemaSelect.Text;
      ConnectDB(True);
      GlobalParams.DBconfig.mslDataBase := oldDb;
    end;
    else
    end;
  finally
    FbInReconnect := False;
  end;
end;  

procedure TF_MAIN.cbbSchemaSelectChange(Sender: TObject);
begin
  ReConnectDB();
end;

procedure TF_MAIN.actExitExecute(Sender: TObject);
begin
  Close;
end;

procedure TF_MAIN.ShowQuickSearch(sKey: string);
begin
  if not pnlQuickSearch.Visible then
    pnlQuickSearch.Visible := True;
  if sKey = #8 then
    edtQuickSearch.Text := Copy(edtQuickSearch.Text, 1,
      Length(edtQuickSearch.Text)-1)
  else
    edtQuickSearch.Text := edtQuickSearch.Text + sKey;
end;   

procedure TF_MAIN.CloseQuickSearch;
begin
  pnlQuickSearch.Visible := False;
  edtQuickSearch.Text := '';
end;

procedure TF_MAIN.ShowStatus(s: string);
var
  nTextWidth: Integer;
begin
  nTextWidth := Canvas.TextWidth(s);
  if nTextWidth > stbMain.Panels[0].Width then
    stbMain.Panels[0].Width := Min(nTextWidth, pnlLeft.Width);

  stbMain.Panels[0].Text := s;
  if stbMain.Panels[0].Width < stbMain.Canvas.TextWidth(stbMain.Panels[0].Text) then
     stbMain.Panels[0].Width := stbMain.Canvas.TextWidth(stbMain.Panels[0].Text);
  Application.ProcessMessages;
end;

procedure TF_MAIN.DoExecSelSql();
var
  oldcur: TCursor;
begin
  oldcur := Screen.Cursor;
  Screen.Cursor := crHourGlass;
  try
    if GetActiveEditor.SelText <> '' then
      ExecuteSql(GetActiveEditor.SelText)
    else
      ExecuteSql(GetActiveEditor.Lines.Text);
  finally
    Screen.Cursor := oldcur;
  end;
end;

procedure TF_MAIN.actExecSelSqlExecute(Sender: TObject);
begin
  DoExecSelSql;
end;

procedure TF_MAIN.actExecAllSqlsExecute(Sender: TObject);
begin
  ExecuteSql(GetActiveEditor.Text);
end;

procedure TF_MAIN.lvwDataClick(Sender: TObject);
begin
  RefreshUIStatus;
end;

procedure TF_MAIN.RefreshUIStatus;
begin
  actRefreshTree.Enabled := g_Global.Connected;
  actManRefreshTree.Enabled := g_Global.Connected;

  actExecScript.Enabled := g_Global.Connected;
  actManExecScript.Enabled := g_Global.Connected;
  actExportDB.Enabled := g_Global.Connected;

  actExport.Enabled := g_Global.Connected;
  actPrint.Enabled := g_Global.Connected;

  if btnSortNode.Down then
    btnSortNode.Hint := '�����˽ڵ��Զ����򣬵����ˢ��������ʵ������ڵ�'
  else
    btnSortNode.Hint := 'δ���ýڵ��Զ�����';

  if g_Global.Connected then
  begin
    actConnect.Caption := '�Ͽ�';
    actManConnect.Caption := '�Ͽ�';
    Caption := Format('���ݿ⣺%s(%s) �û���%s ���棺%s', [
           g_Global.DBConnect.DBEngine.GetDataSource,
           DBTypeToStr(g_Global.DBConnect.DBEngine.GetDBType, False),
           g_Global.DBConnect.DBEngine.GetUserName,
           DBEngineTypeToStr(g_Global.DBConnect.DBEngine.GetDBEngineType, False)
           ]);
  end
  else
  begin
    actConnect.Caption := '����';
    actManConnect.Caption := '����';
    Caption := gC_AppName;
  end;
  LoadDataBaseNames;
end;

procedure TF_MAIN.LoadDataBaseNames;
var
  i: Integer;
begin
  if FbInReconnect then
    Exit;
  if g_Global.Connected then
  begin
    case g_Global.DBConnect.DBType of
    dbtMySql:
    begin
      pnlSchemaSelect.Visible := True;
      cbbSchemaSelect.Items.Clear;
      TDBUtil.LoadMysqlDatabaseNames(g_Global.DBConnect.DBEngine.GetDataSource,
        cbbSchemaSelect.Items);
      pnlLeftTop.Height := pnlSchemaSelect.Height + pnlObjSelect.Height;
      for i := 0 to cbbSchemaSelect.Items.Count - 1 do
      begin
        if cbbSchemaSelect.Items[i] = GlobalParams.DBconfig.mslDataBase then
        begin
          cbbSchemaSelect.ItemIndex := i;
          Break;
        end;
      end;
    end;
    else
    end;
  end
  else
  begin
    pnlLeftTop.Height := pnlObjSelect.Height;
    pnlSchemaSelect.Visible := False;
  end;
end;  

procedure TF_MAIN.LoadPlugin;
var
  actMenuItem: TActionClientItem;
  act: TAction;
  i: Integer;
begin
  FPluginList.LoadPlugins(GetAppRootPath + 'plugin\');
  actmgr1.ActionBars[0].Items[4].Items.Clear;  // ���˵��С����ߡ�һ��
  if FPluginList.Count = 0 then
  begin
    actMenuItem := actmgr1.ActionBars[0].Items[4].Items.Add;
    actMenuItem.Caption := '��';
  end;

  for i := 0 to FPluginList.Count - 1 do
  begin
    actMenuItem := actmgr1.ActionBars[0].Items[4].Items.Add;
    act := TAction.Create(Self);
    act.OnExecute := actOnePluginExecute;
    act.Caption := FPluginList.Items[i].PluginName;
    act.Tag := FPluginList.Items[i].Index;
    actMenuItem.Action := act;
  end;
end; 

procedure TF_MAIN.actOnePluginExecute(Sender: TObject);
begin
  FPluginList.Items[(Sender as TAction).Tag].Run;
end;

procedure TF_MAIN.actOneDBAHelpExecute(Sender: TObject);
begin
  FDBAHelpList.Items[(Sender as TAction).Tag].Run;
end;

function TF_MAIN.ExecuteSql(sSqlText: string):Boolean;
begin
  Result := False;
  if Trim(sSqlText) = '' then Exit;

  actExecSelSql.Enabled := False;
  actManExecSql.Enabled := False; 
  actStopExec.Enabled := not actExecSelSql.Enabled;
  actManStopExec.Enabled := not actManExecSql.Enabled;
  Application.ProcessMessages;
  try
    if not g_Global.Connected then
    begin
      if IDYES = FormatMsgBox('��û���������ݿ⣡�Ƿ��������ӣ�', mbsQuestion,
        Handle) then
      begin
        actConnectExecute(actConnect);
      end;
    end;
    if g_Global.Connected then
    begin
      g_frmDBOperate.ExecThread.DoAfterExecute := TCustomThreadProxy.Create(DoAfterExecSqlThread);
      g_frmDBOperate.ExecThread.DoOnAbort := TCustomThreadProxy.Create(DoAfterExecSqlThread);
      g_frmDBOperate.ExecuteSql(sSqlText);
      Result := g_Global.DBConnect.LastOperSucc;
    end
    else
      Result := False;
  finally
//    actExecSelSql.Enabled := True;
//    actManExecSql.Enabled := True;
//
//    actStopExec.Enabled := False;
//    actManStopExec.Enabled := False;
  end;
end;

procedure TF_MAIN.DoAfterExecSqlThread;
begin   
  actExecSelSql.Enabled := not g_frmDBOperate.ExecIng;
  actManExecSql.Enabled := not g_frmDBOperate.ExecIng;

  actStopExec.Enabled := not actExecSelSql.Enabled;
  actManStopExec.Enabled := not actManExecSql.Enabled;
end;  

procedure TF_MAIN.actConnectExecute(Sender: TObject);
begin
  actConnect.Enabled := False;
  Application.ProcessMessages;
  try
    ConnectDB(not g_Global.Connected);
  finally                     
    actConnect.Enabled := True;
    Application.ProcessMessages;
  end;
  RefreshUIStatus;
end;

procedure TF_MAIN.FormCreate(Sender: TObject);
begin
//  ReportMemoryLeaksOnShutdown := True;

  FslstNodes := TStringList.Create;
  FPluginList := TPluginList.Create(TPluginItem);
  FDBAHelpList := TDBAHelpList.Create(TDBAHelpItem);
  frmDBOperateMain.Align := alClient;
  g_frmDBOperate := frmDBOperateMain;
  g_frmDBOperate.ShowClient(False);

  actEditData.Enabled := False;
  actManEditData.Enabled := False;
  actExport.Enabled := False;
  actPrint.Enabled := False;

  FbSortNode := btnSortNode.Down;

  LoadDBAHelp;
end;

procedure TF_MAIN.FormShow(Sender: TObject);
var
//  nSqlDia: TSQLDialect;
  i: Integer;
begin
  with GlobalParams do
  begin
    actManAutoShowError.Checked := AutoShowError;
    for i := 0 to clbr1.Bands.Count - 1 do
    begin
      if i >= Length(BarOptions) then Break;
      if not ((i = 0) and (BarOptions[i].bBreak = True)) then
        clbr1.Bands[i].Break := BarOptions[i].bBreak;
      if BarOptions[i].nWidth <> 0 then
        clbr1.Bands[i].Width := BarOptions[i].nWidth;
    end;
//    g_frmDBOperate.dbgrdData.MaxFitColCount := MaxFitColCount;
    mniShowMenu.Checked := ShowMenu;
    mniShowBar.Checked := ShowBar;
    actmmb1.Visible := ShowMenu;
    acttb1.Visible := ShowBar;
    ShowTreePanel(ShowTree);
//    g_frmDBOperate.btnSortOnColumnClick.Down := SortOnColumnClick;
  end;

//  cbbHighLight.Items.Clear;
//  for nSqlDia := Low(gC_ARY_HIGHLIGHT_SET) to High(gC_ARY_HIGHLIGHT_SET) do
//    cbbHighLight.Items.Add(gC_ARY_HIGHLIGHT_SET[nSqlDia].DisplayName);
//  cbbHighLight.ItemIndex := 0;

  if FileExists(GetAppRootPath + 'DBNames.xml') then
  begin
    g_Global.DBConfigList.Clear;
    g_Global.DBConfigList.LoadFromFile(GetAppRootPath + 'DBNames.xml');
  end;

  {$IFDEF USECNDEBUG}
    CnDebugger.TraceMsg('����');
  {$ENDIF}
  pgcMain.ActivePageIndex := C_nPgcMainPageSQL;

//   ���ݱ�������ݾ����п�
//  pmAdjust.Items[0].Checked := True;

  RefreshTree;
  RefreshUIStatus;
  AdjustActiveDBFrame;
//  RefreshTabArray;
  LoadPlugin;

//  frmDBOperateMain.pnlGrid.Align := alBottom;
end;

procedure TF_MAIN.FormDestroy(Sender: TObject);
var
  i: Integer;
begin
  with GlobalParams do
  begin
    SetLength(BarOptions, clbr1.Bands.Count);
    for i := clbr1.Bands.Count - 1 downto 0 do
    begin
      BarOptions[i].bBreak := clbr1.Bands[i].Break;
      BarOptions[i].nWidth := clbr1.Bands[i].Width;
    end;
    
    if Assigned(g_Global.DBConnect) then
    begin
      MaxRecord := g_Global.DBConnect.MaxRecords;
    end;
//    SortOnColumnClick := btnSortOnColumnClick.Down;
    SaveSettings;
  end;

  FslstNodes.Free;
  FPluginList.Free;
  FDBAHelpList.Free;
  g_DBTreeFunc.clearDBTree(tvwObjects);
  if Assigned(FThreadEditData) then
    FreeAndNil(FThreadEditData);

end;

procedure TF_MAIN.RefreshTree;
begin
  g_DBTreeFunc.initDBTree(tvwObjects);
end;

procedure TF_MAIN.cbbTableFilterChange(Sender: TObject);
begin
  RefreshTree;
end;

procedure TF_MAIN.actRefreshTreeExecute(Sender: TObject);
begin
  RefreshTree;
end; 

procedure TF_MAIN.actFindAtTreeExecute(Sender: TObject);
begin
  if not Assigned(gFindForm) then
    gFindForm := TF_Find.Create(Self);
  with gFindForm do
  begin
    PassControl(TWinControl(tvwObjects), '�������ݿ����');
    FormStyle := fsStayOnTop;
    Show;
  end;
end;  

procedure TF_MAIN.actEditDataExecute(Sender: TObject);
var
  sFields, sSqlExecute: string;
  parentNode, childNode: TTreeNode;
  sFieldName: string;
  slstFields: TStringList;
  slstQuerydFields: TStringList;
  i, n: Integer;
  node: TTreeNode;
begin
  if tvwObjects.SelectionCount = 0 then
    Exit;

  slstQuerydFields := TStringList.Create;
  try
    for n := 0 to tvwObjects.SelectionCount - 1 do
    begin
      node := tvwObjects.Selections[n];
      case TDBTreeNodeData(node.Data).NodeTag of
      ntgTable:
        sSqlExecute := Format('select * from %s', [node.Text]);
      ntgField:
      begin
        // ��ѡ���ֶ�ʱ�����Ϊ��ѯ�����޶�ֻ��ѯѡ���˵��ֶ�
        // ����һ�ο��ܴ�����ѡ�е��ֶΣ����������ѯ�д�����ˣ�����ѭ����������
        sFieldName := pubFunc.GetFieldShortName(IntToStr(node.AbsoluteIndex));
        if slstQuerydFields.IndexOf(sFieldName) <> -1 then
          Continue;
        parentNode := node.Parent;
        if Assigned(parentNode) and
           (TDBTreeNodeData(parentNode.Data).NodeTag = ntgTable) then
        begin          
          childNode := parentNode.getFirstChild;
          slstFields := TStringList.Create;
          while childNode <> nil do
          begin
            if childNode.Selected then
            begin
              sFieldName := pubFunc.GetFieldShortName(childNode.Text);
              slstFields.Add(sFieldName);
              slstQuerydFields.Add(IntToStr(childNode.AbsoluteIndex));
            end;
            childNode := childNode.getNextSibling;
          end;  
          try
            if FbSortNode then
              slstFields.Sort;
            sFields := '';
            for i := 0 to slstFields.Count - 1 do
            begin
              if sFields = '' then
                sFields := slstFields.Strings[i]
              else
                sFields := sFields + ', ' + slstFields.Strings[i];
            end;
            sSqlExecute := Format('select %s from %s', [
                sFields,  // ȡ����һ���ո�ǰ���ַ���Ϊ�ֶ���������пո�
                parentNode.Text]);
          finally
            slstFields.Free;
          end;
        end;
      end;
      ntgView:
      begin
        sSqlExecute := Format('select * from %s', [node.Text]);
      end;  
  //    ntgProc, ntgView, ntgTrig:
  //    begin
  //      case g_Global.DBConnect.DBType of
  //      dbtOracle:
  //      begin
  //        GetActiveEditor.Lines.Add((
  //          g_Global.DBConnect as TOracleDBConnect).GetSource(selnode.Text));
  //        Exit;
  //      end;
  //      end;
  //    end;
      else
      end;
      if n = 0 then
        ExecuteSql(sSqlExecute)
      else
      begin
        GetDBFrameInPage(AddPage()).ExecuteSql(sSqlExecute);
      end;  
    end;
  finally
    slstQuerydFields.Free;
  end;
end;   

procedure TF_MAIN.actEditObjExecute(Sender: TObject);   
var
  sourceList: TStrings;
  i: Integer;
  node: TTreeNode;
begin
  if tvwObjects.SelectionCount = 0 then
    Exit;
  for i := 0 to tvwObjects.SelectionCount - 1 do
  begin
    node := tvwObjects.Selections[i];
    case TDBTreeNodeData(node.Data).NodeTag of
    ntgProc, ntgTrig:
    begin
      case g_Global.DBConnect.DBEngine.GetDBType of
      dbtOracle:
      begin  
        sourceList := GetDBFrameInPage(AddPage(node.Text)).Editor.Lines;
        g_Global.DBConnect.GetProcSource(node.Text, sourceList);
      end;
      end;
    end;
    ntgView:
    begin 
      case g_Global.DBConnect.DBEngine.GetDBType of
      dbtOracle:
      begin
        sourceList := GetDBFrameInPage(AddPage(node.Text)).Editor.Lines;
        g_Global.DBConnect.GetViewSource(node.Text, sourceList);
      end;
      end;
    end;  
    else
    end;
  end;
end;

procedure TF_MAIN.actExportTableCreateSQLExecute(Sender: TObject);
var
  selnode: TTreeNode;
  slst: TStrings;
begin
  selnode := tvwObjects.Selected;
  if Assigned(selnode) then
  begin
    case TDBTreeNodeData(selnode.Data).NodeTag of
      ntgTable:
      begin
//        with TSaveDialog.Create(Self) do
//        try
//          FileName := 'CreateTable_'+selnode.Text+'.sql';
//          if Execute() then
//          begin
            slst := TStringList.Create;
            try
              g_Global.DBConnect.GenTableCreateSQL(selnode.Text, slst);
              GetActiveEditor.Lines.AddStrings(slst);
//              slst.SaveToFile(FileName);
//              with TF_ExportResult.Create(Self) do
//              try
//                ShowExportFile(FileName);
//                ShowModal;
//              finally
//                Free;
//              end;
            finally
              slst.Free;
            end;
//          end;
//        finally
//          Free;
//        end;
      end;
    end;
  end;
end;

procedure TF_MAIN.tvwObjectsMouseDown(Sender: TObject;
  Button: TMouseButton;Shift: TShiftState;X, Y: Integer);
var
  tndHit: TTreeNode;
//  tag: TNodeTag;
begin
  tndHit := UIUtil.GetTreeViewHittedNode(TTreeView(Sender), X, Y, False);
  FNodeHitted := tndHit;
  if not Assigned(tndHit) then
    Exit;

//  if tndHit <> nil then
//    tag := TNodeTag(tndHit.Data)
//  else
//    tag := ntgNone;
    
//  actEditData.Enabled := g_Global.Connected and not (tag in [ntgNone, ntgTypeRoot]);
//  actManEditData.Enabled := actEditData.Visible;
//  actExportTableCreateSQL.Visible := g_Global.Connected and (tag = ntgTable);
//  actDropTable.Visible := g_Global.Connected and (tag = ntgTable);
//  actDeleteField.Visible := g_Global.Connected and (tag = ntgField);
  
  if pnlQuickSearch.Visible then
  begin
    pnlQuickSearch.Visible := False;
    edtQuickSearch.Text := '';
  end;

  if Button = mbRight then
  begin
    tvwObjects.Selected := tndHit;
    pmTvw.Popup(Mouse.CursorPos.X, Mouse.CursorPos.Y);
  end;
end;

procedure TF_MAIN.actExecScriptExecute(Sender: TObject);
var
  i: Integer;
begin     
  with OpenDialog1 do // TOpenDialog.Create(Self) do
  begin
    if GlobalParams.LastScriptFile <> '' then
      InitialDir := ExtractFilePath(GlobalParams.LastScriptFile)
    else
      InitialDir := ExtractFilePath(Application.ExeName);
    Filter := 'sql�ű�*.sql|*.sql|ȫ���ļ�|*.*';
    Options := Options + [ofAllowMultiSelect];
    if Execute() then
    begin
      AddPage(commonFunc.RemoveFileExt(ExtractFileName(Files[0])), True);
      for i := 0 to Files.Count - 1 do
        GetActiveEditor.Lines.Add(Format('@%s;', [Files[i]]));
      actExecSelSqlExecute(actExecSelSql);
    end;  
  end;

//  if not Assigned(F_ExecScript) then
//    F_ExecScript := TF_ExecScript.Create(Self);
////  try
//    F_ExecScript.FormStyle := fsStayOnTop;
//    F_ExecScript.Show;
////    ShowModal;
////  finally
////    Free;
////  end;
end;

procedure TF_MAIN.mniAutoShowErrClick(Sender: TObject);
begin
  actManAutoShowError.Checked := not actManAutoShowError.Checked;
  GlobalParams.AutoShowError := actManAutoShowError.Checked;
  GlobalParams.SaveSettings;
end;

procedure TF_MAIN.mniCopyNameClick(Sender: TObject);
var
  sFields: string;
  sCopyedName: string;
  selnode, selparent: TTreeNode;
  slstFields: TStringList;
  i: Integer;
begin

  selnode := tvwObjects.Selected;
  if Assigned(selnode) then
  begin
    case TDBTreeNodeData(selnode.Data).NodeTag of
      ntgTable:
        sCopyedName := selnode.Text;
      ntgField:
      begin
        selparent := selnode.Parent;
        if Assigned(selparent) and
           (TDBTreeNodeData(selparent.Data).NodeTag = ntgTable) then
        begin
          slstFields := TStringList.Create;
          try
            for i := 0 to tvwObjects.SelectionCount - 1 do
            begin
              if tvwObjects.Selections[i].Parent <> selparent then
                Continue;
              slstFields.Add(pubFunc.GetFieldShortName(tvwObjects.Selections[i].Text));
            end;
            if FbSortNode then
              slstFields.Sort;
            sFields := '';
            for i := 0 to slstFields.Count - 1 do
            begin
              if sFields = '' then
                sFields := slstFields.Strings[i]
              else
                sFields := sFields + ', ' + slstFields.Strings[i];
            end;
            sCopyedName := sFields;
          finally
            slstFields.Free;
          end;
        end;
      end;
      else
        sCopyedName := selnode.Text;
    end;
    Clipboard.AsText := sCopyedName;
  end;
end;

procedure TF_MAIN.mniTvwGoToClick(Sender: TObject);
begin
  if not Assigned(gFindForm) then
    gFindForm := TF_Find.Create(Self);
  with gFindForm do
  begin
    PassControl(TWinControl(tvwObjects));
    FormStyle := fsStayOnTop;
    Show;
  end;
end;
procedure TF_MAIN.pnlLeftResize(Sender: TObject);
  procedure SetControlWidth(ACtrl: TControl;nWidth: Integer);
  begin
    if nWidth < 0 then
      nWidth := 0;
    ACtrl.Width := nWidth;
  end;
begin
  SetControlWidth(pnlObjSelect, pnlLeft.Width - 5);
  SetControlWidth(pnlSchemaSelect, pnlLeft.Width - 5);
  SetControlWidth(bvlLeft1, pnlLeft.Width - pnlLeftContranCloseBtn.Width-5 );
  SetControlWidth(bvlLeft2, bvlLeft1.Width );
end;

procedure TF_MAIN.pnlObjSelectResize(Sender: TObject);
begin
  //
end;

procedure TF_MAIN.pnlSchemaSelectResize(Sender: TObject);
begin
  //
end;

procedure TF_MAIN.SetHighLight(ASQLDialect: TSQLDialect);
//var
//  i: TSQLDialect;
begin
//  for i := Low(gC_ARY_HIGHLIGHT_SET) to High(gC_ARY_HIGHLIGHT_SET) do
//  begin
//    if gC_ARY_HIGHLIGHT_SET[i].SQLDialect = ASQLDialect then
//    begin
//      cbbHighLight.ItemIndex := Integer(i);
//      Break;
//    end;
//  end;
//  cbbHighLightChange(nil);
end;

procedure TF_MAIN.cbbHighLightChange(Sender: TObject);
begin
//  if cbbHighLight.ItemIndex >= 0 then
//  begin
//    g_frmDBOperate.SynSQLSyn1.SQLDialect :=
//      gC_ARY_HIGHLIGHT_SET[TSQLDialect(cbbHighLight.ItemIndex)].SQLDialect;
//  end;
end;

procedure TF_MAIN.mniOnTopClick(Sender: TObject);
begin
  actManOnTop.Checked := not actManOnTop.Checked;
  if actManOnTop.Checked then
    commonFunc.SetOnTopMost(Handle, True)
  else
    commonFunc.SetOnTopMost(Handle, False);
end;

procedure TF_MAIN.ShowTreePanel(bShow: Boolean);
begin
  pnlLeft.Visible := bShow;
  splLeftAndMain.Visible := bShow;
  splLeftAndMain.Left := pnlLeft.Left + pnlLeft.Width;
  actManShowTree.Checked := bShow;
  mniShowTree.Checked := bShow;
  GlobalParams.ShowTree := bShow;
end;

procedure TF_MAIN.actStopExecExecute(Sender: TObject);
begin
  g_frmDBOperate.StopExec;
end;

procedure TF_MAIN.OpenWithEditor(AFileName: string);
var
  nPageIndex: Integer;
  dbFrame: TFM_DBOperate;
begin
  nPageIndex := AddPage(ExtractFileName(AFileName));
  pgcMain.ActivePageIndex := nPageIndex;  // ����ǰ̨
  AdjustActiveDBFrame;
  dbFrame := GetDBFrameInPage(nPageIndex); //����frame
  dbFrame.FileName := AFileName;
//  dbFrame.Editor.Lines.Clear;
  dbFrame.LoadFromFile(AFileName);
end;

procedure TF_MAIN.actOpenWithEditorExecute(Sender: TObject);
begin
//  if IDCancel = CheckOpenedFileSaved then Exit;  

  with TOpenDialog.Create(Self) do
  try
    Filter := '*.sql|*.sql|*.*|*.*';
    InitialDir := GetAppRootPath;
    if Execute then
    begin
      OpenWithEditor(FileName);
    end;
  finally
    Free;
  end;
end;

procedure TF_MAIN.btnClosePnlLeftClick(Sender: TObject);
begin
  ShowTreePanel(not pnlLeft.Visible);
end;

procedure TF_MAIN.mniShowTreeClick(Sender: TObject);
begin
  ShowTreePanel(not actManShowTree.Checked);
end;

procedure TF_MAIN.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if g_Global.Connected then
  begin
    if IDYES <> FormatMsgBox('�ѽ������ݿ����ӣ��Ƿ�Ҫֱ���˳���',
       mbsQuestion, Handle) then
      CanClose := False;
  end; 
//  if IDCancel = CheckOpenedFileSaved then
//    CanClose := False;
end;

function TF_MAIN.AddPage(sPageName: string; active: Boolean): Integer;
var
  ts : TTabSheet;
  i, nNameIndex: Integer;
  dbframe: TFM_DBOperate;
begin
  ts := TTabSheet.Create(Self);

  dbframe := TFM_DBOperate.Create(Self);
  dbframe.Align := alClient;
  dbframe.Parent := ts;

  nNameIndex := 1;
  if sPageName = '' then
  begin
    for i := 0 to pgcMain.PageCount - 1 do
    begin
      if pgcMain.Pages[i].Caption = '�༭' + IntToStr(nNameindex) then
        Inc(nNameIndex);
    end;    
    ts.Caption := '�༭' + IntToStr(nNameIndex);
  end
  else
    ts.Caption := sPageName;
    
  dbframe.Name := '';//'frmDBOperate' + IntToStr(nNameIndex);

  ts.PageControl := pgcMain;
  actRemovePage.Enabled := pgcMain.PageCount > 1;

  Result := ts.PageIndex;
  if active then
  begin
    pgcMain.ActivePageIndex := ts.PageIndex;
    AdjustActiveDBFrame;
  end
  else
    RefreshTabArray;
end;

procedure TF_MAIN.actAddFieldExecute(Sender: TObject);
begin
// alter table add column_name column_type;
end;  

procedure TF_MAIN.actModifyFieldExecute(Sender: TObject);
begin             
// ����������Ͳ�һ�£���Ҫ���ԭ��
// update table set column_name=null;
// alter table modify(column_name column_type);
end;

procedure TF_MAIN.actAddPageExecute(Sender: TObject);
begin
  AddPage();
end;

procedure TF_MAIN.actRemovePageExecute(Sender: TObject);
var
  nPage: Integer;
begin
  nPage := pgcMain.ActivePageIndex;
  if nPage > 0 then
  begin
    pgcMain.Pages[nPage].Free;
    pgcMain.ActivePageIndex := nPage-1;
  end;
  actRemovePage.Enabled := pgcMain.PageCount > 1;
  AdjustActiveDBFrame;
end;

procedure TF_MAIN.actManShowTreeExecute(Sender: TObject);
begin
  ShowTreePanel(not actManShowTree.Checked);
end;

procedure TF_MAIN.actManAutoShowErrorExecute(Sender: TObject);
begin
  actManAutoShowError.Checked := not actManAutoShowError.Checked;
  GlobalParams.AutoShowError := actManAutoShowError.Checked;
  GlobalParams.SaveSettings;
end;

procedure TF_MAIN.actManOnTopExecute(Sender: TObject);
begin
  actManOnTop.Checked := not actManOnTop.Checked;
  commonFunc.SetOnTopMost(Handle, actManOnTop.Checked)
end;

procedure TF_MAIN.actManRefreshTreeExecute(Sender: TObject);
begin
  RefreshTree;
end;

procedure TF_MAIN.act15Execute(Sender: TObject);
begin
  SetWindowLong(Handle,GWL_HWNDPARENT, GetDesktopWindow);
end;

procedure TF_MAIN.actAboutExecute(Sender: TObject);
begin
  with TF_About.Create(Self) do
  try
    ShowModal;
  finally
    Free;
  end;
end;  

procedure TF_MAIN.pmTvwPopup(Sender: TObject);
var
  selNode: TTreeNode;
begin
  selNode := tvwObjects.Selected;
  if selNode = nil then
    Exit;
end;

procedure TF_MAIN.tvwObjectsChange(Sender: TObject; Node: TTreeNode);
var
  tag: TNodeTag;
begin
  tag := TDBTreeNodeData(Node.Data).NodeTag;
  // �Գ����ڵ�����Ľڵ���Ч
  actEditData.Enabled := g_Global.Connected and
    (tag in [ntgTable, ntgField, ntgView]);
  actManEditData.Enabled := actEditData.Enabled;
  // ����table��Ч
  actDelData.Enabled := g_Global.Connected and (tag in [ntgTable]);
  // ����table��Ч
  actExportTableCreateSQL.Visible := g_Global.Connected and (tag in [ntgTable]);
  actExport.Enabled := g_Global.Connected and GetActiveDBFrame.CanExport;
  // ����table��Ч
  actDropTable.Visible := g_Global.Connected and (tag = ntgTable);
  // ����field��Ч
  actDeleteField.Visible := g_Global.Connected and (tag = ntgField);
  // �༭�洢���̡���ͼ��
  actEditObj.Enabled := g_Global.Connected and (tag in [ntgView, ntgProc]);

  mniGenSql.Enabled := g_Global.Connected and (tag in [ntgTable, ntgField]);
  mniGenCreateSQL.Enabled := g_Global.Connected and (tag in [ntgTable]);
  mniGenSelect.Enabled := g_Global.Connected and (tag in [ntgTable, ntgField]);
  mniGenInsert.Enabled := g_Global.Connected and (tag in [ntgTable, ntgField]);
  mniGenUpdate.Enabled := g_Global.Connected and (tag in [ntgTable, ntgField]);
end;

procedure TF_MAIN.tvwObjectsCollapsed(Sender: TObject; Node: TTreeNode);
begin
  TDBTree.SetNodeImage(Node);
end;

procedure TF_MAIN.tvwObjectsExpanded(Sender: TObject; Node: TTreeNode);
begin
  TDBTree.SetNodeImage(Node);
end;

procedure TF_MAIN.tvwObjectsExpanding(Sender: TObject; Node: TTreeNode;
  var AllowExpansion: Boolean);
begin
  g_DBTreeFunc.AddDBTreeNodeChilds(tvwObjects, Node, IsSystemObjects, FbSortNode);
//  AddDBTreeNodeChilds(Node);
end;

procedure TF_MAIN.actDBParamsExecute(Sender: TObject);
begin
  with TF_DBOption.Create(Self) do
  try
    SetDBConfigList(g_Global.DBConfigList);
    if ShowModal = mrOK then
      if g_Global.Connected then
        actConnect.Execute;
  finally
    Free;
  end;
end;

procedure TF_MAIN.actResetIconCacheExecute(Sender: TObject);
begin
  if IDYES = FormatMsgBox('���ؽ�ͼ�껺�棬�Ƿ������', mbsQuestion, Handle) then
    commonFunc.ResetIconCache;
end;     

procedure TF_MAIN.lbxMsgMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button = mbRight then
  begin
//    pmMenu.Popup(Mouse.CursorPos.X, Mouse.CursorPos.Y);
  end;
end;

procedure TF_MAIN.actSaveEditorTextAsExecute(Sender: TObject);
var
  dbFrame: TFM_DBOperate;
begin
  dbFrame := GetActiveDBFrame;
  with TSaveDialog.Create(Self) do
  try
    if dbFrame.FileName <> '' then
      FileName := dbFrame.FileName
    else
    begin
      FileName := pgcMain.ActivePage.Caption;
    end;  
    Filter := '*.sql|*.sql*.*|*.*';
    InitialDir := GetAppRootPath;
    if Execute then
    begin
      if (FilterIndex = 0)
         and(LowerCase(ExtractFileExt(FileName)) <> '.sql') then
        FileName := FileName + '.sql';
        
      dbFrame.SaveToFile(FileName);
    end;
  finally
    Free;
  end;
end;

procedure TF_MAIN.actSaveEditorTextExecute(Sender: TObject);
var
  dbFrame: TFM_DBOperate;
begin
  dbFrame := GetActiveDBFrame;
//  if dbFrame.Editor then
//  begin
    if (dbFrame.FileName = '') or (not FileExists(dbFrame.FileName)) then
      actSaveEditorTextAsExecute(actSaveEditorTextAs)
    else
    begin
      dbFrame.SaveToFile(dbFrame.FileName);
    end;
//  end;
end;

procedure TF_MAIN.BarMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button = mbRight then
    pmBars.Popup(Mouse.CursorPos.X, Mouse.CursorPos.Y);
end;

procedure TF_MAIN.mniShowMenuClick(Sender: TObject);
begin
  mniShowMenu.Checked := not mniShowMenu.Checked;
  actmmb1.Visible := mniShowMenu.Checked;
  GlobalParams.ShowMenu := mniShowMenu.Checked;
end;

procedure TF_MAIN.mniShowBarClick(Sender: TObject);
begin
  mniShowBar.Checked := not mniShowBar.Checked;
  acttb1.Visible := mniShowBar.Checked;
  GlobalParams.ShowBar := mniShowBar.Checked;
end;

procedure TF_MAIN.btnSearchTextClick(Sender: TObject);
begin
  if not Assigned(gFindForm) then
    gFindForm := TF_Find.Create(Self);
  with gFindForm do
  begin
    PassStrings(GetActiveEditor.Lines, EditorGoToLine, '�����ı�');
    FormStyle := fsStayOnTop;
    Show;
  end;
end;

procedure TF_MAIN.EditorGoToLine(nLine, nIndex, nCount: Integer);
var
  i: Integer;
  nSelStart: Integer;
begin
  if GetActiveEditor <> nil then
  begin
    nSelStart := 0;
    for i := 0 to nLine-1 do
    begin
      nSelStart := nSelStart + Length(GetActiveEditor.Lines[i])+2;
    end;
    GetActiveEditor.SelStart := nSelStart+nIndex-1;
    GetActiveEditor.SelLength := nCount;
  end;
end;

procedure TF_MAIN.EditorStateChanged(oldState, newState: TEditorState);
begin
  case newState of
    esOpened:
    begin
      actSaveEditorText.Enabled := False;
      actManSaveEditorText.Enabled := False;
    end;
    esChaged:
    begin
      actSaveEditorText.Enabled := True;    
      actManSaveEditorText.Enabled := True;
    end;
    esSaved:
    begin     
      actSaveEditorText.Enabled := False;   
      actManSaveEditorText.Enabled := False;
    end;
    else
    begin
      actSaveEditorText.Enabled := False;   
      actManSaveEditorText.Enabled := False;
    end;  
  end;
end;

procedure TF_MAIN.tvwObjectsKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_F3 then
    actFindAtTreeExecute(nil)
  else if Key = VK_ESCAPE then
  begin
    CloseQuickSearch;
    Self.Focused;
  end;
  
  if ((Key = Ord('c')) or (Key = Ord('C'))) and (ssCtrl in Shift) then
  begin
    mniCopyNameClick(mniCopyName);
  end;  
end;

procedure TF_MAIN.tvwObjectsKeyPress(Sender: TObject; var Key: Char);
begin
//  if Key in ['a'..'z', 'A'..'Z', '0'..'9', '_'] then
  if Key > #31 then
  begin
    ShowQuickSearch(Key);
  end
  else if Key = #8 then   // �ظ�
  begin
    if edtQuickSearch.Text = '' then
      CloseQuickSearch
    else        
      ShowQuickSearch(Key);
  end
end;

procedure TF_MAIN.edtQuickSearchChange(Sender: TObject);
begin
  gFindForm.DoFindAtTree(tvwObjects, edtQuickSearch.Text, False, False, True, False);
end; 

procedure TF_MAIN.edtQuickSearchKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_ESCAPE then
  begin
    CloseQuickSearch;
  end;
end;

procedure TF_MAIN.actTipExecute(Sender: TObject);
begin
  if not Assigned(tipForm) then
    tipForm := TF_Tip.Create(Self);
  commonFunc.FormShowOnTop(tipForm);
end;

function TF_MAIN.IsSystemObjects: Boolean;
begin
  case cbbTableFilter.ItemIndex of
    1: Result := True;
    else
      Result := False;
  end;
end;

procedure TF_MAIN.btnSortNodeClick(Sender: TObject);
begin
  FbSortNode := btnSortNode.Down;
  RefreshUIStatus;
end;

procedure TF_MAIN.pgcMainChange(Sender: TObject);
begin
  AdjustActiveDBFrame;
end;   

function TF_MAIN.GetActiveDBFrame: TFM_DBOperate; 
begin
  Result := GetDBFrameInPage(pgcMain.ActivePageIndex);
end;

procedure TF_MAIN.AdjustActiveDBFrame;
begin
  if Assigned(g_frmDBOperate) then
    g_frmDBOperate.EditStateChanged := nil;
  
  g_frmDBOperate := GetActiveDBFrame;
  g_frmDBOperate.EditStateChanged := Self.EditorStateChanged;
  Self.EditorStateChanged(g_frmDBOperate.EditorState,g_frmDBOperate.EditorState);
//  g_frmDBOperate.ForcusEdit;
  RefreshTabArray;
  actExecSelSql.Enabled := not g_frmDBOperate.ExecIng;
  actStopExec.Enabled := not actExecSelSql.Enabled;
end;   

function TF_MAIN.GetDBFrameInPage(pageIndex: Integer): TFM_DBOperate;
var
  i: Integer;
begin
  if pageIndex < 0 then  
    raise Exception.Create('(GetDBFrameInPage)ҳ�Ŵ���');

  Result := nil;
  for i := 0 to pgcMain.Pages[pageIndex].ControlCount - 1 do
  begin
    if pgcMain.Pages[pageIndex].Controls[i] is TFM_DBOperate then
    begin
      Result := TFM_DBOperate(pgcMain.Pages[pageIndex].Controls[i]);
      Break;
    end;
  end;   
  if Result = nil then
    raise Exception.Create('δ�ҵ�TFM_DBOperate�ؼ�');
end;

function TF_MAIN.getTabIndexByPoint(X, Y: Integer): Integer;
var
  i: Integer;
  pt: TPoint;
begin
  Result := -1;
  pt.X := X;
  pt.Y := Y;
  for i := Low(FaryTabs) to High(FaryTabs) do
  begin
    if PtInRect(FaryTabs[i], pt) then
    begin
      Result := i;
      Exit;
    end;  
  end;  
end;       

procedure TF_MAIN.pmTabPopup(Sender: TObject);
begin
  pmiCloseTab.Enabled := pgcMain.ActivePageIndex <> 0;
  pmiRename.Enabled := pgcMain.ActivePageIndex <> 0;
end;

procedure TF_MAIN.pgcMainMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  nTab: Integer;
begin     
  nTab := getTabIndexByPoint(X, Y);
  // ˫��
  if ssDouble in Shift then
  begin
    if (nTab >0)
       and (nTab < pgcMain.PageCount)
       and (pgcMain.PageCount <> 1) // ֻ��һҳʱ���ر�
       then
    begin
      pgcMain.Pages[nTab].Free;
    end;

    if pgcMain.PageCount <= 1 then
      actRemovePage.Enabled := False;
    AdjustActiveDBFrame;
  end
  // �Ҽ��˵�
  else if ssRight in Shift then
  begin
    pmTab.Popup(Mouse.CursorPos.X, Mouse.CursorPos.Y);
  end;    
end;

procedure TF_MAIN.RefreshTabArray;
var
  i: Integer;
begin
  SetLength(FaryTabs, pgcMain.PageCount);
  for i := 0 to pgcMain.PageCount - 1 do
  begin
    FaryTabs[i] := pgcMain.TabRect(i);
  end;
  btnWrap.Down := GetActiveDBFrame.Editor.WordWrap;
end;

procedure TF_MAIN.actExportExecute(Sender: TObject);
begin
  g_frmDBOperate.actExport.Execute;
end;

procedure TF_MAIN.actPrintExecute(Sender: TObject);
begin
  g_frmDBOperate.btnPrint.Click;
end;

procedure TF_MAIN.tvwObjectsCustomDrawItem(Sender: TCustomTreeView;
  Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
  if Sender.Selected = Node then
  begin
    Sender.Canvas.Font.Color := clWhite;
    Sender.Canvas.Font.Style := [];
    Sender.Canvas.Brush.Color := clHighLight;
  end;
end;

procedure TF_MAIN.actDropTableExecute(Sender: TObject);
var
  i: Integer;
  node: TTreeNode;
  bConfirmDelete: Boolean;
  sSql: string;
begin
  if tvwObjects.SelectionCount = 0 then
    Exit;
  if tvwObjects.SelectionCount = 1 then
  begin
    node := tvwObjects.Selections[0];
    bConfirmDelete := IDYES = FormatMsgBox(Format('�Ƿ�ȷ��Ҫ���١�%s��', [
       node.Text]),mbsQuestion)
  end
  else
    bConfirmDelete := IDYES = FormatMsgBox(Format('�Ƿ�ȷ��Ҫ����ѡ�еĶ���', [
       ]),mbsQuestion);
  if bConfirmDelete then
  begin
    for i := tvwObjects.SelectionCount - 1 downto 0 do
    begin
      node := tvwObjects.Selections[i];
      case TDBTreeNodeData(node.Data).NodeTag of
      ntgTable:
        sSql := Format('drop table %s', [node.Text]);
      ntgView:   
        sSql := Format('drop view %s', [node.Text]);
      ntgProc:   
        sSql := Format('drop procedure %s', [node.Text]);
      else  
        sSql := '';
      end;    
      if (sSql <> '') and ExecuteSql(sSql) then
        tvwObjects.Items.Delete(node);
    end;
  end;
end;

procedure TF_MAIN.actDelDataExecute(Sender: TObject);
var
  i: Integer;
  node: TTreeNode;
  bConfirmDelete: Boolean;
begin
  if tvwObjects.SelectionCount = 0 then
    Exit;
  if tvwObjects.SelectionCount = 1 then
  begin
    node := tvwObjects.Selections[0];
    bConfirmDelete := IDYES = FormatMsgBox(Format('�Ƿ�ȷ��Ҫɾ����%s���������', [
       node.Text]),mbsQuestion)
  end
  else
    bConfirmDelete := IDYES = FormatMsgBox(Format('�Ƿ�ȷ��Ҫɾ��ѡ�б�Ķ���', [
       ]),mbsQuestion);
  if bConfirmDelete then
  begin
    for i := 0 to tvwObjects.SelectionCount - 1 do
    begin
      node := tvwObjects.Selections[i];
      case TDBTreeNodeData(node.Data).NodeTag of
      ntgTable:               
        ExecuteSql(Format('delete from %s', [FNodeHitted.Text]));
      end;
    end;  
  end;
end;

procedure TF_MAIN.actDeleteFieldExecute(Sender: TObject);
begin    
  if not Assigned(FNodeHitted) then
    Exit;
  if TDBTreeNodeData(FNodeHitted.Data).NodeTag = ntgField then
  begin        
    if IDYES = FormatMsgBox(Format('�Ƿ�ȷ��Ҫɾ����%s�����ֶΡ�%s��', [
       FNodeHitted.Parent.Text, pubFunc.GetFieldShortName(FNodeHitted.Text)]),
       mbsQuestion) then
    begin
      if 0 <= g_Global.DBConnect.ExecUpdate(Format('alter table %s drop column %s', [
        FNodeHitted.Parent.Text, pubFunc.GetFieldShortName(FNodeHitted.Text)])) then
      begin
        tvwObjects.Items.Delete(FNodeHitted);
      end;  
    end;
  end;
end;

function TF_MAIN.GetActiveEditor: TSynEdit;
begin
  Result := g_frmDBOperate.syndtSql;
end;

procedure TF_MAIN.pmiCloseTabClick(Sender: TObject);
begin
  actRemovePageExecute(actRemovePage);  
end;   

procedure TF_MAIN.pmiRenameClick(Sender: TObject);
begin
  with TF_Rename.Create(Self) do
  try
    OldName := pgcMain.Pages[pgcMain.ActivePageIndex].Caption;
    if mrOK = ShowModal then 
      pgcMain.Pages[pgcMain.ActivePageIndex].Caption := NewName;
  finally
    Free;
  end;   
end;

procedure TF_MAIN.WMUserDBConnectorEditorSave(var msg: TMsg);
begin
  actSaveEditorTextExecute(actSaveEditorText);
end;

procedure TF_MAIN.WMUserDBConnectorMainRefreshOnConnect(var msg: TMsg);
begin
  RefreshUIStatus;    
  ShowStatus('���ӳɹ������ݿ�����' + DBTypeToStr(
    g_Global.DBConnect.DBEngine.GetDBType, False));
end; 

procedure TF_MAIN.WMUserDBConnectorMainRefreshOnDisConnect(var msg: TMsg);
begin
  RefreshUIStatus;
  ShowStatus('�Ѿ��Ͽ����ӣ�');
end;

procedure TF_MAIN.WMUSERDBConnectorOnConnectSucc(var msg: TMsg);
begin
  g_Global.ReRegisterAllFrame;
end;

procedure TF_MAIN.WMUSERDBConnectorOnConnectFail(var msg: TMsg);
var
  i: Integer;
begin
  for i := 0 to g_Global.DBConnect.Log.Count - 1 do
    g_frmDBOperate.DBConnect.AddLog(g_Global.DBConnect.Log[i]);
end;

function TF_MAIN.GetFieldsByTableNode(tableNode: TTreeNode;
  schema: string): string;
var                 
  fieldNode: TTreeNode;
  sFields: string;
  sField: string;
begin         
  // ����ֶ��Ƿ��Ѿ�����
  if tableNode.getFirstChild = nil then
    g_DBTreeFunc.AddDBTreeNodeChilds(tvwObjects, tableNode, IsSystemObjects,
      FbSortNode);
  fieldNode := tableNode.getFirstChild;
  while fieldNode <> nil do
  begin
      sField := pubFunc.GetFieldShortName(fieldNode.Text);
    if schema <> '' then
      sField := Format('%s.%s',[schema, sField]);
    if sFields = '' then
      sFields := sField
    else
      sFields := sFields + ',' + sField;
    fieldNode := fieldNode.getNextSibling;
  end; 
  Result := sFields;
end;

function TF_MAIN.GetFieldsByFieldNode(fieldNode: TTreeNode;
  schema: string): string;
var
  sFields: string;
  sField: string;
  i: Integer;
begin
  for i := tvwObjects.SelectionCount - 1 downto 0 do
  begin
    if tvwObjects.Selections[i].Parent <> fieldNode.Parent then
      Continue;
    sField := pubFunc.GetFieldShortName(tvwObjects.Selections[i].Text);
    if schema <> '' then
      sField := Format('%s.%s',[schema, sField]);
    if sFields = '' then
      sFields := sField
    else
      sFields := sFields + ',' + sField;
  end;
  Result := sFields;
end;

function TF_MAIN.GenSelectSqlByNode(node: TTreeNode; alias: string): string;
var
  tag: TNodeTag;
  sFields: string;
  sTable: string;
begin    
  tag := TDBTreeNodeData(node.Data).NodeTag;    
  case tag of
  ntgTable:
  begin
    sTable := node.Text + ' ' + alias;
    sFields := GetFieldsByTableNode(node, alias);
  end;
  ntgField:
  begin
    sTable := node.Parent.Text + ' ' + alias;;
    sFields := GetFieldsByFieldNode(node);
  end
  else
    Exit;
  end;
  Result := Format('Select %s From %s;', [sFields, sTable]);
end;

procedure TF_MAIN.mniGenSelectClick(Sender: TObject);
var
  selnode: TTreeNode;
  sSql: string;
begin
  selnode := tvwObjects.Selected;
  if not Assigned(selnode) then
    Exit;
  // ���ݱ�ṹ����select
  sSql := GenSelectSqlByNode(selnode);
  GetActiveEditor.Lines.Add(sSql);
end;

function TF_MAIN.GetFieldValuesByTableNode(tableNode: TTreeNode): string;
var
  fieldNode: TTreeNode;
  sFieldValues: string;
begin              
  // ����ֶ��Ƿ��Ѿ�����
  if tableNode.getFirstChild = nil then
    g_DBTreeFunc.AddDBTreeNodeChilds(tvwObjects, tableNode, IsSystemObjects, FbSortNode);
  fieldNode := tableNode.getFirstChild;
  while fieldNode <> nil do
  begin
    if sFieldValues = '' then
      sFieldValues := ':' + pubFunc.GetFieldShortName(fieldNode.Text)
    else
      sFieldValues := sFieldValues + ',' + ':'
        + pubFunc.GetFieldShortName(fieldNode.Text);
    fieldNode := fieldNode.getNextSibling;
  end;
  Result := sFieldValues;
end;

function TF_MAIN.GetFieldValuesByFieldNode(fieldNode: TTreeNode): string;
var
  sFieldValues: string;
  i: Integer;
begin
  for i := 0 to tvwObjects.SelectionCount - 1 do
  begin
    if tvwObjects.Selections[i].Parent <> fieldNode.Parent then
      Continue;
    if sFieldValues = '' then
      sFieldValues := ':' + pubFunc.GetFieldShortName(fieldNode.Text)
    else
      sFieldValues := sFieldValues + ',' + ':'
        + pubFunc.GetFieldShortName(fieldNode.Text);
    fieldNode := fieldNode.getNextSibling;
  end;
  Result := sFieldValues;
end;

function TF_MAIN.GenInsertSqlByNode(node: TTreeNode): string;
var
  tag: TNodeTag;
  sFields, sFieldValues: string;
  sTable: string;
begin    
  tag := TDBTreeNodeData(node.Data).NodeTag;

  case tag of
  ntgTable:
  begin
    // ����ֶ��Ƿ��Ѿ�����
     if node.getFirstChild = nil then
       g_DBTreeFunc.AddDBTreeNodeChilds(tvwObjects, node, IsSystemObjects, FbSortNode);
    sTable := node.Text;
    // fields
    sFields := GetFieldsByTableNode(node);
    // fieldvalues
    sFieldValues := GetFieldValuesByTableNode(node);
  end;
  ntgField:
  begin
    sTable := node.Parent.Text;
    sFields := GetFieldsByFieldNode(node);
    sFieldValues := GetFieldValuesByFieldNode(node);
  end
  else
    Exit;
  end;
  Result := Format('Insert Into %s(%s) Values(%s);', [sTable, sFields, sFieldValues]);
end;

procedure TF_MAIN.mniGenInsertClick(Sender: TObject);
var
  selnode: TTreeNode;
  sSql: string;
begin
  selnode := tvwObjects.Selected;
  if not Assigned(selnode) then
    Exit;
  // ���ݱ�ṹ����Insert
  sSql := GenInsertSqlByNode(selnode);
  GetActiveEditor.Lines.Add(sSql);
end;

function TF_MAIN.GetFieldUpdatesByTableNode(tableNode: TTreeNode): string;
var
  fieldNode: TTreeNode;
  sFieldUpdates: string;
begin
  // ����ֶ��Ƿ��Ѿ�����
  if tableNode.getFirstChild = nil then
    g_DBTreeFunc.AddDBTreeNodeChilds(tvwObjects, tableNode, IsSystemObjects, FbSortNode);
  fieldNode := tableNode.getFirstChild;
  while fieldNode <> nil do
  begin
    if sFieldUpdates = '' then
      sFieldUpdates := pubFunc.GetFieldShortName(fieldNode.Text) + '='
    else
      sFieldUpdates := sFieldUpdates + ','
        + pubFunc.GetFieldShortName(fieldNode.Text) + '=';
    fieldNode := fieldNode.getNextSibling;
  end;
  Result := sFieldUpdates;
end;

function TF_MAIN.GetFieldUpdatesByFieldNode(fieldNode: TTreeNode): string;
var
  sFieldUpdates: string;
  i: Integer;
begin
  for i := 0 to tvwObjects.SelectionCount - 1 do
  begin
    if tvwObjects.Selections[i].Parent <> fieldNode.Parent then
      Continue;
    if sFieldUpdates = '' then
      sFieldUpdates := pubFunc.GetFieldShortName(fieldNode.Text) + '='
    else
      sFieldUpdates := sFieldUpdates + ','
        + pubFunc.GetFieldShortName(fieldNode.Text) + '=';
    fieldNode := fieldNode.getNextSibling;
  end;
  Result := sFieldUpdates;
end;

function TF_MAIN.GenUpdateSqlByNode(node: TTreeNode): string;
var
  tag: TNodeTag;
  sFieldUpdates: string;
  sTable: string;
begin    
  tag := TDBTreeNodeData(node.Data).NodeTag;

  case tag of
  ntgTable:
  begin
    // ����ֶ��Ƿ��Ѿ�����
     if node.getFirstChild = nil then
       g_DBTreeFunc.AddDBTreeNodeChilds(tvwObjects, node, IsSystemObjects, FbSortNode);
    sTable := node.Text;
    sFieldUpdates := GetFieldUpdatesByTableNode(node);
  end;
  ntgField:
  begin
    sTable := node.Parent.Text;
    sFieldUpdates := GetFieldUpdatesByFieldNode(node);
  end
  else
    Exit;
  end;
  Result := Format('Update %s set %s where %s;', [sTable, sFieldUpdates, '']);
end;

procedure TF_MAIN.mniGenUpdateClick(Sender: TObject);  
var
  selnode: TTreeNode;
  tabName: string;
  sSql: string;
  function GetTableFieldsForUpdateOnUI(tableNode: TTreeNode): string;
  var
    fieldNode: TTreeNode;
  begin
    // ����ֶ��Ƿ��Ѿ�����
    if tableNode.getFirstChild = nil then
      g_DBTreeFunc.AddDBTreeNodeChilds(tvwObjects, tableNode, IsSystemObjects, FbSortNode);

    fieldNode := tableNode.getFirstChild;
    while fieldNode <> nil do
    begin
      if Result = '' then
        Result := pubFunc.GetFieldShortName(fieldNode.Text) + '='
      else
        Result := Result + ',' + pubFunc.GetFieldShortName(fieldNode.Text) + '=';
      fieldNode := fieldNode.getNextSibling;
    end;
  end;
begin
  selnode := tvwObjects.Selected;
  if not Assigned(selnode) then
    Exit;
  tabName := selnode.Text;
  sSql := GenUpdateSqlByNode(selnode);
  GetActiveEditor.Lines.Add(sSql);
end;

procedure TF_MAIN.btnWrapClick(Sender: TObject);
var
  bWraped: Boolean;
begin
  bWraped := GetActiveDBFrame.Editor.WordWrap;
  btnWrap.Down := not bWraped;
  GetActiveDBFrame.Editor.WordWrap := not bWraped;
end;

procedure TF_MAIN.actCommandHelpExecute(Sender: TObject);
begin
  ExecuteSql(C_sHeadFlag_DBCommand + ' ' + C_sDBCommand_Help);
end;

procedure TF_MAIN.actShowSettingsExecute(Sender: TObject);
begin
  ExecuteSql(C_sHeadFlag_DBCommand + ' ' + C_sDBCommand_ShowSettings);
end;

procedure TF_MAIN.tvwObjectsShowHint(Sender: TObject; Node: TTreeNode);
var
  sHint: string;
begin
  if Node <> nil then
  begin
    sHint := Node.Text;
    case TDBTreeNodeData(Node.Data).NodeTag of
      ntgTypeRoot:
      begin
      end;
      ntgTable:
      begin
      end;
      ntgProc:
      begin
      end;
      ntgField:
      begin
        if TDBTreeNodeData(Node.Data).IsPrimaryField then
          sHint := '(����)' + Node.Text;
      end;
      ntgTrig:
      begin
      end;
      ntgView:
      begin
      end;
      else
      begin
      end;
    end;
  end
  else
    sHint := 'ֱ�������ַ��������в�ѯ,������*, ?ͨ�����';
  tvwObjects.ShowHint(sHint);
end;

procedure TF_MAIN.actExportDBExecute(Sender: TObject);
begin
  with TF_ExportTables.Create(Self) do
  try
    setDBConnect(g_Global.DBConnect);
    ShowModal;
  finally
    Free;
  end;   
end;

procedure TF_MAIN.RunDBAHelp(item: TDBAHelpItem);
var
  nPageIndex: Integer;
  dbFrame: TFM_DBOperate;
  editor: TSynEdit;
begin
  if dhsoNotOpenTab in item.Option then
  begin
    dbFrame := GetActiveDBFrame;
    editor := GetActiveEditor;
  end
  else
  begin
    nPageIndex := Self.AddPage('DBAHelp_' + item.Name);
    dbFrame := GetDBFrameInPage(nPageIndex);    
    editor := dbFrame.Editor;
  end;
  editor.Lines.AddStrings(item.SQL);

  if dhsoNotRun in item.Option then
  else
    dbFrame.ExecuteSql(item.SQL.Text);
end;

procedure TF_MAIN.LoadDBAHelp;
var              
  actMenuItem: TActionClientItem;
  act: TAction;
  sDir: string;
  i: Integer;
begin
  // actmgr1.ActionBars[0].Items[5] ���˵��С��߼���һ��
  // ��Items[2] �ָ���
  // Items[3]   DHAHelp
  actmgr1.ActionBars[0].Items[5].Items[2].Visible := False;
  actmgr1.ActionBars[0].Items[5].Items[3].Visible := False;
  if not g_Global.Connected then
    Exit;

  if g_Global.DBConnect.DBEngine.GetDBType = dbtOracle then
    sDir := GetAppRootPath + 'DBA\oracle\'
  else
    Exit;  // δʵ��

  FDBAHelpList.Load(sDir);
  if FDBAHelpList.Count = 0 then
    Exit;
      
  actmgr1.ActionBars[0].Items[5].Items[2].Visible := True;
  actmgr1.ActionBars[0].Items[5].Items[3].Visible := True;
  actmgr1.ActionBars[0].Items[5].Items[3].Items.Clear;

  for i := 0 to FDBAHelpList.Count - 1 do
  begin
    FDBAHelpList.Items[i].RunProc := RunDBAHelp;
    actMenuItem := actmgr1.ActionBars[0].Items[5].Items[3].Items.Add;
    act := TAction.Create(Self);
    act.OnExecute := actOneDBAHelpExecute;
    act.Caption := FDBAHelpList.Items[i].Name;
    act.Tag := FDBAHelpList.Items[i].Index;
    actMenuItem.Action := act;
  end;  
end;

end.

