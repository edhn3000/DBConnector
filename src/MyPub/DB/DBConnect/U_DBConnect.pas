{**
 * <p>Title: U_DBConnect  </p>
 * <p>Copyright: Copyright (c) 2006/7/24</p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: 数据库操作相关类，支持多种引擎和多种数据库 </p>
 * <p>Description: TDBConnect 数据库连接及各种基本操作的类 </p>
 * <p>Description: TFieldItem TFieldItemList </p>  
 * 引用外部数据库相关单元:
 *     U_TableInfo, U_FieldInfo,
 *     U_DBEngineInterface
 * 引用外部非数据库相关单元:
 *     U_fStrUtil, U_PlanarList
 *}

unit U_DBConnect;

interface

{$M+}

uses
  ADODB, DB, DBTables, Classes, Contnrs, Variants,
  SqlExpr, SimpleDS,
  U_DBConnectInterface, U_DBEngineInterface,  
  U_TableInfo, U_FieldInfo,
  U_fStrUtil, U_DBCommand,
  U_PlanarList, U_JSON, U_TextFileWriter;

type           
  TEffectRowCounter = class
  public
    Default: Integer;
    Insert: Integer;
    Delete: Integer;
    Update: Integer;
    Select: Integer;
    Error : Integer;
    AryErrorLineNos: array of Integer;
  public
    procedure ReSet;
    procedure AddErrorLineNo(nLineNo: Integer);
    function GetErrorLinesText: string;
    function HasChanged: Boolean;
    function AllLineCount: Integer;
  end;
{ TDBConnect 数据库连接类 }
  TDBConnect = class(TInterfacedObject, IDBConnect)
  private     
    FDBType        : TDBType;                    // 数据库类型
    FDBEngine      : IDBEngine;                  // 数据库引擎对象
    FEngineShared  : Boolean;                    // 数据库引擎是否共享状态，共享的引擎不可被Free
    
    FCurrExecLine  : Integer;                    // 上次产生错误的行号
    FLastSQL       : string;                     // 最近一条sql
    FLastQry       : TDataSet;                   // 最近一个qry
    FLastSQLType   : TDBConnectSQLType;          // 最近一个sql的类型
    FLastOperSucc  : Boolean;                    // 最近一次执行的成功与否
    FLastElapsedMillis: Double;                  // 最后执行查询的消耗时间

    // 标志
    FStopExec      : Boolean;                    // 停止执行的标志
    FStopExeced    : Boolean;                    // 是否已经停止执行
    FSystemObject  : Boolean;                    // 在一些查询中指定为用户对象或系统全部对象
    FDropNoError   : Boolean;                    // drop命令不导致错误提示
    //
    FLog           : TStrings;                   // 错误列表

    // command 与 属性
    FMaxRecords    : Integer;                    // 设置后 查询时加上限制
    FVariables     : TStringList;                // 变量列表，为保存var命令的变量
    FPageIndex     : Integer;                    // 使用FPageIndex个的FnMaxRecords方式显示记录
    FRowCounter    : TEffectRowCounter;          // 行数计算器 

    // 事件
    FOnExecLineChange: TExecLineChangeEvent;     // 事件 执行一行sql时触发
    FOnLog         : TDBLogEvent;                // 事件 写log触发
    FOnConnected   : TNotifyEvent;               // 事件 成功连接数据库时触发
    FOnDisConnected: TNotifyEvent;               // 事件 断开连接时触发
    FOnExecuted    : TNotifyEvent;

    FDBCommandManager: TDBCommandManager;    // 避免单元循环引用暂用TObject
                                
    function GetLog: TStrings;
    function GetLastError: string;
    function GetConnection: TCustomConnection;
    function GetConnected: Boolean;
    procedure SetConnected(value: Boolean);
    function GetDataSet: TDataSet;
    function GetLastQry: TDataSet;
    function GetLastSQL: string;
    function GetLastSQLType: TDBConnectSQLType;
    function GetLastOperSucc: Boolean;

    function GetMaxRecords: Integer;
    procedure SetMaxRecords(value: Integer);
    procedure SetExecLine(value: Integer);
    function GetPageIndex: Integer;
    procedure SetPageIndex(value: Integer);
    function GetDropNoError: Boolean;
    procedure SetDropNoError(value: Boolean);   
    function GetVariables: TStringList;
    procedure SetVariables(value: TStringList);
    function GetSystemObject: Boolean;
    procedure SetSystemObject(value: Boolean);
    function GetLastElapsedMilis: Double;

    function GetDBType(): TDBType;                                                        
    function GetDBEngine(): IDBEngine;
    function GetNewDBEngine(Adbet: TDBEngineType): IDBEngine;
    procedure CheckDBEngine(Adbet: TDBEngineType);
    function GetOnLog: TDBLogEvent;
    procedure SetOnLog(value: TDBLogEvent);
    function GetOnConnected: TNotifyEvent;
    function GetOnDisConnected: TNotifyEvent;
    procedure SetOnConnected(value: TNotifyEvent);
    procedure SetOnDisConnected(value: TNotifyEvent);
    function GetOnExecLineChange: TExecLineChangeEvent;
    procedure SetOnExecLineChange(value: TExecLineChangeEvent);
    function GetOnExecuted: TNotifyEvent;
    procedure SetOnExecuted(value: TNotifyEvent);
    function GetExportValidFieldValue(Field: TField;
      const sValueDef: string): string;
    procedure ProcessLobExport(qry: TDataSet; tableInfo: TTableInfo;
      sInsertSqlPrefix: string; slstPlanar: TPlanarStringList;
      sExportFileName: string; ClobInFile: Boolean; RemoveBreakLine: Boolean);
  
  protected        
    // last items 都是protected
    FsLastScript   : string;                     // 上次执行的脚本
    FbInScript     : Boolean;                    // 是否正在执行脚本

    function GetSqlFromQuery(ds: TDataSet; var sql: string): Boolean;
    procedure SetDBType(value: TDBType);
    
    procedure FillListByField(List: TStrings; ADataSet: TDataSet;
      FieldNames: array of string);     
    procedure FillListByOwnerName(List: TStrings; ADataSet: TDataSet);
    procedure GetObjectNames(List: TStrings; sSql: string; AQry: TDataSet = nil);

    function IsBlobField(Field: TField): Boolean;virtual;
    function IsClobField(Field: TField): Boolean;virtual;
    function GetNewTableInfo(): TTableInfo;virtual; // 各数据库返回特有Item
//    function GetNewFieldItem(Field: TField): TFieldItem; // 各数据库返回特有Item
    function SaveBlobs(blobFields: TFieldItemList; tableInfo: TTableInfo;
      sDir: string; var json: TJSON): Boolean;virtual;
    function SaveClobs(clobFields: TFieldItemList; tableInfo: TTableInfo;
      sDir: string; var json: TJSON): Boolean;virtual;

    // 使用多个引擎尝试
    function TestOpenDB(sDataSource: string;sUser: string;sPwd: string;
      dbt: TDBType; dbetAry: array of TDBEngineType): Boolean;

    // 功能方法
    function IsAbsoluteSql(const sSql: string):Boolean;
    function IsQryInvalid(Ads: TDataSet): Boolean;
    function GetMaxRecordSQLCondition(nMaxRecord: Integer;
      sSql: string): string;virtual;
  public
    property DBType: TDBType read GetDBType;
    property DBEngine: IDBEngine read GetDBEngine;

    property LastError: string read GetLastError;                        
    property LastSQL: string read GetLastSQL;
    property LastSQLType: TDBConnectSQLType read GetLastSQLType;
    property LastOperSucc: Boolean read GetLastOperSucc;
    property LastQry: TDataSet read GetLastQry;
    property Log: TStrings read GetLog;
    property SystemObject: Boolean read GetSystemObject write SetSystemObject;

    property DataSet: TDataSet read GetDataSet;
    property Connection: TCustomConnection read GetConnection;

    property Connected: Boolean read GetConnected write SetConnected;
    property MaxRecords: Integer read GetMaxRecords write SetMaxRecords;
    property PageIndex: Integer read GetPageIndex write SetPageIndex;
    property DropNoError: Boolean read GetDropNoError write SetDropNoError;
    property Variables: TStringList read GetVariables write SetVariables;
  public
    constructor Create; virtual;
    destructor Destroy; override;
    // 运行期可改参数重置，同时会用于setdefault命令
    procedure ResetParams;

    procedure ClearLog;      
    procedure AddLog(sLog: string); 
    procedure AddErrorLog(sLog: string);
    procedure AddInfoLog(sLog: string);
    procedure AddSqlExecLog(nEffectRows: Integer; sCommandContent: string;
       relData: string);

    // 可使多个Connect对象共用一个Engine而不必重复创建新的Engine并连接数据库
    procedure ShareEngine(DestDBEngine: IDBEngine; bCopy: Boolean = False);

    { 数据库操作 }
    function OpenDB(sDataSource: string;sUser: string;sPwd: string;
      dbt: TDBType; dbet: TDBEngineType): Boolean;overload;virtual;
    function OpenDB(AIniFile, ADBSection: string): Boolean;overload;virtual;
    function CloseDB: Boolean;
    function CloseQuery: Boolean;
    function GetNewQuery: TDataSet; virtual;
//    function GetCanShowDataSet(ds: TDataSet): TDataSet;virtual;
    procedure StopExec(); 

    //*************** virtual ***************
    procedure GetDBs(List: TStrings);virtual;
    procedure GetTableNames(List: TStrings;
      AQry: TDataSet = nil);virtual;
    // 获得对象列表，可得到更多信息        
    function GetTableInfo(sTableName: string): TTableInfo;virtual;
    function GetTableInfos(AQry: TDataSet = nil):TTableInfoList;virtual;
    function GetFieldInfos(const tableName: string; AQry: TDataSet = nil): TFieldItemList;virtual;
    procedure GetViewNames(List: TStrings; AQry: TDataSet = nil);virtual;
    procedure GetFieldNames(const tableName: string; List: TStrings;
      option: TFieldShowOptions; AQry: TDataSet = nil);virtual;
    procedure GetProcedureNames(List: TStrings; AQry: TDataSet = nil);virtual;
    procedure GetTriggerNames(List: TStrings; AQry: TDataSet = nil);virtual; 
    function GetPrimaryField(TableName: string; AQry: TDataSet = nil): TFieldItem; virtual;
    function IsFieldUnique(tableName, fieldName: string;
      AQry: TDataSet = nil): Boolean; virtual;
    procedure GetIndexNames(List: TStrings);virtual;
    procedure CheckFieldsUnique(tableName: string; fields: TFieldItemList); virtual;
    procedure CheckFieldUnique(tableName: string; field: TFieldItem);

    function GenTableCreateSQL(sTableName: string; slst: TStrings):Boolean;virtual;
    function ExportQuery(AQry: TDataSet; exportFile: TTextFileWriter;
      ClobInFile: Boolean; RemoveBreakLine: Boolean): Boolean;overload;virtual;
    function ExportQuery(AQry: TDataSet; sExportFileName: string):Boolean;overload; virtual;
    function ExportTableData(tableName, sExportFileName: string;
      option: TDBExportOption): Boolean;

    function GetValidFieldValue(FieldItem: TFieldItem;
      const sValueDef: string = ''): string;overload;virtual;
    function GetValidFieldValue(Field: TField;
      const sValueDef: string = ''): string;overload;
    function GetFieldsKeyValuesStr(Fields: TFieldItemList): string;virtual;
    function GetFieldsValuesStr(Fields: TFieldItemList): string;virtual;
    function TableExists(sTableName: string): Boolean;
    //*************** virtual ***************

    { sql处理 }
    function RemoveHeadFlag(Asql: string): string;
    procedure SetVariableValue(var s: string);

    { 执行sql的方法 }
    function ExecQuery(sSql: string):Boolean; overload;
    function ExecQuery(AQ: TDataSet; sSql: string):Boolean; overload;virtual;
    function ExecQueryWithParams(sSqlWithParams: string;
      Params: array of Variant):Boolean; overload;
    function ExecQueryWithParams(AQ: TDataSet; sSqlWithParams: string;
        aryParams: array of Variant):Boolean;overload;virtual;
    function ExecUpdate(sSql: string): Integer; overload;
    function ExecUpdate(AQ: TDataSet; sSql: string): Integer; overload;virtual;
    function ExecUpdateWithParams(sSqlWithParams: string;
      Params: array of Variant): Integer; overload;
    function ExecUpdateWithParams(AQ: TDataSet; sSqlWithParams: string;
        aryParams: array of Variant): Integer;overload;virtual;

    function ExecOneSql(Asql: string; aQry: TDataSet = nil): Integer;virtual;
    // 执行多条sql
    function ExecSqlList(slstSqlList:TStrings; aQry: TDataSet = nil):Integer;
    // 不要用来执行脚本，会去除所有换行重新分行
    function ExecSqlText(sSqlText: string; cDelimiter: Char = ';';
      aQry: TDataSet = nil): Integer;
    // 专门用于执行脚本文件
    function ExecScriptFile(AScriptFile: string; aQry: TDataSet = nil):Integer;

    function GetProcSource(sName: string; list: TStrings): string;virtual;
    function GetTrigSource(sName: string; list: TStrings): string;virtual;
    function GetViewSource(sName: string; list: TStrings): string;virtual;
     
    function UpdateBlobs(JsonStr: string): Integer;virtual;
    function UpdateClobs(JsonStr: string): Integer;virtual;

    // access
    procedure GetMoudleNames(List: TStrings; AQry: TDataSet = nil);virtual;
    procedure GetReportNames(List: TStrings; AQry: TDataSet = nil);virtual;
    procedure GetMacroNames(List: TStrings; AQry: TDataSet = nil);virtual;
    procedure GetFormNames(List: TStrings; AQry: TDataSet = nil);virtual;
    procedure GetPageNames(List: TStrings; AQry: TDataSet = nil);virtual;

    // oracle    
    procedure GetUsers(List: TStrings);virtual;
    procedure GetSynonymNames(List: TStrings);virtual;
    procedure GetDBLinkNames(List: TStrings);virtual;

  published                                
    property OnLog: TDBLogEvent read GetOnLog write SetOnLog;
    property OnConnected: TNotifyEvent read GetOnConnected write SetOnConnected;
    property OnDisConnected: TNotifyEvent read GetOnDisConnected write SetOnDisConnected;
    property OnExecLineChange: TExecLineChangeEvent read GetOnExecLineChange
      write SetOnExecLineChange;
    property OnExecuted: TNotifyEvent read GetOnExecuted write SetOnExecuted;
  end;

  // 传入连接串，将其拆分，要求形式为如下
  // oracle kbuser/tusc@tnsname|SID, IP, Port, Protocol
  // access tax/tax@db,Secured
//  function SplitConnectParams(sConnectCmd: string; var dbt: TDBType;
//    var sUser, sPass, sDB, sOtherDBParam: string; var sErrMsg: string): Boolean;


implementation

uses
  SysUtils, StrUtils, Forms,
  U_DBEngineFactory, U_SqlUtils;

function GetMillisFromDateTime(dt: TDateTime): Double;
var
  fractional: Double;
  fsResult: Double;
begin
  fractional := dt - Trunc(dt);
  fsResult := fractional * 24 * 3600;  // 现在单位是秒
  Result := fsResult*1000;
end;


function SplitConnectParams(sConnectCmd: string; var dbt: TDBType;
    var sUser, sPass, sDB, sOtherDBParam: string; var sErrMsg: string): Boolean;
var
  params: TStrings;
  sParam: string;
  nPos,nPos2,nPos3: Integer;
begin
  Result := False;
  params := TStringList.Create;
  try
    fStrUtil.SplitEscapeQuote(sConnectCmd, ' ', '"', '\', params);
    if params.Count < 2 then
    begin
      sErrMsg := 'Conn命令参数格式：DbTypeStr UserName/Password@DataSourceStr';
      Exit;
    end;
    
    // 第一个参数 是数据库类型
    dbt := StrToDBType(params[0]);
    if dbt = dbtUnKnown then
    begin
      sErrMsg := '无法从参数1分析出数据库类型';
      Exit;
    end;

    sParam := params[1];
    // 第二个参数是User/Pass@DB
    nPos  := Pos('/', sParam);
    nPos2 := Pos('@', sParam);
    nPos3 := Pos(C_sSEPARATOR_DATASOURCE, sParam);     // 可以没有
    if nPos3 = 0 then
    begin
      nPos3 := MaxInt;
      sOtherDBParam := '';
    end
    else
    begin
      sOtherDBParam := Copy(sParam, nPos3+1, MaxInt);
      if Copy(sOtherDBParam, 1, 1) = '"' then
        sOtherDBParam := Copy(sOtherDBParam, 2, MaxInt);
      if Copy(sOtherDBParam, Length(sOtherDBParam), 1) = '"' then
        sOtherDBParam := Copy(sOtherDBParam, 1, Length(sOtherDBParam)-1);
    end;
    if (nPos = 0) or (nPos2 = 0) then
    begin
      sErrMsg := '无法从参数2分析出用户信息';
      Exit;
    end;    
    sUser := Copy(sParam, 1, nPos-1);
    sPass := Copy(sParam, nPos+1, nPos2-nPos-1);
    sDB   := Copy(sParam, nPos2+1, nPos3-nPos2-1);
    
    if Copy(sDB, 1, 1) = '"' then
      sDB := Copy(sDB, 2, MaxInt);
    if Copy(sDB, Length(sDB), 1) = '"' then
      sDB := Copy(sDB, 1, Length(sDB)-1);
  finally
    params.Free;
  end;
  Result := True;
end;

{ TEffectRowCounter }

procedure TEffectRowCounter.AddErrorLineNo(nLineNo: Integer);
begin
  SetLength(AryErrorLineNos, Length(AryErrorLineNos)+1);
  AryErrorLineNos[High(AryErrorLineNos)] := nLineNo;
end;

function TEffectRowCounter.AllLineCount: Integer;
begin
  Result := Default + Insert + Delete + Update + Select + Error;
end;

function TEffectRowCounter.GetErrorLinesText: string;
var
  i: Integer;
begin
  for i := Low(AryErrorLineNos) to High(AryErrorLineNos) do
    if Result = '' then
      Result := IntToStr(AryErrorLineNos[i])
    else
      Result := Result + ',' + IntToStr(AryErrorLineNos[i]);
end;

function TEffectRowCounter.HasChanged: Boolean;
begin
  Result := (Insert <> 0) or (Delete <> 0) or (Update <> 0) or (Default <> 0)
    or (Error <> 0) or (Length(AryErrorLineNos) <> 0);
end;

procedure TEffectRowCounter.ReSet;
begin          
  Default := 0;
  Insert  := 0;
  Delete  := 0;
  Update  := 0;
  Select  := 0;
  Error   := 0;
  SetLength(AryErrorLineNos, 0);
end;    

{ TDBConnect }

// private

function TDBConnect.IsBlobField(Field: TField): Boolean;
begin
  Result := Field.DataType in [ftOraBlob, ftBlob];
end;

function TDBConnect.IsClobField(Field: TField): Boolean;
begin
  Result := Field.DataType in [ftOraClob, ftMemo];
end;  

function TDBConnect.GetNewTableInfo: TTableInfo;
begin
  Result := TTableInfo.Create;
end;

function TDBConnect.IsAbsoluteSql(const sSql: string):Boolean;
var
  nPos, nLen: Integer;
  s: string;
begin
  Result := False;
  s := Trim(sSql);
  nLen := Length(C_sHeadFlag_Script);
  if (Copy(s, 1, nLen) = C_sHeadFlag_Script) and
     (Copy(s, 1+nLen, nLen) <> C_sHeadFlag_Script) then     // 脚本
  begin
    Result := True;
    Exit;
  end;
  nPos := Pos(' ', s);
  if nPos <> 0 then
    Result := TSqlUtils.IsSqlBeginCommand(Trim(Copy(s, 1, nPos-1)));
end;  

function TDBConnect.RemoveHeadFlag(Asql: string): string;
begin
  if Copy(Asql, 1, Length(C_sHeadFlag_Force))= C_sHeadFlag_Force then
    Result := Copy(Asql, Pos(' ', Asql), Length(Asql))
  else
    Result := Asql;
end;

constructor TDBConnect.Create;
begin       
  FLastOperSucc := True;
  FLastSQLType := dbcstCommand;
  FLastElapsedMillis := 0;

  FStopExec := False;
  FStopExeced := True;
  FbInScript := False;
  FEngineShared := False;
   
  FLog := TStringList.Create;
              
  FVariables := TStringList.Create;

  FRowCounter := TEffectRowCounter.Create;
  FRowCounter.ReSet;
                 
  FPageIndex := 1;

  FDBCommandManager := TDBCommandManager.Create;
  FDBCommandManager.DBConnect := Self;

  ResetParams;
  FDBEngine := nil;
end;   

procedure TDBConnect.ResetParams;
begin  
  FDropNoError := False;
  MaxRecords := 0;
end;

destructor TDBConnect.Destroy;
begin                
  FDBCommandManager.Free;

  if Assigned(FRowCounter) then
    FreeAndNil(FRowCounter);
  if Assigned(FVariables) then
    FreeAndNil(FVariables);
  if Assigned(FDBEngine) and not FEngineShared then
    FDBEngine := nil;
  if Assigned(FLog) then
    FreeAndNil(FLog);
  inherited;
end;

function TDBConnect.ExecQuery(sSql: string):Boolean;
begin
  Result := DBEngine.ExecQuery(sSql); 
  FLastSQLType := dbcstQuery;
  FLastSQL := sSql;
  if not Result then
    AddErrorLog(DBEngine.GetLastError);
end;

function TDBConnect.ExecQueryWithParams(sSqlWithParams: string;
  Params: array of Variant):Boolean;
begin
  Result := DBEngine.ExecQueryWithParams(sSqlWithParams, Params);
  FLastSQLType := dbcstQuery;  
  FLastSQL := sSqlWithParams;
  if not Result then
    AddErrorLog(DBEngine.GetLastError);
end;

function TDBConnect.ExecUpdate(sSql: string): Integer;
begin
  Result := DBEngine.ExecUpdate(sSql);
  FLastSQLType := dbcstUpdate; 
  FLastSQL := sSql;
  if Result = C_nTDBE_ERROR_EXECFAIL then
    AddErrorLog(DBEngine.GetLastError);
end;

function TDBConnect.ExecUpdateWithParams(sSqlWithParams: string;
  Params: array of Variant): Integer;
begin 
  Result := DBEngine.ExecUpdateWithParams(sSqlWithParams, Params);
  FLastSQLType := dbcstUpdate;
  FLastSQL := sSqlWithParams;
  if Result = C_nTDBE_ERROR_EXECFAIL then
    AddErrorLog(DBEngine.GetLastError);
end;

procedure TDBConnect.ShareEngine(DestDBEngine: IDBEngine; bCopy: Boolean);
var
  event: TNotifyEvent;
begin
  if bCopy then
  begin
    if Assigned(Self.FDBEngine) then    
      Self.FDBEngine := nil;
    CheckDBEngine(DestDBEngine.GetDBEngineType);
    Self.FDBEngine.SetOnConnected(DestDBEngine.GetOnConnected);
    Self.FDBEngine.SetOnConnected(DestDBEngine.GetOnDisConnected);

    if DestDBEngine.IsConnected then
      OpenDB(DestDBEngine.GetDataSource, DestDBEngine.GetUserName,
        DestDBEngine.GetPassword, DestDBEngine.GetDBType,
        DestDBEngine.GetDBEngineType);
    FEngineShared := False;
  end
  else
  begin
    if Self.FDBEngine <> DestDBEngine then
    begin
      if Assigned(Self.FDBEngine) then
        Self.FDBEngine := nil;
      Self.FDBEngine := DestDBEngine;     // 直接指向传入的DBEngine

      if Connected then
      begin
        if Assigned(FOnConnected) then
        begin
          FOnConnected(Self);
        end
        else if Assigned(DBEngine.GetOnConnected) then
        begin
          event := DBEngine.GetOnConnected();
          event(Self);
        end
      end
      else
      begin 
        if Assigned(FOnDisConnected) then
        begin
          FOnDisConnected(Self);
        end
        else if Assigned(DBEngine.GetOnConnected) then
        begin
          event := DBEngine.GetOnConnected();
          event(Self);
        end
      end;
      FEngineShared := True;
    end;
  end;
end;

procedure TDBConnect.AddSqlExecLog(nEffectRows: Integer;
  sCommandContent: string; relData: string);
var
  sLogPrefix: string;
begin
  if FbInScript then
    sLogPrefix := 'ExecLogInScript'
  else
    sLogPrefix := 'ExecLog';
  if SameText(sCommandContent, 'delete') then begin
    Inc(FRowCounter.Delete, nEffectRows);
    AddLog(Format('%s: 删除的行数%d',[sLogPrefix, nEffectRows]));
  end else if SameText(sCommandContent, 'insert') then begin
    Inc(FRowCounter.Insert, nEffectRows);
    AddLog(Format('%s: 插入的行数%d',[sLogPrefix, nEffectRows]));
  end else if SameText(sCommandContent, 'update') then begin
    Inc(FRowCounter.Update, nEffectRows);
    AddLog(Format('%s: 更新的行数%d',[sLogPrefix, nEffectRows]));
  end else if SameText(sCommandContent, 'create') then begin
    if (nEffectRows >= 0) then
      AddLog(Format('%s: 创建成功。表名=%s',[sLogPrefix, relData]))
    else
      AddLog(Format('%s: 创建失败。表名=%s',[sLogPrefix, relData]));
  end else if SameText(sCommandContent, 'drop') then begin
    if (nEffectRows >= 0) then
      AddLog(Format('%s: 销毁成功。表名=%s',[sLogPrefix, relData]))
    else begin
      if FDropNoError then
        // 开启此开关，失败时不记录日志（但也不会记录成功日志）
      else
        AddLog(Format('%s: 销毁失败。表名=%s'
          +'打开DropNoError开关可屏蔽此提示。',[sLogPrefix, relData]));
    end;
  end else if SameText(sCommandContent, C_sHeadFlag_Blob) then begin 
    if (nEffectRows >= 0) then
      AddLog(Format('%s: 更新Blob成功。影响的数据行数%d',[sLogPrefix, nEffectRows]))
    else
      AddLog(Format('%s: 更新Blob失败。',[sLogPrefix, nEffectRows]));
  end else begin  // 其他影响行数
    if (nEffectRows >= 0) then
      AddLog(Format('%s: 执行成功。影响的数据行数%d',[sLogPrefix, nEffectRows]))
    else
      AddLog(Format('%s: 执行失败。',[sLogPrefix, nEffectRows]));
    Inc(FRowCounter.Default, nEffectRows);
  end;
end;

function TDBConnect.ExecOneSql(Asql: string; aQry: TDataSet): Integer;
var
  bOpen: Boolean;
  sRealSql: string;
  bError: Boolean;
  nOneExecResult: Integer;
  theQry: TDataSet;
  sError: string;
  sSqlCommand: string;
  sRelData: string;
begin
  nOneExecResult := 0;
  FPageIndex := 1;
  if aQry = nil then
    theQry := DataSet
  else
    theQry := aQry;
  Result := 0;
  bOpen :=  not AnsiStartsText(C_sHeadFlag_ForceUpdate, Asql)
            and (TSqlUtils.IsQuerySql(Asql)
                 or AnsiStartsText(C_sHeadFlag_ForceQuery, Asql)
                );

  // RemoveHeadFlag
  if Copy(Asql, 1, Length(C_sHeadFlag_Force))= C_sHeadFlag_Force then
    sRealSql := Copy(Asql, Pos(' ', Asql), Length(Asql))
  else
    sRealSql := Asql;
  
  try
    if AnsiStartsText(C_sHeadFlag_Script, sRealSql) then begin   // 执行脚本
      nOneExecResult := ExecScriptFile(Copy(sRealSql,
        1+Length(C_sHeadFlag_Script), MaxInt), theQry);
      bError := nOneExecResult < 0;
      Result := nOneExecResult;
    end else if AnsiStartsText('exec',sRealSql) then  begin      // 存储过程
      //TODO:
      bError := False;
    end else begin                                                   // 一般sql
      // 变量替换处理
      if Pos(C_sDBCmdParam_Var_Start, sRealSql) > 0 then begin
        SetVariableValue(sRealSql);
      end;
      FLastQry := theQry; 
      if bOpen then begin
        if FMaxRecords > 0 then begin
          sRealSql := GetMaxRecordSQLCondition(FMaxRecords*FPageIndex, sRealSql);
          bError := not DBEngine.ExecQuery(theQry, sRealSql)
        end else begin
          bError := not DBEngine.ExecQuery(theQry, sRealSql);
        end;
        FLastSQLType := dbcstQuery;
      end else begin
        nOneExecResult := DBEngine.ExecUpdate(theQry, sRealSql);
        bError := nOneExecResult = C_nTDBE_ERROR_EXECFAIL;
        Result := nOneExecResult;
        FLastSQLType := dbcstUpdate;
      end;
    end;
  except
    on ex: Exception do begin
      bError := True;
      sError := ex.Message;
    end;
  end;    

  sSqlCommand := TSqlUtils.GetSqlCommand(sRealSql, sRelData);
  FLastSQL := sRealSql;     
  FLastOperSucc := not bError;

  if bError then begin
    if (LowerCase(sSqlCommand) = 'drop') and FDropNoError then
      //
    else begin
      Inc(FRowCounter.Error);
      FRowCounter.AddErrorLineNo(FCurrExecLine);
      Result := C_nTDBE_ERROR_EXECFAIL;
      sError := DBEngine.GetLastError;
      if not FbInScript then
        AddErrorLog(Format('InScript(%s) Line(%d) '+ sError,[
          FsLastScript, FCurrExecLine]))
      else
        AddErrorLog(Format('Line(%d) ' + sError, [FCurrExecLine]));
    end;
    
    nOneExecResult := TDBC_ERROR_EXECFAIL;
  end;

  if (not bOpen) then begin
    AddSqlExecLog(nOneExecResult, sSqlCommand, sRelData);
  end;
end;

function TDBConnect.ExecSqlList(slstSqlList: TStrings; aQry: TDataSet): Integer;
var
  i, nEffectRows: Integer;
  TheSql: string;
  sSummeryLog: string;
  sOneLine: string;
  execTime: Double;

  function GetCombineSql(sqlList: TStrings; var nStartIndex: Integer): string;
  var
    sCombinedSql: string;
    nLine: Integer;  
    sLine: string;
    isInBlock: Boolean;
  begin
    nLine := nStartIndex;
    isInBlock := False;
    while nLine <= sqlList.Count - 1 do
    begin
      sLine := Trim(sqlList[nLine]);      
      // check if has 'begin'
      if not isInBlock then
        isInBlock :=
          fStrUtil.MatchRegxSubStr('[A,a][S,s]\s*$', sLine) or
          (fStrUtil.PosFrom('Begin', sLine, 1, True) = 1);

      // combine sql 
      if sCombinedSql = '' then
        sCombinedSql := TSqlUtils.DeleteSQLComment(sLine)
      else if Copy(sCombinedSql, Length(sCombinedSql), 1) = ' ' then
        sCombinedSql := sCombinedSql + TSqlUtils.DeleteSQLComment(sLine)
      else
        sCombinedSql := sCombinedSql + ' ' + TSqlUtils.DeleteSQLComment(sLine);
      sCombinedSql := Trim(sCombinedSql);

      // check end of sql, or end of proc
      if isInBlock then
      begin
        if (fStrUtil.PosFrom('End', sLine, 1, True) = 1)
           and (fStrUtil.PosFrom('End IF', sLine, 1, True) = 0) then
        begin
          Break;
        end;
      end
      else
      begin
        if C_sSQLEnd = Copy(sLine, Length(sLine), 1) then
          Break;
//        nEndPos := fStrUtil.PosEscapeQuote(C_sSQLEnd, sCombinedSql, '''', '''');
      end;
      
      Inc(nLine);
    end;    

    nStartIndex := nLine;
    Result := sCombinedSql;
  end;  
//// ExecSqlList Begin
begin
  nEffectRows := 0;
  SetExecLine(1);
  FStopExec := False;
  FStopExeced := False;
  execTime := Now;
//  ClearLog;       // 错误列表清空
  i := 0;
  FRowCounter.ReSet;
  while i <= slstSqlList.Count - 1 do begin
    Application.ProcessMessages;
    
    SetExecLine(i);
    if FStopExec then begin  // 退出标志
      FStopExeced := True;
      Break;
    end;

    sOneLine := Trim(slstSqlList[i]);
    if sOneLine = '' then begin
      Inc(i);
      Continue;                      // 脚本中的空行
    end;
    // 单行注释直接跳过
    if TSqlUtils.IsSingleLineComment(sOneLine) then begin
      TheSql := Trim(TSqlUtils.DeleteSQLComment(sOneLine, True));
      if FDBCommandManager.IsDBConnectCommand(TheSql) then
      begin
        FLastQry := aQry;
        FLastSQLType := dbcstCommand;
        nEffectRows := nEffectRows + FDBCommandManager
          .ProcessCommand(TheSql);
      end;
      Inc(i);
      Continue;
    end;

    TheSql := Trim(GetCombineSql(slstSqlList, i));

    if TheSql = '' then begin
      Inc(i);
      Continue;                             // 有效内容为空
    end;

    // ********特有数据库命令******** //
    if FDBCommandManager.IsDBConnectCommand(TheSql) then begin
      FLastQry := aQry;
      nEffectRows := nEffectRows + FDBCommandManager.ProcessCommand(TheSql);
      Inc(i);
      Continue;
    end;

    if not Connected then begin
      nEffectRows := TDBC_ERROR_NOTCONNECT;
      Break;
    end;

    // ********开始执行sql******** //
    TheSql := Trim(TheSql);  // 这个去空格只在给SQL.Text赋值之前做，确保后面判断执行方式时 无行首空格

    // 分号只用于分析结束位置，在实际执行的时候去掉
    if Copy(TheSql, Length(TheSql), 1) = ';' then begin
      TheSql := Copy(TheSql, 1, Length(TheSql)-1);
    end;
    nEffectRows := nEffectRows + ExecOneSql(TheSql, aQry);
    Inc(i);
  end; // of while

  // 如果恰好为-2会误认为是未连接的标志
  if Connected and (nEffectRows<0) then
    nEffectRows := TDBC_ERROR_EXECFAIL;
    
  if FRowCounter.HasChanged then begin
    sSummeryLog := '';
    if FRowCounter.Delete <> 0 then
      sSummeryLog := sSummeryLog + Format(' 删除%d行',[FRowCounter.Delete]);
    if FRowCounter.Insert <> 0 then
      sSummeryLog := sSummeryLog + Format(' 插入%d行',[FRowCounter.Insert]);
    if FRowCounter.Update <> 0 then
      sSummeryLog := sSummeryLog + Format(' 更新%d行',[FRowCounter.Update]);
    if FRowCounter.Default <> 0 then
      sSummeryLog := sSummeryLog + Format(' 其他影响行数%d行',[FRowCounter.Default]);
    if FRowCounter.Error <> 0 then begin
      sSummeryLog := sSummeryLog + Format(' 错误%d行',[FRowCounter.Error]);
      sSummeryLog := sSummeryLog + Format(' 错误行号=%s',[FRowCounter.GetErrorLinesText]);
    end;
    if sSummeryLog <> '' then begin
      if FbInScript then         
        AddLog('Summery InScript'+': '+sSummeryLog)
      else
        AddLog('Summery: '+sSummeryLog);
    end;
  end;

  FLastElapsedMillis := GetMillisFromDateTime(Now - execTime);
  Result := nEffectRows;
end;

function TDBConnect.ExecSqlText(sSqlText: string;
  cDelimiter: Char;aQry: TDataSet): Integer;
var
  slst: TStringList;
begin
  slst := TStringList.Create;
  try
    // 根据换行符拆分成列表
    slst.Text := sSqlText;
    Result := ExecSqlList(slst, aQry);
  finally
    slst.Free;
  end;
  if Assigned(FOnExecuted) then
    FOnExecuted(Self);
end;

function TDBConnect.ExecScriptFile(AScriptFile: string; aQry: TDataSet): Integer;
var
  slstScript: TStringList;
  function GetValidFullPath(Str: string):string;
  begin
    if (Trim(Str) = '') or (Pos(':', Str) <> 0)
       or (Pos('\\', Str) = 1) then
      Result := Str
    else
      Result := ExtractFilePath(Application.ExeName)+Str;
  end;
begin
  slstScript := TStringList.Create;
  try
    try
      AScriptFile := GetValidFullPath(AScriptFile);
      if not FileExists(AScriptFile) then begin
        Result := TDBC_ERROR_SCRIPT_FILENOTEXISTS;  
        AddErrorLog('找不到脚本文件' + AScriptFile);
        Exit;
      end;
      FsLastScript := AScriptFile;
  //    RemoveNewLine(FsLastScript, '');     //文件名中不能有换行符，否则会有异常
      slstScript.LoadFromFile(FsLastScript);
      FbInScript := True;
      Result := ExecSqlList(slstScript,aQry);
    except
      on ex: Exception do begin
        Result := C_nTDBE_ERROR_EXECFAIL;
        AddErrorLog('执行脚本文件出错。' + ex.Message);
      end;
    end;
  finally
    FbInScript := False;
    slstScript.Free;
  end;
end;  

function TDBConnect.IsQryInvalid(Ads: TDataSet): Boolean;
begin
  Result := True;
  if not Assigned(Ads) then
    Exit;

  if Ads is TADOQuery then
    Result := TADOQuery(Ads).Connection = nil
  else if Ads is TQuery then
    Result := TQuery(Ads).Database = nil
  else if Ads is TSQLDataSet then
    Result := TSQLDataSet(Ads).SQLConnection = nil;
end;   

function TDBConnect.GetMaxRecordSQLCondition(nMaxRecord: Integer;
  sSql: string): string;
begin
  Result := sSql; // to be inherited
end;

function TDBConnect.OpenDB(AIniFile, ADBSection: string): Boolean;
begin
  CheckDBEngine(dbetADO);
  Result := DBEngine.OpenDBWithIni(AIniFile, ADBSection);
end;

function TDBConnect.TestOpenDB(sDataSource, sUser, sPwd: string;
  dbt: TDBType; dbetAry: array of TDBEngineType): Boolean;
var
  i: Integer;
begin
  Result := False;
  for i := Low(dbetAry) to High(dbetAry) do begin
    CheckDBEngine(dbetAry[i]);
    if DBEngine.IsEngineInstalled then begin
      Result := DBEngine.OpenDB(sDataSource, sUser, sPwd, dbt);
      if Result then
        Break
      else
        AddErrorLog(DBEngineTypeToStr(dbetAry[i], False) + ' engine error:'
          + DBEngine.GetLastError);
    end;
  end;
end;

function TDBConnect.OpenDB(sDataSource, sUser, sPwd: string;
  dbt: TDBType; dbet: TDBEngineType): Boolean;
begin
  FDBType := dbt;
  if dbet <> dbetAuto then begin
    Result := TestOpenDB(sDataSource, sUser, sPwd, dbt, [dbet]);
  end else begin
    case dbt of
    dbtAccess, dbtAccess2007:
    begin
      Result := TestOpenDB(sDataSource, sUser, sPwd, dbt, [dbetADO, dbetDBExpress, dbetBDE]);
    end;
    dbtOracle:
    begin
      Result := TestOpenDB(sDataSource, sUser, sPwd, dbt, [dbetADO, dbetDBExpress, dbetBDE]);
    end;
    dbtMySQL:
    begin
      Result := TestOpenDB(sDataSource, sUser, sPwd, dbt, [dbetADO, dbetDBExpress]);
    end;
    dbtSybaseASA:
    begin
      Result := TestOpenDB(sDataSource, sUser, sPwd, dbt, [dbetDBExpress, dbetADO]);
    end;
    dbtSqlite:
    begin
      Result := TestOpenDB(sDataSource, sUser, sPwd, dbt, [dbetSqlite]);
    end;
    else begin
      Result := TestOpenDB(sDataSource, sUser, sPwd, dbt, [dbetADO,
        dbetDBExpress, dbetBDE]);
    end;
    end;
    if not Result then
      AddErrorLog(DBEngine.GetLastError);
  end;
end;

function TDBConnect.GetNewQuery: TDataSet;
begin
  if DBEngine = nil then begin
    Result := nil;
    Exit;
  end;
  Result := DBEngine.GetNewQuery;
end;

procedure TDBConnect.StopExec();
begin
  FStopExec := True;
end;

procedure TDBConnect.SetDBType(value: TDBType);
begin
  FDBType := value;
end;

procedure TDBConnect.SetExecLine(value: Integer);
begin
  if FCurrExecLine <> value then begin
    FCurrExecLine := value;
    if Assigned(FOnExecLineChange) then
      FOnExecLineChange(value);
  end;
end;

procedure TDBConnect.FillListByField(List: TStrings; ADataSet: TDataSet;
  FieldNames: array of string);
var
  i: Integer;
  sValues: string;
  fld: TField;
begin
  with ADataSet do begin
    List.BeginUpdate;
    List.Clear;
    while not Eof do begin
      sValues := '';
      for i := Low(FieldNames) to High(FieldNames) do begin
        fld := FieldByName(FieldNames[i]);
        if sValues = '' then
          sValues := fld.AsString
        else
          sValues := sValues + ' ' + fld.AsString;
      end;  
      List.Add(sValues);
      Next;
    end;
    List.EndUpdate;
  end;
end;

procedure TDBConnect.FillListByOwnerName(List: TStrings;
  ADataSet: TDataSet);
var
  sValues: string;
  sOwner, sName: string;
begin
  with ADataSet do begin
    List.BeginUpdate;
    List.Clear;
    while not Eof do begin
      sValues := '';
      if SystemObject then begin
        sOwner := FieldByName('OWNER').AsString;
        sName := FieldByName('NAME').AsString;
        if LowerCase(sOwner) <> LowerCase(DBEngine.GetUserName) then
          sValues := sOwner + '.' + sName
        else        
          sValues := sName;
      end else begin
        sName := FieldByName('NAME').AsString;
        sValues := sName;
      end;
          
      List.Add(sValues);
      Next;
    end;
    List.EndUpdate;
  end;
end;    

procedure TDBConnect.GetObjectNames(List: TStrings; sSql: string; AQry: TDataSet);
begin
  if IsQryInvalid(AQry) then
    AQry := DBEngine.GetDataSet;
  if ExecQuery(AQry, sSql) then
    FillListByOwnerName(List, AQry);
end;

procedure TDBConnect.GetProcedureNames(List: TStrings; AQry: TDataSet = nil);
begin
  if IsQryInvalid(AQry) then AQry := DBEngine.GetDataSet;
  if AQry is TADOQuery then
    TADOQuery(AQry).Connection.GetProcedureNames(List);
end;

function TDBConnect.GetValidFieldValue(FieldItem: TFieldItem;
  const sValueDef: string): string;
var
  nPos: Integer;
begin
  if sValueDef <> '' then          // 默认值
    Result := sValueDef
  else
    Result := FieldItem.AsString;

  case FieldItem.DataType of
    ftString, ftMemo, ftFmtMemo{$IFDEF VER180}, ftWideMemo{$ENDIF}, ftWideString, ftOraClob:
    begin
      nPos := fStrUtil.PosFrom('''', Result, 1);
      while nPos <> 0 do
      begin
        if Copy(Result, nPos+1, 1) = '''' then
          nPos := fStrUtil.PosFrom('''', Result, nPos+2)
        else
        begin
          Result := Copy(Result, 1, nPos)+''''+Copy(Result, nPos+1, MaxInt);
          nPos := fStrUtil.PosFrom('''', Result, nPos+2)
        end;
      end;
      Result := Format('''%s''', [Result]);
    end;
    ftDate, ftTime, ftDateTime:
    begin
      // to be override
      // 日期仅仅加上引号,一般化处理.oracle需覆盖此处理,access可不覆盖此处理
      if Result <> '' then
        Result := ''''+Result+''''  //'datevalue(''' + Result + ''')';
      else
        Result := 'Null';
    end;
    ftBlob, ftOraBlob:
    begin
      // oracle需覆盖此处理
      Result := 'Null';
    end;
    ftInteger, ftSmallint, ftWord, ftAutoInc:
    begin      
      if Result = '' then
        Result := 'Null';
    end;
    else begin
      if Result = '' then
        Result := 'Null';
    end;  
  end; // of case
end;  

function TDBConnect.GetValidFieldValue(Field: TField;
  const sValueDef: string): string;
var
  fieldItem: TFieldItem;
begin
  fieldItem := TFieldItem.Create;
  fieldItem.InitByField(Field);
  try
    Result := GetValidFieldValue(fieldItem, sValueDef);
  finally
    fieldItem.Free;
  end;   
end; 

function TDBConnect.GetFieldsKeyValuesStr(Fields: TFieldItemList): string;
var
  i: Integer;
  sCondition: string;
begin
  sCondition := '';
  for i := 0 to Fields.Count - 1 do begin
    if sCondition = '' then
      sCondition := Format('%s=%s',[Fields.Items[i].FieldName,
        GetValidFieldValue(Fields.Items[i])])
    else
      sCondition := sCondition + ','+ Format('%s=%s',[
        Fields.Items[i].FieldName,
        GetValidFieldValue(Fields.Items[i])]);
  end;
  Result := sCondition;
end;

function TDBConnect.GetFieldsValuesStr(Fields: TFieldItemList): string;
var
  i: Integer;
  s: string;
begin  
  for i := 0 to Fields.Count - 1 do begin
    if s = '' then
      s := GetValidFieldValue(Fields.Items[i])
    else
      s := s + ','+ GetValidFieldValue(Fields.Items[i]);
  end;
  Result := s;
end;  

procedure TDBConnect.GetViewNames(List: TStrings; AQry: TDataSet = nil);
begin
//  raise Exception.Create('not support GetViewNames');
end;

function TDBConnect.TableExists(sTableName: string): Boolean;
var
  qry: TDataSet;
  list: TStrings;
  i: Integer;
begin
  Result := False;
  qry := DBEngine.GetNewQuery;
  list := TStringList.Create;
  try
    GetTableNames(list, qry);
    for i := 0 to list.Count - 1 do
      if SameText(sTableName, list[i]) then begin
        Result := True;
        Break;
      end;
  finally
    list.Free;
    qry.Free;
  end;
end;

function TDBConnect.GetTableInfos(AQry: TDataSet): TTableInfoList;
var
  StrList: TStrings;
  i: Integer;
  tableInfo: TTableInfo;
  List: TTableInfoList;
begin
  StrList := TStringList.Create;  
  List := TTableInfoList.Create;
  try
    GetTableNames(StrList, AQry);
    for i := 0 to StrList.Count - 1 do begin
      tableInfo := GetTableInfo(StrList[i]);
      List.Add(tableInfo);
    end;  
  finally
    StrList.Free;
  end;
  Result := List;
end;

procedure TDBConnect.GetTableNames(List: TStrings;
  AQry: TDataSet = nil);
begin
  if IsQryInvalid(AQry) then
    AQry := DBEngine.GetDataSet;
  if AQry is TADOQuery then
    TADOQuery(AQry).Connection.GetTableNames(List, FSystemObject)
  else if AQry is TQuery then
    TQuery(AQry).Database.GetTableNames(List, FSystemObject)
  else if AQry is TSQLQuery then
    TSQLQuery(AQry).SQLConnection.GetTableNames(List, FSystemObject)
  else if AQry is TSQLDataSet then
    TSQLDataSet(AQry).SQLConnection.GetTableNames(List, FSystemObject)
  else if AQry is TSimpleDataSet then
    TSimpleDataSet(AQry).Connection.GetTableNames(List, FSystemObject);
end;  

procedure TDBConnect.GetIndexNames(List: TStrings);
begin

end;

procedure TDBConnect.GetFieldNames(const TableName: string; List: TStrings;
    option: TFieldShowOptions; AQry: TDataSet);
var
  i: Integer;    
  sFmt: string;
  fldList: TFieldItemList;
  sNullable: string;
begin
  if fsoName in option then
    sFmt := '%s';
  if fsoType in option then
    sFmt := sFmt + ' %s';
  if fsoNullable in option then
    sFmt := sFmt + ' %s';
  if (Length(sFmt)>1) and (sFmt[1]=' ') then
    sFmt := Copy(sFmt, 2, MaxInt);
    
  List.BeginUpdate;
  List.Clear;    
  fldList := GetFieldInfos(UpperCase(TableName), AQry);
  try
    for i := 0 to fldList.Count - 1 do begin
      if not fldList.Items[i].Nullable then
        sNullable := 'Not Null'
      else
        sNullable := '';
      List.Add(Format(sFmt, [fldList.Items[i].FieldName,
        fldList.Items[i].DataTypeStr, sNullable]))
    end;
  finally
    List.EndUpdate;
    fldList.Free;
  end;   
end;  

function TDBConnect.GetFieldInfos(const TableName: string;
  AQry: TDataSet): TFieldItemList;   
var
  List: TFieldItemList;
begin
  List := TFieldItemList.Create;
  List.SystemObject := Self.SystemObject;
  List.DB := Self.DBEngine;
  List.InitFieldInfos(tableName);
  Result := List;
end;

procedure TDBConnect.GetTriggerNames(List: TStrings; 
  AQry: TDataSet = nil);
begin
//  raise Exception.Create('not support GetTriggerNames');
end;

function TDBConnect.GetPrimaryField(TableName: string;
   AQry: TDataSet = nil): TFieldItem;
begin
  raise Exception.Create('not support GetPrimaryField');
end;

function TDBConnect.GetTableInfo(sTableName: string): TTableInfo;
var
  table: TTableInfo;
begin
  table := GetNewTableInfo;
  table.SystemObject := Self.SystemObject;
  table.DB := Self.DBEngine;
  table.CheckUniqueFields := False;   //为了不影响速度，不自动检查唯一字段，在需要检查的地方再创建tableinfo对象完成
  if table.InitTableInfo(sTableName) then
    Result := table
  else begin
    table.Free;
    Result := nil;
  end;
end;

function TDBConnect.IsFieldUnique(tableName, fieldName: string;
  AQry: TDataSet): Boolean;
var
  sSql: string;
begin
  Result := False;
  sSql := SQL_COMMON_CHECK_FIELD_UNIQUE;
  if ExecQuery(AQry, Format(sSql, [fieldName, tableName])) then begin
    Result := AQry.Eof;
  end;
end;

procedure TDBConnect.CheckFieldsUnique(tableName: string;
  fields: TFieldItemList);
var
  sSql: string;
  sFieldname: string;
  i: Integer;
  field: TFieldItem;
  qry: TDataSet;
  qry2: TDataSet;
  function InQry2(AFieldname: string): Boolean; begin
    Result := False;
    qry2.First;
    while not qry2.Eof do
    begin
      if SameText(qry2.FieldByName('FieldName').AsString, AFieldname) then
      begin
        Result := True;
        Break;
      end;
      qry2.Next;
    end;  
  end;  
begin
  qry := GetNewQuery;
  qry2:= GetNewQuery;
  try
    for i := 0 to fields.Count - 1 do begin
      // 这里仅仅对int字段进行唯一检查
      if not (fields.Items[i].DataType in
         [ftInteger, ftSmallint, ftWord]) then
        Continue;
      if sSql = '' then
        sSql := Format(SQL_COMMON_CHECK_FIELD_UNIQUE, [
          fields.Items[i].FieldName, tableName])
      else
        sSql := sSql +
          Format(' union all ' + SQL_COMMON_CHECK_FIELD_UNIQUE, [
          fields.Items[i].FieldName, tableName]);
    end;
    if sSql = '' then
      Exit;
    if ExecQuery(qry, 'select distinct * from (' + sSql + ')')
       and ExecQuery(qry2, 'select distinct * from (' + sSql + ') where getcount > 1') then
    begin
      // 只需处理第一个记录
      while not qry.Eof do begin
        sFieldname := qry.FieldByName('FieldName').AsString;
        // findfielditem传入的参数和查询时起的别名有关，这里仍用了原字段名
        field := fields.FindFieldItem(sFieldname);
        if field <> nil then begin
          if not InQry2(sFieldname) then
            field.IsUnique := True;
        end;
        qry.Next;
      end;  
    end;
  finally
    qry.Free;
    qry2.Free;
  end;
end;

procedure TDBConnect.CheckFieldUnique(tableName: string; field: TFieldItem);
var
  fieldList: TFieldItemList;
begin
  fieldList := TFieldItemList.Create();
  try
    fieldList.Add(field);
    CheckFieldsUnique(tableName, fieldList);
    fieldList.Extract(field);
  finally
    fieldList.Free;
  end;   
end;

function TDBConnect.GenTableCreateSQL(sTableName: string;
  slst: TStrings):Boolean;
var
  i: Integer;
  fldList: TFieldItemList;
  tabInfo: TTableInfo;
  PlanarStringList: TPlanarStringList;
  qry: TDataSet;
begin
  fldList := TFieldItemList.Create;
  qry := GetNewQuery;
  try
    tabInfo := GetTableInfo(sTableName);
    fldList := tabInfo.Fields; //GetFieldInfos(sTableName, qry);
    slst.Add(Format('create table %s(', [tabInfo.TableName]));
    PlanarStringList := TPlanarStringList.Create;
    try
      for i := 0 to fldList.Count - 1 do begin
        if i <> fldList.Count - 1 then begin
          PlanarStringList.Strs[i, 0] := '  ' + fldList.Items[i].FieldName + ' ';
          PlanarStringList.Strs[i, 1] := fldList.Items[i].GetDataTypeInSQL + ',';
        end else begin
          PlanarStringList.Strs[i, 0] := '  ' + fldList.Items[i].FieldName + ' ';
          PlanarStringList.Strs[i, 1] := fldList.Items[i].GetDataTypeInSQL;
        end;
      end;
      PlanarStringList.JustifyMode := pjmBothLeft;
      PlanarStringList.FormatItemLengths;
      for i := 0 to PlanarStringList.Count-1 do begin
        slst.Add(PlanarStringList.ItemStr[i]);
      end;  
    finally
      PlanarStringList.Free;
    end;
    // 主键sql
    case FDBType of
    dbtOracle:
    begin
      if fldList.GetPrimaryField <> nil then
      begin
        // 补充上逗号
        slst[slst.Count - 1] := slst[slst.Count - 1] + ',';

        slst.Add(Format('constraint %s primary key (%s)',
          [fldList.GetPrimaryField.ConstraintName, fldList.GetPrimaryField.FieldName]));
      end;
    end;  
    else
    end;
    
    slst.Add(');');
    Result := True;
  finally
    fldList.Free;
    qry.Free;
  end;   
end;

function TDBConnect.ExportTableData(tableName, sExportFileName: string;
  option: TDBExportOption): Boolean;
var
  qry: TDataSet;
  exportFile: TTextFileWriter;
begin
  Result := False;
  qry := GetNewQuery;
  exportFile := TTextFileWriter.Create(sExportFileName);
  try
    // process options
    if option.DisableTriggers and (DBEngine.GetDBType = dbtOracle) then
      exportFile.WriteLine(Format('Alter table %s disable all triggers;',[
        tableName]));
    if option.DeleteTables then
      exportFile.WriteLine(Format('Delete from %s;',[tableName]));
    if option.ExportData then begin
      if ExecQuery(qry, Format('select * from %s', [tableName])) then begin
        Result := ExportQuery(qry, exportFile, option.ClobInFile,
          option.RemoveBreakLine);
      end;
    end;   
    // pose process options
    if option.DisableTriggers and (DBEngine.GetDBType = dbtOracle) then
      exportFile.WriteLine(Format('Alter table %s enable all triggers;',[
        tableName]));
  finally
    qry.Free;
    exportFile.Free;
  end;   
end;

function TDBConnect.GetExportValidFieldValue(Field: TField;
  const sValueDef: string): string;
var
  s: string;
begin
  if IsClobField(Field) or IsBlobField(Field) then
    s := 'null'
  else begin
    s := Self.GetValidFieldValue(Field, sValueDef);
  end;
  Result := s;
end;

procedure TDBConnect.ProcessLobExport(qry: TDataSet; tableInfo: TTableInfo;
    sInsertSqlPrefix: string; slstPlanar: TPlanarStringList;
    sExportFileName: string; ClobInFile: Boolean; RemoveBreakLine: Boolean);
var
  json: TJSON;
  blobFields, clobFields: TFieldItemList;
  n: Integer;   
  slstItem: TStringList;
  sFieldValue: string;
  fldItem: TFieldItem;
begin
  blobFields := TFieldItemList.Create();
  clobFields := TFieldItemList.Create();
  json := TJSON.Create;
  slstItem := TStringList.Create;
  try
    slstItem.Add(sInsertSqlPrefix);
    tableInfo.Fields.ReadValuesFromFields(qry.Fields);
    for n := 0 to qry.FieldCount - 1 do begin
      if FStopExec then
        Exit;       
      // blob 处理
      if IsBlobField(qry.Fields[n]) and not qry.Fields[n].IsNull then begin
        fldItem := tableInfo.GetNewFieldItem;
        fldItem.InitByField(qry.Fields[n]);
        blobFields.Add(fldItem);
        sFieldValue := GetExportValidFieldValue(qry.Fields[n], '');
      end
      // clob 处理
      else if IsClobField(qry.Fields[n]) then begin
        if ClobInFile then
        begin
          sFieldValue := GetExportValidFieldValue(qry.Fields[n], '');
          if not qry.Fields[n].IsNull
             and (qry.Fields[n].AsString <> '') then
          begin
            fldItem := tableInfo.GetNewFieldItem;
            fldItem.InitByField(qry.Fields[n]);
            clobFields.Add(fldItem);
          end;
        end
        else     
          sFieldValue := GetValidFieldValue(qry.Fields[n], '');
      end
      else
        sFieldValue := GetExportValidFieldValue(qry.Fields[n], '');

      if RemoveBreakLine then
        sFieldValue := StringReplace(sFieldValue, #$D#$A, ' ', [rfReplaceAll]);
      if n = qry.FieldCount - 1 then
        slstItem.Add(sFieldValue)
      else
        slstItem.Add(sFieldValue + ', ');
    end;
    slstItem.Add(');');
    slstPlanar.AddItem(slstItem);
    if blobFields.Count > 0 then begin
      if SaveBlobs(blobFields, tableInfo, Format('%s%s%s', [
        ExtractFilePath(sExportFileName), tableInfo.TableName, C_sBlob_Dir_Name]),
        json) then begin
        slstPlanar.Add('--'+BuildCommandPrefix(C_sDBCommand_SQL)
          + ' ' + C_sHeadFlag_Blob + ' ' + json.ToString);
        // 这里给slstPlanar的行不需要进行二维排列,需要包含在排除行中
        slstPlanar.AddExcludeRow(slstPlanar.Count-1);
      end;
    end;    
    if clobFields.Count > 0 then begin
      if SaveClobs(clobFields, tableInfo, Format('%s%s%s', [
        ExtractFilePath(sExportFileName), tableInfo.TableName, C_sClob_Dir_Name]),
        json) then
      begin
        slstPlanar.Add('--'+BuildCommandPrefix(C_sDBCommand_SQL)
          + ' ' + C_sHeadFlag_Clob + ' ' + json.ToString);
        // 这里给slstPlanar的行不需要进行二维排列,需要包含在排除行中
        slstPlanar.AddExcludeRow(slstPlanar.Count-1);
      end;
    end;
  finally 
    blobFields.Free;
    clobFields.Free;
    json.Free;
  end;
end;    

function TDBConnect.GetSqlFromQuery(ds: TDataSet; var sql: string): Boolean;
begin
  Result := True;
  if ds is TQuery then
    sql := StringReplace((ds as TQuery).SQL.Text, #$D#$A, '', [rfReplaceAll])
  else if ds is TADOQuery then
    sql := StringReplace((ds as TADOQuery).SQL.Text, #$D#$A, '', [rfReplaceAll])
  else if ds is TSQLQuery then
    sql := StringReplace((ds as TSQLQuery).SQL.Text, #$D#$A, '', [rfReplaceAll])
  else if ds is TSQLDataSet then
    sql := StringReplace((ds as TSQLDataSet).CommandText,
      #$D#$A, '', [rfReplaceAll])
  else if ds is TSimpleDataSet then
    sql := StringReplace((ds as TSimpleDataSet).DataSet.CommandText,
      #$D#$A, '', [rfReplaceAll])
  else
    Result := False; 
end;    

function TDBConnect.ExportQuery(AQry: TDataSet;
  exportFile: TTextFileWriter; ClobInFile: Boolean;
  RemoveBreakLine: Boolean): Boolean;
var
  sSql: string;
  sTableName: string;
  tableInfo: TTableInfo;
  slstPlanar: TPlanarStringList;
  slstItem: TStringList;
  sFieldValue: string;
  sInsertFields, sInsertSql: string;
  i: Integer;
  bHasBlob: Boolean;
  bHasClob: Boolean;
begin
  Result := False;
  sInsertFields := '';
  bHasBlob := False;
  bHasClob := False;
  FStopExec := False;
  if not GetSqlFromQuery(AQry, sSql) then
    raise Exception.Create('不支持的Query类型');
  slstPlanar := TPlanarStringList.Create;
  try
    sTableName := TSqlUtils.GetTableNameBySQL(sSql);  
    tableInfo := GetTableInfo(sTableName);
    if tableInfo = nil then begin    // 获取tableinfo失败
      AddErrorLog(Format('获取表%s的信息失败！', [sTableName]));
      Exit;
    end;
    // get fieldnames
    for i := 0 to AQry.FieldCount - 1 do begin
      if FStopExec then
        Exit;
      if sInsertFields = '' then
        sInsertFields := AQry.Fields[i].DisplayName
      else
        sInsertFields := sInsertFields + ', ' + AQry.Fields[i].DisplayName;
      if IsBlobField(AQry.Fields[i]) then
        bHasBlob := True;
      if IsClobField(AQry.Fields[i]) then
        bHasClob := True;
    end;
    sInsertSql := Format('Insert into %s(%s) Values(', [sTableName, sInsertFields]);
    // get fieldvalues
    AQry.DisableControls;
    try
      AQry.First;
      while not AQry.Eof do begin
        if FStopExec then
          Exit;
        if (not bHasBlob) and (not bHasClob) then begin
          slstItem := TStringList.Create;
          slstItem.Add(sInsertSql);
          for i := 0 to AQry.FieldCount - 1 do begin
            sFieldValue := GetExportValidFieldValue(AQry.Fields[i], '');
            if RemoveBreakLine then
              sFieldValue := StringReplace(sFieldValue, #$D#$A, ' ', [rfReplaceAll]);
            if i = AQry.FieldCount - 1 then
              slstItem.Add(sFieldValue)
            else
              slstItem.Add(sFieldValue + ', ');
          end;
          slstItem.Add(');');
          slstPlanar.AddItem(slstItem);
        end else begin
          ProcessLobExport(AQry, tableInfo, sInsertSql, slstPlanar,
            exportFile.GetFileName, ClobInFile, RemoveBreakLine);
        end;
        AQry.Next;
      end;
      // format and write to file
      slstPlanar.JustifyMode := pjmBothRight;
      slstPlanar.FormatItemLengths;
      for i := 0 to slstPlanar.Count - 1 do
        exportFile.WriteLine(slstPlanar.ItemStr[i]);
      // commit at tail
      if DBEngine.GetDBType = dbtOracle then
        exportFile.WriteLine('commit;');
      Result := True;
    finally
      AQry.EnableControls;
    end;
  finally
    slstPlanar.Free;
  end;
end;

//function TDBConnect.ExportQuery(AQry: TDataSet; sExportFileName: string;
//  option: TDBExportOption): Boolean;
//var
//  resultFile: TTextFileWriter;
//begin
//  resultFile := TTextFileWriter.Create(sExportFileName);
//  try
//    Result := ExportQuery(AQry, resultFile, option.ClobInFile,
//      option.RemoveBreakLine);
//  finally
//    resultFile.Free;
//  end;
//end;

function TDBConnect.ExportQuery(AQry: TDataSet; sExportFileName: string): Boolean;
var
  defOption : TDBExportOption;
  resultFile: TTextFileWriter;
begin
  defOption := TDBExportOption.Create;
  resultFile := TTextFileWriter.Create(sExportFileName);
  try
    Result := ExportQuery(AQry, resultFile, defOption.ClobInFile,
      defOption.RemoveBreakLine);
  finally
    defOption.Free; 
    resultFile.Free;
  end;
end;

procedure TDBConnect.ClearLog;
begin
  FLog.Clear;
end;

procedure TDBConnect.AddLog(sLog: string);
begin
  FLog.Add(sLog);
  if Assigned(FOnLog) then
    FOnLog(sLog);
end;  

procedure TDBConnect.AddErrorLog(sLog: string);
begin
  if (LowerCase(TSqlUtils.GetSqlCommand(FLastSQL)) = 'drop')
     and FDropNoError then
    // 此时不产生错误日志
  else
    AddLog('Error: '+ sLog);
end;

procedure TDBConnect.AddInfoLog(sLog: string);
begin
  AddLog('Info: ' +sLog);
end;

function TDBConnect.GetLastElapsedMilis: Double;
begin
  Result := FLastElapsedMillis;
end;

function TDBConnect.GetLastError: string;
begin
  if FLog.Count = 0 then
    Result := ''
  else
    Result := FLog[FLog.Count - 1];
end;

function TDBConnect.GetConnection: TCustomConnection;
begin      
  if not Assigned(FDBEngine) then
    Result := nil
  else
    Result := DBEngine.GetConnection;
end;

function TDBConnect.GetConnected: Boolean;
begin
  if not Assigned(FDBEngine) then
    Result := False
  else
    Result := DBEngine.IsConnected;
end;

procedure TDBConnect.SetConnected(value: Boolean);
begin
  if Assigned(DBEngine) then
    DBEngine.SetConnected(value);
end;

function TDBConnect.GetDataSet: TDataSet;
begin  
  if not Assigned(FDBEngine) then
    Result := nil
  else
    Result := DBEngine.GetDataSet;
end;  

function TDBConnect.GetLastQry: TDataSet;
begin
  Result := FLastQry;
end;  

function TDBConnect.GetLastOperSucc: Boolean;
begin
  Result := FLastOperSucc;
end;

function TDBConnect.GetLastSQL: string;
begin
  Result := FLastSQL;
end;

function TDBConnect.GetLastSQLType: TDBConnectSQLType;
begin
  Result := FLastSQLType;
end;

function TDBConnect.GetMaxRecords: Integer;
begin
  Result := FMaxRecords;
end;  

function TDBConnect.GetLog: TStrings;
begin
  Result := FLog;
end; 

function TDBConnect.GetPageIndex: Integer;
begin
  Result := FMaxRecords;
end;

procedure TDBConnect.SetMaxRecords(value: Integer);
begin
  FMaxRecords := value;
  if Assigned(DBEngine) then begin
    DBEngine.SetMaxRecords(value);
    if Assigned(FLastQry) then
      DBEngine.SetDataSetMaxRecords(FLastQry, value);
  end;
end;

function TDBConnect.GetSystemObject: Boolean;
begin
  Result := FSystemObject;
end;

procedure TDBConnect.SetSystemObject(value: Boolean);
begin
  FSystemObject := value;
end; 

function TDBConnect.GetOnConnected: TNotifyEvent;
begin
  Result := FOnConnected;
end;

function TDBConnect.GetOnDisConnected: TNotifyEvent;
begin
  Result := FOnDisConnected;
end;

function TDBConnect.GetOnLog: TDBLogEvent;
begin
  Result := FOnLog;
end;

procedure TDBConnect.SetOnLog(value: TDBLogEvent);
begin
  FOnLog := value;
end;

procedure TDBConnect.SetOnConnected(value: TNotifyEvent);
begin
  FOnConnected := value;
end;

procedure TDBConnect.SetOnDisConnected(value: TNotifyEvent);
begin
  FOnDisConnected := value;
end;   

function TDBConnect.GetOnExecLineChange: TExecLineChangeEvent;
begin
  Result := FOnExecLineChange; 
end;

procedure TDBConnect.SetOnExecLineChange(value: TExecLineChangeEvent);
begin
  FOnExecLineChange := value;
end;   

function TDBConnect.GetOnExecuted: TNotifyEvent;
begin
  Result := FOnExecuted;
end;

procedure TDBConnect.SetOnExecuted(value: TNotifyEvent);
begin
  FOnExecuted := value;
end;

procedure TDBConnect.SetPageIndex(value: Integer);
begin
  if value > 0 then begin
    if (FPageIndex <> value)
       and (FLastQry <> nil) and (FLastSQL <> '')  then begin
      FPageIndex := value;
      ExecQuery(FLastQry, GetMaxRecordSQLCondition(FMaxRecords*FPageIndex,
        FLastSQL));
    end;
  end;
end; 

function TDBConnect.GetDropNoError: Boolean;
begin
  Result := FDropNoError;
end;

procedure TDBConnect.SetDropNoError(value: Boolean);
begin
  FDropNoError := value;
end;  

function TDBConnect.GetVariables: TStringList;
begin
  Result := FVariables;
end;

procedure TDBConnect.SetVariables(value: TStringList);
begin
  FVariables := value;
end;  

function TDBConnect.GetDBType: TDBType;
begin
  Result := FDBType;
end;

function TDBConnect.GetDBEngine: IDBEngine;
begin
  Result := FDBEngine;
end;   

function TDBConnect.GetNewDBEngine(Adbet: TDBEngineType): IDBEngine;
var
  engine: IDBEngine;
begin
  engine := TDBEngineFactory.GetNewDBEngine(Adbet);
  engine.SetMaxRecords(Self.MaxRecords);
  Result := engine;
end;

procedure TDBConnect.CheckDBEngine(Adbet: TDBEngineType);
begin
  if Assigned(FDBEngine) and (DBEngine.GetDBEngineType <> Adbet) then
    FDBEngine := nil;
  if not Assigned(FDBEngine) then
    FDBEngine := GetNewDBEngine(Adbet);
  FDBEngine.SetOnConnected(FOnConnected);
  FDBEngine.SetOnDisConnected(FOnDisConnected);
end;

function TDBConnect.CloseDB: Boolean;
begin
  Result := DBEngine.CloseDB; 
  FLastSQL := '';
end;

function TDBConnect.CloseQuery: Boolean;
begin
  Result := DBEngine.CloseQuery;
end;  

function TDBConnect.ExecQuery(AQ: TDataSet; sSql: string): Boolean;
begin
  Result := DBEngine.ExecQuery(AQ, sSql); 
  FLastSQLType := dbcstQuery; 
  FLastSQL := sSql;
  if not Result then    
    AddErrorLog(DBEngine.GetLastError);
end;

function TDBConnect.ExecUpdate(AQ: TDataSet; sSql: string): Integer;
begin
  Result := DBEngine.ExecUpdate(AQ, sSql); 
  FLastSQLType := dbcstUpdate; 
  FLastSQL := sSql;
  if Result = C_nTDBE_ERROR_EXECFAIL then
    AddErrorLog(DBEngine.GetLastError);
end;

function TDBConnect.ExecQueryWithParams(AQ: TDataSet;
  sSqlWithParams: string; aryParams: array of Variant): Boolean;
begin
  Result := DBEngine.ExecQueryWithParams(AQ, sSqlWithParams, aryParams);
  FLastSQLType := dbcstQuery;
  FLastSQL := sSqlWithParams;
  if not Result then    
    AddErrorLog(DBEngine.GetLastError);
end;

function TDBConnect.ExecUpdateWithParams(AQ: TDataSet;
  sSqlWithParams: string; aryParams: array of Variant): Integer;
begin
  Result := DBEngine.ExecUpdateWithParams(AQ, sSqlWithParams, aryParams); 
  FLastSQLType := dbcstUpdate;
  FLastSQL := sSqlWithParams;
  if Result = C_nTDBE_ERROR_EXECFAIL then
    AddErrorLog(DBEngine.GetLastError);
end;

procedure TDBConnect.SetVariableValue(var s: string);
var
  i: Integer;
  sName, sValue: string;
  varList: TStrings;
  function getValue(varname: string): string; begin
    // TODO: 仅为临时写法，有空改为SysVaList保存变量名与获取方法的方式
    // 系统变量
    if SameText(varname, C_sDBCmdParam_Var_Start
        + C_sDBCmdParam_Var_SysVar_DateTime14
        + C_sDBCmdParam_Var_End) then
      Result := FormatDateTime('yyyymmddhhnnss', Now)
    else if i < Variables.Count then
      Result := Variables.Values[Variables.Names[i]]
    else
      Result := varname;
  end;  
begin
  varList := fStrUtil.GetSubStrListWithRegx(C_sDBCmdParam_Regx_GetVar,s);
  for i := 0 to varList.Count - 1 do begin
    sName := varList[i];
    sValue := getValue(sName);
    if sName <> sValue then
      s := StringReplace(s, sName, sValue, [rfReplaceAll, rfIgnoreCase]);
  end;  
end;

function TDBConnect.UpdateBlobs(JsonStr: string): Integer;
begin
  Result := 0;
end;

function TDBConnect.UpdateClobs(JsonStr: string): Integer;
begin
  Result := 0;
end;

procedure TDBConnect.GetDBs(List: TStrings);
begin
  
end;

function TDBConnect.GetProcSource(sName: string; list: TStrings): string;
begin
  Result := '';
end;

function TDBConnect.GetTrigSource(sName: string; list: TStrings): string;
begin
  Result := '';
end;

function TDBConnect.GetViewSource(sName: string; list: TStrings): string;
begin
  Result := '';
end;

function TDBConnect.SaveBlobs(blobFields: TFieldItemList;
  tableInfo: TTableInfo; sDir: string; var json: TJSON): Boolean;
var
  sSql: string;
  qry: TDataSet;
  num: Integer;
  jsonField: TJSON;
  i: Integer;
  sBlobFileName: string;  
  blobfield: TBlobField; 
  stream: TMemoryStream;
  sCondition: string;
begin
  Result := False;
  qry := GetNewQuery;
  try
    if tableInfo.PrimaryKeys.Count <> 0 then
      sCondition := GetFieldsKeyValuesStr(tableInfo.PrimaryKeys)
    else begin
      sCondition := GetFieldsKeyValuesStr(tableInfo.GetUniqueFields);
    end;
    if sCondition = '' then begin
      AddErrorLog('缺乏定位Blob字段的主键或唯一键。');
      Exit;
    end;
    sSql := Format('select * from %s where %s', [tableInfo.TableName,
      sCondition]);
    if ExecQuery(qry, sSql) then begin
      num := 1;
      if not qry.Eof then begin
        json.AddField(C_JSON_KEY_SQL, sSql);
        // 每个字段增加一个json子对象
        for i := 0 to blobFields.Count - 1 do begin
          blobfield := qry.FieldByName(blobFields.Items[i].FieldName) as TBlobField;
          stream := TMemoryStream.Create;
          try
            blobfield.SaveToStream(stream);
            sBlobFileName := IncludeTrailingPathDelimiter(sDir)
              + blobFields.Items[i].FieldName
              + '('+sCondition+')';
            if not DirectoryExists(sDir) then
              ForceDirectories(sDir);
            stream.SaveToFile(sBlobFileName);
          finally
            stream.Free;
          end;

          jsonField := TJSON.Create;
          jsonField.AddField(C_JSON_KEY_NAME,
            blobFields.Items[i].FieldName);
          jsonField.AddField(C_JSON_KEY_PATH,
            tableInfo.TableName + C_sBlob_Dir_Name + ExtractFileName(sBlobFileName));
                                                                                         
          // 是field+序号数字的形式,没添加一个序号增1
          json.AddFieldObject(C_JSON_KEY_FIELD
            + IntToStr(num), jsonField);
          Inc(num);
        end;
        Result := True;
      end;
    end;
  finally
    qry.Free;
  end;
end;

function TDBConnect.SaveClobs(clobFields: TFieldItemList;
  tableInfo: TTableInfo; sDir: string; var json: TJSON): Boolean;
var
  sSql: string;
  qry: TDataSet;
  num: Integer;
  jsonField: TJSON;
  i: Integer;
  sBlobFileName: string;  
  clobfield: TBlobField;
  slst: TStrings;
  sCondition: string;
begin
  Result := False;
  qry := GetNewQuery;
  try
    if tableInfo.PrimaryKeys.Count <> 0 then
      sCondition := GetFieldsKeyValuesStr(tableInfo.PrimaryKeys)
    else begin
      sCondition := GetFieldsKeyValuesStr(tableInfo.GetUniqueFields);
    end;
    if sCondition = '' then begin
      AddErrorLog('缺乏定位Clob字段的主键或唯一键。');
      Exit;
    end;
    sSql := Format('select * from %s where %s', [tableInfo.TableName,
      sCondition]);
    if ExecQuery(qry, sSql) then begin
      num := 1;
      if not qry.Eof then begin
        json.AddField(C_JSON_KEY_SQL, sSql);
        // 每个字段增加一个json子对象
        for i := 0 to clobFields.Count - 1 do begin
          clobfield := qry.FieldByName(clobFields.Items[i].FieldName) as TBlobField;
          slst := TStringList.Create;
          try
            slst.Text := clobfield.AsString;
            sBlobFileName := IncludeTrailingPathDelimiter(sDir)
              + clobFields.Items[i].FieldName
              + '('+sCondition+')';
            if not DirectoryExists(sDir) then
              ForceDirectories(sDir);
            slst.SaveToFile(sBlobFileName);
          finally
            slst.Free;
          end;

          jsonField := TJSON.Create;
          jsonField.AddField(C_JSON_KEY_NAME, clobFields.Items[i].FieldName);
          jsonField.AddField(C_JSON_KEY_PATH, tableInfo.TableName
            + C_sClob_Dir_Name + ExtractFileName(sBlobFileName));
                                                                                         
          // 是field+序号数字的形式,没添加一个序号增1
          json.AddFieldObject(C_JSON_KEY_FIELD + IntToStr(num), jsonField);
          Inc(num);
        end;
        Result := True;
      end;
    end;
  finally
    qry.Free;
  end;
end;

procedure TDBConnect.GetDBLinkNames(List: TStrings);
begin     
  // implement in oracle dbconnect
end;

procedure TDBConnect.GetFormNames(List: TStrings; AQry: TDataSet);
begin
  // implement in access dbconnect
end;

procedure TDBConnect.GetMacroNames(List: TStrings; AQry: TDataSet);
begin
  // implement in access dbconnect
end;

procedure TDBConnect.GetMoudleNames(List: TStrings; AQry: TDataSet);
begin
  // implement in access dbconnect
end;

procedure TDBConnect.GetPageNames(List: TStrings; AQry: TDataSet);
begin
  // implement in access dbconnect
end;

procedure TDBConnect.GetReportNames(List: TStrings; AQry: TDataSet);
begin
  // implement in access dbconnect
end;

procedure TDBConnect.GetSynonymNames(List: TStrings);
begin
  // implement in oracle dbconnect
end;

procedure TDBConnect.GetUsers(List: TStrings);
begin
  // implement in oracle dbconnect
end;

initialization

finalization

end.
