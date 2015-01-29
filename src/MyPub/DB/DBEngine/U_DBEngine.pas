{**
 * <p>Title: TDBEngine  </p>
 * <p>Copyright: Copyright (c) 2006/7/24</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 数据库引擎类
 * <p>Description: 封装数据库连接, 执行sql，为了消除数据库的不同</p>
 *}
unit U_DBEngine;

interface

{$M+}

uses
  ADODB, DB, Classes, Variants, DBTables, U_DBEngineInterface;

type   

{ TDBEngine 该类有抽象方法，需要用实际的DBEngine类来实例化 }
  TDBEngine = class(TInterfacedObject, IDBEngine)
  private
    FLastExecTime  : TDateTime;
    FDBKeepInterval: Integer;
    FLastElapsedMillis : Double;

  protected               
    FDBType    : TDBType;                    // 数据库类型
    FDataSource: string;
    FUserName  : string;
    FPassword  : string;
    FLastError : string;

    FConnection: TCustomConnection;
    FDataSet   : TDataSet;
    FStoredProc: TStoredProc; 
    FDBEngineType  : TDBEngineType; 
    FMaxRecords    : Integer;
    FPageSize      : Integer;
    FOnConnected   : TNotifyEvent;
    FOnDisConnected: TNotifyEvent;

    function GetConnection: TCustomConnection;virtual;
    function GetConnected: Boolean;virtual;
    procedure SetConnected(value: Boolean);virtual;
    function GetDataSet: TDataSet;virtual;

    function GetMaxRecords: Integer;virtual;
    procedure SetMaxRecords(value: Integer);virtual;
    function GetDBHandle: Pointer;virtual;

    //  ExtractStrings 不添加空字符串到列表，实际有为空的可能  所以需要自己写分割字符串的方法    
    function Split(sText: String; sSeparator: String; ssParts: TStrings):TStrings;
    // abstract
    function doCloseConnection: Boolean;virtual;abstract;
    procedure doCloseQuery;virtual;abstract;
    function OpenDBWithDBEngine(sDataSource: string;sUser: string;sPwd: string;
      dbt: TDBType): Boolean;virtual;abstract;
    
    function doExecOneSql(Ads: TDataSet; Asql: string; bOpen: Boolean): Integer;virtual;
    procedure doExecQuery(Ads: TDataSet; sSql: string);virtual;abstract;
    procedure doExecQueryWithParams(Ads: TDataSet; sSql: string;aryParams: array of Variant);virtual;abstract;
    function doExecUpdate(Ads: TDataSet; sSql: string):Integer;virtual;abstract;
    function doExecUpdateWithParams(Ads: TDataSet; sSql: string;aryParams: array of Variant):Integer;virtual;abstract;

  public
    property DBType: TDBType read FDBType;
    property DBEngineType: TDBEngineType read FDBEngineType; 
    property DBHandle: Pointer read GetDBHandle;

    property Connection: TCustomConnection read GetConnection;
    property DataSet: TDataSet read GetDataSet;

    property Connected: Boolean read GetConnected write SetConnected;
    property LastError: string read FLastError;
    property MaxRecords: Integer read GetMaxRecords write SetMaxRecords;

    property LastExecTime: TDateTime read FLastExecTime;

    property DBKeepInterval: Integer read FDBKeepInterval write FDBKeepInterval;
    property DataSource: string read FDataSource;
    property UserName: string read FUserName;
    property Password: string read FPassword;
  public
    constructor Create; virtual;
    destructor Destroy; override;

    function GetLastError: string;
    function GetLastElapsedMilis: Double;
    function GetDBType: TDBType;
    function GetDBEngineType: TDBEngineType; 
    function GetOnConnected(): TNotifyEvent;
    function GetOnDisConnected(): TNotifyEvent;
    procedure SetOnConnected(event: TNotifyEvent); 
    procedure SetOnDisConnected(event: TNotifyEvent);
    function IsConnected: Boolean; virtual;
    function GetDataSource: string;  
    function GetUserName: string;
    function GetPassword: string;

    { 数据库操作 }
    // 参照 TADODBEngine.BuildConnectionString  和 TBDEDBEngine.OpenDBWithDBEngine
    function OpenDB(sDataSource: string;sUser: string;sPwd: string;
      dbt: TDBType): Boolean;overload;virtual;
    function OpenDBWithIni(AIniFile, ADBSection: string):Boolean;overload;virtual;
    function ReOpenDB: Boolean;
    function CloseDB: Boolean; 
    
    function CloseQuery: Boolean;
    function GetNewQuery: TDataSet;virtual;

    { 执行sql的方法 }
    function ExecQuery(sSql: string):Boolean; overload;
    function ExecQuery(AQ: TDataSet; sSql: string):Boolean; overload;virtual;
    function ExecQueryWithParams(sSqlWithParams: string; Params: array of Variant):Boolean; overload;
    function ExecQueryWithParams(AQ: TDataSet; sSqlWithParams: string;
        aryParams: array of Variant):Boolean;overload;virtual;

    function ExecUpdate(sSql: string): Integer; overload;
    function ExecUpdate(AQ: TDataSet; sSql: string): Integer; overload;virtual;
    function ExecUpdateWithParams(sSqlWithParams: string; Params: array of Variant): Integer; overload;
    function ExecUpdateWithParams(AQ: TDataSet; sSqlWithParams: string;
        aryParams: array of Variant): Integer;overload;virtual;

    function ExecProc(sProcName: String; params: array of Variant):Boolean;overload;
    function ExecProc(sProcName: String; keyvalues: TStrings):Boolean;overload;

    { 事务 }
    function BeginTrans: Integer; virtual;
    procedure CommitTrans;virtual;
    procedure RollbackTrans;virtual;

    function KeepDBOpen: Boolean;

    function IsEngineInstalled:Boolean;virtual;
    procedure SetDataSetMaxRecords(Ads: TDataSet; nMaxRecords: Integer);virtual;
    
  published
    property OnConnected: TNotifyEvent read FOnConnected write FOnConnected;
    property OnDisConnected: TNotifyEvent read FOnDisConnected write FOnDisConnected;
  end;

implementation  

uses
  SysUtils, Forms;

{ function }

function GetMillisFromDateTime(dt: TDateTime): Double;
var
  fractional: Double;
  fsResult: Double;
begin
  fractional := dt - Trunc(dt);
  fsResult := fractional * 24 * 3600;  // 现在单位是秒
  Result := fsResult*1000;
end;

function FindValueInList(List: TStrings; Key, delimiter: string): string;
var
  nIndex: Integer;
  nPosItem: Integer;
begin
  Result := '';
  for nIndex := 0 to List.Count - 1 do
  begin
    nPosItem := Pos(LowerCase(Key), LowerCase(List[nIndex]));
    if nPosItem > 0 then
    begin
      nPosItem := Pos(delimiter, List[nIndex]);
      if nPosItem > 0 then
        Result := Trim(Copy(List[nIndex], nPosItem+1, Length(List[nIndex])));
      Break;
    end;
  end;
end;   

{ TDBEngine }

//function TDBEngine.GetLastError: string;
//begin
//  if FErrors.Count = 0 then
//    Result := ''
//  else
//    Result := FErrors[FErrors.Count-1];
//end;

function TDBEngine.GetConnection: TCustomConnection;
begin
  Result := FConnection;
end;    

function TDBEngine.GetConnected: Boolean;
begin
  Result := FConnection.Connected;
end;        

procedure TDBEngine.SetConnected(value: Boolean);
begin
  if not value then
    FConnection.Connected := value
  else
    OpenDB(Self.DataSource, Self.UserName, Self.Password, Self.FDBType);
end;

function TDBEngine.GetDataSet: TDataSet;
begin
  Result := FDataSet;
end;

function TDBEngine.Split(sText, sSeparator: String;
  ssParts: TStrings): TStrings;
var
  nPos: Integer;
  sPart, sRemain: String;
begin
  Result := ssParts;
  if Result = nil then
    Result := TStringList.Create;
  Result.Clear;
  sRemain:= sText;
  nPos:= Pos(sSeparator, sRemain);

  while nPos > 0 do
  begin
    sPart:= Copy(sRemain, 1, nPos- 1);
    Result.Add(sPart);
    sRemain:= Copy(sRemain, nPos+ Length(sSeparator), Length(sRemain)- nPos+ 1);
    nPos:= Pos(sSeparator, sRemain);
  end;

  if sRemain <> '' then
    Result.Add(sRemain);
end;

function TDBEngine.OpenDB(sDataSource, sUser, sPwd: string;
  dbt: TDBType): Boolean;
begin
  FDBType := dbt;
  FDataSource := sDataSource;
  FUserName := sUser;
  FPassword := sPwd;
  try
    FLastError := '';
    FLastExecTime := Now;
    Result := OpenDBWithDBEngine(sDataSource, sUser, sPwd, FDBType);
    if Result then
    begin
      if Assigned(FOnConnected) then
        FOnConnected(Self);
    end;
    FLastElapsedMillis := GetMillisFromDateTime(Now - FLastExecTime);
  except
    on ex: Exception do
    begin    
      Result := False;
      FLastError := ex.Message;
//      raise Exception.Create('(TDBEngine.OpenDB) ' + ex.Message);
    end;
  end;
end; 

function TDBEngine.OpenDBWithIni(AIniFile, ADBSection: string): Boolean;
begin
  Result := False;
end;

function TDBEngine.ReOpenDB: Boolean;
begin
  if Connected then
    CloseDB;
  if (FDataSource = '') or (FUserName = '' ) then
  begin
    Result := False;
    Exit;
  end;  
  Result := OpenDB(FDataSource, FUserName, FPassword, FDBType);
end;

//procedure TDBEngine.ClearLog;
//begin
//  FErrors.Clear;
//end;

//procedure TDBEngine.AddLog(sLog: string);
//begin
//  FErrors.Add(sLog);
//  if Assigned(FOnLog) then
//    FOnLog(sLog);
//end;

//procedure TDBEngine.AddInfoLog(sLog: string);
//begin
//  FErrors.Add(sLog);
//  if Assigned(FOnLog) then
//    FOnLog(sLog);
//end;

function TDBEngine.CloseDB: Boolean;
var
  bOldConnected: Boolean;
begin
  try     
    FLastError := '';
    bOldConnected := Connected;
    Result := doCloseConnection;   
    if bOldConnected and Result then
    begin
      if Assigned(FOnDisConnected) then
        FOnDisConnected(Self);
    end;
  except
    on ex: Exception do
    begin     
      Result := False;
      FLastError := ex.Message;
//      raise Exception.Create('(TDBEngine.CloseConnection) ' + ex.Message);
    end;
  end;
end;

function TDBEngine.CloseQuery: Boolean;
begin
  try
    FLastError := '';
//    ClearLog;
    Result := True;
    doCloseQuery;
  except
    on ex: Exception do
    begin
      Result := False;
      FLastError := ex.Message;
//      raise Exception.Create('(TDBEngine.CloseQuery) ' + ex.Message);
    end;
  end;
end;      

function TDBEngine.ExecProc(sProcName: String; params: array of Variant): Boolean;
var
  i: Integer;
begin
  FStoredProc.Close; 
  FLastError := '';
  try
    FStoredProc.StoredProcName := sProcName;
    for i := 0 to High(params) do
    begin
      FStoredProc.Params[i].AsString := params[i];
    end;
    FStoredProc.ExecProc;
    Result := True;
  except
    on ex: Exception do
    begin       
      FLastError := ex.Message;
      Result := False;
    end;
  end;
end;

function TDBEngine.ExecProc(sProcName: String;
  keyvalues: TStrings): Boolean;
var
  i: Integer;
begin
  FStoredProc.Close;
  FLastError := '';
  try
    FStoredProc.StoredProcName := sProcName;
    if Assigned(keyvalues) then
    begin
      for i := 0 to keyvalues.Count - 1 do
      begin
        FStoredProc.ParamByName(keyvalues.Names[i]).AsString := keyvalues
          .Values[keyvalues.Names[i]];
      end;
    end;
    FStoredProc.ExecProc;
    Result := True;
  except
    on ex: Exception do
    begin                   
      FLastError := ex.Message;
      Result := False;
    end;
  end;
end;

function TDBEngine.ExecQuery(AQ: TDataSet; sSql: string): Boolean;
begin
  try                
    FLastError := '';
    FLastExecTime := Now;
//    ClearLog;
    doExecQuery(AQ, sSql); 
    FLastElapsedMillis := GetMillisFromDateTime(Now - FLastExecTime);
    Result := True;
  except
    on ex: Exception do
    begin
      Result := False;
      FLastError := ex.Message;
//      raise Exception.Create('(TDBEngine.ExecQuery) ' + ex.Message);
    end;
  end;
end;

function TDBEngine.ExecUpdate(AQ: TDataSet; sSql: string): Integer;
begin
  try
    FLastError := '';
    FLastExecTime := Now;
//    ClearLog;
    Result := doExecUpdate(AQ, sSql);
    FLastElapsedMillis := GetMillisFromDateTime(Now - FLastExecTime);
  except
    on ex: Exception do
    begin
      Result := C_nTDBE_ERROR_EXECFAIL;
      FLastError := ex.Message;
//      raise Exception.Create('(TDBEngine.ExecUpdate) ' + ex.Message);
    end;
  end;
end;

function TDBEngine.ExecQueryWithParams(AQ: TDataSet;
  sSqlWithParams: string; aryParams: array of Variant): Boolean;
begin
  try
    FLastError := '';
    FLastExecTime := Now;
//    ClearLog;
    doExecQueryWithParams(AQ, sSqlWithParams, aryParams);
    FLastElapsedMillis := GetMillisFromDateTime(Now - FLastExecTime);
    Result := True;
  except
    on ex: Exception do
    begin
      Result := False;
      FLastError := ex.Message;
//      raise Exception.Create('(TDBEngine.ExecQueryWithParams) ' + ex.Message);
    end;
  end;
end;

function TDBEngine.ExecUpdateWithParams(AQ: TDataSet;
  sSqlWithParams: string; aryParams: array of Variant): Integer;
begin
  try    
    FLastError := '';
    FLastExecTime := Now;
//    ClearLog;
    Result := doExecUpdateWithParams(AQ, sSqlWithParams, aryParams); 
    FLastElapsedMillis := GetMillisFromDateTime(Now - FLastExecTime);
  except
    on ex: Exception do
    begin     
      Result := C_nTDBE_ERROR_EXECFAIL;
      FLastError := ex.Message;
//      raise Exception.Create('(TDBEngine.ExecUpdateWithParams) ' + ex.Message);
    end;
  end;
end;

constructor TDBEngine.Create;
begin
  FDBType := dbtUnKnown;
  FDBKeepInterval := 60;
  FStoredProc := TStoredProc.Create(nil);
end;

destructor TDBEngine.Destroy;
begin
  FStoredProc.Free;
  inherited;
end;   

function TDBEngine.doExecOneSql(Ads: TDataSet; Asql: string;
  bOpen: Boolean): Integer;
begin
  Result := 0;
  if bOpen then
    doExecQuery(Ads, Asql)
  else
    Result := doExecUpdate(Ads, Asql);
end;

function TDBEngine.ExecQuery(sSql: string):Boolean;
begin
  Result := ExecQuery(FDataSet, sSql);
end;

function TDBEngine.ExecQueryWithParams(sSqlWithParams: string; Params: array of Variant):Boolean;
begin
  Result := ExecQueryWithParams(FDataSet, sSqlWithParams, Params);
end;

function TDBEngine.ExecUpdate(sSql: string): Integer;
begin
  Result := ExecUpdate(FDataSet, sSql);
end;

function TDBEngine.ExecUpdateWithParams(sSqlWithParams: string; Params: array of Variant): Integer;
begin
  Result := ExecUpdateWithParams(FDataSet, sSqlWithParams, Params);
end;

function TDBEngine.KeepDBOpen: Boolean;
begin
  Result := False;
  if FConnection = nil then
    Exit;
  if not Connected then
  begin
    Result := ReOpenDB;
    Exit;
  end
  // 最后操作数据库的时间小于FDBKeepInterval 默认60秒 且Connected=True
  else if (Now-FLastExecTime) * 86400 < FDBKeepInterval then
  begin
    Result := True;
  end
  else
  begin
    case FDBType of
      dbtAccess,dbtAccess2007:
        Result := Connected;
      dbtOracle:
      begin
        Result := ExecQuery('select 1 from dual');
        if not Result then
          Result := ReOpenDB;
      end;
      dbtSyBase: Result := Connected;
      else
        Result := Connected;
    end;
  end;
end;

function TDBEngine.GetDBHandle: Pointer;
begin
  Result := nil;
end;

function TDBEngine.GetLastError: string;
begin
  Result := FLastError;
end;   

function TDBEngine.GetLastElapsedMilis: Double;
begin
  Result := FLastElapsedMillis;
end;

procedure TDBEngine.SetOnConnected(event: TNotifyEvent);
begin
  FOnConnected := event;
end;

procedure TDBEngine.SetOnDisConnected(event: TNotifyEvent);
begin
  FOnDisConnected := event;
end;

function TDBEngine.GetOnConnected: TNotifyEvent;
begin
  Result := FOnConnected;
end;

function TDBEngine.GetOnDisConnected: TNotifyEvent;
begin
  Result := FOnDisConnected;
end; 

function TDBEngine.GetDBType: TDBType;
begin
  Result := FDBType;
end;

function TDBEngine.GetDBEngineType: TDBEngineType;
begin
  Result := FDBEngineType;
end;

function TDBEngine.IsConnected: Boolean;
begin
  Result := FConnection.Connected;
end;

function TDBEngine.GetDataSource: string;
begin
  Result := FDataSource;
end;

function TDBEngine.GetPassword: string;
begin
  Result := FPassword;
end;

function TDBEngine.GetUserName: string;
begin
  Result := FUserName;
end;

function TDBEngine.BeginTrans: Integer;
begin
  Result := 0;
end;

procedure TDBEngine.CommitTrans;
begin

end;

function TDBEngine.GetMaxRecords: Integer;
begin
  Result := FMaxRecords;
end;

function TDBEngine.GetNewQuery: TDataSet;
begin
  Result := nil;
end;

procedure TDBEngine.RollbackTrans;
begin

end;

procedure TDBEngine.SetMaxRecords(value: Integer);
begin
  FMaxRecords := value;
end;

function TDBEngine.IsEngineInstalled: Boolean;
begin
  Result := False;
end;

procedure TDBEngine.SetDataSetMaxRecords(Ads: TDataSet;
  nMaxRecords: Integer);
begin
  FMaxRecords := nMaxRecords;
end;

end.
