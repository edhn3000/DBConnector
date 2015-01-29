{**
 * <p>Title: U_MySqlLibEngine.pas  </p>
 * <p>Copyright: Copyright (c) 2011-9-10</p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: mysqllib引擎 基于mysql.pas再封装 </p>
 *}
unit U_MySqlLibEngine;

interface

uses
  Classes, Variants, SqlExpr, SysUtils, IniFiles, DB,
  U_DBEngine, mysql;

type

{ TMySqlLibEngine }
  TMySqlLibEngine = class(TDBEngine)
  private
    FLibHandle: PMYSQL;
    FCurrDB: string;
    FmySQL_Res: PMYSQL_RES;
    FMYSQL_ROW: PMYSQL_ROW;
//    FDataBase: TDataBase;
//    FBDEQry: TQuery;

    FHost: string;
    FPort: Integer;
    procedure AnalyzeDataSource;
  protected      
    function GetMaxRecords: Integer;override;
    procedure SetMaxRecords(value: Integer);override;
//    function GetConnection: TCustomConnection;override;
    function GetConnected: Boolean;override;
//    function GetDataSet: TDataSet;override;
    function GetDBHandle: Pointer; override;

    function OpenDBWithDBEngine(sDataSource: string;sUser: string;sPwd: string;
      dbt: TDBType): Boolean;override;

    function doCloseConnection: Boolean;override;
    procedure doCloseQuery;override;

    procedure doExecQuery(Ads: TDataSet; sSql: string);override;
//    procedure doExecQueryWithParams(Ads: TDataSet; sSqlWithParams: string;aryParams: array of Variant);override;
//    function doExecUpdate(Ads: TDataSet; sSql: string):Integer;override;
//    function doExecUpdateWithParams(Ads: TDataSet; sSqlWithParams: string;aryParams: array of Variant):Integer;override;   
    
    procedure GetDBs(List: TStrings);
    procedure GetFieldNames(const tableName: string; List: TStrings);
    procedure GetTableNames(List: TStrings);
    procedure SetCurrDB(db: string);
  public
//    property Connection: TDataBase read FDataBase;
//    property DataSet: TQuery read FBDEQry;
  public
    constructor Create; override;
    destructor Destroy; override;
                                   
//    function GetNewQuery: TDataSet;override;
    { 事务 }
    function BeginTrans: Integer;override;
    procedure CommitTrans;override;
    procedure RollbackTrans;override;

    function IsEngineInstalled:Boolean;override;  
//    procedure SetDataSetMaxRecords(Ads: TDataSet; nMaxRecords: Integer);override;
  end;


implementation

{ TMySqlLibEngine }

function TMySqlLibEngine.BeginTrans: Integer;
begin  
  mysql_autocommit(FLibHandle, False);
end;

procedure TMySqlLibEngine.CommitTrans;
begin
  inherited;
  mysql_autocommit(FLibHandle, True);
end; 

procedure TMySqlLibEngine.RollbackTrans;
begin
  inherited;
  mysql_rollback(FLibHandle);
end;

constructor TMySqlLibEngine.Create;
begin
  inherited;
end;

destructor TMySqlLibEngine.Destroy;
begin

  inherited;
end;

function TMySqlLibEngine.doCloseConnection: Boolean;
begin               
  if FLibHandle <> nil then
  begin
    mysql_close(FLibHandle);
    FLibHandle := nil;
  end;
end;

procedure TMySqlLibEngine.doCloseQuery;
begin
  if FmySQL_Res <> nil then
  begin
    mysql_free_result(FmySQL_Res);
    FmySQL_Res := nil;
  end;
end;

function TMySqlLibEngine.GetConnected: Boolean;
begin
  Result := FLibHandle <> nil;
end;

function TMySqlLibEngine.GetMaxRecords: Integer;
begin

end;

function TMySqlLibEngine.IsEngineInstalled: Boolean;
begin

end;

procedure TMySqlLibEngine.AnalyzeDataSource;
var
  nPos: Integer;
begin
  nPos := Pos(':',FDataSource);
  if nPos > 0 then
  begin
    FHost := Copy(FDataSource, 1, nPos-1);
    FPort := StrToIntDef(Copy(FDataSource, nPos+1, MaxInt), 0);
  end
  else
  begin
    FHost := FDataSource;
    FPort := 0;
  end;    
end;

function TMySqlLibEngine.OpenDBWithDBEngine(sDataSource, sUser,
  sPwd: string; dbt: TDBType): Boolean;
begin
  FLibHandle := mysql_init(nil);
  if FLibHandle=nil then
    raise Exception.Create('mysql_init failed');

  AnalyzeDataSource;
  if (mysql_real_connect(FLibHandle,
                         PAnsiChar(AnsiString(FHost)),
                         PAnsiChar(AnsiString(sUser)),
                         PAnsiChar(AnsiString(sPwd)),
                         nil, FPort, nil, 0)=nil)
  then
    raise Exception.Create(mysql_error(FLibHandle));
end;

procedure TMySqlLibEngine.SetMaxRecords(value: Integer);
begin
  inherited;

end;

function TMySqlLibEngine.GetDBHandle: Pointer;
begin
  Result := FLibHandle;
end;

procedure TMySqlLibEngine.doExecQuery(Ads: TDataSet; sSql: string);
var
  mySQL_Res: PMYSQL_RES;
begin
  if mysql_real_query(FLibHandle, PAnsiChar(sSql), Length(sSql))<>0 then
    raise Exception.Create(mysql_error(FLibHandle)); 
  mySQL_Res := mysql_store_result(FLibHandle);
end;

procedure TMySqlLibEngine.GetDBs(List: TStrings);
var
  mySQL_Res: PMYSQL_RES;
  MYSQL_ROW: PMYSQL_ROW;   
  LibHandle: PMYSQL;  
begin
  LibHandle := FLibHandle;
  mySQL_Res := mysql_list_dbs(LibHandle, nil);
  if mySQL_Res=nil then   
    FLastError := mysql_error(LibHandle);
  try
    repeat
      MYSQL_ROW := mysql_fetch_row(mySQL_Res);
      if MYSQL_ROW<>nil
      then begin
        List.Add(MYSQL_ROW^[0]);
      end;
    until MYSQL_ROW=nil;
  finally
    mysql_free_result(mySQL_Res);
    mySQL_Res := nil;
  end;
end;

procedure TMySqlLibEngine.GetFieldNames(const tableName: string;
  List: TStrings);  
var
  i, field_count, row_count: Integer;
  mySQL_Field: PMYSQL_FIELD;
  sql: AnsiString;
  ticks: Cardinal;
  mySQL_Res: PMYSQL_RES;  
  MYSQL_ROW: PMYSQL_ROW;
  LibHandle: PMYSQL;
begin     
  LibHandle := FLibHandle;
  mySQL_Res := mysql_list_fields(LibHandle,  PAnsiChar(tableName), nil);

end;

procedure TMySqlLibEngine.GetTableNames(List: TStrings);
var       
  mySQL_Res: PMYSQL_RES;
  MYSQL_ROW: PMYSQL_ROW;
  LibHandle: PMYSQL;  
begin
  LibHandle := FLibHandle;
  if mySQL_Res<>nil then
    mysql_free_result(mySQL_Res);

  mySQL_Res := mysql_list_tables(LibHandle, nil);
  if mySQL_Res=nil then
  begin
    FLastError := mysql_error(LibHandle);
    Exit;
  end;  
  try
    repeat
      MYSQL_ROW := mysql_fetch_row(mySQL_Res);
      if MYSQL_ROW<>nil then
      begin
        List.Add(MYSQL_ROW^[0]);
      end;
    until MYSQL_ROW=nil;
  finally
    mysql_free_result(mySQL_Res);
    mySQL_Res := nil;
  end;
end;

procedure TMySqlLibEngine.SetCurrDB(db: string);
var
  LibHandle: PMYSQL;    
  MyResult: Integer;
begin
  LibHandle := FlibHandle;
  FCurrDB := db;     
  MyResult := mysql_select_db(LibHandle, PAnsiChar(FCurrDB));
  if MyResult<>0 then
    FLastError := mysql_error(LibHandle);
end;

end.
