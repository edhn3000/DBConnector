{**
 * <p>Title: U_DBEngineInterface  </p>
 * <p>Copyright: Copyright (c) 2012/6/10</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 数据库操作类的接口
 *}
unit U_DBEngineInterface;

interface

uses
  DB, Classes; 

type     
  { TDBType 数据库类型,构造连接串时判断 }
  TDBType = (dbtUnKnown, dbtAccess, dbtAccess2007, dbtOracle, dbtSybase,
             dbtMySql, dbtSybaseASA, dbtSQLite, dbtDBF);

  { TDBEngineType }
  TDBEngineType = (dbetAuto, dbetAdo, dbetBde, dbetDBExpress, dbetSQLite);
             
  IDBEngine = interface
  ['{9337BA96-BADD-43A2-8878-D4D3F418F481}']

    function GetLastError: string;
    function GetLastElapsedMilis: Double;
    function GetDBType: TDBType;
    function GetDBEngineType: TDBEngineType;      
    function GetOnConnected(): TNotifyEvent;
    function GetOnDisConnected(): TNotifyEvent;
    procedure SetOnConnected(event: TNotifyEvent);
    procedure SetOnDisConnected(event: TNotifyEvent);
    function IsConnected: Boolean;
    procedure SetConnected(value: Boolean);
    function GetDataSource: string;
//    function SetDataSource(value: string);
    function GetUserName: string;
    function GetPassword: string;
    function GetDataSet: TDataSet;
    function GetConnection: TCustomConnection;
    function GetMaxRecords: Integer;
    procedure SetMaxRecords(value: Integer);
    
    { 数据库操作 }
    // 参照 TADODBEngine.BuildConnectionString  和 TBDEDBEngine.OpenDBWithDBEngine
    function OpenDB(sDataSource: string;sUser: string;sPwd: string;
      dbt: TDBType): Boolean;overload;
    function OpenDBWithIni(AIniFile, ADBSection: string):Boolean;overload;  
    function ReOpenDB: Boolean;
    function CloseDB: Boolean; 
    
    function CloseQuery: Boolean;
    function GetNewQuery: TDataSet;

    { 执行sql的方法 }
    function ExecQuery(sSql: string):Boolean; overload;
    function ExecQuery(AQ: TDataSet; sSql: string):Boolean; overload;
    function ExecQueryWithParams(sSqlWithParams: string; Params: array of Variant):Boolean; overload;
    function ExecQueryWithParams(AQ: TDataSet; sSqlWithParams: string;
        aryParams: array of Variant):Boolean;overload;

    function ExecUpdate(sSql: string): Integer; overload;
    function ExecUpdate(AQ: TDataSet; sSql: string): Integer; overload;
    function ExecUpdateWithParams(sSqlWithParams: string; Params: array of Variant): Integer; overload;
    function ExecUpdateWithParams(AQ: TDataSet; sSqlWithParams: string;
        aryParams: array of Variant): Integer;overload;

    function ExecProc(sProcName: String; params: array of Variant):Boolean;overload;
    function ExecProc(sProcName: String; keyvalues: TStrings):Boolean;overload;

    { 事务 }
    function BeginTrans: Integer;
    procedure CommitTrans;
    procedure RollbackTrans;

    function KeepDBOpen: Boolean;

    function IsEngineInstalled:Boolean;
    procedure SetDataSetMaxRecords(Ads: TDataSet; nMaxRecords: Integer);
  end;      
                   
  { 枚举to字符串 }
  function DBTypeToStr(Adbt: TDBType; bWithPrefix: Boolean = True): string;
  function StrToDBType(s: string): TDBType;

  function DBEngineTypeToStr(Adbet: TDBEngineType; bWithPrefix: Boolean = True): string;
  function StrToDBEngineType(s: string): TDBEngineType;   

  function BuildAccessDataSource(sDb, sSecured: string): string;
  function BuildOracleDataSource(SID, IP, Port, Protocol: string;
    bLocalTns: Boolean): string;
  function BuildSybaseDataSource(ServerName, Port, DataBaseName: string): string;
  function BuildMySqlDataSource(Host, DataBaseName, CharSet: string): string;

const
  C_nTDBE_ERROR_EXECFAIL = -1;
//  C_sSEPARATOR_ORACLE_TNS = ',';
  C_sSEPARATOR_DATASOURCE = ',';

  ERROR_ACCESS_DATASOURCE =
    'Access数据源格式错误！正确的格式为：[数据源, 工作组]';
  ERROR_ORACLE_DATASOURCE = 'Oracle数据源格式错误！正确的格式为：'
    +'[服务名, 主机, 端口, 协议]';
  ERROR_DBTYPE = '数据库类型错误！';

  CONNSTR_ACCESS =
    'Provider=Microsoft.Jet.OLEDB.4.0;'
    + 'Data Source=%s;'
    + '%s'                                // Jet OLEDB:System database=%s;
//    + 'Persist Security Info=False;'
    + 'User ID=%s;'
    + 'Jet OLEDB:Database Password=%s;';
//    + 'Password=%s;';
  SYSDB_ACCESS = 'Jet OLEDB:System database=%s;';

  // access 2007
  CONNSTR_ACCESS2007 =
    'Provider=Microsoft.ACE.OLEDB.12.0;'
     + 'Data Source=%s;'
     + '%s'                                 // Jet OLEDB:System database=%s;
//     + 'Persist Security Info=False;'
     + 'User ID=%s;'
     + 'Password=%s;';
  SYSDB_ACCESS2007 = SYSDB_ACCESS;

  CONNSTR_DBF =
//    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s;Extended Properties=dBASE IV;User ID=Admin;Password=;';
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s;Extended Properties=dBase 5.0;Persist Security Info=False';

//  CONNSTR_DBF2 = 'Provider=ADsDSOObject;Encrypt Password=False;Data Source=%s;Mode=Read;Bind Flags=0;ADSI Flag=-2147483648';

  CONNSTR_ORACLE =
    'Provider=OraOLEDB.Oracle.1;' +
    'Password=%s;' +
//    'Persist Security Info=True;' +
    'User ID=%s;' +
    'Data Source=%s';

  CONNSTR_SYBASE =
//  'Provider=MSDASQL.1;Driver={SYBASE SYSTEM 12};srvr=%s;db=%s;uid=%s;pwd=%s';
  'Provider=Sybase.ASEOLEDBProvider.2;' +
  'Initial Catalog=%s;Language=us_english;Character Set=cp936;' +
  'Server Name=%s;Server Port Address=%s;User ID=%s;Password=%s';

  TNSNAME_ORACLE =
      '(DESCRIPTION = '
//    +   '(ADDRESS_LIST = '
    +     '(ADDRESS = (PROTOCOL = %s)(HOST = %s)(PORT = %s))'
//    +    ')'
    +   '(CONNECT_DATA = '
    +     '(SERVER = DEDICATED)'
    +     '(SERVICE_NAME = %s) '
    +   ')'
    + ')';

  CONNSTR_MYSQL3 =
  'DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=;'
  +'USER=root;PASSWORD=mysql;OPTION=3;';

  CONNSTR_MYSQL5 =
  'Provider=MSDASQL.1;Persist Security Info=False;Data Source=%s;%s%s%s';
  CHARSET_MYSQL  = 'CHARSET=%s;';
  USERINFO_MYSQL = 'User ID=%s;Password=%s;';
  DataBase_MySQL = 'DATABASE=%s';

implementation

uses
  TypInfo, SysUtils;

function DBTypeToStr(Adbt: TDBType; bWithPrefix: Boolean = True): string;
var
  ti: PTypeInfo;
  td: PTypeData;
  i: Integer;
begin
  ti := TypeInfo(TDBType);
  td := GetTypeData(ti);
  for i := td^.MinValue to td^.MaxValue do
    if Integer(Adbt) = i then
    begin
      Result := GetEnumName(ti, i);
      if not bWithPrefix then
        Result := Copy(Result, 4, Length(Result));
      Break;
    end;
end;

function StrToDBType(s: string): TDBType;
var
  nDBT, nCode: Integer;
  ti: PTypeInfo;
  td: PTypeData;
  dbt: TDBType;
begin  
  Result := dbtUnKnown;
  Val(s, nDBT, nCode);
  if (nCode = 0) then
  begin
    ti := TypeInfo(TDBType);
    td := GetTypeData(ti);
    if (nDBT <= td^.MaxValue) and (nDBT >= td^.MinValue) then
      Result := TDBType(nDBT);
  end
  else
  begin
    for dbt := Low(TDBType) to High(TDBType) do
      if SameText(DBTypeToStr(dbt, False), s) then
      begin
        Result := dbt;
        Break;
      end;  
  end;  
end;

function DBEngineTypeToStr(Adbet: TDBEngineType; bWithPrefix: Boolean = True): string;
var
  ti: PTypeInfo;
  td: PTypeData;
  i: Integer;
begin
  ti := TypeInfo(TDBEngineType);
  td := GetTypeData(ti);
  for i := td^.MinValue to td^.MaxValue do
    if Integer(Adbet) = i then
    begin
      Result := GetEnumName(ti, i);
      if not bWithPrefix then
        Result := Copy(Result, 5, Length(Result));
      Break;
    end;
end;

function StrToDBEngineType(s: string): TDBEngineType;
var
  nDBT, nCode: Integer;
  ti: PTypeInfo;
  td: PTypeData;
  dbet: TDBEngineType;
begin
  Result := dbetAuto;
  Val(s, nDBT, nCode);
  if (nCode = 0) then
  begin
    ti := TypeInfo(TDBEngineType);
    td := GetTypeData(ti);
    if (nDBT <= td^.MaxValue) and (nDBT >= td^.MinValue) then
      Result := TDBEngineType(nDBT);
  end
  else
  begin
    for dbet := Low(TDBEngineType) to High(TDBEngineType) do
      if SameText(DBEngineTypeToStr(dbet, False), s) then
      begin
        Result := dbet;
        Break;
      end;  
  end;  
end;  

function BuildAccessDataSource(sDb, sSecured: string): string;
begin
  Result := Format('%1:s%0:s%2:s',[
          C_sSEPARATOR_DATASOURCE,sDb, sSecured])
end;

function BuildOracleDataSource(SID, IP, Port, Protocol: string;
  bLocalTns: Boolean): string;
begin
  if bLocalTns then
    Result := SID
  else
    Result := Format('%1:s%0:s%2:s%0:s%3:s%0:s%4:s',
        [C_sSEPARATOR_DATASOURCE, SID, IP, Port, Protocol]);
end;

function BuildSybaseDataSource(ServerName, Port, DataBaseName: string): string;
begin
  Result := Format('%1:s%0:s%2:s%0:s%%3:s%',
            [C_sSEPARATOR_DATASOURCE, ServerName, Port, DataBaseName]);
end;

function BuildMySqlDataSource(Host, DataBaseName, CharSet: string): string;
begin
  Result := Format('%1:s%0:s%2:s%0:s%3:s', [
            C_sSEPARATOR_DATASOURCE, Host, DataBaseName, CharSet])
end;


end.
