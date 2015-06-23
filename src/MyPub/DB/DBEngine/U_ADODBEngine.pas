{**
 * <p>Title: U_ADODBEngine.pas  </p>
 * <p>Copyright: Copyright (c) 2009-4-12</p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: ado引擎 </p>
 *}
unit U_ADODBEngine;

interface

uses
  ADODB, DB, Classes, Variants, SysUtils, IniFiles,
  U_DBEngine, U_DBEngineInterface;

type
{ TADODBEngine }
  TADODBEngine = class(TDBEngine)
  private  
    FAdoConn : TADOConnection;
    FAdoQry  : TADOQuery;
  protected         
    function GetMaxRecords: Integer;override;
    procedure SetMaxRecords(value: Integer);override;
    function GetConnection: TCustomConnection;override;
    function GetConnected: Boolean;override;
    function GetDataSet: TDataSet;override;
    
    function OpenDBWithDBEngine(sDataSource: string;sUser: string;sPwd: string;
      dbt: TDBType):Boolean;override;

    function doCloseConnection: Boolean;override;
    procedure doCloseQuery;override;
    
    procedure doExecQuery(Ads: TDataSet; sSql: string);override;
    procedure doExecQueryWithParams(Ads: TDataSet; sSqlWithParams: string;
      aryParams: array of Variant);override;
    function doExecUpdate(Ads: TDataSet; sSql: string):Integer;override;
    function doExecUpdateWithParams(Ads: TDataSet; sSqlWithParams: string;
      aryParams: array of Variant):Integer;override;
  public
    property Connection: TADOConnection read FAdoConn;
    property DataSet: TADOQuery read FAdoQry;
  public
    constructor Create; override;
    destructor Destroy; override;
                                         
    function OpenDBWithIni(AIniFile, ADBSection: string):Boolean;override;
    function BuildConnectionString(sDataSource, sUser, sPwd: string;
        dbt: TDBType):string;overload;
    function BuildConnectionString(AIniName, ASection: string):string;overload;

    function GetNewQuery: TDataSet;override;

    { 事务 }
    function BeginTrans: Integer;override;
    procedure CommitTrans;override;
    procedure RollbackTrans;override;

    function IsEngineInstalled:Boolean;override;
    procedure SetDataSetMaxRecords(Ads: TDataSet; nMaxRecords: Integer);override;
  end;

implementation


{ TADODBEngine }

function TADODBEngine.BeginTrans: Integer;
begin
  Result := FAdoConn.BeginTrans;
end;

procedure TADODBEngine.CommitTrans;
begin
  if FAdoConn.InTransaction then
    FAdoConn.CommitTrans;
end;

procedure TADODBEngine.RollbackTrans;
begin
  FAdoConn.RollbackTrans;
end;

function TADODBEngine.GetNewQuery: TDataSet;
var
  adoqry: TADOQuery;
begin
  adoqry := TADOQuery.Create(nil);
  adoqry.Connection := FAdoConn;
  Result := adoqry;
end;

constructor TADODBEngine.Create;
begin
  inherited;
  FAdoConn := TADOConnection.Create(nil);   
  FAdoConn.LoginPrompt := False;
  FAdoQry := TADOQuery.Create(nil);
  // 对父类对象赋值
  FConnection := FAdoConn;
  FDataSet := FAdoQry;
  FDBEngineType := dbetADO;
end;

destructor TADODBEngine.Destroy;
begin
  if Assigned(FAdoQry) then
  begin
    FAdoQry.Close;
    FreeAndNil(FAdoQry);
  end;
  if Assigned(FAdoConn) then
  begin
    FAdoConn.Close;
    FreeAndNil(FAdoConn);
  end;    
  inherited;
end;

function TADODBEngine.doCloseConnection: Boolean;
begin
  inherited;
  FAdoConn.Close;
  Result := not FAdoConn.Connected;
end;

procedure TADODBEngine.doCloseQuery;
begin
  inherited;
  FAdoQry.Close;
end;    

function TADODBEngine.BuildConnectionString(AIniName,
  ASection: string): string;
var
  nKey: Integer;
  sValue: string;
begin
  with TIniFile.Create(AIniName) do
  begin
    try
      Result := Trim(ReadString(ASection, 'ConnStr', ''));
      if Result = '' then Exit;
      
      nKey := 0;
      while True do
      begin
        sValue := ReadString(ASection, IntToStr(nKey), '');
        if sValue = '' then Break;

        // 使用{0}的形式表示要替换的参数
        Result := StringReplace(Result, '{' + IntToStr(nKey) + '}', sValue, [
          rfReplaceAll]);
        Inc(nKey);
      end;
    finally
      Free;
    end;
  end;
end;

function TADODBEngine.BuildConnectionString(sDataSource, sUser, sPwd: string;
  dbt: TDBType): string;
var
  sSec, sDSource: string;
  sMySQLDB: string;
  sCharset: string;
  sUserInfo: string;
  slst: TStrings;
begin
  // 构造Oracle的DataSource
  slst := TStringList.Create;
  try
//    ExtractStrings([C_sSEPARATOR_ORACLE_TNS], [], PChar(sDataSource), slst);
    Split(sDataSource, C_sSEPARATOR_DATASOURCE, slst);
    case dbt of
      dbtAccess:
      begin    
        while slst.Count < 2 do
          slst.Add('');
        sSec := slst[1];
//        if '' <> sSec then
          sSec := Format(SYSDB_ACCESS, [sSec]);
        Result := Format(CONNSTR_ACCESS, [slst[0], sSec,
                                                sUser, sPwd, sPwd]);
      end;
      dbtAccess2007:
      begin        
        while slst.Count < 2 do
          slst.Add('');
        sSec := slst[1];
        if '' <> sSec then
          sSec := Format(SYSDB_ACCESS2007, [sSec]);
        Result := Format(CONNSTR_ACCESS2007, [slst[0], sSec,
                                                    sUser, sPwd]);
      end;
      dbtOracle:
      begin    
        while slst.Count < 2 do
          slst.Add('');
        if Pos(C_sSEPARATOR_DATASOURCE, sDataSource) = 0 then    // 使用本地数据库别名时传入的参数不带逗号
          sDSource := sDataSource
        else                                                     // 使用远程数据库服务名
        begin
          // 后两个参数可使用默认值
          if slst.Count = 2 then
            slst.Add('1521');
          if slst.Count = 3 then
            slst.Add('tcp');
//            if slst.Count < 4 then
//              raise Exception.Create(ERROR_ORACLE_DATASOURCE)
          // TNSNAME_ORACLE中顺序:服务名，主机，端口，协议 // slst[3]协议，slst[1]主机，端口，服务名
          if Trim(slst[1]) <> '' then
            sDSource := Format(TNSNAME_ORACLE, [slst[3], slst[1], slst[2],slst[0]])
          else // 不填写ip即可使用本地服务名
            sDSource := slst[0];
        end;

        Result := Format(CONNSTR_ORACLE, [sPwd, sUser, sDSource]);
      end;
      dbtMySQL:
      begin      
        while slst.Count < 3 do
          slst.Add('');
        sMySQLDB := slst[1];
        if sMySQLDB <> '' then
          sMySQLDB := Format(DataBase_MySQL, [sMySQLDB, sMySQLDB]);
        if (slst.Count > 2) and (slst[2] <> '') then
          sCharset := Format(CHARSET_MYSQL, [slst[2]])
        else
          sCharset := Format(CHARSET_MYSQL, ['GBK']);
        if sUser <> '' then    // 如果在ODBC里写了用户名密码 则连接串不再加
          sUserInfo := Format(USERINFO_MYSQL, [sUser, sPwd]);
        Result := Format(CONNSTR_MYSQL5, [slst[0], sCharset, sUserInfo, sMySQLDB]);
      end;
      dbtSybase:
      begin    
        while slst.Count < 3 do
          slst.Add('');
        Result := Format(CONNSTR_SYBASE, [slst[0], slst[1], slst[2],
                  sUser, sPwd]);
      end;
      dbtDBF:
      begin
        Result := Format(CONNSTR_DBF, [slst[0]]);
      end;
      else
      begin
        raise Exception.Create(ERROR_DBTYPE);
        Result := sDSource;
      end;
    end;
  finally
    slst.Free;
  end;
end;

function TADODBEngine.OpenDBWithDBEngine(sDataSource, sUser, sPwd: string;
  dbt: TDBType): Boolean;
begin    
  FAdoConn.ConnectionString := BuildConnectionString(sDataSource,sUser,sPwd,dbt);
  FAdoConn.Connected := True;
  FAdoQry.Connection := FAdoConn;
   
  FConnection := FAdoConn;
  FDataSet := FAdoQry;
  Result := FAdoConn.Connected;
end;

function TADODBEngine.OpenDBWithIni(AIniFile, ADBSection: string): Boolean;
begin   
  FAdoConn.ConnectionString := BuildConnectionString(AIniFile, ADBSection);
  FAdoConn.Connected := True;
  FAdoQry.Connection := FAdoConn;
   
  FConnection := FAdoConn;
  FDataSet := FAdoQry;
  Result := FAdoConn.Connected;
end;
procedure TADODBEngine.doExecQuery(Ads: TDataSet; sSql: string);
begin
  inherited;
  with TADOQuery(Ads) do
  begin
    if Connection <> FConnection then
      Connection := TADOConnection(FConnection);
    Close;
    SQL.Text := sSql;
    Open;
  end;
end;

procedure TADODBEngine.doExecQueryWithParams(Ads: TDataSet; sSqlWithParams: string;
  aryParams: array of Variant);
var
  i: Integer;
begin
  inherited; 

  with TADOQuery(Ads) do
  begin
    if Connection <> FAdoConn then
      Connection := FAdoConn;

    Close;
    SQL.Text := sSqlWithParams;
    for i := Low(aryParams) to High(aryParams) do
      Parameters[i].Value := aryParams[i];
    Open;
  end;
end;

function TADODBEngine.doExecUpdate(Ads: TDataSet; sSql: string):Integer;
begin
  inherited;
  with TADOQuery(Ads) do
  begin
    Close;
    SQL.Text := sSql;
    Result := ExecSQL;
  end;
end;

function TADODBEngine.doExecUpdateWithParams(Ads: TDataSet; sSqlWithParams: string;
  aryParams: array of Variant):Integer;
var
  i: Integer;
begin
  inherited;

  with TADOQuery(Ads) do
  begin
    if Connection <> FAdoConn then
      Connection := FAdoConn;

    Close;
    SQL.Text := sSqlWithParams;
    for i := Low(aryParams) to High(aryParams) do
      Parameters[i].Value := aryParams[i];
    Result := ExecSQL;
  end;
end;

function TADODBEngine.GetConnected: Boolean;
begin
  Result := FAdoConn.Connected;
end;

function TADODBEngine.GetConnection: TCustomConnection;
begin
  Result := FAdoConn;
end;

function TADODBEngine.GetMaxRecords: Integer;
begin
  Result := FAdoQry.MaxRecords;
end;

procedure TADODBEngine.SetMaxRecords(value: Integer);
begin
  inherited;
  FMaxRecords := value;
  FAdoQry.MaxRecords := value;
end;

function TADODBEngine.GetDataSet: TDataSet;
begin
  Result := FAdoQry;
end;

function TADODBEngine.IsEngineInstalled: Boolean;
begin
  Result := True;
end;

procedure TADODBEngine.SetDataSetMaxRecords(Ads: TDataSet; nMaxRecords: Integer);
begin
  inherited;
  TADOQuery(Ads).MaxRecords := nMaxRecords;
end;

end.
