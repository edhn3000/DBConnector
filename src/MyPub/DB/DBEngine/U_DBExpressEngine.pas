{**
 * <p>Title: U_DBExpressEngine.pas  </p>
 * <p>Copyright: Copyright (c) 2009-4-12</p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: dbexpress引擎 </p>
 *}
unit U_DBExpressEngine;

interface

uses
  ADODB, DB, Classes, Variants, SqlExpr, SysUtils, 
  U_DBEngine{$IFDEF VER180}, DBXCommon{$ENDIF}, SimpleDS, U_DBEngineInterface;

type
  
  { TDBExpressEngine }
  TDBExpressEngine = class(TDBEngine)
  private
//    FSQLConnection: TSQLConnection;
//    FSQLQuery: TSQLQuery;
//    FClientDataSet: TClientDataSet;
//    FDataSetProvider: TDataSetProvider;
    FSimpleDataSet: TSimpleDataSet;
//    FSimpleDataSet: TSimple
//    FSQLStoredProc:TSQLStoredProc;
//    FSQLTable:TSQLTable;
//    FSQLDataSet: TSQLDataSet;
    {$IFDEF VER180}
    FTransDesc: TDBXTransaction;
    {$ENDIF}
    function GetSQLConnection: TSQLConnection;
    function GetSimpleDataSet: TSimpleDataSet;
  protected
    function GetMaxRecords: Integer;override;
    procedure SetMaxRecords(value: Integer);override;
    function GetConnection: TCustomConnection;override;
    function GetConnected: Boolean;override;
    function GetDataSet: TDataSet;override;
    
    function OpenDBWithDBEngine(sDataSource: string;sUser: string;sPwd: string;
      dbt: TDBType): Boolean;override;
      
    function doCloseConnection: Boolean;override;
    procedure doCloseQuery;override;

    procedure doExecQuery(Ads: TDataSet; sSql: string);override;
    procedure doExecQueryWithParams(Ads: TDataSet; sSqlWithParams: string;
      aryParams: array of Variant);override;
    function doExecUpdate(Ads: TDataSet; sSql: string):Integer;override;
    function doExecUpdateWithParams(Ads: TDataSet; sSqlWithParams: string;
      aryParams: array of Variant):Integer;override;   
  public
    property Connection: TSQLConnection read GetSQLConnection;
    property DataSet: TSimpleDataSet read GetSimpleDataSet;
  public    
    constructor Create; override;
    destructor Destroy; override;

    function GetNewQuery: TDataSet;override;
    { 事务 }
    function BeginTrans: Integer;override;
    procedure CommitTrans;override;
    procedure RollbackTrans;override;

    function IsEngineInstalled:Boolean;override;  
    procedure SetDataSetMaxRecords(Ads: TDataSet; nMaxRecords: Integer);override;
  end;


implementation

{ TDBExpressEngine }

constructor TDBExpressEngine.Create;
begin
  inherited;
//  FSQLConnection := TSQLConnection.Create(nil);
//  FSQLConnection.LoginPrompt := False;
//  FSQLQuery := TSQLQuery.Create(nil);
//  FClientDataSet := TClientDataSet.Create(nil);
//  FDataSetProvider := TDataSetProvider.Create(nil);
//  FClientDataSet.ProviderName := 'FDataSetProvider';

  FSimpleDataSet := TSimpleDataSet.Create(nil);
  FSimpleDataSet.Connection.LoginPrompt := False;

  FConnection := FSimpleDataSet.Connection;
  FDataSet := FSimpleDataSet;
  FDBEngineType := dbetDBExpress; 
end;

destructor TDBExpressEngine.Destroy;
begin
  FSimpleDataSet.Free;
//  FSQLQuery.Free;
//  FSQLConnection.Free;
//  FClientDataSet.Free;
//  FDataSetProvider.Free;
  inherited;
end;

function TDBExpressEngine.OpenDBWithDBEngine(sDataSource, sUser,
  sPwd: string; dbt: TDBType): Boolean;
var
  slst: TStringList;
begin
  Result := False;

  slst := TStringList.Create;
  try       
    FSimpleDataSet.Connection.Close;
      
    case dbt of
      dbtOracle:
      begin
        FSimpleDataSet.Connection.ConnectionName := 'ORACLE';
        FSimpleDataSet.Connection.DriverName := 'ORACLE';
        FSimpleDataSet.Connection.GetDriverFunc := 'getSQLDriverORACLE';
        FSimpleDataSet.Connection.LibraryName := 'dbexpora.dll';
        FSimpleDataSet.Connection.VendorLib := 'OCI.DLL';
      end;
      dbtMySQL:
      begin
        FSimpleDataSet.Connection.ConnectionName := 'MSConnection';
        FSimpleDataSet.Connection.DriverName := 'MYSQL';
        FSimpleDataSet.Connection.GetDriverFunc := 'getSQLDriverMYSQL';
        FSimpleDataSet.Connection.LibraryName := 'dbexpmys.dll';
        FSimpleDataSet.Connection.VendorLib := 'LIBMYSQL.dll';
      end;
      dbtSybaseASA:
      begin
        FSimpleDataSet.Connection.ConnectionName := 'ASAConnection';
        FSimpleDataSet.Connection.DriverName := 'ASA';
        FSimpleDataSet.Connection.GetDriverFunc := 'getSQLDriverASA';
        FSimpleDataSet.Connection.LibraryName := 'dbxasa30.dll';
        FSimpleDataSet.Connection.VendorLib := 'dbodbc9.dll';
      end;
      else
        Exit;
    end;
    
    ExtractStrings([C_sSEPARATOR_DATASOURCE], [], PChar(sDataSource), slst);
    while slst.Count < 2 do
      slst.Add('');
      
    FSimpleDataSet.Connection.Params.Clear;
    FSimpleDataSet.Connection.Params.Add('User_Name='+sUser);
    FSimpleDataSet.Connection.Params.Add('Password='+sPwd);
    FSimpleDataSet.Connection.Params.Add('Database='+slst[0]);
    FSimpleDataSet.Connection.Open;
//    FSQLQuery.SQLConnection := FSQLConnection;

    FConnection := FSimpleDataSet.Connection;
    FDataSet := FSimpleDataSet.DataSet;
    Result := FConnection.Connected;
  finally
    slst.Free;
  end;
end;

function TDBExpressEngine.doCloseConnection: Boolean;
begin
  FConnection.Close;
  Result := not FConnection.Connected;
end;

procedure TDBExpressEngine.doCloseQuery;
begin
  inherited;
  FDataSet.Close;
end;

procedure TDBExpressEngine.doExecQuery(Ads: TDataSet; sSql: string);
begin
  inherited;
  with TSimpleDataSet(Ads) do
  begin
//    if SQLConnection <> FConnection then
//      SQLConnection := FConnection;   
    Close;
//    dataset.CanModify := True
    DataSet.CommandType := ctQuery;
    DataSet.CommandText := sSql;
    Open;
  end;
end;

procedure TDBExpressEngine.doExecQueryWithParams(Ads: TDataSet;
  sSqlWithParams: string; aryParams: array of Variant);
var
  i: Integer;
begin
  inherited;
  with TSimpleDataSet(Ads) do
  begin
//    if SQLConnection <> FSQLConnection then
//      SQLConnection := FSQLConnection; 
    Close;
//    SQL.Text := sSqlWithParams;
    DataSet.CommandType := ctQuery;
    DataSet.CommandText := sSqlWithParams;
    for i := Low(aryParams) to High(aryParams) do
      Params[i].Value := aryParams[i];
    Open;
  end;
end;

function TDBExpressEngine.doExecUpdate(Ads: TDataSet;
  sSql: string): Integer;
begin  
  with TSimpleDataSet(Ads) do
  begin
//    if SQLConnection <> FSQLConnection then
//      SQLConnection := FSQLConnection;
    Close;
//    SQL.Text := sSql;
    DataSet.CommandType := ctQuery;
    DataSet.CommandText := sSql;
    Execute;
    Result := 1;
    // TODO: 暂时没有找到可以获得影响列数的属性
//    Result := DataSet.ExecSQL(True);
  end;
end;

function TDBExpressEngine.doExecUpdateWithParams(Ads: TDataSet;
  sSqlWithParams: string; aryParams: array of Variant): Integer;
var
  i: Integer;
begin
  with TSimpleDataSet(Ads) do
  begin
//    if SQLConnection <> FSQLConnection then
//      SQLConnection := FSQLConnection;  
    Close;
//    SQL.Text := sSqlWithParams;
    DataSet.CommandType := ctQuery;
    DataSet.CommandText := sSqlWithParams;
    for i := Low(aryParams) to High(aryParams) do
      Params[i].Value := aryParams[i];
    Execute;
    Result := 1;
    // TODO: 暂时没有找到可以获得影响列数的属性
//    Result := ExecSQL(True);
  end;
end;

function TDBExpressEngine.GetConnected: Boolean;
begin
  Result := FConnection.Connected;
end;

function TDBExpressEngine.GetConnection: TCustomConnection;
begin
  Result := TSQLConnection(FConnection);
end;

function TDBExpressEngine.GetMaxRecords: Integer;
begin
  Result := FMaxRecords;
end;

function TDBExpressEngine.GetNewQuery: TDataSet;
var
  dst: TSimpleDataSet;
begin
  dst := TSimpleDataSet.Create(nil);
  dst.Connection := FSimpleDataSet.Connection;
//  dst.SQLConnection := FSQLConnection;
  Result := dst;
end;

procedure TDBExpressEngine.SetMaxRecords(value: Integer);
begin
  inherited;
  FMaxRecords := value;
end;

function TDBExpressEngine.BeginTrans: Integer;
begin
  Result := 0;
  if FSimpleDataSet.Connection.InTransaction then       // 暂时控制为单事物
    Exit;
  {$IFDEF VER180}
  FTransDesc := FSimpleDataSet.Connection.BeginTransaction;
  {$ENDIF}
  if FSimpleDataSet.Connection.InTransaction then
    Result := 1;
end;

procedure TDBExpressEngine.CommitTrans;
begin
  inherited;
  {$IFDEF VER180}
  FSQLConnection.CommitFreeAndNil(FTransDesc); 
  {$ENDIF}
end;

procedure TDBExpressEngine.RollbackTrans;
begin
  inherited;         
  {$IFDEF VER180}
  FSQLConnection.RollbackFreeAndNil(FTransDesc); 
  {$ENDIF}
end;

function TDBExpressEngine.GetDataSet: TDataSet;
begin
  Result := FSimpleDataSet;
end;

function TDBExpressEngine.IsEngineInstalled: Boolean;
begin
  Result := True;
end;

procedure TDBExpressEngine.SetDataSetMaxRecords(Ads: TDataSet;
  nMaxRecords: Integer);
begin
  inherited;
  
end;

function TDBExpressEngine.GetSQLConnection: TSQLConnection;
begin
  Result := FSimpleDataSet.Connection;
end;

function TDBExpressEngine.GetSimpleDataSet: TSimpleDataSet;
begin
  Result := FSimpleDataSet;
end;

end.
