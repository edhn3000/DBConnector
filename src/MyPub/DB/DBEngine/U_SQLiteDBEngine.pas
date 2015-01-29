unit U_SQLiteDBEngine;

interface

uses
  ASGSQLite3, DB, Classes, Variants, SysUtils,
  U_DBEngine, U_DBEngineInterface;

type
  TSQLiteDBEngine = class(TDBEngine)
  private
    FSqlConn: TASQLite3DB;
    FSqlQry: TASQLite3Query;
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
    property Connection: TASQLite3DB read FSqlConn;
    property DataSet: TASQLite3Query read FSqlQry;
  public
    constructor Create; override;
    destructor Destroy; override;
    function IsConnected: Boolean;override;

    function GetNewQuery: TDataSet;override;

    { ÊÂÎñ }
    function BeginTrans: Integer; override;
    procedure CommitTrans;override;
    procedure RollbackTrans;override;

    function IsEngineInstalled:Boolean;override;
    procedure SetDataSetMaxRecords(Ads: TDataSet; nMaxRecords: Integer);override;
  end;


implementation

{ TSQLiteDBEngine }

constructor TSQLiteDBEngine.Create;
begin
  inherited;
  FSqlConn := TASQLite3DB.Create(nil);
  FSqlConn.DriverDLL := 'sqlite3.dll';
  FSqlQry := TASQLite3Query.Create(nil);
  FSqlQry.Connection := FSqlConn;

  FConnection := TCustomConnection(FSqlConn);
  FDataSet := FSqlQry;
end;

destructor TSQLiteDBEngine.Destroy;
begin
  FSqlQry.Free;
  FSqlConn.Free;
  inherited;
end;

function TSQLiteDBEngine.BeginTrans: Integer;
begin
  FSqlConn.StartTransaction;
  Result := 1;
end;

procedure TSQLiteDBEngine.CommitTrans;
begin
  inherited;
  FSqlConn.Commit;
end;

procedure TSQLiteDBEngine.RollbackTrans;
begin
  inherited;
  FSqlConn.RollBack;
end;

function TSQLiteDBEngine.doCloseConnection: Boolean;
begin
  FSqlConn.Close;
  Result := True;
end;

procedure TSQLiteDBEngine.doCloseQuery;
begin
  inherited;
  FSqlQry.Close;
end;

procedure TSQLiteDBEngine.doExecQuery(Ads: TDataSet; sSql: string);
begin
  inherited;
  with TASQLite3Query(Ads) do
  begin
    if Connection <> FSqlConn then
      Connection := FSqlConn;
    Close;
    SQL.Text := sSql;
    Open;
  end;
end;

procedure TSQLiteDBEngine.doExecQueryWithParams(Ads: TDataSet;
  sSqlWithParams: string; aryParams: array of Variant);
var
  i: Integer;
begin
  inherited;
  with TASQLite3Query(Ads) do
  begin
    if Connection <> FSqlConn then
      Connection := FSqlConn;

    Close;
    SQL.Text := sSqlWithParams;
    for i := Low(aryParams) to High(aryParams) do
      Params[i].Value := aryParams[i];
    Open;
  end;
end;

function TSQLiteDBEngine.doExecUpdate(Ads: TDataSet; sSql: string): Integer;
begin
  with TASQLite3Query(Ads) do
  begin
    if Connection <> FSqlConn then
      Connection := FSqlConn;

    Close;
    SQL.Text := sSql;
    ExecSQL;
    Result := FSqlConn.RowsAffected;
  end;
end;

function TSQLiteDBEngine.doExecUpdateWithParams(Ads: TDataSet;
  sSqlWithParams: string; aryParams: array of Variant): Integer;
var
  i: Integer;
begin
  with TASQLite3Query(Ads) do
  begin
    if Connection <> FSqlConn then
      Connection := FSqlConn;

    Close;
    SQL.Text := sSqlWithParams;
    for i := Low(aryParams) to High(aryParams) do
      Params[i].Value := aryParams[i];
    Result := FSqlConn.RowsAffected;
  end;
end;

function TSQLiteDBEngine.GetConnected: Boolean;
begin
  Result := FSqlConn.Connected;
end;

function TSQLiteDBEngine.GetConnection: TCustomConnection;
begin
  Result := TCustomConnection(FSqlConn);
end;

function TSQLiteDBEngine.GetDataSet: TDataSet;
begin
  Result := FSqlQry;
end;

function TSQLiteDBEngine.GetMaxRecords: Integer;
begin
  Result := FSqlQry.MaxResults;
end;

procedure TSQLiteDBEngine.SetDataSetMaxRecords(Ads: TDataSet;
  nMaxRecords: Integer);
begin
  inherited;
  TASQLite3Query(Ads).MaxResults := nMaxRecords;
end;

procedure TSQLiteDBEngine.SetMaxRecords(value: Integer);
begin
  inherited;
  FMaxRecords := value;
  FSqlQry.MaxResults := value;
end;

function TSQLiteDBEngine.GetNewQuery: TDataSet;
var
  q: TASQLite3Query;
begin
  q := TASQLite3Query.Create(nil);
  q.Connection := FSqlConn;
  Result := q;
end;

function TSQLiteDBEngine.IsConnected: Boolean;
begin
  Result := FSqlConn.Connected;
end;

function TSQLiteDBEngine.IsEngineInstalled: Boolean;
begin
  // TODO check sqlite3.dll
  Result := True;
end;

function TSQLiteDBEngine.OpenDBWithDBEngine(sDataSource, sUser, sPwd: string;
  dbt: TDBType): Boolean;
var
  sDir, sFileName, sExt: String;
begin
  sDir := ExtractFilePath(sDataSource);
  sFileName := ExtractFileName(sDataSource);
  sExt := ExtractFileExt(sDataSource);
  FSqlConn.DefaultExt := sExt;
  FSqlConn.DefaultDir := sDir;
  FSqlConn.Database := sFileName;
  FSqlConn.Open;
  Result := FSqlConn.Connected;
end;

end.
