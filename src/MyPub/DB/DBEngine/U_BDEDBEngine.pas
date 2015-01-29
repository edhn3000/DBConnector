{**
 * <p>Title: U_BDEDBEngine.pas  </p>
 * <p>Copyright: Copyright (c) 2009-4-12</p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: bde引擎 </p>
 *}
unit U_BDEDBEngine;

interface

uses
  ADODB, DB, Classes, Variants, DBTables, SysUtils, 
  U_DBEngine, U_DBEngineInterface;

type

{ TBDEDBEngine }
  TBDEDBEngine = class(TDBEngine)
  private
    FDataBase: TDataBase;
    FBDEQry  : TQuery;
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
    property Connection: TDataBase read FDataBase;  
    property DataSet: TQuery read FBDEQry;
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

{ TBDEDBEngine }

function TBDEDBEngine.BeginTrans: Integer;
begin
  FDataBase.StartTransaction;
  Result := 1;
end;

procedure TBDEDBEngine.CommitTrans;
begin
  FDataBase.Commit;
end;    

procedure TBDEDBEngine.RollbackTrans;
begin
  FDataBase.Rollback;
end;

function TBDEDBEngine.GetNewQuery: TDataSet;
var
  bdeqry: TQuery;
begin
  bdeqry := TQuery.Create(nil);
  bdeqry.DatabaseName := FDataBase.DatabaseName;
  bdeqry.SessionName := FDataBase.SessionName;
  Result := bdeqry;
end;

constructor TBDEDBEngine.Create;
begin
  inherited;
  FDataBase := TDatabase.Create(nil);
  FDataBase.LoginPrompt := False;
  FBDEQry := TQuery.Create(nil);
  
  // 对父类对象赋值
  FConnection := FDataBase;
  FDataSet := FBDEQry;
  FDBEngineType := dbetBDE;
end;

destructor TBDEDBEngine.Destroy;
begin
  if Assigned(FBDEQry) then
  begin
    FBDEQry.Close;
    FreeAndNil(FBDEQry);
  end;   
  if Assigned(FDataBase) then
  begin
    FDataBase.Close;
    FreeAndNil(FDataBase);
  end;
  inherited;
end;

function TBDEDBEngine.doCloseConnection: Boolean;
begin
  inherited;
  FDataBase.Close;
  Result := not FDataBase.Connected;
end;

procedure TBDEDBEngine.doCloseQuery;
begin
  inherited;
  FBDEQry.Close;
end;

function TBDEDBEngine.OpenDBWithDBEngine(sDataSource, sUser, sPwd: string;
  dbt: TDBType):Boolean;
var
  slst: TStrings;
  sDriveName: string;
begin
  Result := False;
  case dbt of
    dbtSyBase:
      sDriveName := 'SYBASE';
//DATABASE NAME=
//SERVER NAME=SYB_SERVER
//USER NAME=MYNAME
//OPEN MODE=READ/WRITE
//SCHEMA CACHE SIZE=8
//BLOB EDIT LOGGING=
//LANGDRIVER=
//SQLQRYMODE=
//SQLPASSTHRU MODE=SHARED AUTOCOMMIT
//DATE MODE=0
//SCHEMA CACHE TIME=-1
//MAX QUERY TIME=300
//MAX ROWS=-1
//BATCH COUNT=200
//ENABLE SCHEMA CACHE=FALSE
//SCHEMA CACHE DIR=
//HOST NAME=
//APPLICATION NAME=
//NATIONAL LANG NAME=
//ENABLE BCD=FALSE
//TDS PACKET SIZE=512
//BLOBS TO CACHE=64
//BLOB SIZE=32
//CS CURSOR ROWS=1
//PASSWORD=
    dbtOracle:
      sDriveName := 'ORACLE';
//DATABASE NAME=
//USER NAME=
//ODBC DSN=
//OPEN MODE=READ/WRITE
//SCHEMA CACHE SIZE=8
//SQLQRYMODE=
//LANGDRIVER=
//SQLPASSTHRU MODE=SHARED AUTOCOMMIT
//SCHEMA CACHE TIME=-1
//MAX ROWS=-1
//BATCH COUNT=200
//ENABLE SCHEMA CACHE=FALSE
//SCHEMA CACHE DIR=
//ENABLE BCD=FALSE
//ROWSET SIZE=20
//BLOBS TO CACHE=64
//PASSWORD=

    dbtAccess, dbtAccess2007:
      sDriveName := 'MSACCESS';
    else
      Exit;
  end;


  slst := TStringList.Create;
  try
    ExtractStrings([C_sSEPARATOR_DATASOURCE], [], PChar(sDataSource), slst);
    while slst.Count < 2 do
      slst.Add('');
    with FDataBase do
    begin
      FDataBase.Close;
      FDataBase.DatabaseName := 'DB_'+ slst[0] + '_' + sUser; //'ADatabaseName';//slst[1];
      FDataBase.DriverName:=sDriveName;
      FDataBase.Params.Clear;
      FDataBase.Params.Add('USER NAME='+sUser);
      FDataBase.Params.Add('PASSWORD='+sPwd);
      case dbt of
        dbtAccess, dbtAccess2007:
        begin  
          FDataBase.Params.Add('DATABASE NAME='+slst[0]);
          FDataBase.Params.Add('SYSTEM DATABASE='+slst[1]);
        end;
        else
        begin
          FDataBase.Params.Add('SERVER NAME='+slst[0]);
  //    FDataBase.Params.Add('BLOB SIZE='+IntToStr(iBlobSize));
        end;
      end;
      FDataBase.Open;
      FBDEQry.DatabaseName := FDataBase.DatabaseName;
      FBDEQry.SessionName := FDataBase.SessionName;
      Result := FDataBase.Connected;
    end;
  finally
    slst.free;
  end;
end;

procedure TBDEDBEngine.doExecQuery(Ads: TDataSet; sSql: string);
begin
  inherited;
  with TQuery(Ads) do
  begin
    if DatabaseName <> FDataBase.DatabaseName then
      DatabaseName := FDataBase.DatabaseName;
    Close;
    SQL.Text := sSql;
    Open;
  end;
end;

procedure TBDEDBEngine.doExecQueryWithParams(Ads: TDataSet; sSqlWithParams: string;
  aryParams: array of Variant);
var
  i: Integer;
begin
  inherited;
  with TQuery(Ads) do
  begin
    if DatabaseName <> Database.DatabaseName then
      DatabaseName := Database.DatabaseName;

    Close;
    SQL.Text := sSqlWithParams;
    for i := Low(aryParams) to High(aryParams) do
      Params[i].Value := aryParams[i];
    Open;
  end;
end;

function TBDEDBEngine.doExecUpdate(Ads: TDataSet; sSql: string): Integer;
begin
  with TQuery(Ads) do
  begin
    Close;
    SQL.Text := sSql;
    ExecSQL;
    Result := RowsAffected;
  end;
end;

function TBDEDBEngine.doExecUpdateWithParams(Ads: TDataSet; sSqlWithParams: string;
  aryParams: array of Variant): Integer;
var
  i: Integer;
begin
  with TQuery(Ads) do
  begin
    if DatabaseName <> Database.DatabaseName then
      DatabaseName := Database.DatabaseName;

    Close;
    SQL.Text := sSqlWithParams;
    for i := Low(aryParams) to High(aryParams) do
      Params[i].Value := aryParams[i];
    ExecSQL;
    Result := RowsAffected;
  end;
end;


function TBDEDBEngine.GetConnected: Boolean;
begin
  Result := FDataBase.Connected;
end;

function TBDEDBEngine.GetConnection: TCustomConnection;
begin
  Result := FDataBase;
end;

function TBDEDBEngine.GetMaxRecords: Integer;
begin
  Result := FMaxRecords;
end;

procedure TBDEDBEngine.SetMaxRecords(value: Integer);
begin
  inherited;
  FMaxRecords := value;
end;

function TBDEDBEngine.GetDataSet: TDataSet;
begin
  Result := FBDEQry;
end;

function TBDEDBEngine.IsEngineInstalled: Boolean;
begin
  Result := True;
end;

procedure TBDEDBEngine.SetDataSetMaxRecords(Ads: TDataSet; nMaxRecords: Integer);
begin
  inherited;

end;

end.
