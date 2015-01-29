unit U_SQLiteDBConnect;

interface

uses
  DB, Classes, Variants,
  Dialogs, SysUtils, U_TableInfo,

  U_DBConnect;
type
  TSQLiteDBConnect = class(TDBConnect)
  protected
    function GetNewTableInfo(): TTableInfo;override;
  public
    constructor Create;override;
    destructor Destroy;override;

    procedure GetTableNames(List: TStrings; AQry: TDataSet = nil);override;
    procedure GetViewNames(List: TStrings; AQry: TDataSet = nil);override;
    procedure GetIndexNames(List: TStrings);override;
    procedure GetProcedureNames(List: TStrings; AQry: TDataSet = nil);override;
    procedure GetTriggerNames(List: TStrings; AQry: TDataSet = nil);override;
  end;

implementation

{ TSQLiteDBConnect }

constructor TSQLiteDBConnect.Create;
begin
  inherited;

end;

destructor TSQLiteDBConnect.Destroy;
begin

  inherited;
end;

function TSQLiteDBConnect.GetNewTableInfo: TTableInfo;
begin
  Result := TTableInfo.Create;
end;

procedure TSQLiteDBConnect.GetTableNames(List: TStrings; AQry: TDataSet);
var
  sGetTables: string;
begin
  inherited;
  if IsQryInvalid(AQry) then
    AQry := DBEngine.GetDataSet;
  sGetTables := 'SELECT name FROM sqlite_master where type=''table''';
  if ExecQuery(AQry, sGetTables) then
    FillListByOwnerName(List, AQry);
end;

procedure TSQLiteDBConnect.GetViewNames(List: TStrings; AQry: TDataSet);
var
  sGetViews: string;
begin
  sGetViews := 'SELECT name FROM sqlite_master where type=''view''';
  if ExecQuery(AQry, sGetViews) then
    FillListByOwnerName(List, AQry);
end;

procedure TSQLiteDBConnect.GetTriggerNames(List: TStrings; AQry: TDataSet);
var
  sGetViews: string;
begin
  sGetViews := 'SELECT name FROM sqlite_master where type=''trigger''';
  if ExecQuery(AQry, sGetViews) then
    FillListByOwnerName(List, AQry);
end;

procedure TSQLiteDBConnect.GetIndexNames(List: TStrings);
var
  sGetViews: string;
  AQry: TDataSet;
begin
  AQry := GetNewQuery;
  sGetViews := 'SELECT name FROM sqlite_master where type=''index''';
  if ExecQuery(AQry, sGetViews) then
    FillListByOwnerName(List, AQry);
end;

procedure TSQLiteDBConnect.GetProcedureNames(List: TStrings; AQry: TDataSet);
var
  sGetViews: string;
begin
  sGetViews := 'SELECT name FROM sqlite_master where type=''view''';
  if ExecQuery(AQry, sGetViews) then
    FillListByOwnerName(List, AQry);
end;

end.
