unit U_TableInfoFactory;

interface        

uses
  SysUtils,    
  U_DBEngineInterface,
  U_TableInfo,
  U_AccessTableInfo, U_OracleTableInfo, U_SybaseTableInfo, U_MySqlTableInfo;

type
  TTableInfoFactory = class
  public
    class function CreateTableInfo(dbt: TDBType): TTableInfo;
  end;

implementation

{ TTableInfoFactory }

class function TTableInfoFactory.CreateTableInfo(dbt: TDBType): TTableInfo;
begin
  case dbt of
  dbtAccess, dbtAccess2007:
    Result := TAccessTableInfo.Create;
  dbtOracle:
    Result := TOracleTableInfo.Create;
  dbtSybase:
    Result := TSybaseTableInfo.Create;
  dbtMySql:
    Result := TMySqlTableInfo.Create;
  else
    Result := TTableInfo.Create;
  end;
end;

end.
