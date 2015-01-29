unit U_FieldInfoFactory;

interface

uses
  SysUtils,    
  U_DBEngineInterface,
  U_FieldInfo,
  U_AccessFieldInfo, U_OracleFieldInfo, U_SybaseFieldInfo, U_MySqlFieldInfo;

type
  TFieldInfoFactory = class
  public
    class function CreateFieldItemList(dbt: TDBType): TFieldItemList;
  end;

implementation

{ TFieldInfoFactory }

class function TFieldInfoFactory.CreateFieldItemList(
  dbt: TDBType): TFieldItemList;
begin
  case dbt of
  dbtAccess, dbtAccess2007:
    Result := TAccessFieldItemList.Create;
  dbtOracle:
    Result := TOracleFieldItemList.Create;
  dbtSybase:
    Result := TSybaseFieldItemList.Create;
  dbtMySql:
    Result := TMySQLFieldItemList.Create;
  else
    Result := TFieldItemList.Create;
  end;
end;

end.
