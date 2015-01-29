{**
 * <p>Title: U_DBConnnectManager.pas  </p>
 * <p>Copyright: Copyright (c) 2009-4-12</p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: 此单元的目的在于尽量减少外部直接操作DBConnect的各继承类 </p>
 *}
unit U_DBConnnectManager;

interface

uses
  SysUtils,
  U_DBEngineInterface, U_DBConnectInterface,
  U_DBConnect, U_AccessDBConnect, U_OracleDBConnect, U_SybaseDBConnect,
  U_MySQLDBConnect, U_DBConfig, U_SQLiteDBConnect;

type
  TDBConnectManager = class
  public
    class function CreateDBConnect(dbt: TDBType): IDBConnect;
    class function OpenDB(sDataSource, sUser, sPwd: string;
      dbt: TDBType; dbet: TDBEngineType): IDBConnect;overload;
    class function OpenDB(dbCfg: TDBConfig): IDBConnect;overload;    
  end;

implementation

{ TDBConnnectManager }

class function TDBConnectManager.CreateDBConnect(
  dbt: TDBType): IDBConnect;
begin
  case dbt of
  dbtAccess, dbtAccess2007:
    Result := TAccessDBConnect.Create;
  dbtOracle:
    Result := TOracleDBConnect.Create;
  dbtSybase:
    Result := TSybaseDBConnect.Create;
  dbtMySql:
    Result := TMySQLDBConnect.Create;
  dbtSQLite:
    Result := TSQLiteDBConnect.Create;
  else
  begin
    Result := TDBConnect.Create;
  end;
  end;
end;

class function TDBConnectManager.OpenDB(sDataSource, sUser, sPwd: string;
  dbt: TDBType; dbet: TDBEngineType): IDBConnect;
var
  dbconnect: IDBConnect;
begin
  dbconnect := CreateDBConnect(dbt);
  dbconnect.OpenDB(sDataSource, sUser, sPwd, dbt, dbet);
  Result := dbconnect;
end; 

class function TDBConnectManager.OpenDB(dbCfg: TDBConfig): IDBConnect;
var
  dbconnect: IDBConnect;
begin
  dbconnect := CreateDBConnect(dbCfg.DBType);
  dbconnect.OpenDB(dbCfg.GetSeparatedDataSource, dbCfg.UserName, dbCfg.Password,
    dbCfg.DBType, dbCfg.DBEngineType);
  Result := dbconnect;
end;

end.
