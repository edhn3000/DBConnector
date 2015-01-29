unit U_DBEngineFactory;

interface

uses
  ADODB, DB, Classes, Variants, SysUtils,
  U_DBEngineInterface, U_ADODBEngine, U_BDEDBEngine, U_DBExpressEngine,
  U_DBEngine, U_SQLiteDBEngine;

    
type
{ TDBEngineFactory }
  TDBEngineFactory = class

  public
    class function GetNewDBEngine(Adbet: TDBEngineType): IDBEngine;

  end;

implementation  

{ TDBEngineFactory }

class function TDBEngineFactory.GetNewDBEngine(Adbet: TDBEngineType): IDBEngine;
begin     
  case Adbet of
    dbetBDE:
      Result := TBDEDBEngine.Create;
    dbetDBExpress:
      Result := TDBExpressEngine.Create;
    dbetAdo:
      Result := TADODBEngine.Create;
    dbetSqlite:
      Result := TSQLiteDBEngine.Create;
    else
      raise Exception.Create(
        'UnKnown DBEngine Type of' + IntToStr(Integer(Adbet)));
  end;
end;

end.
