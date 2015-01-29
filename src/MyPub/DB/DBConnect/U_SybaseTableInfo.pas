unit U_SybaseTableInfo;

interface

uses
  ADODB, DB, Classes, Contnrs, Variants, SysUtils,
  U_SybaseFieldInfo, U_TableInfo, U_FieldInfo;

type
  TSybaseTableInfo = class(TTableInfo)
  protected
    procedure initObjects();override;
  public
    function GetNewFieldItem: TFieldItem;override;
    function InitTableInfo(tableName: string): Boolean;override;
  end;

implementation

{ TSybaseTableInfo }

function TSybaseTableInfo.GetNewFieldItem: TFieldItem;
begin
  Result := TSybaseFieldItem.Create;
end;

procedure TSybaseTableInfo.initObjects;
begin   
  Fields := TSybaseFieldItemList.Create();
  FUniqueFields := TSybaseFieldItemList.Create();
  FPrimaryKeys := TSybaseFieldItemList.Create(False);
end;

function TSybaseTableInfo.InitTableInfo(tableName: string): Boolean;
begin
  Result := inherited InitTableInfo(tableName);
end;

end.
