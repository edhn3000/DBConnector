unit U_MySqlTableInfo;

interface

uses
  ADODB, DB, Classes, Contnrs, Variants,
  U_MySqlFieldInfo, U_TableInfo, SysUtils, U_FieldInfo;

type
  TMySqlTableInfo = class(TTableInfo)
  protected
    procedure initObjects();override;
  public
    function GetNewFieldItem: TFieldItem;override;
    function InitTableInfo(tableName: string): Boolean;override;
  end;


implementation

{ TMySqlTableInfo }

function TMySqlTableInfo.GetNewFieldItem: TFieldItem;
begin
  Result := TMySQLFieldItem.Create;
end;

procedure TMySqlTableInfo.initObjects;
begin
  Fields := TMySQLFieldItemList.Create();
  FUniqueFields := TMySQLFieldItemList.Create();
  FPrimaryKeys := TMySQLFieldItemList.Create(False);
end;

function TMySqlTableInfo.InitTableInfo(tableName: string): Boolean;
begin
  Result := inherited InitTableInfo(tableName);
end;

end.
