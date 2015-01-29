unit U_AccessTableInfo;

interface

uses
  ADODB, DB, Classes, Contnrs, Variants, SysUtils,
  U_AccessFieldInfo, U_TableInfo, U_FieldInfo;

type
  TAccessTableInfo = class(TTableInfo)
  protected
    procedure initObjects();override;
  public
    function GetNewFieldItem: TFieldItem;override;
    function InitTableInfo(tableName: string): Boolean;override;
  end;

implementation

{ TAccessTableInfo }

function TAccessTableInfo.GetNewFieldItem: TFieldItem;
begin
  Result := TAccessFieldItem.Create;
end;

procedure TAccessTableInfo.initObjects;
begin
  Fields := TAccessFieldItemList.Create();
  FUniqueFields := TAccessFieldItemList.Create();
  FPrimaryKeys := TAccessFieldItemList.Create(False);
end;

function TAccessTableInfo.InitTableInfo(tableName: string): Boolean;
begin
  Result := inherited InitTableInfo(tableName);
end;

end.
