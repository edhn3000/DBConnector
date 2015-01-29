unit U_MySqlFieldInfo;

interface       

uses
  ADODB, DB, Classes, Contnrs, Variants,
  U_FieldInfo, SysUtils;

type
  TMySQLFieldItem = class (TFieldItem)
  public
    function GetDataTypeInSQL: string;override;
  end;    

  TMySQLFieldItemList = class(TFieldItemList)
  public
    function CreateFieldItem: TFieldItem;override;

  end;

implementation  

{ TMySQLFieldItem }

function TMySQLFieldItem.GetDataTypeInSQL: string;
begin
  Result := inherited GetDataTypeInSQL;
end;


{ TMySQLFieldItemList }

function TMySQLFieldItemList.CreateFieldItem: TFieldItem;
begin
  Result := TMySQLFieldItem.Create;
end;

end.
