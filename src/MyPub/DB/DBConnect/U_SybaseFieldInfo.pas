unit U_SybaseFieldInfo;

interface   

uses
  ADODB, DB, Classes, Contnrs, Variants,
  U_FieldInfo, SysUtils;

type
  TSybaseFieldItem = class (TFieldItem)
  public
    function GetDataTypeInSQL: string;override;
  end; 

  TSybaseFieldItemList = class(TFieldItemList)
  public
    function CreateFieldItem: TFieldItem;override;

  end;

implementation

{ TSybaseFieldItem }

function TSybaseFieldItem.GetDataTypeInSQL: string;
begin
  Result := inherited GetDataTypeInSQL;
end;

{ TSybaseFieldItemList }

function TSybaseFieldItemList.CreateFieldItem: TFieldItem;
begin
  Result := TSybaseFieldItem.Create;
end;

end.
