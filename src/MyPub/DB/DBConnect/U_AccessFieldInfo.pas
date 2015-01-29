unit U_AccessFieldInfo;

interface         

uses
  ADODB, DB, Classes, Contnrs, Variants,
  U_FieldInfo, SysUtils;

type
  TAccessFieldItem = class (TFieldItem)
  public
    function GetDataTypeInSQL: string;override;
    procedure InitByField(AField: TField);override;
  end;   
        
  TAccessFieldItemList = class(TFieldItemList)
  public
    function CreateFieldItem: TFieldItem;override;

  end;

implementation 

{ TAccessFieldItem }

function TAccessFieldItem.GetDataTypeInSQL: string;
begin
  Result := inherited GetDataTypeInSQL;
end;

procedure TAccessFieldItem.InitByField(AField: TField);
begin
  inherited;
  if AField.DataType in [ftAutoInc] then
    Self.IsUnique := True;

  // accessû��ֱ�ӻ���Ƿ������İ취��һ��readonly������
  Self.IsPrimary := AField.ReadOnly;
end;

{ TAccessFieldItemList }

function TAccessFieldItemList.CreateFieldItem: TFieldItem;
begin
  Result := TAccessFieldItem.Create;
end;

end.
