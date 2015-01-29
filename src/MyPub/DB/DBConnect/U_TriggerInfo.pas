unit U_TriggerInfo;

interface

uses
  Contnrs;

type
  TTriggerInfo = class
  public
    Name: string;
    TableName: string;
    Owner: string;
  end;

  TTriggerInfoList = class(TObjectList)
  protected        
    FSystemObject: Boolean;
    function GetItem(Index: Integer): TTriggerInfo;
    procedure SetItem(Index: Integer; const Value: TTriggerInfo);

  public
    property Items[Index: Integer]: TTriggerInfo read GetItem write SetItem;  
    property SystemObject: Boolean read FSystemObject write FSystemObject;

  end;

implementation

{ TTriggerInfoList }

function TTriggerInfoList.GetItem(Index: Integer): TTriggerInfo;
begin
  Result := inherited GetItem(Index) as TTriggerInfo;
end;

procedure TTriggerInfoList.SetItem(Index: Integer;
  const Value: TTriggerInfo);
begin
  inherited SetItem(Index, Value);
end;

end.
