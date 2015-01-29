unit U_IndexInfo;

interface
uses
  Contnrs;

type
  TIndexInfo = class
  public
    Name: string;
    TableName: string;
    Owner: string;
  end;   

  TIndexInfoList = class(TObjectList)
  protected        
    FSystemObject: Boolean;
    function GetItem(Index: Integer): TIndexInfo;
    procedure SetItem(Index: Integer; const Value: TIndexInfo);

  public
    property Items[Index: Integer]: TIndexInfo read GetItem write SetItem;  
    property SystemObject: Boolean read FSystemObject write FSystemObject;

  end;

implementation

{ TIndexInfoList }

function TIndexInfoList.GetItem(Index: Integer): TIndexInfo;
begin
  Result := inherited GetItem(Index) as TIndexInfo;
end;

procedure TIndexInfoList.SetItem(Index: Integer; const Value: TIndexInfo);
begin
  inherited SetItem(Index, Value);
end;

end.
