{**
 * <p>Title: U_HashTable </p>
 * <p>Copyright: Copyright (c) 2007-11-20 </p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 哈希表</p>
 *}
unit U_HashTable;
   
interface

uses SysUtils, Classes;

type       
  TVariantObject = class
  public
    Value: Variant;
  public
    constructor Create(v: Variant);
  end;
  
  PPHashItem = ^PHashItem;
  PHashItem = ^THashItem;
  THashItem = record
    Next: PHashItem;
    Key: string;
    Value: TObject;
    Index: Integer;
  end;

{ THashTable }
  THashTable = class
  private
    FEnumIndex: Integer;
    FCount: Integer;
    Buckets: array of PHashItem;
    FFreeObjectOnRemove: Boolean;
    function getSize: Integer;
    procedure setSize(value: Integer);
  protected
    function Find(const Key: string): PPHashItem;
    function HashOf(const Key: string): Cardinal; virtual;
    function Modify(const Key: string; Value: TObject): Boolean;
  public
    property Count: Integer read FCount;
    property Size: Integer read getSize write setSize;
    property FreeObjectOnRemove: Boolean read FFreeObjectOnRemove;
  public
    constructor Create(InitialSize: Integer = 256);
    destructor Destroy; override;
    procedure Put(const Key: string; Value: TObject);virtual;
    function Get(const Key: string): TObject;virtual;
    procedure Clear;virtual;
    procedure Remove(const Key: string);virtual;

    // 用于按照顺序列举
    procedure ResetEnum;
    function EnumNext: TObject;
  end;

  { 元素有顺序的Map }
  TListMap = class(THashTable)
  private
    FList: TStringList;
  protected         
    function GetKey(index: Integer): String;
    function GetValue(index: Integer): TObject;
    procedure SetValue(index: Integer; value: TObject);
    function GetCount: Integer;
  public
    property Keys[index: Integer]: String read GetKey;
    property Values[index: Integer]: TObject read GetValue write SetValue;
    property Count: Integer read GetCount;
  public
    constructor Create();
    destructor Destroy();override;

    procedure Sort; virtual;
    procedure CustomSort(Compare: TStringListSortCompare); virtual;

    function Add(const S: string): Integer;
    function AddObject(const S: string; AObject: TObject): Integer;
    procedure Put(const Key: string; Value: TObject);override;
    function Get(const Key: string): TObject;override;
    procedure Clear;override;
    procedure Remove(const Key: string);override;
  end;
   
implementation

{ TVariantObject }

constructor TVariantObject.Create(v: Variant);
begin
  Value := v;
end;
 
{ THashTable }  

 { private }
function THashTable.getSize: Integer;
begin
  Result := Length(Buckets);
end;

procedure THashTable.setSize(value: Integer);
begin
  SetLength(Buckets, value)
end;

 { protected }
function THashTable.Find(const Key: string): PPHashItem;
var
  Hash: Integer;
begin
  Hash := HashOf(Key) mod Cardinal(Length(Buckets));
  Result := @Buckets[Hash];
  while Result^ <> nil do
  begin
  if Result^.Key = Key then
    Exit
  else
    Result := @Result^.Next;
  end;
end;
   
function THashTable.HashOf(const Key: string): Cardinal;
var
  I: Integer;
begin
  Result := 0;
  for I := 1 to Length(Key) do
  Result := ((Result shl 2) or (Result shr (SizeOf(Result) * 8 - 2))) xor
    Ord(Key[I]);
end;
   
function THashTable.Modify(const Key: string; Value: TObject): Boolean;
var
  P: PHashItem;
begin
  P := Find(Key)^;
  if P <> nil then
  begin
    Result := True;
    P^.Value := Value;
  end
  else
    Result := False;
end;

 { public }  
constructor THashTable.Create(InitialSize: Integer);
begin
  inherited Create;
  FCount := 0;
  FFreeObjectOnRemove := False;
  Self.Size := InitialSize;
end;
   
destructor THashTable.Destroy;
begin
  Clear;
  inherited;
end;

procedure THashTable.Put(const Key: string; Value: TObject);
var
  Hash: Integer;
  Bucket: PHashItem;
begin
  if not Modify(Key, Value) then
  begin
    if FCount = Size-1 then
      SetLength(Buckets, Size*2+1);
    Hash := HashOf(Key) mod Cardinal(Length(Buckets));
    New(Bucket);
    Bucket^.Key := Key;
    Bucket^.Value := Value;
    Bucket^.Index := Hash;
    Bucket^.Next := Buckets[Hash];
    Buckets[Hash] := Bucket;
    Inc(FCount);
  end;
end; 
   
procedure THashTable.Clear;
var
  I: Integer;
  P, N: PHashItem;
begin
  for I := 0 to Length(Buckets) - 1 do
  begin
  P := Buckets[I];
  while P <> nil do
  begin
    N := P^.Next;
    if Assigned(p^.Value) then
      if FFreeObjectOnRemove then
        FreeAndNil(p^.Value)
      else
        p^.Value := nil;
    Dispose(P);
    P := N;
  end;
  Buckets[I] := nil;
  end;
  FCount := 0;
end;    
   
procedure THashTable.Remove(const Key: string);
var
  P: PHashItem;
  Prev: PPHashItem;
begin
  Prev := Find(Key);
  P := Prev^;
  if P <> nil then
  begin
    Prev^ := P^.Next;
    if Assigned(p^.Value) then
      if FFreeObjectOnRemove then
        FreeAndNil(p^.Value)
      else
        p^.Value := nil;
    Dispose(P);
    Dec(FCount);
  end;
end;
   
function THashTable.Get(const Key: string): TObject;
var
  P: PHashItem;
begin
  P := Find(Key)^;
  if P <> nil then
    Result := P^.Value
  else
    Result := nil;
end;

procedure THashTable.ResetEnum;
begin
  FEnumIndex := 0;
end;

function THashTable.EnumNext: TObject;
var
  i: Integer;
begin
  Result := nil;
  for i := FEnumIndex+1 to High(Buckets) do
  begin
    if Buckets[i] <> nil then
    begin
      Result := Buckets[i]^.Value;
      FEnumIndex := i;
      Break;
    end;
  end;
end;

{ TListMap }

procedure TListMap.CustomSort(Compare: TStringListSortCompare);
begin
  FList.CustomSort(Compare);
end;

procedure TListMap.Sort;
begin
  FList.Sort;
end;

function TListMap.Add(const S: string): Integer;
begin
  Result := Self.AddObject(S, TObject(S));
end;

function TListMap.AddObject(const S: string;
  AObject: TObject): Integer;
var
  i: Integer;
begin
  if Get(S) = nil then
  begin
    i := FList.AddObject(S, AObject);
  end
  else
  begin
    i := FList.IndexOf(S);
    if i <> -1 then
      FList.Objects[i] := AObject
    else
      i := FList.AddObject(S, AObject);
  end;
  Result := i;
  Put(S, AObject);
end;

procedure TListMap.Clear;
begin        
  inherited Clear;
  FList.Clear;
end;

constructor TListMap.Create;
begin
  inherited Create;
  FFreeObjectOnRemove := False;
  FList := TStringList.Create;
end;

destructor TListMap.Destroy;
begin
  FList.Clear;
  inherited;
end;

function TListMap.GetValue(index: Integer): TObject;
begin
  Result := FList.Objects[index];
end;

procedure TListMap.SetValue(index: Integer; value: TObject);
begin
  FList.Objects[index] := value;
end;

function TListMap.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TListMap.Get(const Key: string): TObject;
begin
  Result := inherited Get(Key);
end;

procedure TListMap.Put(const Key: string; Value: TObject);
var
  n: Integer;
begin
  inherited;   
  n := FList.IndexOf(Key);
  if n = - 1 then
    FList.AddObject(Key, Value)
  else
    FList.Objects[n] := Value;
end;

procedure TListMap.Remove(const Key: string);
var
  n: Integer;
begin
  inherited;
  n := FList.IndexOf(Key);
  if n <> - 1 then
    FList.Delete(n);
end;

function TListMap.GetKey(index: Integer): String;
begin
  Result := FList.Strings[index];
end;

end.
