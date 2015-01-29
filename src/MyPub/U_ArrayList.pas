{**
 * <p>Title: U_ArrayList </p>
 * <p>Copyright: Copyright (c) 2007-11-26 </p>
 * @author fengyq
 * @version 1.0
 * <p>Description: ArrayList</p>
 *}
unit U_ArrayList;

interface
               
uses SysUtils, Classes, SyncObjs, Windows;

type
  TArrayList = class(TList)
  private
    FOwnsObjects: Boolean;
  protected
    procedure Notify(Ptr: Pointer; Action: TListNotification); override;
    function GetItem(Index: Integer): TObject;
    procedure SetItem(Index: Integer; AObject: TObject);
  public
    constructor Create; overload;
    constructor Create(AOwnsObjects: Boolean); overload;

    function Add(AObject: TObject): Integer;
    function Extract(Item: TObject): TObject;
    function Remove(AObject: TObject): Integer;
    function IndexOf(AObject: TObject): Integer;
    function FindInstanceOf(AClass: TClass; AExact: Boolean = True; AStartAt: Integer = 0): Integer;
    procedure Insert(Index: Integer; AObject: TObject);
    function First: TObject;
    function Last: TObject;
    property OwnsObjects: Boolean read FOwnsObjects write FOwnsObjects;
    property Items[Index: Integer]: TObject read GetItem write SetItem; default;
  end;

  TLockArrayList = class(TArrayList)
  private
    FLock: TRTLCriticalSection;
  public
    constructor Create;virtual;
    destructor Destroy;override;

    procedure Lock;
    procedure UnLock;
  end;

implementation

{ TArrayList }

function TArrayList.Add(AObject: TObject): Integer;
begin
  Result := inherited Add(AObject);
end;

constructor TArrayList.Create;
begin
  inherited Create;
  FOwnsObjects := True;
end;

constructor TArrayList.Create(AOwnsObjects: Boolean);
begin
  inherited Create;
  FOwnsObjects := AOwnsObjects;
end;

function TArrayList.Extract(Item: TObject): TObject;
begin
  Result := TObject(inherited Extract(Item));
end;

function TArrayList.FindInstanceOf(AClass: TClass; AExact: Boolean;
  AStartAt: Integer): Integer;
var
  I: Integer;
begin
  Result := -1;
  for I := AStartAt to Count - 1 do
    if (AExact and
        (Items[I].ClassType = AClass)) or
       (not AExact and
        Items[I].InheritsFrom(AClass)) then
    begin
      Result := I;
      break;
    end;
end;

function TArrayList.First: TObject;
begin
  Result := TObject(inherited First);
end;

function TArrayList.GetItem(Index: Integer): TObject;
begin
  Result := inherited Items[Index];
end;

function TArrayList.IndexOf(AObject: TObject): Integer;
begin
  Result := inherited IndexOf(AObject);
end;

procedure TArrayList.Insert(Index: Integer; AObject: TObject);
begin
  inherited Insert(Index, AObject);
end;

function TArrayList.Last: TObject;
begin
  Result := TObject(inherited Last);
end;

procedure TArrayList.Notify(Ptr: Pointer; Action: TListNotification);
begin
  if OwnsObjects then
//    if Action = lnDeleted then
//      TObject(Ptr).Free;
  inherited Notify(Ptr, Action);
end;

function TArrayList.Remove(AObject: TObject): Integer;
begin
  Result := inherited Remove(AObject);
end;

procedure TArrayList.SetItem(Index: Integer; AObject: TObject);
begin
  inherited Items[Index] := AObject;
end;  

{ TLockArrayList }

constructor TLockArrayList.Create;
begin
  InitializeCriticalSection(FLock);
end;

destructor TLockArrayList.Destroy;
begin
  Lock;
  try
    inherited Destroy;
  finally
    UnLock;
    DeleteCriticalSection(FLock);
  end;
end;

procedure TLockArrayList.Lock;
begin
  EnterCriticalSection(FLock);
end;

procedure TLockArrayList.UnLock;
begin
  LeaveCriticalSection(FLock);
end;

end.
