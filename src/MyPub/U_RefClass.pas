unit U_RefClass;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  ActiveX, Contnrs, TypInfo;

const
  MaxCustomListSize = 40000000;

type
   TSetPropertyProc = procedure (const Value: Variant) of object;
   TGetPropertyFunction = function : Variant of object;

   TCustomList = class;

   PCustomItem = ^TCustomItem;
   TCustomItem = record
     FString: string;
     FVariant: Variant;
   end;                                       

  PCustomItemList = ^TCustomItemList;
  TCustomItemList = array[0..MaxCustomListSize] of TCustomItem;

  TCustomList = class
  private
    FList: PCustomItemList;

    function GetName(index: Integer): string;
    function Get(key: string):Variant;
    procedure Put(key: string; value: Variant);
    function GetItem(index: Integer):Variant;
    procedure SetItem(index: Integer; value: Variant);
  public
    property Names[Index: Integer]: string read GetName;
    property List[key: string]: Variant read Get write Put ;default;
    property Items[index:Integer]: Variant read GetItem write SetItem;
  public
    constructor Create;
    destructor Destroy;override;

    procedure Add(S: string; V: Variant); overload;
    function IndexOf(sKey: string): Integer;
  end;

{ TRefClass }

  TRefClass = class(TInterfacedObject, IDispatch)
  private
    FOwner : TComponent;

    function CallFunction(ProcAddress: Pointer; var Params: TDispParams): variant;
  protected
    FList: TCustomList;
    function GetFieldValue(DataField: String): Variant; virtual;
    procedure SetFieldValue(DataField: String; Value: Variant); virtual;
    function ClassMapp: variant;
    procedure AddProperty(Attribute: String; Default: variant);
    function GetProperty(Attribute: String): variant;
    function FindProperty(DataField: String): boolean;
    procedure AddProcedure(ProName: string; Address: Pointer; OfObject: TObject);

    {IDispatch}
    function GetTypeInfoCount(out Count: Integer): HResult; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; stdcall;
    function GetIDsOfNames(const IID: TGUID; Names: Pointer;
      NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; virtual; stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; stdcall;

  public
    constructor Create(AOwner: TComponent); virtual;
    destructor Destroy; override;
    class function CreateAsAutoObject(AOwner: TComponent): variant;
  end;

  function PointerToVariant(P: Pointer):Variant;
  function VariantToPointer(V: Variant):Pointer;

implementation

function PointerToVariant(P: Pointer):Variant;
begin
  Result := Variant(P^);
end;

function VariantToPointer(V: Variant):Pointer;
begin
  Result := Pointer(@V);
end;

{ TRefClass }
procedure TRefClass.AddProcedure(ProName: string; Address: Pointer;
  OfObject: TObject);
var
  pMethod: ^TSetPropertyProc;
begin
  pMethod := AllocMem(SizeOf(TSetPropertyProc));
  TMethod(pMethod^).Data := OfObject;
  TMethod(pMethod^).Code := Address;

  FList.Add(ProName, PointerToVariant(pMethod));
end;

procedure TRefClass.AddProperty(Attribute: String; Default: variant);
begin
  FList.Add(Attribute, Default);
end;
function TRefClass.CallFunction(ProcAddress: Pointer;
  var Params: TDispParams): variant;

  function GetParamAdrs(Value: Variant): Pointer;
  var
    s: PString;
    k: PInteger;
    F: PExtended;
  begin
    case TVarData(Value).VType of
      varInteger, varSmallint, varSingle:
        begin
          New(k);
          K^ := Value;
          Result := K;
        end;
      varOleStr, varString:
        begin
          New(S);
          s^ := Value;
          Result := S;
        end;
      varDouble, varCurrency:
        begin
          New(F);
          F^ := Value;
          Result := F;
        end;
      else
        Result := nil;
    end;
  end;
var
  i: Integer;
  pParams: array of Pointer;
  p: TMethod;
  pCount: Integer; 
begin
  pCount := Params.cArgs;
  SetLength(pParams, pCount);
  FillChar(pParams[0], pCount*Sizeof(Pointer), 0);
  p := TMethod(ProcAddress^);

  for i:=pCount-1 downto 0 do
  begin
    pParams[pCount-1-i] := GetParamAdrs(Variant(DispParams(Params).rgvarg^[i]));
  end;

  asm
//    push eax
//    push ecx
//    push edx
//    //push p.Data
    cmp pCount, 1
    JB @exec
    JE @One
    cmp pCount, 2
    JE @two
    @ThreeUp:
      CLD
      mov ecx, pCount
      sub ecx, 2
      mov edx, 4
      add edx, 4
    @loop:
      mov eax, [pParams]
      mov eax, [eax]+edx
      mov eax, [eax]
      push eax
      add edx, 4
      Loop @loop
    @Two:
      mov ecx, [pParams]
      mov ecx, [ecx]+4
      mov ecx, [ecx]
    @One:
      mov edx, [pParams]
      mov edx, [edx]
      mov edx, [edx]
    @exec:
      mov eax, p.Data
      test eax, eax
      je @1
      jne @call
      @1:
        mov eax, edx
        mov edx, ecx
        pop ecx
        jmp @call
      @call:
        call P.Code
//    pop edx
//    pop ecx
//    pop eax
  end;

  for i:=0 to pCount-1 do
    Dispose(pParams[i]);
end;

function TRefClass.ClassMapp: variant;
begin
  result := Self as IDispatch;
end;

constructor TRefClass.Create(AOwner: TComponent);
begin
  inherited Create;
  FOwner := AOwner;
  FList := TCustomList.Create();
end;

class function TRefClass.CreateAsAutoObject(AOwner: TComponent): variant;
begin
  result := Create(AOwner).ClassMapp;
end;

destructor TRefClass.Destroy;
begin
  FList.free;
  inherited;
end;

function TRefClass.FindProperty(DataField: String): boolean;
begin
  result := false;
end;

function TRefClass.GetFieldValue(DataField: String): Variant;
var
  P: PPropInfo;
  V: variant;
  TypeData: PTypeData;
  PFunction: ^TGetPropertyFunction;
  k: integer;
begin
  V := FList[DataField];
  if TVarData(V).Reserved1=1 then
  begin
    P := VariantToPointer(V);
    case P^.PropType^.Kind of
      tkInteger, tkChar, tkWChar, tkClass:
        result := GetOrdProp(FOwner, P);
      tkEnumeration:
        begin
          TypeData := GetTypeData(P^.PropType^);
          if TypeData^.BaseType^ = TypeInfo(Boolean) then
            Result := Boolean(GetOrdProp(FOwner, P))
          else
            Result := GetOrdProp(FOwner, P);
        end;
      tkFloat:
        Result := GetFloatProp(FOwner, P);
      tkString, tkLString, tkWString:
        Result := GetStrProp(FOwner, P);
      tkSet:
        Result := GetSetProp(FOwner, P);
      tkMethod:
        Result := P^.PropType^.Name;
      tkVariant:
        Result := GetVariantProp(FOwner, P);
      tkInt64:
        Result := GetInt64Prop(FOwner, P) + 0.0;
      tkDynArray,
      tkArray,
      tkRecord,
      tkInterface:;
    end;
  end
  else
  begin
    k := FList.IndexOf(Format('Get%s', [DataField]));
    if k<>-1 then
    begin
      PFunction := VariantToPointer(FList.Items[k]);
      V := PFunction^;
    end;
    result := V;
  end;
end;

function TRefClass.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;
const
  FUNCLIST: array[1..4] of String =
    ('AddProperty', 'GetProperty', 'FindProperty', 'AddProcedure');
var
  S: String;
  DispID: integer;
  i: integer;
begin
  s := WideString(POleStrList(Names)^[0]);
  DispID := 0;
  for i := 1 to 4 do
    if CompareText(S, FUNCLIST[i])=0 then
    begin
      DispID := -1*i;
      break;
    end;
  if DispID = 0 then
  begin
    DispID := FList.IndexOf(S);

    if DispID = -1 then
    begin
      result := E_NOTIMPL;
      exit;
    end;
  end;
  PDispIdList(DispIDs)^[0] := DispID;
  result := S_OK;
end;

function TRefClass.GetProperty(Attribute: String): variant;
begin
  result := GetFieldValue(Attribute);
end;

function TRefClass.GetTypeInfo(Index, LocaleID: Integer;
  out TypeInfo): HResult;
begin
  result := S_OK;
end;

function TRefClass.GetTypeInfoCount(out Count: Integer): HResult;
begin
  result := S_OK;
end;

function TRefClass.Invoke(DispID: Integer; const IID: TGUID;
  LocaleID: Integer; Flags: Word; var Params; VarResult, ExcepInfo,
  ArgErr: Pointer): HResult;
Var
  V: Variant;
  P: Pointer;
begin
  case DispID of
    -1:
      begin
        if PDispParams(@Params)^.cArgs=1 then
          AddProperty(OleVariant(PDispParams(@Params).rgvarg^[0]), NULL)
        else
          AddProperty(OleVariant(PDispParams(@Params).rgvarg^[1]), OleVariant(PDispParams(@Params).rgvarg^[0]))
      end;
    -2:
      V := FindProperty(OleVariant(PDispParams(@Params).rgvarg^[0]));
    -3:
      V := GetProperty(OleVariant(PDispParams(@Params).rgvarg^[0]));
    -4:
      if PDispParams(@Params)^.cArgs=2 then
        AddProcedure(OleVariant(PDispParams(@Params).rgvarg^[1]),
          VariantToPointer(OleVariant(PDispParams(@Params).rgvarg^[0])), nil)
      else if PDispParams(@Params)^.cArgs=3 then
        AddProcedure(OleVariant(PDispParams(@Params).rgvarg^[2]),
          VariantToPointer(OleVariant(PDispParams(@Params).rgvarg^[1])),
          VariantToPointer(OleVariant(PDispParams(@Params).rgvarg^[0])));
    else
    begin
      if (Flags and DISPATCH_PROPERTYGET) = DISPATCH_PROPERTYGET then
        V := GetFieldValue(FList.Names[DispID])
      else if (Flags and DISPATCH_PROPERTYPUT) = DISPATCH_PROPERTYPUT then
        SetFieldValue(FList.Names[DispID], OleVariant(PDispParams(@Params).rgvarg^[0]))
      else
      begin
        V := FList.Items[DispID];
        P := VariantToPointer(V);
        V := CallFunction(P, TDispParams(Params));
      end;
    end;
  end;
  if assigned(VarResult) then
    PVariant(VarResult)^ := V;
  result := S_OK;
end;

procedure TRefClass.SetFieldValue(DataField: String; Value: Variant);
  function RangedValue(const AMin, AMax: Int64): Int64;
  begin
    Result := Trunc(Value);
    if Result < AMin then
      Result := AMin;
    if Result > AMax then
      Result := AMax;
  end;
var
  PropInfo: PPropInfo;
  TypeData: PTypeData;
  V: variant;
  PProcedure: ^TSetPropertyProc;
  k: integer;
begin
  V := FList[DataField];

  if TVarData(V).Reserved1=1 then
  begin
    PropInfo := VariantToPointer(V);

    if PropInfo <> nil then
    begin
      case PropInfo.PropType^^.Kind of
        tkInteger, tkChar, tkWChar:
          begin
            TypeData := GetTypeData(PropInfo^.PropType^);
            SetOrdProp(FOwner, PropInfo, RangedValue(TypeData^.MinValue, TypeData^.MaxValue));
          end;
        tkEnumeration:
          if VarType(Value) = varString then
            SetEnumProp(FOwner, PropInfo, VarToStr(Value))
          else
          begin
            TypeData := GetTypeData(PropInfo^.PropType^);
            SetOrdProp(FOwner, PropInfo, RangedValue(TypeData^.MinValue, TypeData^.MaxValue));
          end;
        tkSet:
          if VarType(Value) = varInteger then
            SetOrdProp(FOwner, PropInfo, Value)
          else
            SetSetProp(FOwner, PropInfo, VarToStr(Value));
        tkFloat:
          SetFloatProp(FOwner, PropInfo, Value);
        tkString, tkLString, tkWString:
          SetStrProp(FOwner, PropInfo, VarToStr(Value));
        tkVariant:
          SetVariantProp(FOwner, PropInfo, Value);
        tkInt64:
          begin
            TypeData := GetTypeData(PropInfo^.PropType^);
            SetInt64Prop(FOwner, PropInfo, RangedValue(TypeData^.MinInt64Value, TypeData^.MaxInt64Value));
          end;
      else
      end;
    end;
  end
  else
  begin
    k := FList.IndexOf(Format('Set%s', [DataField]));
    if k<>-1 then
    begin
      PProcedure := VariantToPointer(FList.Items[k]);
      PProcedure^(Value);
    end;
    FList.Add(DataField, Value);
  end;
end;

{ TCustomList }

procedure TCustomList.Add(S: string; V: Variant);
begin
  Put(S, V);
end;

constructor TCustomList.Create;
begin

end;

destructor TCustomList.Destroy;
begin

  inherited;
end;

function TCustomList.Get(key: string): Variant;
var
  theindex: Integer;
begin
  theindex := IndexOf(Key);
  if theindex = -1 then
    Result := null
  else
    Result := FList^[theindex].FVariant;
end;

function TCustomList.GetItem(index: Integer): Variant;
begin
  Result := FList^[Index].FVariant;
end;

function TCustomList.GetName(index: Integer): string;
begin
  Result := FList^[index].FString;
end;

function TCustomList.IndexOf(sKey: string): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := Low(FList^) to High(FList^) do
  begin
    if FList^[i].FString = '' then
      Break;
    if FList^[i].FString = sKey then
    begin
      Result := i;
      Break;
    end;
  end;
end;

procedure TCustomList.Put(key: string; value: Variant);
var
  i, nFind: Integer;
begin
  nFind := IndexOf(key);
  if nFind <> -1 then
    FList^[nFind].FVariant := value
  else
  begin
    for i := Low(FList^) to High(FList^) do
    begin
      if FList^[i].FString = '' then
      begin
        FList^[i].FString := value;
      end;
    end;
  end;
end;

procedure TCustomList.SetItem(index: Integer; value: Variant);
begin
  FList^[index].FString := value;
end;

end.

