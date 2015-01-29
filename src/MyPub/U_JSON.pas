{**
 * <p>Title: U_JSON </p>
 * <p>Copyright: Copyright (c) 2009/2/13</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: DSON </p> 
 * <p>Description: 仿JSON的对象表达式解析，解析结果是保存在List中的键值对 </p>
 *}
 unit U_JSON;

interface

uses
  Classes, U_fStrUtil;


const
  JSON_BRACKET_LEFT = '{';
  JSON_BRACKET_LEFT_LENGTH = 1;
  JSON_BRACKET_RIGHT = '}';
  JSON_BRACKET_RIGHT_LENGTH = 1;
  JSON_SEPARATOR_ITEM = ',';
  JSON_SEPARATOR_KEYVALUE = ':';
  JSON_SEPARATOR_KEYVALUE_LENGTH = 1;
  JSON_VALUE_ISOBJECT = 'object';
  JSON_SEPARATOR_NAMELEVEL = '.';   // name1.name2.name3

type
  TJSON = class
  private
    FList: TStrings;
    
    function GetKey(index: Integer): string;
    function GetValue(index: Integer): string;
    procedure SetValue(index: Integer; value: string);
    function GetObject(index: Integer): TJSON;
    procedure SetObject(index: Integer; value: TJSON);
    function GetValueByName(key: string): string;
    procedure SetValueByName(key: string; value: string);
    function GetObjectByName(key: string): TJSON;
    procedure SetObjectByName(key: string; value: TJSON);
  public
    property Keys[index: Integer]: string read GetKey;
    property Values[index: Integer]: string read GetValue write SetValue;
    property Objects[index: Integer]: TJSON read GetObject write SetObject;
    property ValueByName[key: string]: string read GetValueByName write SetValueByName;
    property ObjectByName[key: string]: TJSON read GetObjectByName write SetObjectByName;
  public
    constructor Create();
    destructor Destroy();override;
    procedure Clear;
    function Init(const DSONStr: string): Boolean;
    function ValueIsObject(index: Integer): Boolean;overload;
    function ValueIsObject(key: string): Boolean;overload;

    function AddField(sKey, sValue: string): TJSON;
    function AddFieldObject(sKey: string; dson: TJSON): TJSON;
    function ToString(): string;
  end;

implementation

uses
  SysUtils;

{ TDSON }  

function TJSON.GetKey(index: Integer): string;
begin
  Result := FList.Names[index];
end;

function TJSON.GetValue(index: Integer): string;
begin
  Result := FList.Values[FList.Names[index]];
end;

function TJSON.GetValueByName(key: string): string;
var
  sKey, sKeyLeft: string;
  nPos: Integer;
begin
  nPos := Pos(JSON_SEPARATOR_NAMELEVEL, key);
  if nPos = 0 then
  begin
    if key = '' then
      Exit;
    if ValueIsObject(key) then
      Result := ObjectByName[key].ToString
    else
      Result := FList.Values[key];
  end
  else
  begin
    sKey := Copy(key, 1, nPos - 1);
    sKeyLeft := Copy(key, nPos + 1, MaxInt);
    if ValueIsObject(sKey) then
      Result := ObjectByName[sKey].GetValueByName(sKeyLeft)
    else
      Result := FList.Values[sKey];
  end;  
end;

procedure TJSON.SetValue(index: Integer; value: string);
begin
  FList.Values[FList.Names[index]] := value;
end;

procedure TJSON.SetValueByName(key, value: string);
begin     
  FList.Values[key] := value;
end;

function TJSON.GetObject(index: Integer): TJSON;
begin
  Result := FList.Objects[index] as TJSON;
end;

function TJSON.GetObjectByName(key: string): TJSON;
var
  nIndex: Integer;
begin
  nIndex := FList.IndexOfName(key);
  if nIndex <> -1 then
    Result := FList.Objects[nIndex] as TJSON
  else
    Result := nil;
end;

procedure TJSON.SetObject(index: Integer; value: TJSON);
begin
  FList.Objects[index] := value;
end;

procedure TJSON.SetObjectByName(key: string; value: TJSON);
var
  nIndex: Integer;
begin
  nIndex := FList.IndexOfName(key);
  if nIndex <> -1 then
    FList.Objects[nIndex] := value;
end;

constructor TJSON.Create;
begin
  FList := TStringList.Create;
end;

destructor TJSON.Destroy;
var
  i: Integer;
begin
  for i := FList.Count - 1 downto 0 do
  begin
    if FList.Objects[i] <> nil then
    begin
      TJSON(FList.Objects[i]).Free;
    end;  
  end;
  FList.Free;
  inherited;
end;    

function TJSON.AddField(sKey, sValue: string): TJSON;
begin
  FList.Add(sKey+'='+sValue);
  Result := Self;
end;

function TJSON.AddFieldObject(sKey: string; dson: TJSON): TJSON;
begin
  FList.AddObject(sKey+'='+JSON_VALUE_ISOBJECT, dson);
  Result := Self;
end;

procedure TJSON.Clear;
var
  i: Integer;
begin 
  for i := FList.Count - 1 downto 0 do
  begin
    if FList.Objects[i] <> nil then
    begin
      TJSON(FList.Objects[i]).Free;
    end;  
  end;
  FList.Clear;
end;

function TJSON.Init(const DSONStr: string): Boolean;
var
  sDSON: string;
  sItemFullStr: string;
  sOneItem: string;
  nIndex1, nIndex2: Integer;
  slstItems: TStrings;
  i: Integer;
  sKey, sValue: string;
  dson: TJSON;
begin
  Result := True;
  sDSON := Trim(DSONStr);
  // {开始的位置
  nIndex1 := fStrUtil.PosFrom(JSON_BRACKET_LEFT, sDSON);
  if nIndex1 = 0 then
    Exit;
  // 找到与开始的 "{" 对应的 "}"  因为可能有嵌套 
  nIndex2 := fStrUtil.ParseMatchedBracket(nIndex1,sDSON,
      JSON_BRACKET_LEFT,JSON_BRACKET_RIGHT);
  if nIndex2 = 0 then
    Exit;
  // 包含在{}中的整个DSON串的内容
  sItemFullStr := Copy(sDSON, nIndex1+JSON_BRACKET_LEFT_LENGTH,
      nIndex2-nIndex1-JSON_BRACKET_LEFT_LENGTH);
  if sItemFullStr = '' then
    Exit;
  slstItems := TStringList.Create;
  try
    // 根据分隔符","  拆分item
    fStrUtil.SplitEscapeBracket(sItemFullStr, JSON_SEPARATOR_ITEM,
      JSON_BRACKET_LEFT, JSON_BRACKET_RIGHT, slstItems);
    Clear;
    // 逐个分析item并添加  
    for i := 0 to slstItems.Count - 1 do
    begin
      sOneItem := slstItems[i];
      fStrUtil.StrToKeyValue(sOneItem, JSON_SEPARATOR_KEYVALUE, sKey, sValue);
      // 增加到列表
      if Copy(sValue, 1, JSON_BRACKET_LEFT_LENGTH) = JSON_BRACKET_LEFT then   // 对象
      begin
        dson := TJSON.Create;
        dson.Init(sValue);
        FList.AddObject(sKey + '=' + JSON_VALUE_ISOBJECT, dson);
      end  
      else                                                                    // 字符串
        FList.Add(sKey + '=' + sValue);
    end;
  finally
    slstItems.Free;
  end;
  Result := True;
end;

function TJSON.ValueIsObject(index: Integer): Boolean;
begin
  Result := (FList.Values[FList.Names[index]] = JSON_VALUE_ISOBJECT)
      and (FList.Objects[index] <> nil);
end;

function TJSON.ValueIsObject(key: string): Boolean;
var
  nIndex: Integer;
begin
  nIndex := FList.IndexOfName(key);
  if nIndex <> -1 then
    Result := (FList.Values[key] = JSON_VALUE_ISOBJECT)
        and (FList.Objects[nIndex] <> nil)
  else
    Result := False;
end;

function TJSON.ToString: string;
var
  i: Integer;
begin
  for i := 0 to FList.Count - 1 do
  begin
    if i = 0 then
    begin
      if ValueIsObject(i) then
        Result := Keys[i] + JSON_SEPARATOR_KEYVALUE + Objects[i].ToString
      else
        Result := Keys[i] + JSON_SEPARATOR_KEYVALUE + Values[i];
    end
    else
    begin
      if ValueIsObject(i) then
        Result := Result + JSON_SEPARATOR_ITEM + Keys[i]
          + JSON_SEPARATOR_KEYVALUE + Objects[i].ToString
      else
        Result := Result + JSON_SEPARATOR_ITEM + Keys[i]
          + JSON_SEPARATOR_KEYVALUE + Values[i];
    end; 
  end;   
  Result := JSON_BRACKET_LEFT + Result + JSON_BRACKET_RIGHT;
end;

end.
