{**
 * <p>Title: U_FieldInfo.pas  </p>
 * <p>Copyright: Copyright (c) 2009-4-12</p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: 自定义字段对象类，提高字段可操作性 </p>
 *}
unit U_FieldInfo;

interface  

uses
  ADODB, DB, Classes, Contnrs, Variants,
  SysUtils, U_DBEngineInterface;

type

{ TFieldItem 字段信息类 }
  TFieldItem = class
  protected
    FFieldName: string;
    FDataType: TFieldType;
    FDataTypeStr: string;
    FAsString: string;
    FIsPrimary: Boolean;  
    FIsUnique: Boolean;
    FNullable: Boolean;
    FDataSize: Integer;      // 就是长度
    FComment: string;
    FDecimal: Integer;       // 精度
    FConstraintName: string; // 约束名（Oracle定义主键时）
    
    FDB: IDBEngine;

    procedure SetFieldName(value: string);
    procedure SetFieldType(value: TFieldType);
    procedure SetAsString(value: string);
    procedure SetIsPrimary(value: Boolean);
    procedure SetDataTypeStr(value: string);   
    function GetDataTypeStr: string;

  public
    property FieldName: string read FFieldName write SetFieldName;
    property DataType: TFieldType read FDataType write SetFieldType;
    property DataTypeStr: string read GetDataTypeStr write SetDataTypeStr;
    property AsString: string read FAsString write SetAsString;
    property IsPrimary: Boolean read FIsPrimary write SetIsPrimary;   
    property IsUnique: Boolean read FIsUnique write FIsUnique;
    property Nullable: Boolean read FNullable write FNullable;
    property DataSize: Integer read FDataSize write FDataSize;   
    property Decimal: Integer read FDecimal write FDecimal;
    property Comment: string read FComment write FComment;
    property ConstraintName: string read FConstraintName write FConstraintName;
    
    property DB: IDBEngine read FDB write FDB ;
  public
    constructor Create();overload;
    constructor Create(AField: TField);overload;
    procedure InitByField(AField: TField);virtual;
    procedure Assign(source: TFieldItem);
    function GetDataTypeInSQL: string;virtual;
    function GetNullableStr: string;
    function IsSimilarField(AField: TFieldItem): Boolean;
  end;

{ TFieldItemList }
  TFieldItemList = class(TObjectList)
  private
    function GetItem(index: Integer): TFieldItem;
    procedure SetItem(index: Integer; value: TFieldItem);
  protected
    FDB: IDBEngine;
    FSystemObject: Boolean;
    function CreateFieldItemList: TFieldItemList; virtual;
  public
    constructor Create(Fields: TFields);overload;
    destructor Destroy;override;
    procedure InitByFields(Fields: TFields);               
    function GetPrimaryField(): TFieldItem;
    procedure SetPrimaryField(sFieldName: string);

    property Items[index: Integer]: TFieldItem read GetItem write SetItem;
    function FindFieldItem(sFieldName: string): TFieldItem;

    procedure FillPrimaryFieldList(fields: TFieldItemList);
    function IsSimilarFieldList(AFields: TFieldItemList): Boolean;
    // 获得第一个唯一字段
    function GetFirstUniqueField(): TFieldItem;
    function GetFieldNamesStr: string;
    procedure ReadValuesFromFields(fields: TFields);
    procedure AddFieldList(fields: TFieldItemList);   
    procedure SortByFieldName();

    // 初始化字段列表，不同数据库有不同初始化方式
    function InitFieldInfos(sTableName: string): Boolean;virtual;
    // 根据不同数据库创建不同的FieldItem
    function CreateFieldItem: TFieldItem; virtual;     
    function Clone: TFieldItemList;    

  public
    property DB: IDBEngine read FDB write FDB;
    property SystemObject: Boolean read FSystemObject write FSystemObject;
  end;

  function FieldTypeToStr(Aft: TFieldType; bWithPrefix: Boolean = True): string;
  function StrToFieldType(s: string): TFieldType;

implementation

uses
  TypInfo, U_fStrUtil;

function CompareFieldItemName(Item1, Item2: Pointer): Integer;
begin
  if TFieldItem(item1).FieldName = TFieldItem(item2).FieldName then
    Result := 0
  else if TFieldItem(item1).FieldName < TFieldItem(item2).FieldName then
    Result := -1
  else
    Result := 1;
end;  

function FieldTypeToStr(Aft: TFieldType; bWithPrefix: Boolean = True): string;
var
  ti: PTypeInfo;
  td: PTypeData;
  i: Integer;
begin
  ti := TypeInfo(TFieldType);
  td := GetTypeData(ti);
  for i := td^.MinValue to td^.MaxValue do
    if Integer(Aft) = i then
    begin
      Result := GetEnumName(ti, i);
      if not bWithPrefix then
        Result := Copy(Result, 3, Length(Result));
      Break;
    end;
end;
//var
//  ti: PTypeInfo;
//  i: TFieldType;
//begin
//  ti := TypeInfo(TFieldType);
//  for i := Low(TFieldType) to High(TFieldType) do
//  begin
//    if Aft = i then
//    begin
//      Result := GetEnumName(ti, Integer(i));
//      Break;
//    end;
//  end;
//end;

function StrToFieldType(s: string): TFieldType;
var
  nFT, nCode: Integer;
  ti: PTypeInfo;
  td: PTypeData;
  ft: TFieldType;
begin  
  Result := ftUnknown;
  Val(s, nFT, nCode);
  if (nCode = 0) then
  begin
    ti := TypeInfo(TFieldType);
    td := GetTypeData(ti);
    if (nFT <= td^.MaxValue) and (nFT >= td^.MinValue) then
      Result := TFieldType(nFT);
  end
  else
  begin
    for ft := Low(TFieldType) to High(TFieldType) do
      if SameText(FieldTypeToStr(ft, False), s) then
      begin
        Result := ft;
        Break;
      end;  
  end;

end;

{ TFieldItem }

constructor TFieldItem.Create();
begin        
  inherited Create;
  IsPrimary := False;
  DataType := ftString;
  FAsString := '';
end;      

constructor TFieldItem.Create(AField: TField);
begin
  Self.Create;
  Self.InitByField(AField);
end;

procedure TFieldItem.Assign(source: TFieldItem);
begin
  Self.FieldName  := source.FieldName;
  Self.DataType   := source.DataType;
  Self.AsString   := source.AsString;
  Self.DataSize   := source.DataSize;
  Self.Nullable   := source.Nullable;
  Self.FIsUnique  := source.IsUnique;
  Self.FIsPrimary := source.IsPrimary;
  Self.FDecimal   := source.Decimal;
  Self.FComment   := source.Comment;
end;

procedure TFieldItem.InitByField(AField: TField);
begin
  Self.FieldName := AField.FieldName;
  Self.DataType  := AField.DataType;
  Self.AsString  := AField.AsString;
  if AField.Size <> 0 then
    Self.DataSize  := AField.Size
  else
    Self.DataSize  := AField.DataSize;
  Self.Nullable  := not AField.Required;
end;

function TFieldItem.GetDataTypeInSQL(): string;
begin
  case DataType of
  ftString,ftWideString: Result := Format('Varchar(%d)', [FDataSize]);
  ftSmallint,ftInteger,ftBytes: Result := 'Int';
  ftLargeint:Result := 'Long';
  ftBoolean: Result := 'YesNo';
  ftDate: Result := 'Date';
  ftTime,ftDateTime,ftTimeStamp: Result := 'DateTime';
  ftAutoInc: Result := 'AutoIncrement';
  ftMemo,ftFmtMemo: Result := 'Text';
  ftFloat:   Result := 'Float';
  else
    Result := FieldTypeToStr(DataType, False);
  end;
end;

procedure TFieldItem.SetAsString(value: string);
begin
  FAsString := value;
end;

procedure TFieldItem.SetFieldType(value: TFieldType);
begin
  FDataType := value;
  if FDataType = ftAutoInc then
  begin
    FIsUnique := True;
  end;  
end;

procedure TFieldItem.SetFieldName(value: string);
begin
  FFieldName := value;
end;

procedure TFieldItem.SetIsPrimary(value: Boolean);
begin
  FIsPrimary := value;
end;      

function TFieldItem.GetDataTypeStr: string;
begin
  case DataType of
    ftString,ftWideString: Result := Format('String,%d', [FDataSize]);
    ftSmallint,ftInteger,ftBytes, ftLargeint:
    begin
      if FDataSize <> 0 then
        Result := Format('Integer,%d', [FDataSize])
      else
        Result := 'Integer';
    end;
    ftBoolean: Result := 'YesNo';
    ftDate: Result := 'Date';
    ftTime,ftDateTime,ftTimeStamp: Result := 'datetime';
    ftAutoInc: Result := 'AutoIncrement';
    // VER180=Delphi2007
    ftMemo,ftFmtMemo{$IFDEF VER180}, ftWideMemo{$ENDIF}: Result := 'Text';
    ftFloat:   Result := Format('Float,%d,%d', [FDataSize, FDecimal]);
    ftOraClob: Result := 'Clob';
    ftOraBlob: Result := 'Blob';
    else
      Result := 'unknown'
  end;
end;

function TFieldItem.GetNullableStr: string;
begin
  if Nullable then
    Result := ''
  else
    Result := 'Not Null'
end;

procedure TFieldItem.SetDataTypeStr(value: string);
var
  slst: TStrings;
  nMatch: Integer;
begin
  FDataTypeStr := value;
  slst := TStringList.Create;
  fStrUtil.Split(value, ',', slst);
  try
    if fStrUtil.PosArrayFrom(value, ['string', 'CHAR', 'VARCHAR2'],
      nMatch, 1, True) > 0 then
    begin
      FDataType := ftString;
      if slst.Count >= 2 then
        FDataSize := StrToInt(slst[1]);
    end  
    else if fStrUtil.PosArrayFrom(value, ['integer', 'NUMBER'],
      nMatch, 1, True) > 0 then
    begin
      if slst.Count >= 2 then
        FDataSize := StrToInt(slst[1]);
      if FDecimal <> 0 then
        FDataType := ftFloat
      else            
        FDataType := ftInteger
    end
    else if fStrUtil.PosFrom('float',value, 1, True) > 0 then
    begin
      FDataType := ftFloat;
      if slst.Count >= 3 then
      begin
        FDataSize := StrToInt(slst[1]);
        FDecimal := StrToInt(slst[2]);
      end;
    end
    else if fStrUtil.PosArrayFrom(value, ['datetime', 'datetime'],
      nMatch, 1, True) > 0 then
    begin
      FDataType := ftDateTime;
    end
    else if fStrUtil.PosFrom('text', value, 1, True) > 0 then
      FDataType := ftMemo
    else if fStrUtil.PosFrom('clob', value, 1, True) > 0 then
      FDataType := ftOraClob  
    else if fStrUtil.PosFrom('blob', value, 1, True) > 0 then
      FDataType := ftOraBlob;
  finally
    slst.Free;
  end;
end;

function TFieldItem.IsSimilarField(AField: TFieldItem): Boolean;
begin
  Result := (Self.FFieldName = AField.FFieldName)
    and (Self.FDataType = AField.FDataType)
    and (Self.FAsString = AField.FAsString);
end;

{ TFieldItemList }

procedure TFieldItemList.AddFieldList(fields: TFieldItemList);
var
  i: Integer;
  field: TFieldItem;
begin
  for i := 0 to fields.Count - 1 do
  begin
    field := CreateFieldItem;
    field.Assign(fields.Items[i]);
    Add(field);
  end;
end;

constructor TFieldItemList.Create(Fields: TFields);
begin
  inherited Create;
  InitByFields(Fields);
end;

procedure TFieldItemList.InitByFields(Fields: TFields);
var
  i: Integer;
  fldItem: TFieldItem;
begin
  for i := 0 to Fields.Count - 1 do
  begin
    fldItem := CreateFieldItem;
    fldItem.InitByField(Fields[i]);
    Add(fldItem);
  end;
end;

function TFieldItemList.FindFieldItem(sFieldName: string): TFieldItem;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Count - 1 do
  begin
    if SameText(sFieldName, Items[i].FieldName) then
    begin
      Result := Items[i];
      Break;
    end;  
  end;  
end;

function TFieldItemList.GetItem(index: Integer): TFieldItem;
begin
  Result := inherited GetItem(index) as TFieldItem;
end;

procedure TFieldItemList.SetItem(index: Integer; value: TFieldItem);
begin
  inherited SetItem(index, value);
end;
 
procedure TFieldItemList.SortByFieldName;
begin
  Sort(CompareFieldItemName);
end;

function TFieldItemList.GetPrimaryField: TFieldItem;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Count - 1 do
  begin
    if Items[i].IsPrimary then
    begin
      Result := Items[i];
    end;
  end;
end; 

procedure TFieldItemList.SetPrimaryField(sFieldName: string);
var
  f: TFieldItem;
begin
  f := FindFieldItem(sFieldName);
  if f <> nil then
    f.IsPrimary := True;
end;

procedure TFieldItemList.FillPrimaryFieldList(fields: TFieldItemList);
var
  i: Integer;
begin
  for i := 0 to Count - 1 do
  begin
    if Items[i].IsPrimary then
    begin
      fields.Add(Items[i]);
    end;
  end;
end;

destructor TFieldItemList.Destroy;
begin

  inherited;
end;

function TFieldItemList.IsSimilarFieldList(AFields: TFieldItemList): Boolean;
var
  i: Integer;
  fieldFind: TFieldItem;
begin
  for i := 0 to Count - 1 do
  begin
    fieldFind := AFields.FindFieldItem(Items[i].FieldName);
    if fieldFind = nil then
      Break;
    if Items[i].IsSimilarField(fieldFind) then
      Break;
  end;            
  for i := 0 to AFields.Count - 1 do
  begin
    fieldFind := FindFieldItem(AFields.Items[i].FieldName);
    if fieldFind = nil then
      Break;
    if AFields.Items[i].IsSimilarField(fieldFind) then
      Break;
  end;
  Result := True;
end;

procedure TFieldItemList.ReadValuesFromFields(fields: TFields);
var
  i: Integer;
  field: TField;
begin
  for i := 0 to Count - 1 do
  begin
    field := fields.FindField(Self.Items[i].FFieldName);
    if field <> nil then
    begin
      Items[i].AsString := field.AsString;
    end;  
  end;  
end;

function TFieldItemList.GetFirstUniqueField: TFieldItem;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Count - 1 do
  begin
    if Items[i].IsUnique then
    begin
      Result := Items[i];
      Break;
    end;  
  end;  
end;

function TFieldItemList.GetFieldNamesStr: string;
var
  i: Integer;
  s: string;
begin
  for i := 0 to Count - 1 do
  begin
    if s = '' then
      s := Items[i].FFieldName
    else
      s := s + ',' + Items[i].FFieldName;
  end; 
  Result := s; 
end;

//function TFieldItemList.GetValuesStr: string;
//var
//  i: Integer;
//  s: string;
//begin
//  for i := 0 to Count - 1 do
//  begin
//    if s = '' then
//      s := Items[i].AsString
//    else
//      s := s + ',' + Items[i].AsString;
//  end;
//  Result := s;
//end;


function TFieldItemList.InitFieldInfos(sTableName: string): Boolean;
var
  i: Integer;
  fld: TFieldItem;
  qry: TDataSet;
begin
  Result := False;
  qry := FDB.GetNewQuery;
  try
    Self.Clear;
    if not FDB.ExecQuery(qry,
           'select * from ' + sTableName + ' where 1=2') then
      Exit;       
    for i := 0 to qry.FieldCount - 1 do
    begin
      fld := Self.FindFieldItem(qry.Fields[i].FieldName);
      if fld = nil then
      begin
        fld := CreateFieldItem;
        fld.InitByField(qry.Fields[i]);
        Self.Add(fld);
      end;
    end;
    qry.Close;
    Result := True;
  finally
    qry.Free;
  end;
end;

function TFieldItemList.CreateFieldItem: TFieldItem;
begin
  Result := TFieldItem.Create;
end;  

function TFieldItemList.CreateFieldItemList: TFieldItemList;
begin
  Result := TFieldItemList.Create;
end;

function TFieldItemList.Clone: TFieldItemList;
var
  list: TFieldItemList;
  i: Integer;
  item: TFieldItem;
begin
  list := CreateFieldItemList;
  for i := 0 to Count - 1 do
  begin
    item := CreateFieldItem;
    item.Assign(items[i]);
    list.Add(item);
  end;  
  Result := list;
end;

end.
