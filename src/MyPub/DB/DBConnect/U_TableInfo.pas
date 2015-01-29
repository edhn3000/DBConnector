{**
 * <p>Title: U_TableInfo.pas  </p>
 * <p>Copyright: Copyright (c) 2009-4-12</p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: 数据库表的类，便于分析表 </p>
 *}
unit U_TableInfo;

interface  

uses
  ADODB, DB, Classes, Contnrs, XMLDoc, XMLIntf, Variants,
  U_FieldInfo, SysUtils, U_DBEngineInterface,
  U_TriggerInfo, U_IndexInfo;

const
  // %0:s field, %1:s table
  SQL_COMMON_CHECK_FIELD_UNIQUE =
    'select ''%0:s'' as FieldName, count(*) as getcount '
      +'from %1:s group by %0:s';

type

// TODO整理 FPrimaryKeys  FUniqueFields  Fields
{ TTableInfo 表信息类 记录表名和主要的字段，不一定是表的主键，但至少应该唯一 }
  TTableInfo = class
  protected
    FTableName: string;
    FSystemObject: Boolean;
    FPrimaryKeys: TFieldItemList;
    FUniqueFields: TFieldItemList;
    FCheckUniqueFields: Boolean;
    FTriggers: TTriggerInfoList;
    FIndexes: TIndexInfoList;
    FDB: IDBEngine;
    
    function GetPrimaryKeys: TFieldItemList;
    procedure CheckFieldsUnique();
    procedure SetTableName(value: string);

    procedure initObjects();virtual;
  public
    Comment: string;
    OWNER: string;
    Fields: TFieldItemList;   
    property TableName: string read FTableName write SetTableName;
    property PrimaryKeys: TFieldItemList read GetPrimaryKeys;
    property SystemObject: Boolean read FSystemObject write FSystemObject;
    property DB: IDBEngine read FDB write FDB;
    property CheckUniqueFields: Boolean read FCheckUniqueFields write FCheckUniqueFields;
  public
    constructor Create();
    destructor Destroy;override;
    function SetPrimaryKey(sKey: string; sValue: string = '';
      bCanCreate: Boolean = True): Boolean;
    procedure Assign(Source: TTableInfo);
    function GetUniqueFields: TFieldItemList;
    function GetNewFieldItem: TFieldItem;virtual;

    function InitTableInfo(tableName: string): Boolean;virtual;
  end;

{ TTableInfoList 多个表的信息 }
  TTableInfoList = class(TObjectList)
  private

  protected        
    FSystemObject: Boolean;
    function GetItem(Index: Integer): TTableInfo;
    procedure SetItem(Index: Integer; const Value: TTableInfo);

  public
    property Items[Index: Integer]: TTableInfo read GetItem write SetItem;  
    property SystemObject: Boolean read FSystemObject write FSystemObject;
  public
    procedure Add(ATableName, APrimaryKey: string);overload;
    function Find(ATableName: string): TTableInfo;
    procedure LoadFromFile(CompareSetXml: string);
  end;

implementation



{ TTableInfo }

procedure TTableInfo.Assign(Source: TTableInfo);
begin
  Self.TableName := Source.TableName;
  Self.PrimaryKeys.Assign(Source.PrimaryKeys);
  Self.Fields.Assign(Source.Fields);
end;   

procedure TTableInfo.initObjects;
begin    
  Fields := TFieldItemList.Create();
  FUniqueFields := TFieldItemList.Create();
  FPrimaryKeys := TFieldItemList.Create(False);  // not own the objects
  FTriggers := TTriggerInfoList.Create();
  FIndexes := TIndexInfoList.Create();
end;

constructor TTableInfo.Create;
begin       
  Self.TableName := 'DEF_TAB';
  initObjects;                
end;

procedure TTableInfo.CheckFieldsUnique;
var
  sSql: string;
  sFieldname: string;
  i: Integer;
  field: TFieldItem;
  qry: TDataSet;
  qry2: TDataSet;
  function InQry2(AFieldname: string): Boolean;
  begin
    Result := False;
    qry2.First;
    while not qry2.Eof do
    begin
      if SameText(qry2.FieldByName('FieldName').AsString, AFieldname) then
      begin
        Result := True;
        Break;
      end;
      qry2.Next;
    end;  
  end;  
begin
  qry := FDB.GetNewQuery;
  qry2:= FDB.GetNewQuery;
  try
    for i := 0 to Fields.Count - 1 do
    begin
      // 这里仅仅对int字段进行唯一检查
      if not (fields.Items[i].DataType in
         [ftInteger, ftSmallint, ftWord]) then
        Continue;
      if sSql = '' then
        sSql := Format(SQL_COMMON_CHECK_FIELD_UNIQUE, [
          fields.Items[i].FieldName, tableName])
      else
        sSql := sSql +
          Format(' union all ' + SQL_COMMON_CHECK_FIELD_UNIQUE, [
          fields.Items[i].FieldName, tableName]);
    end;
    if sSql = '' then
      Exit;
    if FDB.ExecQuery(qry, 'select distinct * from (' + sSql + ')')
       and FDB.ExecQuery(qry2, 'select distinct * from (' + sSql + ') where getcount > 1') then
    begin
      // 只需处理第一个记录
      while not qry.Eof do
      begin
        sFieldname := qry.FieldByName('FieldName').AsString;
        // findfielditem传入的参数和查询时起的别名有关，这里仍用了原字段名
        field := fields.FindFieldItem(sFieldname);
        if field <> nil then
        begin
          if not InQry2(sFieldname) then
            field.IsUnique := True;
        end;
        qry.Next;
      end;  
    end;
  finally
    qry.Free;
    qry2.Free;
  end;
end;

destructor TTableInfo.Destroy;
begin
  Fields.Free;
  FUniqueFields.Free;
  FTriggers.Free;
  FIndexes.Free;
  inherited;
end;       

function TTableInfo.GetPrimaryKeys: TFieldItemList;
begin
  FPrimaryKeys.Clear;    // it will not free objects
  Fields.FillPrimaryFieldList(FPrimaryKeys);
  Result := FPrimaryKeys;
end;

function TTableInfo.GetUniqueFields: TFieldItemList;
var
  item: TFieldItem;
  i: Integer;
begin
  FUniqueFields.Clear;
  for i := 0 to Self.Fields.Count - 1 do
  begin
    if Self.Fields.Items[i].IsUnique then
    begin
      item := Fields.CreateFieldItem;
      item.Assign(Self.Fields.Items[i]);
      FUniqueFields.Add(item);
    end;  
  end;
  Result := FUniqueFields;
end;

function TTableInfo.InitTableInfo(tableName: string): Boolean;
var
  i: Integer;
begin
  Result := False;
  Self.TableName := tableName;
  Fields.SystemObject := Self.SystemObject;
  Fields.DB := FDB;
  if not Fields.InitFieldInfos(tableName) then
    Exit;
  Self.FPrimaryKeys.Clear;
  for i := 0 to Fields.Count - 1 do
  begin
    if Fields.Items[i].IsPrimary then
      FPrimaryKeys.Add(Fields.Items[i]);
  end;   
  Self.FUniqueFields.Clear;
  for i := 0 to Fields.Count - 1 do
  begin  
    if Fields.Items[i].IsUnique then
      FPrimaryKeys.Add(Fields.Items[i]);
  end;

  // 无主键时，可以尝试获取唯一键，用于删除、更新数据
  if (FPrimaryKeys.Count = 0)
     and (FUniqueFields.Count = 0)
     and (FCheckUniqueFields) then
  begin
    CheckFieldsUnique();
  end;
  Result := True;
end;

function TTableInfo.SetPrimaryKey(sKey, sValue: string;
  bCanCreate: Boolean): Boolean;
var
  field: TFieldItem;
begin
  Result := False;
  field := Fields.FindFieldItem(sKey);
  if (field = nil) then
  begin
    if bCanCreate then
    begin
      field := Fields.CreateFieldItem;
      field.AsString := sValue;
      field.DataType := ftString;
      field.FieldName := sKey;
      field.IsPrimary := True;
      Fields.Add(field);      
      Result := True;
    end
    else
//      raise Exception.Create('为表设置了无效主键！');
  end
  else
  begin
    if sValue <> '' then
     field.AsString := sValue;
    field.IsPrimary := True;
    Result := True;
  end;
end;

procedure TTableInfo.SetTableName(value: string);
var
  nPos: Integer;
begin
  nPos := Pos('.', value);
  if nPos > 0 then
  begin
    Self.OWNER := Copy(value, 1, nPos-1);
    Self.FTableName := Copy(value, nPos+1, MaxInt);
  end
  else
  begin
    Self.FTableName := value;
  end;
end;

function TTableInfo.GetNewFieldItem: TFieldItem;
begin
  Result := TFieldItem.Create;
end;

{ TTableInfoList }

procedure TTableInfoList.Add(ATableName, APrimaryKey: string);
var
  tabInfo: TTableInfo;
begin
  tabInfo := TTableInfo.Create;
  tabInfo.TableName := ATableName;
  tabInfo.SetPrimaryKey(APrimaryKey, '', True);
  inherited Add(tabInfo);
end;

function TTableInfoList.Find(ATableName: string): TTableInfo;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Count - 1 do
  begin
    if SameText(Items[i].TableName, ATableName) then
    begin
      Result := Items[i];
      Break;
    end;  
  end;  
end;

function TTableInfoList.GetItem(Index: Integer): TTableInfo;
begin
  Result := inherited GetItem(index) as TTableInfo;
end;

procedure TTableInfoList.LoadFromFile(CompareSetXml: string);
var
  i: Integer;
  nodeTables, node: IXmlNode;
  xml: IXMLDocument;

  function GetChildNodeText(const Anode: IXMLNode; sChildNode: string): string;
  var
    node: IXMLNode;
  begin
    node := Anode.ChildNodes.FindNode(sChildNode);
    if (node <> nil) and (node.HasChildNodes) and
      (node.ChildNodes[0].NodeType = ntText) then
      Result := node.Text
    else
      Result := '';
  end;

//// LoadFromFile Begin
begin
  if ExtractFileExt(CompareSetXml) <> '.xml' then Exit;
  xml := NewXMLDocument();
  try
    xml.Encoding := 'GBK';
    xml.LoadFromFile(CompareSetXml);
    nodeTables := xml.DocumentElement.ChildNodes.FindNode('tables');
    if nodeTables = nil then Exit;
    for i := 0 to nodeTables.ChildNodes.Count - 1 do
    begin
      node := nodeTables.ChildNodes[i];
      if node <> nil then
        Add(GetChildNodeText(node, 'name'), GetChildNodeText(node, 'primarykey'));
    end;
  finally
    xml := nil;
  end;
end;

procedure TTableInfoList.SetItem(Index: Integer; const Value: TTableInfo);
begin
  inherited SetItem(index, Value);
end;

end.
