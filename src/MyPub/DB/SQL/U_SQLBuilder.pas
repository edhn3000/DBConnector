{**
 * <p>Title: U_SQLBuilder.pas  </p>
 * <p>Copyright: Copyright (c) 2009-4-12</p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: TSQLBuilder ����sql�Ĵ����� </p>
 *}
unit U_SQLBuilder;

interface

uses
  ADODB, DB, Classes, Contnrs, Variants, SysUtils,
  Forms, U_PlanarList, U_fStrUtil, U_DBConnnectManager,
  U_FieldInfo, U_TableInfo, U_DBEngineInterface, U_DBConnectInterface;

type
{ TSQLBuilder sql��������� }
  TSQLBuilder = class
  private
    FFields: TFieldItemList;
    FPrimaryFields: TFieldItemList;
    FSQLCommandType: TSQLCmdType;
    FDBType: TDBType;
    FTableInfo: TTableInfo;
    FsCondition: string;
    FnPageRecord: Integer;
    FnCurrPage: Integer;
    FLastError: string;

//    function GetPrimaryField: TFieldItem;
//    procedure SetPrimaryField(value: TFieldItem);
    function GetCondition: string;
    function GetFields: TFieldItemList;
    function GetFieldsKeyValuesWithAnd(Fields: TFieldItemList): string;

  protected
    function GetItem(Index: Integer): TFieldItem;
    procedure SetItem(Index: Integer; const Value: TFieldItem);

    function BuildInsertSQL(): string;
    function BuildDeleteSQL(): string;
    function BuildUpdateSQL(): string;
    function BuildSelectSQL(): string;
    function BuildCreateSQL(): string;
    function BuildDropSQL: String;
                                                    
    function InitByInsertSQL(sSql: string): Boolean;
    function InitByDeleteSQL(sSql: string): Boolean;
    function InitByUpdateSQL(sSql: string): Boolean;
    function InitBySelectSQL(sSql: string): Boolean;

  public
    property Fields: TFieldItemList read GetFields;
    property PrimaryKeys: TFieldItemList read FPrimaryFields write FPrimaryFields;
    property Field[Index: Integer]: TFieldItem read GetItem write SetItem;
    property SQLCommandType: TSQLCmdType read FSQLCommandType write FSQLCommandType;
    property TableInfo: TTableInfo read FTableInfo;
    property Condition: string read GetCondition write FsCondition; // ����������ã�������ȷ����ʹ�ù���
    property PageRecord: Integer read FnPageRecord write FnPageRecord;
    property CurrPage: Integer read FnCurrPage write FnCurrPage;
  public
    constructor Create(dbt: TDBType);
    destructor Destroy;override;
    // Ŀǰֻ��ɶ�select�ķ���  20071024
    function InitBySql(ASql: string): Boolean;
    function BuildSQL(): string;
    function FieldByName(sFieldName: string): TFieldItem;
    function AddField(AField: TFieldItem): Integer;
    procedure AddFields(fields: TFieldItemList);
  end;

implementation

uses
  U_TableInfoFactory, U_FieldInfoFactory, U_SqlUtils;

{ TSQLBuilder }

function TSQLBuilder.BuildInsertSQL(): string;
var
  sFieldNames, sValues: string; 
  db: IDBConnect;
begin                        
  db := TDBConnectManager.CreateDBConnect(FDBType);
  try
    FSQLCommandType := sctInsert;
    sFieldNames := Fields.GetFieldNamesStr;
    sValues := db.GetFieldsValuesStr(Fields);
    Result := Format('insert into %s(%s) values(%s);', [
      FTableInfo.TableName, sFieldNames, sValues]);
  finally
    db := nil;
  end;
end;

function TSQLBuilder.BuildDeleteSQL(): string;
var
  field: TFieldItem;       
  db: IDBConnect;
begin
  FSQLCommandType := sctDelete;                   
  db := TDBConnectManager.CreateDBConnect(FDBType);
  try
    // ����ʹ������ɾ�������û��������Ϣ����ʹ��һ��Ψһ��ɾ���������û�У�ʹ��ȫ����ɾ��
    if PrimaryKeys.Count <> 0 then begin
      Result := Format('delete from %s where %s;', [FTableInfo.TableName,
        GetFieldsKeyValuesWithAnd(PrimaryKeys) + Condition]);
    end else if Fields.GetFirstUniqueField <> nil then begin
      field := Fields.GetFirstUniqueField;
      Result := Format('delete from %s where %s=%s;', [FTableInfo.TableName,
        field.FieldName, db.GetValidFieldValue(field)]) + Condition;
    end else begin
      // ȫ����ɾ�� 
      Result := Format('delete from %s where %s;', [FTableInfo.TableName,
        GetFieldsKeyValuesWithAnd(Fields)]);
    end;
  finally
    db := nil;
  end;
end;

function TSQLBuilder.BuildSelectSQL(): string;
var
  sSql: string;
  sFields: string;
  sPageSql: string;
  sAsName: string;
begin
  FSQLCommandType := sctSelect;
  sAsName := '';    
  sSql := '';
  
  if Fields.Count = 0 then
    sFields := '*'
  else
    sFields := Fields.GetFieldNamesStr;
  // ��ҳsql����
  sPageSql := '';
  if FnPageRecord > 0 then
  begin
    case FDBType of
      dbtOracle:
      begin
        sPageSql := ' where RN >'+IntToStr((FnCurrPage - 1)*FnPageRecord)
           + ' and RN <=' + IntToStr(FnCurrPage*FnPageRecord);
        if sFields = '*' then
        begin
          sAsName := TableInfo.TableName +'_as';
          sFields := sAsName + '.*';
        end;
        sSql := 'select * from '
            + '(select rownum RN,'+ sFields + ' from ' + TableInfo.TableName
            + ' ' + sAsName + ' ' + Condition + ')'
            + sPageSql;
      end;
      else
        sSql := 'select ' + sFields + ' from ' + TableInfo.TableName + Condition;
    end;
    Result := sSql;
  end;
end;

function TSQLBuilder.BuildUpdateSQL(): string;
var
  sSql: string;
  sCondition: string;         
  db: IDBConnect;
begin
  FSQLCommandType := sctUpdate;                   
  db := TDBConnectManager.CreateDBConnect(FDBType);
  try

    if PrimaryKeys.Count > 0 then
    begin
      sCondition := db.GetFieldsKeyValuesStr(PrimaryKeys);
    end;
    if sCondition + Condition = '' then
    begin
      raise Exception.Create('(TSQLBuilder.BuildUpdateSQL)����update���ʱδ����where�Ӿ������');
    end;
    sSql := '';
    sSql := db.GetFieldsKeyValuesStr(Fields);
    Result := Format('update %s set %s where %s;',
                     [FTableInfo.TableName, sSql]) + sCondition + Condition; 
  finally
    db := nil;
  end;
end;

function TSQLBuilder.BuildDropSQL():String;
begin
  FSQLCommandType := sctDrop;
  Result := Format('Drop Table %s;', [TableInfo.TableName]);
end;

function TSQLBuilder.BuildCreateSQL(): string;
var
  i: Integer;
  PlanarStringList: TPlanarStringList;
  slstResult: TStringList;
begin
  FSQLCommandType := sctCreate;
  slstResult := TStringList.Create;
  try
    slstResult.Add(Format('Create Table %s(', [TableInfo.TableName]));
    PlanarStringList := TPlanarStringList.Create;
    try
      for i := 0 to Fields.Count - 1 do
      begin
        if i <> Fields.Count - 1 then
        begin
          PlanarStringList.Strs[i, 0] := '  ' + Fields.Items[i].FieldName + ' ';
          PlanarStringList.Strs[i, 1] := Fields.Items[i].GetDataTypeInSQL + ',';
        end
        else
        begin   
          PlanarStringList.Strs[i, 0] := '  ' + Fields.Items[i].FieldName + ' ';
          PlanarStringList.Strs[i, 1] := Fields.Items[i].GetDataTypeInSQL;
        end;  
      end;                    
      PlanarStringList.JustifyMode := pjmBothLeft;
      PlanarStringList.FormatItemLengths;
      for i := 0 to PlanarStringList.Count-1 do
      begin
        slstResult.Add(PlanarStringList.ItemStr[i]);
      end;  
    finally
      PlanarStringList.Free;
    end;
    slstResult.Add(');');
    Result := slstResult.Text;
  finally
    slstResult.Free;
  end;             
end;

function TSQLBuilder.BuildSQL(): string;
begin
  if not Assigned(TableInfo) then
  begin
    raise Exception.Create('����sqlʱû�и���TableInfo��Ϣ��');
    Exit;
  end;  
  case SQLCommandType of
    sctInsert: Result := BuildInsertSQL();
    sctUpdate: Result := BuildUpdateSQL();
    sctDelete: Result := BuildDeleteSQL();
    sctSelect: Result := BuildSelectSQL();
    sctCreate: Result := BuildCreateSQL();
    sctDrop :  Result := BuildDropSQL();
  end;
end;

constructor TSQLBuilder.Create(dbt: TDBType);
begin
  FDBType := dbt;
  FFields :=  TFieldInfoFactory.CreateFieldItemList(dbt);
  FPrimaryFields := TFieldInfoFactory.CreateFieldItemList(dbt);
  FTableInfo := TTableInfoFactory.CreateTableInfo(dbt);
  
  FnPageRecord := 0;
  FnCurrPage := 1;
end;

destructor TSQLBuilder.Destroy;
begin
  if Assigned(FFields) then
    FreeAndNil(FFields);
  if Assigned(FPrimaryFields) then
    FreeAndNil(FPrimaryFields);
  if Assigned(FTableInfo) then
    FreeAndNil(FTableInfo);
  inherited;
end;

function TSQLBuilder.FieldByName(sFieldName: string): TFieldItem;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Fields.Count - 1 do
  begin
    if Field[i].FieldName = sFieldName then
    begin
      Result := Field[i];
      Break;
    end;
  end;
end;

function TSQLBuilder.GetCondition: string;
begin
  Result := FsCondition;
end;

function TSQLBuilder.GetFields: TFieldItemList;
begin
  Result := FFields;
end;

function TSQLBuilder.GetItem(Index: Integer): TFieldItem;
begin
  Result := FFields.Items[index];
end;

function TSQLBuilder.InitByInsertSQL(sSql: string): Boolean;
var
  nPos, nPos2: Integer;
//  nPos3: Integer;
  sFieldNames: string;
  sFieldValues: string;
  slstFieldNames, slstFieldValues: TStrings;
  i: Integer;
  fieldItem: TFieldItem;
begin
  Result := False;
  nPos := Pos('into', LowerCase(sSql));
  nPos2 := Pos('(', sSql);
//  nPos3 := Pos('values', LowerCase(sSql));
  if (nPos = 0) then
  begin
    FLastError := 'δ�ҵ�into�ؼ���';
    Exit;
  end;                            
  if (nPos2 = 0) then
  begin
    FLastError := 'δ�ҵ�(';
    Exit;
  end;
  TableInfo.TableName := Trim(Copy(sSql, nPos+4, nPos2-1));
  nPos := fStrUtil.ParseMatchedBracket(nPos2, sSql);
  if nPos = 0 then
  begin
    FLastError := '��һ�������޽�β';
    Exit;
  end;  
  sFieldNames := Copy(sSql, nPos2+1, nPos-1-nPos2-1);
  nPos := fStrUtil.PosFrom('(', sSql, nPos+1);
  if nPos = 0 then
  begin
    FLastError := '�Ҳ����ڶ�������';
    Exit;
  end;
  nPos2 := fStrUtil.ParseMatchedBracket(nPos, sSql);
  sFieldValues := Copy(sSql, nPos+1, nPos2-1-nPos-1); 
  SQLCommandType := sctInsert;
  Fields.Clear;
  slstFieldNames  := TStringList.Create;
  slstFieldValues := TStringList.Create;
  try
    ExtractStrings([','], [' '], PChar(sFieldNames), slstFieldNames);
    ExtractStrings([','], [' '], PChar(sFieldValues), slstFieldValues);
    while slstFieldValues.Count < slstFieldNames.Count do
      slstFieldValues.Add('');
      
    for i := 0 to slstFieldNames.Count - 1 do
    begin
      fieldItem := fields.CreateFieldItem;
      fieldItem.FieldName := Trim(slstFieldNames[i]);
      fieldItem.AsString  := Trim(slstFieldValues[i]);
      FFields.Add(fieldItem);
    end;
  finally
    slstFieldNames.Free;
    slstFieldValues.Free;
  end;   
end;

function TSQLBuilder.InitByDeleteSQL(sSql: string): Boolean;
//var
//  nPos, nPos2: Integer;
begin
  Result := False;
  SQLCommandType := sctDelete;
  Fields.Clear;
  TableInfo.TableName := TSQLUtils.GetTableNameBySQL(sSql);
end;

function TSQLBuilder.InitByUpdateSQL(sSql: string): Boolean; 
begin
  Result := False;
  SQLCommandType := sctUpdate;
  Fields.Clear; 
end;

function TSQLBuilder.InitBySelectSQL(sSql: string): Boolean;
var
  sFieldName: string;
  sFields: string;
  sTable: string;
  nPos, nPos2: Integer;
  fld: TFieldItem;
begin
  Result := False;
  nPos := Pos('from', LowerCase(sSql));
  if nPos = 0 then
  begin
    FLastError := 'δ�ҵ�from�ؼ���';
    Exit;
  end;
  sFields := Copy(sSql, 7, nPos-1-7);
  sTable := Trim(Copy(sSql, nPos+4, MaxInt));
  nPos2 := Pos(' ', sTable);
  if nPos2 = 0 then
    nPos2 := MaxInt;
  sTable := Copy(sTable, 1, nPos2);
  TableInfo.TableName := sTable;
  SQLCommandType := sctSelect;
  Fields.Clear;

  while sFields <> '' do
  begin
    nPos := Pos(',', sFields);
    if nPos = 0 then
    begin
      sFieldName := sFields;
      sFields := ''
    end
    else
    begin
      sFieldName := Copy(sFields, 1, nPos - 1);
      sFields := Copy(sFields, nPos + 1, MaxInt);
    end;
    if Trim(sFieldName) <> '*' then
    begin
      fld := fields.CreateFieldItem;
      fld.FieldName := Trim(sFieldName);
      Fields.Add(fld);
    end;
  end;
end;

function TSQLBuilder.InitBySql(ASql: string): Boolean;
var
  sSql: string;
begin
  sSql := Trim(ASql);
  if Pos('select', LowerCase(sSql)) = 1 then
  begin
    Result := InitBySelectSQL(sSql);
  end
  else if Pos('insert', LowerCase(sSql)) = 1 then
  begin
    Result := InitByInsertSQL(sSql);
  end
  else if Pos('update', LowerCase(sSql)) = 1 then
  begin
    Result := InitByUpdateSQL(sSql);
  end
  else if Pos('delete', LowerCase(sSql)) = 1 then
  begin
    Result := InitByDeleteSQL(sSql);
  end
  else
    Result := False;
end;

function TSQLBuilder.GetFieldsKeyValuesWithAnd(Fields: TFieldItemList): string;
var
  i: Integer;
  sCondition: string;
  db: IDBConnect;
begin
  sCondition := '';
  db := TDBConnectManager.CreateDBConnect(FDBType);
  try
    for i := 0 to Fields.Count - 1 do
    begin
      if sCondition = '' then
        sCondition := Format('%s=%s',[Fields.Items[i].FieldName,
          db.GetValidFieldValue(Fields.Items[i])])
      else
        sCondition := sCondition + ' and '+ Format('%s=%s',[
          Fields.Items[i].FieldName,
          db.GetValidFieldValue(Fields.Items[i])]);
    end;
    Result := sCondition;
  finally
    db := nil;
  end;
end;

procedure TSQLBuilder.SetItem(Index: Integer; const Value: TFieldItem);
begin
  FFields.Items[Index] := Value;
end;

function TSQLBuilder.AddField(AField: TFieldItem): Integer;
begin
  Result := Fields.Add(AField);
end;     

procedure TSQLBuilder.AddFields(fields: TFieldItemList);
var
  i: Integer;
  field: TFieldItem;
begin
  for i := 0 to fields.Count - 1 do
  begin
    field := fields.CreateFieldItem;
    field.Assign(fields.Items[i]);
    AddField(field);
  end;  
end;       

end.
