{**
 * <p>Title: U_OracleDBConnect.pas  </p>
 * <p>Copyright: Copyright (c) 2009-3-25</p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: 继承自TDBConnect类，实现一些针对oracle的数据获取方法 </p>
 *}
unit U_OracleDBConnect;

interface

uses
  Windows, ADODB, DB, Classes, Contnrs, Variants,
  Dialogs, SysUtils, forms,
  U_TableInfo, U_OracleTableInfo, U_FieldInfo, U_OracleFieldInfo,
  U_DBConnect, U_JSON, U_DBEngineInterface, U_DBConnectInterface;

const
  SQL_ORACLE_GETTABLES_USER   =
    'select tab.TABLE_NAME as NAME, tabcom.COMMENTS'
    +' from USER_TABLES tab, USER_TAB_COMMENTS tabcom'
    +' where tab.TABLE_NAME = tabcom.TABLE_NAME'
    +' order by tab.TABLE_NAME';
//    'select distinct TABLE_NAME as NAME from USER_TABLES';
  SQL_ORACLE_GETTEBLES_ALL    =
    'select tab.OWNER, tab.TABLE_NAME as NAME, tabcom.COMMENTS'
    +' from ALL_TABLES tab, ALL_TAB_COMMENTS tabcom'
    +' where tab.TABLE_NAME = tabcom.TABLE_NAME '
    +       'and tab.owner = tabcom.owner '
    +' order by tab.TABLE_NAME';
//    'select distinct TABLE_NAME as NAME from ALL_TABLES';

  SQL_ORACLE_GETTABLEINFOS_USER   =
    'select  tab.OWNER, tab.TABLE_NAME as NAME,'
          +'tabcom.COMMENTS,'
          +'conscol.COLUMN_NAME PRIMARYKEY,'
          +'conscol.OWNER '
    +' from USER_TABLES tab, '
          +'USER_TAB_COMMENTS tabcom, '
          +'USER_CONS_COLUMNS conscol '
    +' where tab.TABLE_NAME = tabcom.TABLE_NAME '
          +' and tab.TABLE_NAME = conscol.TABLE_NAME(+)';
    
  SQL_ORACLE_GETTABLEINFOS_ALL    =
    'select  tab.OWNER, tab.TABLE_NAME as NAME,'
          +'tabcom.COMMENTS,'
          +'conscol.COLUMN_NAME PRIMARYKEY,'
          +'conscol.OWNER '
    +' from ALL_TABLES tab, '
          +'ALL_TAB_COMMENTS tabcom, '
          +'ALL_CONS_COLUMNS conscol '
    +' where tab.TABLE_NAME = tabcom.TABLE_NAME '
          +' and tab.table_name = conscol.table_name(+)';


  SQL_ORACLE_GETVIEWS_USER    =
    'select distinct VIEW_NAME as NAME from USER_VIEWS';
  SQL_ORACLE_GETVIEWS_ALL     =
    'select OWNER, VIEW_NAME as NAME from ALL_VIEWS';

  SQL_ORACLE_GETINDEXES_USER  =
    'select distinct INDEX_NAME as NAME from USER_INDEXES';
  SQL_ORACLE_GETINDEXES_ALL   =
    'select OWNER, INDEX_NAME as NAME from ALL_INDEXES';

  SQL_ORACLE_GETSYNONYMS_USER =
    'select distinct SYNONYM_NAME as NAME from USER_SYNONYMS';
  SQL_ORACLE_GETSYNONYMS_ALL  =
    'select OWNER, SYNONYM_NAME as NAME from ALL_SYNONYMS';

  SQL_ORACLE_GETDBLINKS_USER  =
    'select distinct DB_LINK as NAME from USER_DB_LINKS';
  SQL_ORACLE_GETDBLINKS_ALL   =
    'select OWNER, DB_LINK as NAME from ALL_DB_LINKS';


  SQL_ORACLE_GETTABLE_FIELD_USER =
    'select col.COLUMN_NAME, col.DATA_TYPE, col.NULLABLE, col.DATA_PRECISION, '
          +'col.DATA_SCALE, colcom.COMMENTS col_comment'
    +' from USER_TABLES tab'
        + ' USER_TAB_COLUMNS col, USER_COL_COMMENTS colcom'
    +' where tab.TABLE_NAME=col.TABLE_NAME'
        + ' and tab.TABLE_NAME = colcom.TABLE_NAME'
        + ' and col.COLUMN_NAME=colcom.COLUMN_NAME'
        + ' and tab.TABLE_NAME = ''%s'''
        + ' and col.COLUMN_NAME = ''%s''';


  SQL_ORACLE_GETTABLE_FIELD_ALL =
    'select col.COLUMN_NAME, col.DATA_TYPE, col.NULLABLE, col.DATA_PRECISION, '
          +'col.DATA_SCALE, colcom.COMMENTS col_comment'
    +' from all_tables tab'
        + ' ALL_TAB_COLUMNS col, ALL_COL_COMMENTS colcom'
    +' where tab.TABLE_NAME=col.TABLE_NAME'
        + ' and tab.TABLE_NAME = colcom.TABLE_NAME'
        + ' and col.COLUMN_NAME=colcom.COLUMN_NAME'
        + ' and tab.TABLE_NAME = ''%s'''
        + ' and col.COLUMN_NAME = ''%s''';

  SQL_ORACLE_GETPROCS_USER  =
    'select distinct OBJECT_NAME as NAME from USER_PROCEDURES';
  SQL_ORACLE_GETPROCS_ALL   =
    'select OBJECT_NAME as NAME, OWNER from ALL_PROCEDURES '
    + 'where procedure_name is null';

  SQL_ORACLE_GETTRIGS_USER  =
    'select TRIGGER_NAME as NAME, TABLE_OWNER as OWNER, TABLE_NAME from USER_TRIGGERS';
  SQL_ORACLE_GETTRIGS_ALL   =
    'select TRIGGER_NAME as NAME, OWNER, TABLE_NAME from ALL_TRIGGERS';

  SQL_ORACLE_GETCHARSET     =
    'select userenv(''language'') from dual';

  SQL_ORACLE_GETPRIMARYFIELD_USER =
    'Select conscol.COLUMN_NAME,cons.CONSTRAINT_NAME'
    +' From USER_CONSTRAINTS cons,USER_CONS_COLUMNS conscol'
    +' Where cons.CONSTRAINT_NAME=conscol.CONSTRAINT_NAME'
          +' and cons.CONSTRAINT_TYPE=''P'''
          +' and cons.TABLE_NAME=''%s''';

  SQL_ORACLE_GETPRIMARYFIELD_ALL  =   
    'Select conscol.COLUMN_NAME,cons.CONSTRAINT_NAME'
    +' From ALL_CONSTRAINTS cons,ALL_CONS_COLUMNS conscol'
    +' Where cons.CONSTRAINT_NAME=conscol.CONSTRAINT_NAME'
          +' and cons.CONSTRAINT_TYPE=''P'''
          +' and cons.TABLE_NAME=''%s''';

  SQL_ORACLE_GETUSERS_USER =
    'select * from USER_USERS';
  SQL_ORACLE_GETUSERS_ALL  =
    'select * from ALL_USERS';

  
  SQL_ORACLE_GETPROC_SOURCE_USER =
    'select LINE, TEXT from USER_SOURCE where NAME=''%s'''; 
  SQL_ORACLE_GETPROC_SOURCE_ALL =
    'select LINE, TEXT from ALL_SOURCE where NAME=''%s''';

  SQL_ORACLE_GETVIEW_SOURCE_USER =
    'select TEXT from USER_VIEWS where VIEW_NAME=''%s''';
  SQL_ORACLE_GETVIEW_SOURCE_ALL =
    'select TEXT from ALL_VIEWS where VIEW_NAME=''%s''';

type
  TOracleDBConnect = class(TDBConnect)
  private
  protected        
    function GetMaxRecordSQLCondition(nMaxRecord: Integer;
      sSql: string): string;override;
    function GetNewTableInfo(): TTableInfo;override;
  public
    constructor Create;override;
    destructor Destroy;override;
    // inherited
    procedure GetTableNames(List: TStrings; AQry: TDataSet = nil);override;
//    procedure GetTableInfos(List: TTableInfoList; AQry: TDataSet = nil);override;
    function GetFieldInfos(const TableName: string; AQry: TDataSet = nil): TFieldItemList;override;
    procedure GetProcedureNames(List: TStrings; AQry: TDataSet = nil);override;
    procedure GetTriggerNames(List: TStrings; AQry: TDataSet = nil);override;
    function GetPrimaryField(TableName: string; AQry: TDataSet = nil): TFieldItem;override;
    procedure GetViewNames(List: TStrings; AQry: TDataSet = nil);override;
                                             
    function GetTableInfo(sTableName: string): TTableInfo;override; 

    function GetValidFieldValue(FieldItem: TFieldItem;
      const sValueDef: string = ''): string;override;

//    function ExportTableCreateSQL(sTableName: string; slstResult: TStrings):Boolean;override;
//    procedure ExportTableToSql(tableName: string; sExportFileName: string);override;
//    function ExportQuery(AQry: TDataSet; sExportFileName: string): Boolean;override;

    // own
    procedure GetIndexNames(List: TStrings);override;
    procedure GetUsers(List: TStrings);override;
    procedure GetSynonymNames(List: TStrings);override;
    procedure GetDBLinkNames(List: TStrings);override;

    function GetTableFieldInfo(TableName, FieldName: string;
      AQry: TDataSet = nil): TFieldItem;

    function GetCharset: string;
                                    
    function UpdateBlobs(JsonStr: string): Integer;override;  
    function UpdateClobs(JsonStr: string): Integer;override;
//    function SaveBlobs(blobFields: TFieldItemList; tableInfo: TTableInfo;
//      sDir: string; var json: TJSON): Boolean;override;
//    function SaveClobs(clobFields: TFieldItemList; tableInfo: TTableInfo;
//      sDir: string; var json: TJSON): Boolean;override;

    function GetProcSource(sName: string; list: TStrings): string;override;
    function GetTrigSource(sName: string; list: TStrings): string;override;
    function GetViewSource(sName: string; list: TStrings): string;override;
  end;

implementation



{ TOracleDBConnect } 

constructor TOracleDBConnect.Create;
begin
  inherited;
  SetDBType(dbtOracle);
end;

destructor TOracleDBConnect.Destroy;
begin

  inherited;
end;

function TOracleDBConnect.GetMaxRecordSQLCondition(nMaxRecord: Integer;
  sSql: string): string;
var
  nPosOrder, nPosGroup, nPosWhere: Integer;
begin
  nPosOrder := Pos('order', LowerCase(sSql));
  nPosGroup := Pos('group', LowerCase(sSql));
  nPosWhere := Pos('where', LowerCase(sSql));
  if (nPosOrder = 0) and (nPosGroup = 0) then
  begin
    if nPosWhere = 0 then
      Result := sSql + ' where rownum > ' + IntToStr(nMaxRecord)
    else
      Result := sSql + ' rownum > ' + IntToStr(nMaxRecord);
  end
  else if (nPosOrder > 0) then //
    Result := Copy(sSql, 1, nPosOrder - 1) + ' rownum > ' + IntToStr(nMaxRecord)
      + Copy(sSql, nPosOrder, MaxInt)
  else if (nPosGroup > 0) then
    Result := Copy(sSql, 1, nPosGroup - 1) + ' rownum > ' + IntToStr(nMaxRecord)
      + Copy(sSql, nPosGroup, MaxInt)
  else
    Result := sSql;
end; 

function TOracleDBConnect.GetTableInfo(sTableName: string): TTableInfo;
var
  qry, qry2: TDataSet;
  table: TTableInfo;
begin
  qry := GetNewQuery;
  qry2 := GetNewQuery;
  try
    table := TOracleTableInfo.Create;
    table.TableName := sTableName;
    table.SystemObject := Self.SystemObject;
    table.DB := Self.DBEngine;          
    // 先初始化字段列表，之后再设置主键
    table.InitTableInfo(table.TableName);
    Result := table;
  finally
    qry.Free;
    qry2.Free;
  end;
end;

procedure TOracleDBConnect.GetTableNames(List: TStrings; 
  AQry: TDataSet);    
var
  sGetTables: string;
begin
  if IsQryInvalid(AQry) then
    AQry := DBEngine.GetDataSet;
  if SystemObject then
    sGetTables := SQL_ORACLE_GETTEBLES_ALL
  else
    sGetTables := SQL_ORACLE_GETTABLES_USER;
  if ExecQuery(AQry, sGetTables) then
    FillListByOwnerName(List, AQry);
end;

function TOracleDBConnect.GetFieldInfos(const TableName: string;
  AQry: TDataSet): TFieldItemList;
var
  fld: TFieldItem;
  sSql: string;
  List: TFieldItemList;
begin     
  if IsQryInvalid(AQry) then
    AQry := DBEngine.GetDataSet;
  if SystemObject then
    sSql := Format(SQL_ORACLE_GETFIELDS_ALL, [TableName])
  else
    sSql := Format(SQL_ORACLE_GETFIELDS_USER, [TableName]); 
  List := TOracleFieldItemList.Create;
  if ExecQuery(AQry, sSql) then
  with AQry do
  begin
    while not Eof do
    begin
      fld := List.FindFieldItem(FieldByName('COLUMN_NAME').AsString);
      if fld = nil then
      begin
        fld := TOracleFieldItem.Create;  
        List.Add(fld);
      end;
      fld.FieldName   := FieldByName('COLUMN_NAME').AsString;   
      fld.Nullable    := FieldByName('NULLABLE').AsBoolean;
      if SameText(FieldByName('DATA_TYPE').AsString, 'NUMBER')then
      begin     
        fld.DataSize := FieldByName('DATA_PRECISION').AsInteger;
      end
      else if SameText(FieldByName('DATA_TYPE').AsString, 'VARCHAR')
         or SameText(FieldByName('DATA_TYPE').AsString, 'VARCHAR2') then
      begin
        fld.DataSize := FieldByName('DATA_LENGTH').AsInteger;
      end
      else
      begin
        fld.DataSize := FieldByName('DATA_LENGTH').AsInteger;
      end;          
      fld.Decimal     := FieldByName('DATA_SCALE').AsInteger;

      // DataTypeStr的赋值必须在Decimal之后，其中可据此检查integer和float
      fld.DataTypeStr := FieldByName('DATA_TYPE').AsString;
      fld.Comment     := FieldByName('col_comment').AsString;
      Next;
    end;
  end;
  Result := List;
end;

function TOracleDBConnect.GetPrimaryField(TableName: string;
   AQry: TDataSet): TFieldItem;
var
  sGetPrimFld: string;
  sFieldName: string;
begin
  Result := nil;
  if IsQryInvalid(AQry) then
    AQry := DBEngine.GetDataSet;
  if SystemObject then
    sGetPrimFld := SQL_ORACLE_GETPRIMARYFIELD_ALL
  else
    sGetPrimFld := SQL_ORACLE_GETPRIMARYFIELD_USER;
  if ExecQuery(AQry, Format(sGetPrimFld, [TableName])) then
  begin
    if not AQry.Eof then
    begin
      sFieldName := AQry.FieldByName('COLUMN_NAME').AsString;
      Result := GetTableFieldInfo(TableName, sFieldName, AQry);
    end;
  end;
end;

procedure TOracleDBConnect.GetProcedureNames(List: TStrings;
   AQry: TDataSet);  
var
  sGetProcs: string;
begin   
  if IsQryInvalid(AQry) then
    AQry := DBEngine.GetDataSet;
  if SystemObject then
    sGetProcs := SQL_ORACLE_GETPROCS_ALL
  else
    sGetProcs := SQL_ORACLE_GETPROCS_USER;
  if ExecQuery(AQry, sGetProcs) then
    FillListByOwnerName(List, AQry);
end;

procedure TOracleDBConnect.GetTriggerNames(List: TStrings;
   AQry: TDataSet);
var
  sGetTrigs: string;   
begin   
  if IsQryInvalid(AQry) then
    AQry := DBEngine.GetDataSet;
  if SystemObject then
    sGetTrigs := SQL_ORACLE_GETTRIGS_ALL
  else
    sGetTrigs := SQL_ORACLE_GETTRIGS_USER;
  if ExecQuery(AQry, sGetTrigs) then
    FillListByOwnerName(List, AQry);
end;

//////////////////////////////////own

procedure TOracleDBConnect.GetIndexNames(List: TStrings);
var
  AQry: TDataSet;
begin
  AQry := GetNewQuery;
  try
    if SystemObject then
      GetObjectNames(List, SQL_ORACLE_GETINDEXES_ALL, AQry)
    else
      GetObjectNames(List, SQL_ORACLE_GETINDEXES_USER, AQry);
  finally
    AQry.Free;
  end;
end;

procedure TOracleDBConnect.GetDBLinkNames(List: TStrings);
var
  AQry: TDataSet;
begin
  AQry := GetNewQuery;
  try
    if SystemObject then
      GetObjectNames(List, SQL_ORACLE_GETDBLINKS_ALL, AQry)
    else
      GetObjectNames(List, SQL_ORACLE_GETDBLINKS_USER, AQry);
  finally
    AQry.Free;
  end;
end;

procedure TOracleDBConnect.GetSynonymNames(List: TStrings);
var
  AQry: TDataSet;
begin
  AQry := GetNewQuery;
  try
    if SystemObject then
      GetObjectNames(List, SQL_ORACLE_GETSYNONYMS_ALL, AQry)
    else
      GetObjectNames(List, SQL_ORACLE_GETSYNONYMS_USER, AQry);
  finally
    AQry.Free;
  end;
end;

function TOracleDBConnect.GetValidFieldValue(FieldItem: TFieldItem;
  const sValueDef: string): string;
var
  sFieldValue: string;
begin
  sFieldValue := FieldItem.AsString;
  case FieldItem.DataType of
    ftDate, ftTime, ftDateTime:
    begin
      if Length(sFieldValue) <= 10 then
        Result := Format('to_date(''%s'',''yyyy-mm-dd'')', [sFieldValue])
      else if Length(sFieldValue) <= 19 then
        Result := Format('to_date(''%s'',''yyyy-mm-dd hh24:mi:ss'')', [sFieldValue])
      else if Length(sFieldValue) > 19 then
        Result := Format('to_date(''%s'',''yyyy-mm-dd hh24:mi:ss.ff3'')', [sFieldValue])
      else
        Result := sValueDef;
    end; 
    ftBlob, ftOraBlob:
    begin
      //todo:
      Result := 'Null';
    end;
    else
      Result := inherited GetValidFieldValue(FieldItem);
  end;
end;

procedure TOracleDBConnect.GetViewNames(List: TStrings; 
  AQry: TDataSet);
var
  sGetViews: string;
begin
  if SystemObject then
    sGetViews := SQL_ORACLE_GETVIEWS_ALL
  else
    sGetViews := SQL_ORACLE_GETVIEWS_USER;
  if ExecQuery(AQry, sGetViews) then
    FillListByOwnerName(List, AQry);
end;

function TOracleDBConnect.GetTableFieldInfo(TableName, FieldName: string;
   AQry: TDataSet): TFieldItem;
var
  sGetTableField: string;
  fld: TFieldItem;
begin
  Result := nil;
  if SystemObject then
    sGetTableField := SQL_ORACLE_GETTABLE_FIELD_ALL
  else
    sGetTableField := SQL_ORACLE_GETTABLE_FIELD_USER;
  if ExecQuery(AQry, Format(sGetTableField, [TableName, FieldName])) then
  begin
    if not AQry.Eof then
    begin
      fld := TOracleFieldItem.Create;
      fld.FieldName   := AQry.FieldByName('COLUMN_NAME').AsString;
      fld.DataTypeStr := AQry.FieldByName('DATA_TYPE').AsString;
      fld.Nullable    := AQry.FieldByName('NULLABLE').AsBoolean;
      fld.DataSize    := AQry.FieldByName('DATA_PRECISION').AsInteger;
      fld.Decimal     := AQry.FieldByName('DATA_SCALE').AsInteger;
      fld.Comment     := AQry.FieldByName('col_comment').AsString;
      Result := fld;
    end;  
  end;  
end;

function TOracleDBConnect.GetCharset: string;
var
  qry: TDataSet;
begin
  qry := GetNewQuery;
  try
    if ExecQuery(qry, SQL_ORACLE_GETCHARSET) then
    begin
      if not qry.Eof then
        Result := qry.Fields[0].AsString;
    end;
  finally
    qry.Free;
  end;
end;  

function TOracleDBConnect.UpdateBlobs(JsonStr: string): Integer;
var
  sSql: string;
  sFieldJson: string;
  sFieldName, sFilePath: string;
  sFileFullPath: string;
  json: TJSON;
  num: Integer;
  qry: TDataSet;
//  stream: TMemoryStream;
  slst: TStringList;
//  fileStream: TFileStream;
  blobfield: TBlobField;
begin
  Result := 0;
  json := TJSON.Create;
  qry := GetNewQuery;
  try
    json.Init(JsonStr);
    sSql := json.ValueByName[C_JSON_KEY_SQL];
    // 查询sql出错  
    if not ExecQuery(qry, sSql) then
    begin
      Result := TDBC_ERROR_EXECFAIL;
      Exit;
    end;
    qry.Edit;       // qry状态 可编辑 才可以更新select获得的数据
    // 查询sql无结果，则无更新
    if qry.Eof then
    begin
      Exit;
    end;
    
    num := 1;
    while True do
    begin
      // 根据顺序取field，字母必须是小写并且数字的顺序连续
      sFieldJson := json.ValueByName[C_JSON_KEY_FIELD + IntToStr(num)];
      if sFieldJson = '' then
        Break;
      sFieldName := json.ValueByName[C_JSON_KEY_FIELD
        + IntToStr(num) + '.' + C_JSON_KEY_NAME];
      sFilePath := json.ValueByName[C_JSON_KEY_FIELD
        + IntToStr(num) + '.' + C_JSON_KEY_PATH];
      if Copy(sFilePath,2,1) = ':' then         // 第二个字符是冒号说明是一个盘符
        // 绝对路径
        sFileFullPath := sFilePath
      else if Copy(sFilePath,1,2) = '\\' then   // 网络路径
        // 网络路径
        sFileFullPath := sFilePath
      else
      begin
        if FbInScript and (FsLastScript <> '') then
          sFileFullPath := ExtractFilePath(FsLastScript) + sFilePath
        else
          sFileFullPath := ExtractFilePath(Application.ExeName) + sFilePath;
      end;      
      if not FileExists(sFilePath) then
      begin
        AddErrorLog('找不到文件'+sFileFullPath);
        Exit;
      end;  
           
      blobfield := qry.FieldByName(sFieldName) as TBlobField;
//      stream := TMemoryStream.Create;
      slst := TStringList.Create;
//      fileStream := TFileStream.Create(sFilePath, fmOpenRead);
      try
        slst.LoadFromFile(sFilePath);
//        stream.LoadFromFile(sFilePath);
//        stream.CopyFrom(fileStream, fileStream.Size);
//        blobfield.LoadFromStream(stream);
        blobfield.AsString := slst.Text;
        qry.Post;                               // 此处提交修改 使blob的修改生效
      finally
//        fileStream.Free;
//        stream.Free;
        slst.Free;
      end;
      Result := num;
      Inc(num);
    end;
  finally
    json.Free;
    qry.Free;
  end;
end; 

function TOracleDBConnect.UpdateClobs(JsonStr: string): Integer;
var
  sSql: string;
  sFieldJson: string;
  sFieldName, sFilePath: string;
  sFileFullPath: string;
  json: TJSON;
  num: Integer;
  qry: TDataSet;
  stream: TMemoryStream;
//  fileStream: TFileStream;
  blobfield: TBlobField;
begin
  Result := 0;
  json := TJSON.Create;
  qry := GetNewQuery;
  try
    json.Init(JsonStr);
    sSql := json.ValueByName[C_JSON_KEY_SQL];
    // 查询sql出错  
    if not ExecQuery(qry, sSql) then
    begin
      Result := TDBC_ERROR_EXECFAIL;
      Exit;
    end;
    qry.Edit;       // qry状态 可编辑 才可以更新select获得的数据
    // 查询sql无结果，则无更新
    if qry.Eof then
    begin
      Exit;
    end;
    
    num := 1;
    while True do
    begin
      // 根据顺序取field，字母必须是小写并且数字的顺序连续
      sFieldJson := json.ValueByName[C_JSON_KEY_FIELD + IntToStr(num)];
      if sFieldJson = '' then
        Break;
      sFieldName := json.ValueByName[C_JSON_KEY_FIELD
        + IntToStr(num) + '.' + C_JSON_KEY_NAME];
      sFilePath := json.ValueByName[C_JSON_KEY_FIELD
        + IntToStr(num) + '.' + C_JSON_KEY_PATH];
      if Copy(sFilePath,2,1) = ':' then         // 第二个字符是冒号说明是一个盘符
        // 绝对路径
        sFileFullPath := sFilePath
      else if Copy(sFilePath,1,2) = '\\' then   // 网络路径
        // 网络路径
        sFileFullPath := sFilePath
      else
      begin
        if FbInScript and (FsLastScript <> '') then
          sFileFullPath := ExtractFilePath(FsLastScript) + sFilePath
        else
          sFileFullPath := ExtractFilePath(Application.ExeName) + sFilePath;
      end;      
      if not FileExists(sFilePath) then
      begin
        AddErrorLog('找不到文件'+sFileFullPath);
        Exit;
      end;  
           
      blobfield := qry.FieldByName(sFieldName) as TBlobField;
      stream := TMemoryStream.Create;
//      fileStream := TFileStream.Create(sFilePath, fmOpenRead);
      try
        stream.LoadFromFile(sFilePath);
//        stream.CopyFrom(fileStream, fileStream.Size);
        blobfield.LoadFromStream(stream);
        qry.Post;                               // 此处提交修改 使blob的修改生效
      finally
//        fileStream.Free;
        stream.Free;
      end;
      Result := num;
      Inc(num);
    end;
  finally
    json.Free;
    qry.Free;
  end;
end;

//function TOracleDBConnect.SaveBlobs(blobFields: TFieldItemList;
//  tableInfo: TTableInfo; sDir: string; var json: TJSON): Boolean;
//var
//  sSql: string;
//  qry: TDataSet;
//  num: Integer;
//  jsonField: TJSON;
//  i: Integer;
//  sBlobFileName: string;  
//  blobfield: TBlobField; 
//  stream: TMemoryStream;
//  sCondition: string;
//begin
//  Result := False;
//  qry := GetNewQuery;
//  try
//    if tableInfo.PrimaryKeys.Count <> 0 then
//      sCondition := GetFieldsKeyValuesStr(tableInfo.PrimaryKeys)
//    else
//    begin
//      sCondition := GetFieldsKeyValuesStr(tableInfo.GetUniqueFields);
//    end;
//    if sCondition = '' then
//    begin
//      AddErrorLog('缺乏定位Blob字段的主键或唯一键。');
//      Exit;
//    end;
//    sSql := Format('select * from %s where %s', [tableInfo.TableName,
//      sCondition]);
//    if ExecQuery(qry, sSql) then
//    begin
//      num := 1;
//      if not qry.Eof then
//      begin
//        json.AddField(C_JSON_KEY_SQL, sSql);
//        // 每个字段增加一个json子对象
//        for i := 0 to blobFields.Count - 1 do
//        begin
//          blobfield := qry.FieldByName(blobFields.Items[i].FieldName) as TBlobField;
//          stream := TMemoryStream.Create;
//          try
//            blobfield.SaveToStream(stream);
//            sBlobFileName := IncludeTrailingPathDelimiter(sDir)
//              + blobFields.Items[i].FieldName
//              + '('+sCondition+')';
//            if not DirectoryExists(sDir) then
//              ForceDirectories(sDir);
//            stream.SaveToFile(sBlobFileName);
//          finally
//            stream.Free;
//          end;
//
//          jsonField := TJSON.Create;
//          jsonField.AddField(C_JSON_KEY_NAME,
//            blobFields.Items[i].FieldName);
//          jsonField.AddField(C_JSON_KEY_PATH,
//            tableInfo.TableName + C_sBlob_Dir_Name + ExtractFileName(sBlobFileName));
//                                                                                         
//          // 是field+序号数字的形式,没添加一个序号增1
//          json.AddFieldObject(C_JSON_KEY_FIELD
//            + IntToStr(num), jsonField);
//          Inc(num);
//        end;
//        Result := True;
//      end;
//    end;
//  finally
//    qry.Free;
//  end;
//end; 

//function TOracleDBConnect.SaveClobs(clobFields: TFieldItemList;
//  tableInfo: TTableInfo; sDir: string; var json: TJSON): Boolean;
//var
//  sSql: string;
//  qry: TDataSet;
//  num: Integer;
//  jsonField: TJSON;
//  i: Integer;
//  sBlobFileName: string;  
//  blobfield: TBlobField;
//  slst: TStringList;
//  sCondition: string;
//begin
//  Result := False;
//  qry := GetNewQuery;
//  try
//    if tableInfo.PrimaryKeys.Count <> 0 then
//      sCondition := GetFieldsKeyValuesStr(tableInfo.PrimaryKeys)
//    else
//    begin
//      sCondition := GetFieldsKeyValuesStr(tableInfo.GetUniqueFields);
//    end;
//    if sCondition = '' then
//    begin
//      AddErrorLog('缺乏定位Blob字段的主键或唯一键。');
//      Exit;
//    end;
//    sSql := Format('select * from %s where %s', [tableInfo.TableName,
//      sCondition]);
//    if ExecQuery(qry, sSql) then
//    begin
//      num := 1;
//      if not qry.Eof then
//      begin
//        json.AddField(C_JSON_KEY_SQL, sSql);
//        // 每个字段增加一个json子对象
//        for i := 0 to clobFields.Count - 1 do
//        begin
//          blobfield := qry.FieldByName(clobFields.Items[i].FieldName) as TBlobField;
//          slst := TStringList.Create;
//          try
//            slst.Add(blobfield.AsString);
//            sBlobFileName := IncludeTrailingPathDelimiter(sDir)
//              + clobFields.Items[i].FieldName
//              + '('+sCondition+')';
//            if not DirectoryExists(sDir) then
//              ForceDirectories(sDir);
//            slst.SaveToFile(sBlobFileName);
//          finally
//            slst.Free;
//          end;
//
//          jsonField := TJSON.Create;
//          jsonField.AddField(C_JSON_KEY_NAME,
//            clobFields.Items[i].FieldName);
//          jsonField.AddField(C_JSON_KEY_PATH,
//            tableInfo.TableName + C_sClob_Dir_Name + ExtractFileName(sBlobFileName));
//                                                                                         
//          // 是field+序号数字的形式,没添加一个序号增1
//          json.AddFieldObject(C_JSON_KEY_FIELD
//            + IntToStr(num), jsonField);
//          Inc(num);
//        end;
//        Result := True;
//      end;
//    end;
//  finally
//    qry.Free;
//  end;
//end;

procedure TOracleDBConnect.GetUsers(List: TStrings);
var
  sSql: string;
  AQry: TDataSet;
begin
  AQry := GetNewQuery;
  try
    if SystemObject then
      sSql := SQL_ORACLE_GETUSERS_ALL
    else
      sSql := SQL_ORACLE_GETUSERS_USER;
    if ExecQuery(AQry, sSql) then
    begin
      FillListByField(List, AQry, ['USERNAME']);
    end;
  finally
    AQry.Free;
  end;
end;

function TOracleDBConnect.GetProcSource(sName: string; list: TStrings): string;
    
var
  nPos: Integer;
  objectname: string;
  sSql: string;
  qry: TDataSet;
begin
  nPos := Pos('.', sName);
  if nPos > 0 then
  begin
    objectname := Copy(sName, nPos+1, MaxInt);
  end
  else
  begin
    objectname := sName;
  end;
  if SystemObject then
    sSql := Format(SQL_ORACLE_GETPROC_SOURCE_ALL, [objectname])
  else
    sSql := Format(SQL_ORACLE_GETPROC_SOURCE_USER, [objectname]);
  qry := GetNewQuery;
  try
    if ExecQuery(qry, sSql) then
    begin
      while not qry.Eof do
      begin
        list.Add(qry.FieldByName('TEXT').AsString);
        qry.Next;
      end;  
    end;
    if list.Count > 0 then
      list[0] := 'CREATE OR REPLACE ' + list[0];
    Result := list.Text;
  finally
    qry.Free;
  end;
end;

function TOracleDBConnect.GetViewSource(sName: string; list: TStrings): string;
var
  nPos: Integer;
  objectname: string;
  sSql: string;
  qry: TDataSet;
begin
  nPos := Pos('.', sName);
  if nPos > 0 then
  begin
    objectname := Copy(sName, nPos+1, MaxInt);
  end
  else
  begin
    objectname := sName;
  end;
  if SystemObject then
    sSql := Format(SQL_ORACLE_GETVIEW_SOURCE_ALL, [objectname])
  else
    sSql := Format(SQL_ORACLE_GETVIEW_SOURCE_USER, [objectname]);
  qry := GetNewQuery;
  try           
    if ExecQuery(qry, sSql) then
    begin
      list.Text := qry.FieldByName('TEXT').AsString;
    end; 
    Result := list.Text;
  finally
    qry.Free;
  end;
end;

function TOracleDBConnect.GetTrigSource(sName: string;
  list: TStrings): string;
begin
  Result := GetProcSource(sName, list);
end;

function TOracleDBConnect.GetNewTableInfo: TTableInfo;
begin
  Result := TOracleTableInfo.Create;
end;

end.
