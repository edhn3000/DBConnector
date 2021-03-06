{**
 * <p>Title: U_DBUtil  </p>
 * <p>Copyright: Copyright (c) 2006/7/24</p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: TDBUtil数据库功能类 </p>
 *}
 
unit U_DBUtil;

interface

{$M+}

uses
  Windows, ADODB, DB, Classes, Variants,
  U_DBEngineInterface, U_DBEngineFactory, U_FieldInfo;

type
{ TDBUtil 主要保存一些不需要连接数据库或跨数据库对象的功能}
  TDBUtil = class
  private
  public
    class function getLastError: string;

//    class procedure FillTableUpdateScript(oldDB, newDB: TDBConnect;
//      oldTableName, newTableName: string; resultList: TStrings);
//    // 比较两数据库
//    class procedure FillDBUpdateScript(oldDB, newDB: TDBConnect;
//      resultList: TStrings);
//    class function CopyDB(sourceDB, destDB: TDBConnect;
//      tableInfoList: TTableInfoList);


//    class procedure ExportToSql(AQ: TADOQuery; dbtype: TDBType; sFileName: string);

    // 获取有效字段值，可用于sql中的值
//    class function GetValidFieldValue(FieldItem: TFieldItem; eDBType: TDBType;
//        const sValueDef: string = ''): string;overload;
//    class function GetValidFieldValue(Field: TField; eDBType: TDBType;
//        const sValueDef: string = ''): string;overload;
    // 以key=value,key=value......形式返回结果
//    class function GetFieldsKeyValuesStr(Fields: TFieldItemList;Adbt: TDBType): string;
//    class function GetFieldsValuesStr(Fields: TFieldItemList;Adbt: TDBType): string;
    //
    class function FindFieldsByFields(Aqry: TDataSet; fields: TFieldItemList): TFieldItemList;
    class function FindFieldsByField(Aqry: TDataSet; field: TFieldItem): TFieldItemList;
    // 使用DBConfig对象的信息连接数据库，参数传入DBConnect对象
//    class function OpenDBWithDBConfig(dbConnect: TDBConnect; dbCfg: TDBConfig): Boolean;
                                           
    class procedure LoadODBCNames(slst: TStrings);
    class procedure LoadMysqlDatabaseNames(odbcName: string; slst: TStrings);
    class procedure LoadOracleTnsNames(list: TStrings);
  end;

implementation

uses
  SysUtils, Registry;

var
  DBUtilLastError: string;


{ TDBUtil }

class function TDBUtil.FindFieldsByFields(Aqry: TDataSet;
  fields: TFieldItemList): TFieldItemList;
var
  field: TFieldItem;
  fieldsInQry: TFieldItemList;
begin
  Result := nil;
  field := fields.GetPrimaryField;
  if field <> nil then
  begin
    Result := FindFieldsByField(Aqry, field);
  end
  else
  begin
    begin
      Aqry.First;
      while not Aqry.Eof do
      begin
        fieldsInQry := TFieldItemList.Create(Aqry.Fields);
        if fields.IsSimilarFieldList(fieldsInQry) then
        begin
          Result := fieldsInQry;
          Break;
        end;
        Aqry.Next;
      end;
    end;
  end;
end;

class function TDBUtil.FindFieldsByField(Aqry: TDataSet; field: TFieldItem): TFieldItemList;
var
  fieldFind: TField;
begin
  Result := nil;
  with Aqry do
  begin
    First;
    while not Eof do
    begin
      fieldFind := Fields.FindField(field.FieldName);
      if fieldFind <> nil then
      begin
        if fieldFind.AsString = field.AsString then
        begin
          Result := TFieldItemList.Create(Fields);
          Break;
        end;
      end;
      Next;
    end;
  end;
end;

//class procedure TDBUtil.FillTableUpdateScript(oldDB, newDB: TDBConnect;
//  oldTableName, newTableName: string; resultList: TStrings);
//var
//  qryOld, qryNew: TDataSet;
//  oldField, newField: TFieldItem;
//  oldFields, newFields: TFieldItemList;
//  oldTable, newTable: TTableInfo;
//  sSql: string;
//  sqlbuilder: TSQLBuilder;
//  i: Integer;
//begin
//  qryOld := oldDB.DataSet;
//  qryNew := NewDB.DataSet;
//  resultList.BeginUpdate;
//  sqlbuilder := TSQLBuilder.Create(oldDB.DBEngine.GetDBType);
//  try
//    if newTableName = '' then 
//    begin
//      // drop table
//      resultList.Add(Format('drop table %s;', [oldTableName]));
//      Exit;    
//    end;
//    
//    if oldTableName = '' then
//    begin
//      sqlbuilder.SQLCommandType := sctCreate;
//      sqlbuilder.TableInfo.TableName := newTableName;
//      sqlbuilder.Fields.Clear;      
//      newTable := oldDB.GetTableInfo(newTableName);
//      sqlbuilder.AddFields(newTable.Fields);
//      sqlbuilder.PrimaryKeys := newTable.PrimaryKeys;
//      sqlbuilder.BuildSQL();
//      // create table
//      Exit;
//    end;
//    
//    oldTable := oldDB.GetTableInfo(oldTableName);
//    newTable := oldDB.GetTableInfo(newTableName);
//
//    // 获得新旧数据库所有数据
//    sSql := Format('select * from %s', [oldTable.TableName]);
//    oldDB.ExecQuery(sSql);
//    sSql := Format('select * from %s', [newTable.TableName]);
//    newDB.ExecQuery(sSql);
//    // TODO: 一定要有主键配置，否则不处理
//    if oldTable.PrimaryKeys.Count = 0 then
//    begin
//      oldTable.SetPrimaryKey(oldTable.Fields.Items[0].FieldName, '', False);
//      newTable.SetPrimaryKey(oldTable.Fields.Items[0].FieldName, '', False);
//    end;           
//    sqlbuilder.TableInfo.TableName := newTable.TableName;
//
//    // alter table
//    // alter table add field
//    // alter table drop field
//    // alter table modify field
//
//
//    // deltete from oldtable
//    qryOld.First;
//    while not qryOld.Eof do
//    begin
//      sqlbuilder.Fields.Clear;
//      oldFields := TFieldItemList.Create(qryOld.Fields);
//      oldFields.Items[0].IsPrimary := True;
//      newFields := FindFieldsByFields(qryNew, oldFields);
//      if newFields = nil then
//      begin
//        sqlbuilder.Fields.Clear;
//        sqlbuilder.AddFields(oldFields);
//        // delete没有主键的时候如何做删除
//        sqlbuilder.SQLCommandType := sctDelete;
//        resultList.Add(sqlbuilder.BuildSQL());
//      end;
//      oldFields.Free;
//      if Assigned(newFields) then
//        FreeAndNil(newFields);
//      qryOld.Next;
//    end;
//
//    // insert into oldtable
//    // update oldtable
//    qryNew.First;
//    while not qryNew.Eof do
//    begin
//      sqlbuilder.Fields.Clear;
//      newFields := TFieldItemList.Create(qryNew.Fields);
//      newFields.Items[0].IsPrimary := True;
//      oldFields := FindFieldsByFields(qryNew, newFields);
//
//      // insert
//      if oldFields = nil then
//      begin                   
//        sqlbuilder.Fields.Clear;
//        sqlbuilder.AddFields(newFields);
//        sqlbuilder.SQLCommandType := sctInsert;
//        sqlbuilder.BuildSQL();
//      end
//      // update
//      else
//      begin        
//        // 字段列表比较的update
//        for i := 0 to newFields.Count - 1 do
//        begin
//          newField := newFields.Items[i];
//          oldField := oldFields.FindFieldItem(newField.FieldName);
//          if oldField.AsString <> newField.AsString then
//          begin
//            sqlbuilder.Fields.Add(newField);
//          end;
//        end;
//        sqlbuilder.SQLCommandType := sctUpdate;
//        if sqlbuilder.Fields.Count > 0 then
//          sqlbuilder.BuildSQL();
//      end;
//
//      qryNew.Next;
//    end;
//
//  finally
//    sqlbuilder.Free;
//    ResultList.EndUpdate;
//  end;
//end;

//class procedure TDBUtil.FillDBUpdateScript(oldDB, newDB: TDBConnect;
//  resultList: TStrings);
//var
//  tableListOld, tableListNew: TTableInfoList;
//  qryOld, qryNew: TDataSet;
//  sqlbuilder: TSQLBuilder;
//  i: Integer;
//  tableFind: TTableInfo;
//begin
//  qryOld := oldDB.DataSet;
//  qryNew := NewDB.DataSet;
//  ResultList.BeginUpdate;
//  sqlbuilder := TSQLBuilder.Create(oldDB.DBEngine.GetDBType);
//  oldDB.SystemObject := False;
//  newDB.SystemObject := False;
//
//  tableListOld := oldDB.GetTableInfos(qryOld);
//  tableListNew := newDB.GetTableInfos(qryNew);
//  try
//    // compare and drop table
//    for i := 0 to tableListOld.Count - 1 do
//    begin
//      tableFind := tableListNew.Find(tableListOld.Items[i].TableName);
//      // if not tableFind=nil  the function can return drop table sql
//      FillTableUpdateScript(oldDB, newDB, tableListOld.Items[i].TableName,
//        tableFind.TableName, resultList);
//    end;
//
//    // create table
//    for i := 0 to tableListNew.Count - 1 do
//    begin
//      tableFind := tableListOld.Find(tableListNew.Items[i].TableName);
//      if tableFind = nil then
//        FillTableUpdateScript(oldDB, newDB, '', tableListNew.Items[i].TableName,
//          resultList);
//    end;
//
//    // view
//    // trigger
//    // procedure
//  finally
//    sqlbuilder.Free;
//    tableListOld.Free;
//    tableListNew.Free;
//    ResultList.EndUpdate;
//  end;
//end;

//class function TDBUtil.GetValidFieldValue(FieldItem: TFieldItem; eDBType: TDBType;
//  const sValueDef: string): string;
//var
//  db: IDBConnect;
//begin
//  db := TDBConnectManager.CreateDBConnect(eDBType);
//  try
//    Result := db.GetValidFieldValue(FieldItem);
//  finally
//    db := nil;
//  end;
//end;

//class function TDBUtil.GetValidFieldValue(Field: TField; eDBType: TDBType;
//    const sValueDef: string = ''): string;
//var
//  fld: TFieldItem;
//begin
//  fld := TFieldItem.Create(Field);
//  try
//    Result := GetValidFieldValue(fld, eDBType, sValueDef);
//  finally
//    fld.Free;
//  end;
//end;

//class function TDBUtil.OpenDBWithDBConfig(dbConnect: TDBConnect;
//  dbCfg: TDBConfig): Boolean;
//begin
//  Result := dbConnect.OpenDB(dbCfg.GetSeparatedDataSource,
//   dbCfg.UserName, dbCfg.Password,
//   dbCfg.DBType, dbCfg.DBEngineType);
//end;

//class function TDBUtil.getTableNameBySQL(sSql: string): string;
//var
//  nPos1, nPos2: Integer;
//  sTable: string;
//begin
//  // TODO regexp
//  sSql := StringReplace(sSql, #$D, ' ', [rfReplaceAll]);  
//  sSql := StringReplace(sSql, #$A, ' ', [rfReplaceAll]);
//  nPos1  := fStrUtil.PosFrom('from', sSql);
//  if (nPos1 <> 0) then
//  begin
//    sTable := Trim(Copy(sSql, nPos1+4, MaxInt));
//    nPos2  := fStrUtil.PosFrom(' ', sTable);   // 再找下一个空格
//    if nPos2 = 0 then
//      nPos2 := MaxInt;
//    sTable := Trim(Copy(sTable, 1, nPos2));
//  end
//  else
//    sTable := 'UnknownTable';
//  Result := sTable;
//end;   

class function TDBUtil.getLastError: string;
begin
  Result := DBUtilLastError;
end;

//class function TDBUtil.GetFieldsKeyValuesStr(
//  Fields: TFieldItemList;Adbt: TDBType): string;
//var
//  i: Integer;
//  sCondition: string;
//begin
//  sCondition := '';
//  for i := 0 to Fields.Count - 1 do
//  begin
//    if sCondition = '' then
//      sCondition := Format('%s=%s',[Fields.Items[i].FieldName,
//        TDBUtil.GetValidFieldValue(Fields.Items[i], Adbt)])
//    else
//      sCondition := sCondition + ','+ Format('%s=%s',[
//        Fields.Items[i].FieldName,
//        TDBUtil.GetValidFieldValue(Fields.Items[i], Adbt)]);
//  end;
//  Result := sCondition;
//end;

//class function TDBUtil.GetFieldsValuesStr(Fields: TFieldItemList;
//  Adbt: TDBType): string;
//var
//  i: Integer;
//  s: string;
//begin
//  for i := 0 to Fields.Count - 1 do
//  begin
//    if s = '' then
//      s := TDBUtil.GetValidFieldValue(Fields.Items[i], Adbt)
//    else
//      s := s + ','+ TDBUtil.GetValidFieldValue(Fields.Items[i], Adbt);
//  end;
//  Result := s;
//end;

class procedure TDBUtil.LoadODBCNames(slst: TStrings);
var
  reg: TRegistry;
begin      
  reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKeyReadOnly('Software\ODBC\ODBC.INI\ODBC Data Sources') then
    begin
      Reg.GetValueNames(slst);
    end;
  finally
    reg.Free;
  end;
end;

class procedure TDBUtil.LoadMysqlDatabaseNames(odbcName: string;
  slst: TStrings);
var
  engine: IDBEngine;
  qry: TDataSet;
begin
  engine := TDBEngineFactory.GetNewDBEngine(dbetAdo);
  try
    if not engine.OpenDB(odbcName, '', '', dbtMySql) then
      Exit;
    qry := engine.GetNewQuery;
    if engine.ExecQuery(qry,
      'select schema_name from information_schema.schemata') then
    begin
      while not qry.Eof do
      begin
        slst.Add(qry.Fields[0].AsString);
        qry.Next;
      end;
    end;
    engine.CloseDB;
  finally
    engine := nil;
  end;   
end;

class procedure TDBUtil.LoadOracleTnsNames(list: TStrings);
var
  Reg: TRegistry;
  sOraHome: string;
  sTnsFileName: string;
  slstFile: TStrings;
  sLine: string;
  i: Integer;
  nPos: Integer;
begin
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    // 10g
    if Reg.OpenKeyReadOnly('SOFTWARE\ORACLE\KEY_OraDb10g_home1')
//       or Reg.OpenKeyReadOnly('SOFTWARE\ORACLE\KEY_OraDb10g_home1')
       then
    begin
      sOraHome := Reg.ReadString('ORACLE_HOME');
      sTnsFileName := sOraHome + '\NETWORK\ADMIN\tnsnames.ora';
      if not FileExists(sTnsFileName) then
        Exit;

      slstFile := TStringList.Create;
      try
        slstFile.LoadFromFile(sTnsFileName);
        for i := 0 to slstFile.Count - 1 do
        begin
          sLine := Trim(slstFile[i]);
          if Copy(sLine, 1, 1) = '#' then
            Continue;

          nPos := Pos('=', sLine);
          if (sLine <> '') and (nPos = Length(sLine))
             and (Copy(sLine, 1, 1) <> '(') then
          begin
            list.Add(Trim(Copy(sLine, 1, Length(sLine)-1)));
          end;
        end;
      finally
        slstFile.Free;
      end;
    end;
  finally
    Reg.Free;
  end;
end;

initialization

finalization

end.
