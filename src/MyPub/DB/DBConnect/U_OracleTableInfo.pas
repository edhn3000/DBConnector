unit U_OracleTableInfo;

interface

uses
  ADODB, DB, Classes, Contnrs, Variants, SysUtils,
  U_OracleFieldInfo, U_TableInfo, U_FieldInfo;

const
  // P主键 R外键 C不清楚
  SQL_ORACLE_GETTABLEINFO_USER   =
      'select tab.TABLE_NAME, '
          + 'tabcom.COMMENTS, '
          + 'tcons.COLUMN_NAME PRIMARYKEY, '
          + 'tcons.CONSTRAINT_NAME,'
          + 'tcons.OWNER  '
     + 'from USER_TABLES       tab, '
          + 'USER_TAB_COMMENTS tabcom, '
          + '(select conscol.OWNER, conscol.COLUMN_NAME, conscol.TABLE_NAME, '
          + ' cons.CONSTRAINT_NAME '
          + 'from USER_CONS_COLUMNS conscol, USER_CONSTRAINTS cons '
          + 'where cons.CONSTRAINT_NAME = conscol.CONSTRAINT_NAME '
          + 'and cons.CONSTRAINT_TYPE = ''P'') tcons '
    + 'where tab.TABLE_NAME = tabcom.TABLE_NAME '
         +  'and tab.TABLE_NAME = tcons.TABLE_NAME(+) '
         +  'and tab.TABLE_NAME = ''%s''';

  SQL_ORACLE_GETTABLEINFO_ALL    =
      'select tab.TABLE_NAME, '
          + 'tabcom.COMMENTS, '
          + 'tcons.COLUMN_NAME PRIMARYKEY, '
          + 'tcons.CONSTRAINT_NAME,'
          + 'tcons.OWNER  '
     + 'from ALL_TABLES       tab, '
          + 'ALL_TAB_COMMENTS tabcom, '
          + '(select conscol.OWNER, conscol.COLUMN_NAME, conscol.TABLE_NAME, '
          + ' cons.CONSTRAINT_NAME '
          + 'from ALL_CONS_COLUMNS conscol, ALL_CONSTRAINTS cons '
          + 'where cons.CONSTRAINT_NAME = conscol.CONSTRAINT_NAME '
          + 'and cons.CONSTRAINT_TYPE = ''P'') tcons '
    + 'where tab.TABLE_NAME = tabcom.TABLE_NAME '
         +  'and tab.TABLE_NAME = tcons.TABLE_NAME(+) '
         +  'and tab.TABLE_NAME = ''%s''';

type
  TOracleTableInfo = class(TTableInfo)
  protected
    procedure initObjects();override;
  public
    function GetNewFieldItem: TFieldItem;override;
    function InitTableInfo(tableName: string): Boolean;override;
  end;

implementation

{ TOracleTableInfo }

function TOracleTableInfo.GetNewFieldItem: TFieldItem;
begin
  Result := TOracleFieldItem.Create;
end;

procedure TOracleTableInfo.initObjects;
begin
  Fields := TOracleFieldItemList.Create();
  FUniqueFields := TOracleFieldItemList.Create();
  FPrimaryKeys := TOracleFieldItemList.Create(False);
end;

function TOracleTableInfo.InitTableInfo(tableName: string): Boolean;
var
  qry, qry2: TDataSet;
  sSql: string;
  sPriKey: string;
begin
  Result := False;

  Self.TableName := tableName;  
  qry := FDB.GetNewQuery;
  qry2 := FDB.GetNewQuery;
  try
    Fields.DB := FDB;              
    // 先初始化字段列表，之后再设置主键
    Fields.InitFieldInfos(tableName);

    if SystemObject then
      sSql := Format(SQL_ORACLE_GETTABLEINFO_ALL, [UpperCase(tableName)])
    else
      sSql := Format(SQL_ORACLE_GETTABLEINFO_USER, [UpperCase(tableName)]);
      
    if not FDB.ExecQuery(qry, sSql) then
      Exit;
    if not qry.Eof then
    begin
//      table.TableName  := qry.FieldByName('table_name').AsString;
      Comment    := qry.FieldByName('COMMENTS').AsString;
      OWNER      := qry.FieldByName('OWNER').AsString;
      // 设置主键，有几条记录就是几个主键
      while not qry.Eof do
      begin
        sPriKey := qry.FieldByName('PRIMARYKEY').AsString;
        if sPriKey <> '' then
        begin
          if SetPrimaryKey(sPriKey, '', False) then
          begin
            Fields.FindFieldItem(sPriKey)
              .ConstraintName := qry.FieldByName('CONSTRAINT_NAME').AsString;
          end;
        end;
        qry.Next;
      end;
    end;
    // 无主键
    if PrimaryKeys.Count = 0 then
    begin
      CheckFieldsUnique();
    end;
  finally
    qry.Free;
    qry2.Free;
  end;
end;

end.
