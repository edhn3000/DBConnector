unit U_OracleFieldInfo;

interface   

uses
  ADODB, DB, Classes, Contnrs, Variants,
  SysUtils, U_FieldInfo;
    
const                 
  SQL_ORACLE_GETFIELDS_USER =
    'select col.COLUMN_NAME, col.DATA_TYPE, col.NULLABLE, col.DATA_LENGTH, '
          + 'col.DATA_PRECISION, col.DATA_SCALE,'
          + 'colcom.COMMENTS col_comment '
    +' from USER_TABLES tab, '
        + ' USER_TAB_COLUMNS col, USER_COL_COMMENTS colcom'
    +' where tab.TABLE_NAME=col.TABLE_NAME'
         + ' and tab.TABLE_NAME = colcom.TABLE_NAME'
         + ' and col.COLUMN_NAME=colcom.column_name'
         + ' and tab.TABLE_NAME = ''%s'''
    + 'order by COLUMN_ID';
//    'select * from USER_TAB_COLUMNS '
//    +'where TABLE_NAME=''%s''';

  SQL_ORACLE_GETFIELDS_ALL =
    'select col.COLUMN_NAME, col.DATA_TYPE, col.NULLABLE, col.DATA_LENGTH, '
          +'col.DATA_PRECISION, col.DATA_SCALE,'
          +'colcom.COMMENTS col_comment'
    +' from ALL_TABLES tab, '
         +' ALL_TAB_COLUMNS col, ALL_COL_COMMENTS colcom'
    +' where tab.TABLE_NAME=col.TABLE_NAME'
         + ' and tab.TABLE_NAME = colcom.TABLE_NAME'
         + ' and col.COLUMN_NAME=colcom.COLUMN_NAME'
         + ' and tab.TABLE_NAME = ''%s'''
    + 'order by COLUMN_ID';

type
  TOracleFieldItem = class (TFieldItem)
  public
    function GetDataTypeInSQL: string;override;
  end;   

  TOracleFieldItemList = class(TFieldItemList)
  public
    function CreateFieldItem: TFieldItem;override;

    function InitFieldInfos(sTableName: string): Boolean;override;
  end;

implementation   

{ TOracleFieldItem }

function TOracleFieldItem.GetDataTypeInSQL: string;
begin
  case DataType of
  ftInteger:
  begin
    if FDataSize <> 0 then
      Result := Format('Number(%d)', [FDataSize])
    else
      Result := 'Number';
  end;
  ftFloat:   Result := Format('Number(%d, %d)', [FDataSize, FDecimal]);
  ftDate, ftDateTime: Result := 'Date';
  ftString, ftWideString:  Result := Format('Varchar2(%d)', [FDataSize]);
  ftBlob, ftOraBlob: Result := 'Blob';
  ftMemo, ftOraClob: Result := 'Clob';
  else
    Result := inherited GetDataTypeInSQL;
  end;
end;     

{ TOracleFieldItemList }

function TOracleFieldItemList.CreateFieldItem: TFieldItem;
begin
  Result := TOracleFieldItem.Create;
end;

function TOracleFieldItemList.InitFieldInfos(sTableName: string): Boolean;  
var
  fld: TFieldItem;
  sSql: string;
  qry: TDataSet;
begin
  Result := False;
  qry := FDB.GetNewQuery;
  try
    if SystemObject then
      sSql := Format(SQL_ORACLE_GETFIELDS_ALL, [sTableName])
    else
      sSql := Format(SQL_ORACLE_GETFIELDS_USER, [sTableName]);
      
    if not FDB.ExecQuery(qry, sSql) then
      Exit;
    with qry do
    begin
      Self.Clear;
      while not Eof do
      begin
        fld := Self.FindFieldItem(FieldByName('COLUMN_NAME').AsString);
        if fld = nil then
        begin
          fld := TOracleFieldItem.Create;  
          Self.Add(fld);
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
    qry.Close;
    Result := True
  finally
    qry.Free;
  end;
end;

end.
