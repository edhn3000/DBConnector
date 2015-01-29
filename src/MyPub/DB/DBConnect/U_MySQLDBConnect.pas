{**
 * <p>Title: U_MySQLDBConnect.pas  </p>
 * <p>Copyright: Copyright (c) 2009-4-12</p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: 针对mysql的实现 </p>
 *}
unit U_MySQLDBConnect;

interface 

uses
  ADODB, DB, Classes, Variants,
  Dialogs, SysUtils, U_TableInfo, U_MySqlTableInfo,
  U_FieldInfo, 
  U_DBConnect, U_DBEngineInterface;
type
  TMySQLDBConnect = class(TDBConnect)
  protected
    function GetNewTableInfo(): TTableInfo;override;
    function GetMaxRecordSQLCondition(nMaxRecord: Integer;
      sSql: string): string;override;
  public
    constructor Create;override;
    destructor Destroy;override;    
    function GetValidFieldValue(FieldItem: TFieldItem;
      const sValueDef: string): string;override;      
//      procedure GetTableNames(List: TStrings;
//      AQry: TDataSet = nil);
  end;

implementation

{ TMySQLDBConnect }

constructor TMySQLDBConnect.Create;
begin
  inherited;
  SetDBType(dbtMySql);
end;

destructor TMySQLDBConnect.Destroy;
begin

  inherited;
end;

function TMySQLDBConnect.GetMaxRecordSQLCondition(nMaxRecord: Integer;
  sSql: string): string;
begin
  if nMaxRecord > 0 then
    Result := sSql + ' limit ' + IntToStr(nMaxRecord)
  else
    Result := sSql;
end;

function TMySQLDBConnect.GetNewTableInfo: TTableInfo;
begin
  Result := TMySqlTableInfo.Create;
end;

function TMySQLDBConnect.GetValidFieldValue(FieldItem: TFieldItem;
  const sValueDef: string): string;
var
  sFieldValue: string;
begin
  sFieldValue := FieldItem.AsString;
  case FieldItem.DataType of
    ftDate, ftTime, ftDateTime:
    begin
      if Length(sFieldValue) <= 10 then
        Result := Format('STR_TO_DATE(''%s'',''%Y-%m-%d'')', [sFieldValue])
      else if Length(sFieldValue) <= 19 then
        Result := Format('STR_TO_DATE(''%s'',''%Y-%m-%d %h:%i:%s'')',
          [sFieldValue])
      else if Length(sFieldValue) > 19 then
        Result := Format('STR_TO_DATE(''%s'',''%Y-%m-%d %h:%i:%s'')',
          [sFieldValue])
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

end.
