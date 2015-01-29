unit U_DBExport;

interface
uses
  Windows, Messages, Dialogs, Forms, Controls, Classes, Types,
  Graphics, ADODB, DB, OleServer,
  SysUtils, U_UIUtil, U_DBConnectInterface;

type     
  TDBExportType = (detSql, detXls, detTxt);
  
  TDBExport = class
  private
    FDBConnect: IDBConnect;
    function GetDBConnect: IDBConnect;
    procedure SetDBConnect(value: IDBConnect);
  public
    property DBConnect: IDBConnect read GetDBConnect write SetDBConnect;
  public
    function ExportTableData(tableName:string; exportType: TDBExportType;
      sExportFileName: string): Boolean;
    
  end;
        
  function DBExportTypeToStr(Adet: TDBExportType;
    bWithPrefix: Boolean = True): string;
  function StrToDBExportType(s: string): TDBExportType;

implementation

uses
  TypInfo;

{ TDBExport }

function DBExportTypeToStr(Adet: TDBExportType;
  bWithPrefix: Boolean = True): string;
var
  ti: PTypeInfo;
  td: PTypeData;
  i: Integer;
begin
  ti := TypeInfo(TDBExportType);
  td := GetTypeData(ti);
  for i := td^.MinValue to td^.MaxValue do
    if Integer(Adet) = i then
    begin
      Result := GetEnumName(ti, i);
      if not bWithPrefix then
        Result := Copy(Result, 4, Length(Result));
      Break;
    end;
end;

function StrToDBExportType(s: string): TDBExportType;
var
  nDET, nCode: Integer;
  ti: PTypeInfo;
  td: PTypeData;
  det: TDBExportType;
begin  
  Result := detSql;
  Val(s, nDET, nCode);
  if (nCode = 0) then
  begin
    ti := TypeInfo(TDBExportType);
    td := GetTypeData(ti);
    if (nDET <= td^.MaxValue) and (nDET >= td^.MinValue) then
      Result := TDBExportType(nDET);
  end
  else
  begin
    for det := Low(TDBExportType) to High(TDBExportType) do
      if SameText(DBExportTypeToStr(det, False), s) then
      begin
        Result := det;
        Break;
      end;  
  end;  
end; 

function TDBExport.ExportTableData(tableName: string;
  exportType: TDBExportType; sExportFileName: string): Boolean;
const
  C_sSeparatorLine = '-';
var
  slstExport: TStringList;
  AQry: TDataSet;
begin
  Result := False;
  AQry := FDBConnect.GetNewQuery;
  if not FDBConnect.ExecQuery(AQry, Format('select * from %s', [tableName])) then
    Exit;
  case exportType of
  detXls:  
    Result := UIUtil.ExportToExcel(AQry, sExportFileName, tableName);  
  detSql:
    Result := FDBConnect.ExportQuery(AQry, sExportFileName);
  else
  begin
    slstExport := TStringList.Create;
    try
      UIUtil.FillExportQueryList(AQry, slstExport, '  ', '-');
      slstExport.SaveToFile( sExportFileName );
      Result := True;
    finally
      slstExport.Free;
    end;
  end
  end;
  AQry.Free;
end;

function TDBExport.GetDBConnect: IDBConnect;
begin
  Result := FDBConnect;
end;

procedure TDBExport.SetDBConnect(value: IDBConnect);
begin
  FDBConnect := value;
end;

end.
