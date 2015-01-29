{**
 * <p>Title: U_AccessDBConnect.pas  </p>
 * <p>Copyright: Copyright (c) 2009-4-12</p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: 针对access的实现 </p>
 *}
unit U_AccessDBConnect;

interface

uses
  ADODB, DB, Classes, Variants,
  Dialogs, SysUtils, U_TableInfo, U_AccessTableInfo,
              
  U_DBConnect, U_DBEngineInterface;

const
  SQL_ACCESS_GETQUERY  =
   'SELECT NAME FROM MsysObjects WHERE (Left([Name],1)<>"~") AND (Type)=5';         // 查询
  SQL_ACCESS_GETTABLE  =
   'SELECT NAME FROM MsysObjects WHERE (Left([NAME],1)<>"~") AND (Left$([NAME],4) <> "Msys") AND (TYPE)=1'; // 表
  SQL_ACCESS_GETMOUDLE =
   'SELECT NAME FROM MsysObjects WHERE (Left([Name],1)<>"~") AND (Type)= -32761';   // 模块
  SQL_ACCESS_GETREPORT =
   'SELECT NAME FROM MsysObjects WHERE (Left([Name],1)<>"~") AND (Type)= -32764';   // 报表
  SQL_ACCESS_GETMACRO  =
   'SELECT NAME FROM MsysObjects WHERE (Left([Name],1)<>"~") AND (Type)= -32766';    // 宏
  SQL_ACCESS_GETFORM   =
   'SELECT NAME FROM MsysObjects WHERE (Left([Name],1)<>"~") AND (Type)=-32768';     // 窗体
  SQL_ACCESS_GETPAGE   = '';   // 页
  
type
  TAccessDBConnect = class(TDBConnect)
  protected
    function GetNewTableInfo(): TTableInfo;override;
  public
    constructor Create;override;
    destructor Destroy;override;
    
    procedure GetViewNames(List: TStrings; AQry: TDataSet = nil);override;

    // own
    procedure GetMoudleNames(List: TStrings; AQry: TDataSet = nil);override;
    procedure GetReportNames(List: TStrings; AQry: TDataSet = nil);override;
    procedure GetMacroNames(List: TStrings; AQry: TDataSet = nil);override;
    procedure GetFormNames(List: TStrings; AQry: TDataSet = nil);override;
    procedure GetPageNames(List: TStrings; AQry: TDataSet = nil);override;
  end;

implementation

{ TAccessDBConnect }

constructor TAccessDBConnect.Create;
begin
  inherited;
  SetDBType(dbtAccess);
end;

destructor TAccessDBConnect.Destroy;
begin

  inherited;
end;

procedure TAccessDBConnect.GetMacroNames(List: TStrings;
  AQry: TDataSet);
begin
  GetObjectNames(List, SQL_ACCESS_GETMOUDLE, AQry);
end;

procedure TAccessDBConnect.GetMoudleNames(List: TStrings;
  AQry: TDataSet);
begin
  GetObjectNames(List, SQL_ACCESS_GETMOUDLE, AQry);
end;

procedure TAccessDBConnect.GetReportNames(List: TStrings;
  AQry: TDataSet);
begin
  GetObjectNames(List, SQL_ACCESS_GETREPORT, AQry);
end;

procedure TAccessDBConnect.GetFormNames(List: TStrings;
  AQry: TDataSet);
begin
  if IsQryInvalid(AQry) then
    AQry := DBEngine.GetDataSet;
  GetObjectNames(List, SQL_ACCESS_GETFORM, AQry);
end;  

procedure TAccessDBConnect.GetPageNames(List: TStrings;
  AQry: TDataSet);
begin
  GetObjectNames(List, SQL_ACCESS_GETPAGE, AQry);
end;

procedure TAccessDBConnect.GetViewNames(List: TStrings; AQry: TDataSet);
begin  
  GetObjectNames(List, SQL_ACCESS_GETQUERY, AQry);
end;

function TAccessDBConnect.GetNewTableInfo: TTableInfo;
begin
  Result := TAccessTableInfo.Create;
end;

end.
