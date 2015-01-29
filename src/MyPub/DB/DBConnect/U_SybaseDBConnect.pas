{**
 * <p>Title: U_SybaseDBConnect.pas  </p>
 * <p>Copyright: Copyright (c) 2009-4-12</p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: 针对sybase的实现 </p>
 *}
unit U_SybaseDBConnect;

interface

uses
  ADODB, DB, Classes, Variants,
  Dialogs, SysUtils, U_TableInfo, U_SybaseTableInfo,
  
  U_DBConnect, U_DBEngineInterface;
type
  TSybaseDBConnect = class(TDBConnect)
  protected
    function GetNewTableInfo(): TTableInfo;override;
  public
    constructor Create;override;
    destructor Destroy;override;
  end;

implementation

{ TSybaseDBConnect }

constructor TSybaseDBConnect.Create;
begin
  inherited;
  SetDBType(dbtSybase);
end;

destructor TSybaseDBConnect.Destroy;
begin

  inherited;
end;

function TSybaseDBConnect.GetNewTableInfo: TTableInfo;
begin
  Result := TSybaseTableInfo.Create;
end;

end.
