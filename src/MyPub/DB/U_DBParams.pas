unit U_DBParams;

interface

uses
  U_DBEngine;

type
  TDBParams = class
  private
    FsDataSource: string;
    FsSecured: string;
    FsUser: string;
    FsPwd: string;
  public
    // 数据库类型
    DBType: TDBType; 
    DBEngineType: TDBEngineType;

    // access数据库
    DB: string;
    Secured: string;
    AcsUser: string;
    AcsPwd: string;
    LastDir: string;

    // oracle数据库
    SID: string;
    IP: string;
    Protocol: string;
    Port: string;
    OraUser: string;
    OraPwd: string;
    IsLocalTns: Boolean;

    // Sybase数据库
    ServerName: string;
    DatabaseName: string;
    SybUser: string;
    SybPwd: string;

    // MySQL
    mslHost: string;
    mslDataBase: string;
    mslUser:string;
    mslPwd: string;

  public
    // 根据数据库类型返回可用DataSource，以上public字段皆须被赋值
    // 标准为可被DBConnector使用
    function getDataSource: string;
    // 根据数据库类型 返回可用user
    function getUser: string;    
    // 根据数据库类型 返回可用pwd
    function getPwd: string;
  end;

implementation

{ TDBParams }

function TDBParams.getDataSource: string;
begin
   case DBType of
    dbtAccess:
    begin
      sDataSource := GetValidFullPath(DB);
      if FileExists(DB) then
        FileSetAttr(sDataSource, FileGetAttr(sDataSource) and (not SysUtils.faReadOnly));
      sUser := AcsUser;
      sPwd := AcsPwd;
      sSecured := GetValidFullPath(Secured);
//      ConnectedSQLDialect := sqlMSSQL2K;
    end;
    dbtAccess2007:
    begin
      sDataSource := GetValidFullPath(DB); 
      if FileExists(DB) then
        FileSetAttr(sDataSource, FileGetAttr(sDataSource) and (not SysUtils.faReadOnly));
      sUser := AcsUser;
      sPwd := AcsPwd;
      sSecured := GetValidFullPath(Secured);
//      ConnectedSQLDialect := sqlMSSQL2K;
    end;
    dbtOracle:
    begin
      if IsLocalTns or (IP = '') then
        sDataSource := SID
      else
        sDataSource := Format('%1:s%0:s%2:s%0:s%3:s%0:s%4:s',
                            [C_sSEPARATOR_DATASOURCE, SID, IP, Port, Protocol]);
      sUser := OraUser;
      sPwd := OraPwd;
      sSecured := '';
      ConnectedSQLDialect := sqlOracle;
    end;
    dbtSyBase:
    begin
      sDataSource := ServerName;
      sSecured := DatabaseName;
      sUser := SybUser;
      sPwd := SybPwd;
//      ConnectedSQLDialect := sqlSybase;
    end;
    dbtMySQL:
    begin
      sDataSource := Format('%1:s%0:s%2:s', [
        C_sSEPARATOR_DATASOURCE, mslHost, mslDataBase]);
      sUser := mslUser;
      sPwd := mslPwd;
//      ConnectedSQLDialect := sqlMySQL;
    end;
    else
    begin
//      ConnectedSQLDialect := sqlStandard;
    end;
  end;
end;

function TDBParams.getPwd: string;
begin

end;

function TDBParams.getUser: string;
begin

end;

end.
