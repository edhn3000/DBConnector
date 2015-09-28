{**
 * <p>Title: U_DBConnectorInConsole  </p>
 * <p>Copyright: Copyright (c) 2008/11/30</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: DBConnector控制台</p>
 * <p>Description: 实现结果显示在程序自己创建的控制台或显示在系统控制台</p>
 * <p>Description: 直接调用RunInConsole(True)即可完成对命令行的分析并执行</p>
 * <p>Description: 只有在程序加入 APPTYPE CONSOLE条件编译的时候才能传入参数false</p>
 *}
unit U_DBConnectorInConsole;

interface
uses
  SysUtils, Forms, Windows,
  U_LogManager, U_DBEngineInterface, U_DBConnectInterface;

type
  TWriteConsoleMode = (wcmView, wcmNotView, wcmConsole);

  TDBConnectInConsole = class
  public
    procedure OnLogInConsole(sLog: string);
  end;
  
  procedure RunInConsole(wcm: TWriteConsoleMode = wcmView);

implementation

uses
  UF_ConsoleView, U_Const, U_fStrUtil, U_Convertor, U_DBConnnectManager, U_RegexUtil;

var
  dbConnect: IDBConnect;
  DBConnectInConsole: TDBConnectInConsole;
const
  C_REGX_LOGIN = '(\w*)/(\w*)@(.*)';

// dbconnector oracle [option] kbuser/tusc@tnsname|(SID[,IP,Port,Protocol]) @script.sql
// dbconnector access [option] tax/tax@db,Secured @script.sql
function GetConnectParam(var user: string; var psw: string; var db: string;
  var sql: string; var dbt: TDBType; var sMsg: string): Boolean;
var
  nPos3: Integer;
  sDataSource, sOtherDBParam: string;
  i: Integer;
  sParam: string;
  bEncryptlogin: Boolean;
  mr: TMatchResult;
  paramRequire: Integer;
  function GetValidFullPath(Str: string):string;
  begin
    if (Trim(Str) = '') or (Pos(':', Str) <> 0)
       or (Pos('\\', Str) = 1) then
      Result := Str
    else
      Result := ExtractFilePath(Application.ExeName)+Str;
  end;
begin
  Result := False;
  bEncryptlogin := False;
  i := 0;
  { paramRequire按位操作，记录必须参数是否提供
    从后向前第1位数据库类型，第2位登录串
  }
  paramRequire := 0;
  while True do
  begin       
    Inc(i);
    sParam := ParamStr(i);
    if sParam = '' then
      Break;
             
    // 第一个参数是数据库类型
    if i = 1 then
    begin
      dbt := StrToDBType(sParam);
      if dbt = dbtUnKnown then
      begin
        sMsg := '无法从参数1分析出数据库类型';
        Exit;
      end;
      paramRequire := paramRequire or 1;
    end
    else if SameText(ParamStr(i), '-encryptlogin') then
    begin
      bEncryptlogin := True;
    end
    // 登录串处理
    else if TRegexUtil.MatchFirstStr(C_REGX_LOGIN, 0, ParamStr(i)) <> '' then
    begin
      paramRequire := paramRequire or 2;
      mr := TRegexUtil.MatchFirst(C_REGX_LOGIN, 0, ParamStr(i));
      sDataSource := mr.Groups[3].MatchStr;
      nPos3 := Pos(C_sSEPARATOR_DATASOURCE, sDataSource);     // 可以没有
      if nPos3 = 0 then begin
        db := sDataSource;
        sOtherDBParam := '';
      end else begin
        db := Copy(sDataSource, 1, nPos3);
        sOtherDBParam := Copy(sDataSource, nPos3+1, MaxInt);
      end;
      user := mr.Groups[1].MatchStr; //  Copy(sParam, 1, nPos-1);
      psw  := mr.Groups[2].MatchStr; // Copy(sParam, nPos+1, nPos2-nPos-1);
      if bEncryptlogin then
      begin
        user := Convertor.XXEUUEMIMEDecode(user);  
        psw := Convertor.XXEUUEMIMEDecode(psw);
      end;
      
      if sOtherDBParam <> '' then
      case dbt of
        dbtUnKnown:;
        dbtAccess, dbtAccess2007:
        begin
          db := BuildAccessDataSource(GetValidFullPath(db),
                GetValidFullPath(sOtherDBParam));
        end;
        dbtOracle:
          db := db + C_sSEPARATOR_DATASOURCE + sOtherDBParam;
        dbtSyBase:
          db := db + C_sSEPARATOR_DATASOURCE + sOtherDBParam;
        dbtMySql:
          db := db + C_sSEPARATOR_DATASOURCE + sOtherDBParam;
      end;
    end
    // sql脚本
    else if Pos('@', ParamStr(i)) = 1 then
    begin
      sql := ParamStr(i);
    end;  
  end;

  Result := paramRequire >= 3;
end;

procedure WriteLnInConsole(s: string; wcm: TWriteConsoleMode);
var
  n: Cardinal;
begin
  case wcm of
  wcmView:
  begin
    g_consoleForm.WriteLn(s);
  end;
  wcmNotView:
  begin

  end;
  wcmConsole:
  begin   
    Reset(Input);
    n := Length(s);
    WriteConsole(GetStdHandle(STD_OUTPUT_HANDLE), @s[1], n, n, nil);
  end;
  end;
end;          

procedure TDBConnectInConsole.OnLogInConsole(sLog: string);
begin
  WriteLnInConsole(sLog, wcmConsole);
end; 

procedure RunInConsole(wcm: TWriteConsoleMode);
var
  sUserName: string;
  sPassword: string;
  sDatabase: string;
  sSql: string;
  dbtype: TDBType;
  sMsg: string;
  lp: TLogOption;
  i: Integer;
  sShortLogDir: string;
begin
  sShortLogDir := 'log\';
  lp := TLogOption.Create;
  try
    lp.LogDir := ExtractFilePath(Application.ExeName) + sShortLogDir;
    g_Log := TLogManager.Create(lp);
  finally
    lp.Free;
  end;
  if ParamStr(1) = '/?' then begin
    WriteLnInConsole(gC_AppName+' command dbtype user/password@DBName @script.sql', wcm);
    WriteLnInConsole('For Example: ',wcm);
    WriteLnInConsole('  '+gC_AppName+' conn access user/pass@MdbFile[,MdwFile] @script.sql', wcm);
    WriteLnInConsole('  '+gC_AppName+' conn oracle user/pass@TNSName|(SID[,IP,Port,Protocol]) @script.sql', wcm);
    WriteLnInConsole('  '+gC_AppName+' conn sybase user/pass@ODBCServerName[,DatabaseName] @script.sql', wcm);
    WriteLnInConsole('  '+gC_AppName+' conn mysql  user/pass@ODBCServerName[,DatabaseName] @script.sql', wcm);
    Exit;
  end;
  if not GetConnectParam(sUserName, sPassword, sDatabase, sSql, dbtype, sMsg) then
    Application.MessageBox(PChar(sMsg), '错误', MB_OK+MB_ICONINFORMATION)
  else begin
    try
      dbConnect := TDBConnectManager.CreateDBConnect(dbtype);
      WriteLnInConsole('数据库类型'+DBTypeToStr(dbtype, False) + ',' +
        '数据库'+Format('%s@%s', [sUserName,sDatabase])+ ',' +
        'sql='+sSql,wcm);
      g_Log.AddInfoLog('', ['数据库类型'+DBTypeToStr(dbtype, False),
                            '数据库'+Format('%s@%s', [sUserName,sDatabase]),
                            'sql='+sSql],
                       '', '从命令行运行');
      if not dbConnect.OpenDB(sDatabase, sUserName, sPassword, dbtype, dbetAuto) then
          g_Log.AddInfoLog('', [], '', '连接数据库失败 ' + dbConnect.LastError)
      else begin
        dbConnect.ClearLog;
        case wcm of
        wcmView:
        begin
          dbConnect.OnLog := g_consoleForm.OnDBLog;
        end;
        wcmNotView:
        begin

        end;
        wcmConsole:
        begin
          dbConnect.OnLog := DBConnectInConsole.OnLogInConsole;
        end;
        end;

        if sSql <> '' then begin
          dbConnect.ExecOneSql(sSql);
          for i := 0 to dbConnect.Log.Count - 1 do begin
            g_log.AddLog(dbConnect.Log[i]);
          end;
          WriteLnInConsole(Format('执行完成！以上日志被保存在%s', [
            '%'+ UpperCase(gC_AppName)+'_HOME%\'+ sShortLogDir
            + ExtractFileName(g_log.LogFileName)
            ]), wcm);
        end;
        dbConnect.CloseDB;
        
        if wcm = wcmView then
          g_consoleForm.DelayClosed(5);
      end;
    finally
      if Assigned(g_Log) then
        FreeAndNil(g_Log);
      dBConnect._Release;
      dBConnect := nil;
    end;
  end;
end;

initialization

finalization
  if Assigned(DBConnectInConsole) then
    FreeAndNil(DBConnectInConsole);

end.
