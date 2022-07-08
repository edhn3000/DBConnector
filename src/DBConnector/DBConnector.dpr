program DBConnector;

uses
  Forms,
  Windows,
  Variants,
  SysUtils,
  ComCtrls,
  UF_MAIN in 'UF_MAIN.pas' {F_DBConnectorMain},
  U_Const in 'U_Const.pas',
  U_Pub in 'U_Pub.pas',
  U_CommonFunc in '..\MyPub\common\U_CommonFunc.pas',
  UF_Find in '..\MyPub\UF_Find.pas' {F_Find},
  U_UIUtil in '..\MyPub\UI\U_UIUtil.pas',
  U_FileUtil in '..\MyPub\IO\U_FileUtil.pas',
  U_ThreadUtil in '..\MyPub\thread\U_ThreadUtil.pas',
  UF_About in 'UF_About.pas' {F_About},
  U_Ini in 'U_Ini.pas',
  UF_BaseDialog in 'UF_BaseDialog.pas' {F_BaseDialog},
  UF_DBOption in 'UF_DBOption.pas' {F_DBOption},
  UFM_DBOperate in 'UFM_DBOperate.pas' {FM_DBOperate: TFrame},
  U_fStrUtil in '..\MyPub\common\U_fStrUtil.pas',
  UF_ExportResult in 'UF_ExportResult.pas' {F_ExportResult},
  UF_ExecScript in 'UF_ExecScript.pas' {F_ExecScript},
  UF_Tip in 'UF_Tip.pas' {F_Tip},
  U_PluginMsg in 'U_PluginMsg.pas',
  U_DosConsole in '..\MyPub\Console\U_DosConsole.pas',
  U_Plugin in 'U_Plugin.pas',
  UF_ConsoleView in 'UF_ConsoleView.pas' {F_ConsoleView},
  U_PlanarList in '..\MyPub\common\U_PlanarList.pas',
  U_DBEngine in '..\MyPub\db\DBEngine\U_DBEngine.pas',
  U_ADODBEngine in '..\MyPub\db\DBEngine\U_ADODBEngine.pas',
  U_BDEDBEngine in '..\MyPub\db\DBEngine\U_BDEDBEngine.pas',
  U_DBExpressEngine in '..\MyPub\db\DBEngine\U_DBExpressEngine.pas',
  U_DBConnect in '..\MyPub\db\DBConnect\U_DBConnect.pas',
  U_AccessDBConnect in '..\MyPub\db\DBConnect\U_AccessDBConnect.pas',
  U_OracleDBConnect in '..\MyPub\db\DBConnect\U_OracleDBConnect.pas',
  U_MySQLDBConnect in '..\MyPub\db\DBConnect\U_MySQLDBConnect.pas',
  U_SybaseDBConnect in '..\MyPub\db\DBConnect\U_SybaseDBConnect.pas',
  U_DBConnnectManager in '..\MyPub\db\DBConnect\U_DBConnnectManager.pas',
  U_DBUtil in '..\MyPub\DB\U_DBUtil.pas',
  UF_Rename in 'UF_Rename.pas' {F_Rename},
  U_PubFunc in 'U_PubFunc.pas',
  U_TableInfo in '..\MyPub\db\DBConnect\U_TableInfo.pas',
  U_FieldInfo in '..\MyPub\db\DBConnect\U_FieldInfo.pas',
  U_DBConfig in '..\MyPub\DB\U_DBConfig.pas',
  U_SQLBuilder in '..\MyPub\DB\SQL\U_SQLBuilder.pas',
  U_DBConnectorInConsole in 'U_DBConnectorInConsole.pas',
  UF_ExportTables in 'UF_ExportTables.pas' {F_ExportTables},
  U_OracleTableInfo in '..\MyPub\db\DBConnect\U_OracleTableInfo.pas',
  U_AccessFieldInfo in '..\MyPub\db\DBConnect\U_AccessFieldInfo.pas',
  U_OracleFieldInfo in '..\MyPub\db\DBConnect\U_OracleFieldInfo.pas',
  U_MySqlFieldInfo in '..\MyPub\db\DBConnect\U_MySqlFieldInfo.pas',
  U_SybaseFieldInfo in '..\MyPub\db\DBConnect\U_SybaseFieldInfo.pas',
  U_AccessTableInfo in '..\MyPub\db\DBConnect\U_AccessTableInfo.pas',
  U_MySqlTableInfo in '..\MyPub\db\DBConnect\U_MySqlTableInfo.pas',
  U_SybaseTableInfo in '..\MyPub\db\DBConnect\U_SybaseTableInfo.pas',
  U_TableInfoFactory in '..\MyPub\db\DBConnect\U_TableInfoFactory.pas',
  U_FieldInfoFactory in '..\MyPub\db\DBConnect\U_FieldInfoFactory.pas',
  U_DBCommand in '..\MyPub\DB\DBConnect\U_DBCommand.pas',
  U_DataMoudle in 'U_DataMoudle.pas' {DataModule1: TDataModule},
  U_DBTreeFunc in 'U_DBTreeFunc.pas',
  U_DBAHelp in 'U_DBAHelp.pas',
  U_ExcelObject in '..\MyPub\Office\U_ExcelObject.pas',
  U_DBExport in '..\MyPub\DB\U_DBExport.pas',
  U_Convertor in '..\MyPub\U_Convertor.pas',
  U_TextFileWriter in '..\MyPub\IO\U_TextFileWriter.pas',
  U_TriggerInfo in '..\MyPub\DB\DBConnect\U_TriggerInfo.pas',
  U_IndexInfo in '..\MyPub\DB\DBConnect\U_IndexInfo.pas',
  U_DBConnectInterface in '..\MyPub\DB\U_DBConnectInterface.pas',
  U_DBEngineInterface in '..\MyPub\DB\U_DBEngineInterface.pas',
  U_DBEngineFactory in '..\MyPub\DB\DBEngine\U_DBEngineFactory.pas',
  U_SqlUtils in '..\MyPub\DB\SQL\U_SqlUtils.pas',
  U_SQLiteDBEngine in '..\MyPub\DB\DBEngine\U_SQLiteDBEngine.pas',
  U_SQLiteDBConnect in '..\MyPub\DB\DBConnect\U_SQLiteDBConnect.pas',
  U_RegexUtil in '..\MyPub\regex\U_RegexUtil.pas',
  U_Log in '..\MyPub\Log\U_Log.pas',
  U_Log4d in '..\MyPub\Log\U_Log4d.pas',
  U_ThreadPool in '..\MyPub\thread\U_ThreadPool.pas',
  U_ExportUtil in '..\MyPub\IO\U_ExportUtil.pas';

{$R *.res}
//{$R ..\MyPub\XPStyle\XPStyle.res}     
//{$R *.inc}

var
  hnd: HWND;
begin
  if GlobalParams.SingleInstance then
  begin
    hnd := FindWindow('TF_MAIN', 'DBConnector');
    if hnd <> 0 then
    begin
      SetForegroundWindow(hnd);
      Exit;
    end;
  end;
  Application.Initialize;
  Application.Title := 'DBConnector';

  if ParamCount = 0 then
  begin
  Application.CreateForm(TF_DBConnectorMain, F_DBConnectorMain);
  Application.CreateForm(TDataModule1, DataModule1);
  Application.Run;
  end
  else
  begin
  if GlobalParams.ConsoleModeWithView then       
  begin
  Application.CreateForm(TF_ConsoleView, g_consoleForm);
  Application.Run;  
  end
  else
  RunInConsole(wcmNotView);
  end;
end.
