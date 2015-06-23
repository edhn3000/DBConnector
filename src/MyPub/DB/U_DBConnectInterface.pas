{**
 * <p>Title: U_DBConnectInterface  </p>
 * <p>Copyright: Copyright (c) 2012/6/10</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 数据库引擎类的接口
 *}
unit U_DBConnectInterface;

interface

uses
  DB, Classes, Windows, Messages, SysUtils, Variants, Graphics, Controls, Forms,
  Dialogs, U_TextFileWriter,
  U_TableInfo, U_FieldInfo,
  U_DBEngineInterface;


type
  TFieldShowOption = (fsoName, fsoType, fsoNullable);
  TFieldShowOptions = Set of TFieldShowOption;  
  
  TDBExportOption = class
  public
    DropTables: Boolean;
    CreateTables: Boolean;
    DeleteTables: Boolean;
    ExportData: Boolean;
    DisableTriggers: Boolean;
    RemoveBreakLine: Boolean;
    ClobInFile: Boolean;
  end;   
  TDBConnectSQLType = (dbcstCommand, dbcstQuery, dbcstUpdate); 

  TDBLogEvent = procedure (sLog: string) of object;  
  TExecLineChangeEvent = procedure (nCurrLine: Integer) of object;
  
{ TSQLCmdType sql生成器使用 }
  TSQLCmdType = (sctInsert, sctUpdate, sctDelete, sctSelect, sctCreate);
  
  IDBConnect = interface
  ['{68B0D79A-52B1-46A2-9713-EE4D5F7A6E23}']  
    function GetDBType: TDBType;  
    function GetDataSet: TDataSet;
    function GetConnection: TCustomConnection; 
    function GetConnected: Boolean;
    procedure SetConnected(value: Boolean); 
    function GetLog: TStrings;
    function GetLastError: string;      
    function GetLastQry: TDataSet;
    function GetLastSQL: string;
    function GetLastSQLType: TDBConnectSQLType;
    function GetLastOperSucc: Boolean;
    function GetMaxRecords: Integer;
    procedure SetMaxRecords(value: Integer);  
    function GetPageIndex: Integer;
    procedure SetPageIndex(value: Integer);
    function GetDropNoError: Boolean;
    procedure SetDropNoError(value: Boolean); 
    function GetVariables: TStringList;
    procedure SetVariables(value: TStringList);
    function GetDBEngine: IDBEngine;       
    function GetSystemObject: Boolean;
    procedure SetSystemObject(value: Boolean);
    function GetOnLog: TDBLogEvent;
    procedure SetOnLog(value: TDBLogEvent);
    function GetOnConnected: TNotifyEvent;
    function GetOnDisConnected: TNotifyEvent;
    procedure SetOnConnected(value: TNotifyEvent);
    procedure SetOnDisConnected(value: TNotifyEvent);
    function GetOnExecLineChange: TExecLineChangeEvent;
    procedure SetOnExecLineChange(value: TExecLineChangeEvent); 
    function GetOnExecuted: TNotifyEvent;
    procedure SetOnExecuted(value: TNotifyEvent);
    function GetLastElapsedMilis: Double;

    // propery          
    property DBType: TDBType read GetDBType;
    property DataSet: TDataSet read GetDataSet;
    property Connection: TCustomConnection read GetConnection; 
    property Connected: Boolean read GetConnected write SetConnected;
    property Log: TStrings read GetLog;
    property LastError: string read GetLastError; 
    property LastQry: TDataSet read GetLastQry;
    property LastSQL: string read GetLastSQL;
    property LastSQLType: TDBConnectSQLType read GetLastSQLType;
    property LastOperSucc: Boolean read GetLastOperSucc;
    property DBEngine: IDBEngine read GetDBEngine;
    property SystemObject: Boolean read GetSystemObject write SetSystemObject;
    property MaxRecords: Integer read GetMaxRecords write SetMaxRecords;
    property PageIndex: Integer read GetPageIndex write SetPageIndex;
    property DropNoError: Boolean read GetDropNoError write SetDropNoError;
    property Variables: TStringList read GetVariables write SetVariables;
    // event
    property OnLog: TDBLogEvent read GetOnLog write SetOnLog;
    property OnConnected: TNotifyEvent read GetOnConnected write SetOnConnected;
    property OnDisConnected: TNotifyEvent read GetOnDisConnected write SetOnDisConnected;
    property OnExecLineChange: TExecLineChangeEvent read GetOnExecLineChange
      write SetOnExecLineChange;
    property OnExecuted: TNotifyEvent read GetOnExecuted write SetOnExecuted;

    // function
    // 运行期可改参数重置，同时会用于setdefault命令
    procedure ResetParams;

    procedure ClearLog;      
    procedure AddLog(sLog: string); 
    procedure AddErrorLog(sLog: string);
    procedure AddInfoLog(sLog: string);
    procedure AddSqlExecLog(nEffectRows: Integer; sCommandContent: string;
       relData: string);

    // 可使多个Connect对象共用一个Engine而不必重复创建新的Engine并连接数据库
    procedure ShareEngine(DestDBEngine: IDBEngine; bCopy: Boolean = False);

    { 数据库操作 }
    function OpenDB(sDataSource: string;sUser: string;sPwd: string;
      dbt: TDBType; dbet: TDBEngineType): Boolean;overload;
    function OpenDB(AIniFile, ADBSection: string): Boolean;overload;
    function CloseDB: Boolean;
    function CloseQuery: Boolean;
    function GetNewQuery: TDataSet; 
//    function GetCanShowDataSet(ds: TDataSet): TDataSet;
    procedure StopExec();


    //*************** virtual ***************
    procedure GetDBs(List: TStrings);
    procedure GetTableNames(List: TStrings; AQry: TDataSet = nil);
    // 获得对象列表，可得到更多信息        
    function GetTableInfo(sTableName: string): TTableInfo;
    function GetTableInfos(AQry: TDataSet = nil):TTableInfoList;
    function GetFieldInfos(const tableName: string;
      AQry: TDataSet = nil): TFieldItemList;
    procedure GetViewNames(List: TStrings; AQry: TDataSet = nil);
    procedure GetFieldNames(const tableName: string; List: TStrings;
      option: TFieldShowOptions; AQry: TDataSet = nil);
    procedure GetProcedureNames(List: TStrings; AQry: TDataSet = nil);
    procedure GetTriggerNames(List: TStrings; AQry: TDataSet = nil); 
    function GetPrimaryField(TableName: string; AQry: TDataSet = nil): TFieldItem; 
    function IsFieldUnique(tableName, fieldName: string; AQry: TDataSet = nil): Boolean;
    procedure GetIndexNames(List: TStrings);
    procedure CheckFieldsUnique(tableName: string; fields: TFieldItemList); 
    procedure CheckFieldUnique(tableName: string; field: TFieldItem);

    function GenTableCreateSQL(sTableName: string; slst: TStrings):Boolean;
    function ExportQuery(AQry: TDataSet; exportFile: TTextFileWriter;
      ClobInFile: Boolean; RemoveBreakLine: Boolean): Boolean;overload;
    function ExportQuery(AQry: TDataSet; sExportFileName: string):Boolean;overload; 
    function ExportTableData(tableName, sExportFileName: string;
      option: TDBExportOption): Boolean;

    function GetValidFieldValue(FieldItem: TFieldItem;
      const sValueDef: string = ''): string;overload;
    function GetValidFieldValue(Field: TField;
      const sValueDef: string = ''): string;overload;
    function GetFieldsKeyValuesStr(Fields: TFieldItemList): string;
    function GetFieldsValuesStr(Fields: TFieldItemList): string;
    function TableExists(sTableName: string): Boolean;
    //*************** virtual ***************

    { sql处理 }
    function RemoveHeadFlag(Asql: string): string;
    procedure SetVariableValue(var s: string);

    { 执行sql的方法 }
    function ExecQuery(sSql: string):Boolean; overload;
    function ExecQuery(AQ: TDataSet; sSql: string):Boolean; overload;
    function ExecQueryWithParams(sSqlWithParams: string;
      Params: array of Variant):Boolean; overload;
    function ExecQueryWithParams(AQ: TDataSet; sSqlWithParams: string;
        aryParams: array of Variant):Boolean;overload;
    function ExecUpdate(sSql: string): Integer; overload;
    function ExecUpdate(AQ: TDataSet; sSql: string): Integer; overload;
    function ExecUpdateWithParams(sSqlWithParams: string;
      Params: array of Variant): Integer; overload;
    function ExecUpdateWithParams(AQ: TDataSet; sSqlWithParams: string;
        aryParams: array of Variant): Integer;overload;

    function ExecOneSql(Asql: string; aQry: TDataSet = nil): Integer;
    // 执行多条sql
    function ExecSqlList(slstSqlList:TStrings; aQry: TDataSet = nil):Integer;
    // 不要用来执行脚本，会去除所有换行重新分行
    function ExecSqlText(sSqlText: string; cDelimiter: Char = ';';
      aQry: TDataSet = nil): Integer;
    // 专门用于执行脚本文件
    function ExecScriptFile(AScriptFile: string; aQry: TDataSet = nil):Integer;

    function GetProcSource(sName: string; list: TStrings): string;
    function GetTrigSource(sName: string; list: TStrings): string;
    function GetViewSource(sName: string; list: TStrings): string;
     
    function UpdateBlobs(JsonStr: string): Integer;
    function UpdateClobs(JsonStr: string): Integer;


    // access
    procedure GetMoudleNames(List: TStrings; AQry: TDataSet = nil);
    procedure GetReportNames(List: TStrings; AQry: TDataSet = nil);
    procedure GetMacroNames(List: TStrings; AQry: TDataSet = nil);
    procedure GetFormNames(List: TStrings; AQry: TDataSet = nil);
    procedure GetPageNames(List: TStrings; AQry: TDataSet = nil); 

    // oracle      
    procedure GetUsers(List: TStrings);
    procedure GetSynonymNames(List: TStrings);
    procedure GetDBLinkNames(List: TStrings);

  end;
    
const         
  C_sHeadFlag_Script      = '@';
  C_sHeadFlag_Force       = ':';
  C_sHeadFlag_ForceQuery  = C_sHeadFlag_Force + 'query';
  C_sHeadFlag_ForceUpdate = C_sHeadFlag_Force + 'update';
  C_sHeadFlag_SybaseSP    = 'sp_';
  C_sHeadFlag_Blob        = C_sHeadFlag_Force + 'blob';
  C_sHeadFlag_Clob        = C_sHeadFlag_Force + 'clob';
  
  C_sHeadFlag_DBCommand     = 'DBConnector'; // 命令开头字符串 与程序名一致

  //  增加命令时修改SynHighlighterSQL中的DBConnectorCmd可同步高亮处理
  C_sDBCommand_Help         = 'Help';   
  C_sDBCommand_ShowSettings = 'ShowSettings';
  C_sDBCommand_SetDefault   = 'SetDefault';
  C_sDBCommand_MaxRecords   = 'MaxRecords';
  C_sDBCommand_DropNoError  = 'DropNoError';
  C_sDBCommand_PageIndex    = 'PageIndex'; 
  C_sDBCommand_Export       = 'Export';
  C_sDBCommand_Set          = 'Set';
  C_sDBCommand_Conn         = 'Conn';
  C_sDBCommand_Var          = 'Var';
  C_sDBCommand_SQL          = 'Sql';
  C_sDBCommand_Stop         = 'Stop';
  C_sDBCommand_Exit         = 'Exit';
  C_sDBCommand_Trans        = 'Trans';

  C_sDBOptionSwitch_On  = 'on';
  C_sDBOptionSwitch_Off = 'off';

  C_sDBCmdParam_Var_Start = '$';
  C_sDBCmdParam_Var_End = '$';
  C_sDBCmdParam_Var_SysVar_DateTime14 = 'DateTime14';
  C_sDBCmdParam_Regx_GetVar = '\$[^\s^\$]*\$';

//  C_sDBCommand_Var_Sys_Start = '$';
//  C_sDBCommand_Var_Sys_End = '$';

  C_sSQLEnd = ';';

  C_nDefaultSetting_MaxRecords = 500;

  // 脚本文件前缀
  C_FilePrefix_CreateTable    = 'CT_';
  C_FilePrefix_InsertTable    = 'IT_';
  C_FilePrefix_CreateProcdure = 'CP_';
  C_FilePrefix_CreateView     = 'CV_';
  C_FilePrefix_CreateTrigger  = 'CG_';

  // export
  C_sBlob_Dir_Name = '_BLOB';
  C_sClob_Dir_Name = '_CLOB';

  // JSON项
  C_JSON_KEY_SQL   = 'sql';
  C_JSON_KEY_FIELD = 'field';
  C_JSON_KEY_NAME  = 'name';
  C_JSON_KEY_PATH  = 'path';

const
  TDBC_ERROR_EXECFAIL = C_nTDBE_ERROR_EXECFAIL;
  TDBC_ERROR_NOTCONNECT = -2;
  TDBC_ERROR_SCRIPT_FILENOTEXISTS = -3;

implementation

end.
