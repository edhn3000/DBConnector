{**
 * <p>Title: P_UseAdo  </p>
 * <p>Copyright: Copyright (c) 2006/7/24</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 常量定义</p>
 *}

unit U_Const;

interface

uses
  SynHighlighterSQL, 
  Messages;

const
  WMUSER_DBCONNECTOR_ONCONNECT_SUCC = WM_USER + 1001;
  WMUSER_DBCONNECTOR_ONCONNECT_FAIL = WM_USER + 1002;
  WMUSER_DBCONNECTOR_CONNECTWPARAM_TRUE   = 1;
  WMUSER_DBCONNECTOR_CONNECTWPARAM_FALSE  = 2;
  
  WMUSER_DBCONNECTOR_OPER_NEWTAB = WM_USER + 1011;
  WMUSER_DBCONNECTOR_PLUGIN_RESULTDIR = WM_USER+1021;        // not use this but define it here
  
  WMUSER_DBCONNECTOR_EDITOR_SAVE = 1021;
  
  WMUSER_DBCONNECTOR_MAIN_REFRESHONCONNECT = 1031;
  WMUSER_DBCONNECTOR_MAIN_REFRESHONDISCONNECT = 1032;
  
  WMUSER_CONSOLE_RUN = WM_USER + 1101;

type

  THLName = record
    DisplayName: string;
    SQLDialect: TSQLDialect
  end;

const

  C_sPATH_SKIN = 'skin\';

  gC_Delimiter = ';';
  gC_Error = -1;
  gC_nFail = -1;

  gC_nOnePageRecords = 100;

  {TSQLDialect = (sqlStandard, sqlInterbase6, sqlMSSQL7, sqlMySQL, sqlOracle,
    sqlSybase, sqlIngres, sqlMSSQL2K, sqlPostgres, sqlNexus);}
  gC_ARY_HIGHLIGHT_SET: array[TSQLDialect] of THLName = (
    (DisplayName: 'Standard sql';SQLDialect: sqlStandard),
    (DisplayName: 'InterBase6';SQLDialect: sqlInterbase6),
    (DisplayName: 'MSSQL7';SQLDialect: sqlMSSQL7),
    (DisplayName: 'MySQL';SQLDialect: sqlMySQL),
    (DisplayName: 'Oracle';SQLDialect: sqlOracle),
    (DisplayName: 'SyBase';SQLDialect: sqlSybase),
    (DisplayName: 'Ingres';SQLDialect: sqlIngres),
    (DisplayName: 'MSSQL2K';SQLDialect: sqlMSSQL2K),
    (DisplayName: 'Postgres';SQLDialect: sqlPostgres),
    (DisplayName: 'Nexus';SQLDialect: sqlNexus)
    );

  gC_sLastScript = 'LastScript.sql';

//  gC_Version = '1.0';
  gC_AppName = 'DBConnector';
  C_sFILENAME_ERROR = gC_AppName+'_ExecErrors.txt';

implementation

end.

