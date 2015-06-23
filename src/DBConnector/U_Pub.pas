{**
 * <p>Title: U_Pub.pas  </p>
 * <p>Copyright: Copyright (c) 2006/7/24</p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: ���õ�Ԫ </p>
 *}
 unit U_Pub;

interface

uses
  Contnrs, Windows, Messages, Classes, U_Ini, U_Const,
  UFM_DBOperate, U_LogManager, U_DBConfig,
  U_DBEngineInterface, U_DBConnectInterface, U_DBConnnectManager;

type
  TGlobal = class
  private     
    RegisterFrames: TObjectList;
    function GetConnected: Boolean;
  public
    DBConnect: IDBConnect;
    DBConfigList: TDBConfigList;

    // ��ȡ���ݿ�����״̬���Ƿ�������
    property Connected: Boolean read GetConnected;

  public
    constructor Create;
    destructor Destroy; override;

    // �������ݿ⣬�˷�������ʵ�ʵ����Ӵ�����������Ϣʹ������½���͸�frame������
    function OpenDB(sDataSource, sUser, sPwd: string;
      dbt: TDBType; dbet: TDBEngineType): Boolean;
    function CloseDB(): Boolean;

    // ע��frame��ÿ��dbframe����ʱ��globa��ע�ᣬglobal��������״̬�����ݵĸ���
    procedure RegisterFrame(frame: TFM_DBOperate);
    // ע��frame
    procedure UnRegisterFrame(frame: TFM_DBOperate);
    // global������ע������frame
    procedure ReRegisterAllFrame(exceptObj: IDBConnect = nil);

    // ����frame��sql�༭���е�����������ʱ�����ô˷���ͬ����global
    procedure ShareDB(DB: IDBConnect);

    procedure RefreshMainUIOnConected();   
    procedure RefreshMainUIOnDisConected();
  end;


  function EncodeStr(s: string): string;
  function DecodeStr(s: string): string;

  procedure SaveConfig;
  function GetAppRootPath: string;

var
  g_frmDBOperate: TFM_DBOperate;
  g_Log: TLogManager;
  g_Global: TGlobal = nil;

implementation

uses
  SysUtils, Forms, U_Base64Codec, U_CommonFunc, U_DBTreeFunc;

//  m_DBList: TDBConfigList;


function GetAppRootPath: string;
begin
  Result := ExtractFilePath(Application.ExeName);
end;

function EncodeStr(s: string): string; 
begin
  Result := Base64Encode(s);
end;

function DecodeStr(s: string): string;
begin
  Result := Base64Decode(s);
end;   

procedure SaveConfig;   
var
  dbcfg, dbcfgFind: TDBConfig;
begin   
  dbcfg := TDBConfig.Create;

  with GLobalParams do
  begin
    dbcfg.DBType := DBconfig.DBType;
    dbcfg.DBEngineType := DBconfig.DBEngineType;

    dbcfg.Assign(DBconfig);
    dbcfg.UserName := DBconfig.GetCurrUserName;
    dbcfg.Password := DBconfig.getEncodePwd(dbcfg.DBType); // ����ǰ��ı������
    dbcfg.DataSource := dbcfg.GetSeparatedDataSource;
  end;

  dbcfgFind := g_Global.DBConfigList.FindDBConfig(dbcfg.DataSource,
    dbcfg.UserName);
  if dbcfgFind = nil then
    g_Global.DBConfigList.Insert(0, dbcfg)
  else
  begin
    dbcfgFind.Assign(dbcfg);
    dbcfg.Free;
    
    g_Global.DBConfigList.Extract(dbcfgFind);
    g_Global.DBConfigList.Insert(0, dbcfgFind);
  end;
  g_Global.DBConfigList.SaveToFile(GetAppRootPath + 'DBNames.xml');
end;

{ TGlobal }

constructor TGlobal.Create;
begin                                  
  RegisterFrames := TObjectList.Create;
  DBConfigList := TDBConfigList.Create;
end;

destructor TGlobal.Destroy;
begin
  RegisterFrames.Free;
  if Assigned(DBConnect) then begin
    DBConnect._Release;
    DBConnect := nil;
  end;
  DBConfigList.Free;
  inherited;
end;

function TGlobal.GetConnected: Boolean;
begin
  if not Assigned(DBConnect) then
    Result := False
  else
    Result := DBConnect.Connected;
end;

//procedure TGlobal.DoOnConnect;
//var
//  i: Integer;
//begin
//  if not Assigned(g_Global.DBConnect) then
//    Exit;
//  for i := 0 to RegisterFrames.Count - 1 do
//  begin
//    TDBConnect(RegisterFrames.Items[i]).ShareEngine(
//      g_Global.DBConnect.DBEngine);
//  end;
//end;  

procedure TGlobal.RegisterFrame(frame: TFM_DBOperate);
begin
  RegisterFrames.Add(frame);
  if (DBConnect <> nil) and (DBConnect.Connected) then
  begin
    frame.ReCreateDBConnect(DBConnect.DBEngine.GetDBType);
    frame.ShareDBConnect(DBConnect);
  end
  else
  begin
    frame.ReCreateDBConnect(dbtUnKnown);
  end;
end;  

procedure TGlobal.UnRegisterFrame(frame: TFM_DBOperate);
begin   
  RegisterFrames.Extract(frame);
end;

function TGlobal.OpenDB(sDataSource, sUser, sPwd: string; dbt: TDBType;
  dbet: TDBEngineType): Boolean;
begin
  if Assigned(DBConnect) then begin
    DBConnect._Release;
    DBConnect := nil;
  end;
  try
    DBConnect := TDBConnectManager.OpenDB(sDataSource, sUser, sPwd, dbt, dbet);
    Result := DBConnect.Connected;
    g_DBTreeFunc.SetDBConnect(Self.DBConnect);
    if Result then begin
      DBConnect.MaxRecords := GlobalParams.MaxRecord;
      SendMessage(Application.MainForm.Handle, WMUSER_DBCONNECTOR_ONCONNECT_SUCC,
        0, 0);
    end else begin
      SendMessage(Application.MainForm.Handle, WMUSER_DBCONNECTOR_ONCONNECT_FAIL,
        0, 0);
    end;
  except
    SendMessage(Application.MainForm.Handle, WMUSER_DBCONNECTOR_ONCONNECT_FAIL,
      0, 0);
  end;
end;

function TGlobal.CloseDB: Boolean;
begin
  Result := DBConnect.CloseDB;
  DBConnect._Release;
  DBConnect := nil;
end;

procedure TGlobal.ShareDB(DB: IDBConnect);
begin
  if DBConnect = nil then
  begin
    DBConnect := TDBConnectManager.CreateDBConnect(DB.DBEngine.GetDBType);

    DBConnect.ShareEngine(DB.DBEngine, True);
    // ����DBConnect�����иĶ�������Ҫˢ�µ�����frame
    ReRegisterAllFrame(DB);
  end
  else if (DBConnect.DBEngine.GetDBType <> DB.DBEngine.GetDBType) then
  begin
    DBConnect._Release;
    DBConnect := nil;
    DBConnect := TDBConnectManager.CreateDBConnect(DB.DBEngine.GetDBType);

    DBConnect.ShareEngine(DB.DBEngine, True);
    ReRegisterAllFrame(DB);
  end
  // һ�����ڶ���
  else if (DBConnect.Connected <> DB.Connected) then
  begin
    DBConnect.ShareEngine(DB.DBEngine, True);
    ReRegisterAllFrame(DB);
  end;
  // ������������������������ж������������з�ʽ����conn
end;

procedure TGlobal.RefreshMainUIOnConected;
begin
  SendMessage(Application.MainForm.Handle, WMUSER_DBCONNECTOR_MAIN_REFRESHONCONNECT,
    0, 0);
end;   

procedure TGlobal.RefreshMainUIOnDisConected;
begin
  SendMessage(Application.MainForm.Handle, WMUSER_DBCONNECTOR_MAIN_REFRESHONDISCONNECT,
    0, 0);
end;

procedure TGlobal.ReRegisterAllFrame(exceptObj: IDBConnect);
var
  i: Integer;
begin
  for i := 0 to RegisterFrames.Count - 1 do
  begin
    if exceptObj = TFM_DBOperate(RegisterFrames.Items[i]).DBConnect then
      Continue;
    
    if (DBConnect <> nil) and (DBConnect.Connected) then
    begin
      TFM_DBOperate(RegisterFrames.Items[i]).ReCreateDBConnect(
        DBConnect.DBEngine.GetDBType);
      TFM_DBOperate(RegisterFrames.Items[i]).ShareDBConnect(DBConnect);
    end
    else
      TFM_DBOperate(RegisterFrames.Items[i]).ReCreateDBConnect(dbtUnKnown);
  end;  
end;

initialization
  if not Assigned(g_Global) then
    g_Global := TGlobal.Create;

finalization
  if Assigned(g_Global) then
    FreeAndNil(g_Global);


end.
