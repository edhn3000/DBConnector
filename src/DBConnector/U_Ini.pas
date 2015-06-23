{**
 * <p>Title: TIniOperater </p>
 * <p>Copyright: Copyright (c) 2006/9/7</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: ini读写类, 构造函数传入文件名和读取的段  </p>
 *}
unit U_Ini;

interface

uses
  IniFiles, U_DBEngineInterface, U_DBConfig;

type
  TBarOption = record
    bBreak: Boolean;
    nWidth: Integer;
  end;
  { TIniOperater }
  TIniOperater = class(TIniFile)
  private

  protected

  public 
    AutoShowError: Boolean;
    MaxRecord: Integer;
    MaxFitColCount: Integer;
    RecentDBCount: Integer;
    CombineSql: Boolean; 
    LoadLastScript: Boolean;
    SingleInstance: Boolean;
    ConsoleModeWithView:Boolean;
    SortOnColumnClick: Boolean;
    LastScriptFile: string;

    // 数据库配置对象
    DBconfig: TDBConfig;

    LastDir: string;
    Skin: string;
    ShowMenu: Boolean;
    ShowBar: Boolean;
    ShowTree: Boolean;

    BarOptions: array of TBarOption;
  public
    constructor Create(const FileName: string);
    destructor Destroy; override; // 已在finalization调用

    function LoadSettings: Boolean;
    function SaveSettings: Boolean;
  end;               

  function GlobalParams: TIniOperater;

const
  C_sIniFile = 'Options.ini';
  C_sINI_SectionDef = 'Default';
  C_sINI_SectionNotDef = 'Params'; // 预设 为不同设置而更改段名

  C_sINI_SectionAccess = 'Access';
  C_sINI_SectionOracle = 'Oracle';
  C_sINI_SectionSybase = 'Sybase';
  C_sINI_SectionMySQL  = 'MySQL';
  C_sINI_SectionSQLite = 'SQLite';
  C_sINI_SectionDBF = 'DBF';

  C_sINI_SectionDB = 'Database';
  C_sINI_DBTypeUnKnown = '0';
  C_sINI_DBTypeAcs = '1';
  C_sINI_DBTypeOra = '2';
  C_sINI_DBTypeAcs2007 = '3';    
  C_sFILENAME_INI = 'DBOptions.ini';


implementation

uses
  SysUtils, Forms, U_CommonFunc, U_Base64Codec;

const
  // 真假设置
  C_sTrue     = 'True';
  C_sFalse    = 'False';
  C_sTrueNum  = '1';
  C_sFalseNum = '0';
      
var
  m_DBParams: TIniOperater;

function GlobalParams: TIniOperater;
begin
  if not Assigned(m_DBParams) then
  begin
    m_DBParams := TIniOperater.Create(ExtractFilePath(Application.ExeName) +
        C_sFILENAME_INI);
    m_DBParams.LoadSettings;
//    g_Global.DBConfigList.RecentDBCount := m_DBParams.RecentDBCount;
  end;
  Result := m_DBParams;
end;

{ TIniOperater }

constructor TIniOperater.Create(const FileName: string);
begin
  inherited Create(FileName);
  DBconfig := TDBConfig.Create;
end;

destructor TIniOperater.Destroy;
begin
  DBconfig.Free;
  inherited;
end;

function GetBool(s: string): Boolean;
begin
  if SameText(s, C_sTrue) or SameText(s, C_sTrueNum) then
    Result := True
  else
    Result := False;
end;

function GetStr(b: Boolean): string;
begin
  if b then
    Result := C_sTrue
  else
    Result := C_sFalse;
end;

function TIniOperater.LoadSettings: Boolean;
var
  i: Integer;
begin
  Result := True;
  try
    LastDir := ReadString(C_sINI_SectionDef, 'LaseDir', '');
    AutoShowError := ReadBool(C_sINI_SectionDef, 'AutoShowError', True);
    MaxRecord := ReadInteger(C_sINI_SectionDef, 'MaxRecord', 500);
    MaxFitColCount := ReadInteger(C_sINI_SectionDef, 'MaxFitColCount', 20);
    RecentDBCount := ReadInteger(C_sINI_SectionDef, 'RecentDBCount', 10);
    CombineSql := ReadBool(C_sINI_SectionDef, 'CombineSql', True);
    LoadLastScript := ReadBool(C_sINI_SectionDef, 'LoadLastScript', False);
    SingleInstance := ReadBool(C_sINI_SectionDef, 'SingleInstance', False);
    ConsoleModeWithView := ReadBool(C_sINI_SectionDef, 'ConsoleModeWithView', False);
    SortOnColumnClick := ReadBool(C_sINI_SectionDef, 'SortOnColumnClick', False);
    LastScriptFile := ReadString(C_sINI_SectionDef, 'LastScriptFile', '');

    DBconfig.DBType := StrToDBType(ReadString(C_sINI_SectionDB, 'DBType', '0'));
    DBconfig.DBEngineType := StrToDBEngineType(ReadString(C_sINI_SectionDB, 'DBEngineType', '0'));

    DBconfig.AcsDataSource := ReadString(C_sINI_SectionAccess, 'DB', '');
    DBconfig.AcsSecuredDB  := ReadString(C_sINI_SectionAccess, 'Secured', '');
    DBconfig.AcsUser := Base64Decode(ReadString(C_sINI_SectionAccess, 'User', ''));
    DBconfig.AcsPwd  := Base64Decode(ReadString(C_sINI_SectionAccess, 'Pwd', ''));

    DBconfig.SID      := ReadString(C_sINI_SectionOracle, 'SID', '');
    DBconfig.Host     := ReadString(C_sINI_SectionOracle, 'IP', '');
    DBconfig.Protocol := ReadString(C_sINI_SectionOracle, 'Protocol', '');
    DBconfig.Port     := ReadString(C_sINI_SectionOracle, 'Port', '');
    DBconfig.OraUser  := Base64Decode(ReadString(C_sINI_SectionOracle, 'User', ''));
    DBconfig.OraPwd   := Base64Decode(ReadString(C_sINI_SectionOracle, 'Pwd', ''));
    DBconfig.IsLocalTns := ReadBool(C_sINI_SectionOracle, 'IsLocalTns', False);

    DBconfig.SybServerName := ReadString(C_sINI_SectionSybase, 'ServerName', '');
    DBconfig.SybDataBaseName := ReadString(C_sINI_SectionSybase, 'DataBaseName', '');
    DBconfig.SybPort := ReadString(C_sINI_SectionSybase, 'Port', '');
    DBconfig.SybUser := Base64Decode(ReadString(C_sINI_SectionSybase, 'User', ''));
    DBconfig.SybPwd := Base64Decode(ReadString(C_sINI_SectionSybase, 'Pwd', ''));
//    DBconfig.SybType := ReadInteger(C_sINI_SectionSybase, 'Type', 0);

    DBconfig.mslHost := ReadString(C_sINI_SectionMySQL, 'Host', '');
    DBconfig.mslDataBase := ReadString(C_sINI_SectionMySQL, 'DataBase', '');
    DBconfig.mslCharset := ReadString(C_sINI_SectionMySQL, 'Charset', '');
    DBconfig.mslUser := Base64Decode(ReadString(C_sINI_SectionMySQL, 'User', ''));
    DBconfig.mslPwd := Base64Decode(ReadString(C_sINI_SectionMySQL, 'Pwd', ''));

    DBconfig.sltDB := ReadString(C_sINI_SectionSQLite, 'DB', '');

    DBconfig.dbfDB := ReadString(C_sINI_SectionDBF, 'DB', '');

    Skin := ReadString('UI', 'Skin', '');
    ShowMenu := True;
    ShowBar := True;
    ShowTree := True;
//    ShowMenu := ReadBool('UI', 'ShowMenu', True);
//    ShowBar := ReadBool('UI', 'ShowBar', True);
//    ShowTree := ReadBool('UI', 'ShowTree', True);

    i := 0;
    while True do
    begin
      if Length(BarOptions) < i + 1 then
        SetLength(BarOptions, i+1);
      if i = 0 then
        BarOptions[i].bBreak := ReadBool('Bars', 'Break'+IntToStr(i), False)
      else
        BarOptions[i].bBreak := ReadBool('Bars', 'Break'+IntToStr(i), True);
      BarOptions[i].nWidth := ReadInteger('Bars', 'Width'+IntToStr(i), 0);

      if BarOptions[i].nWidth = 0 then
      begin
        SetLength(BarOptions, i);
        Break;
      end;
      Inc(i);
    end;  
//    for i := Low(BarOptions) to High(BarOptions) do
//    begin
//      if i = 0 then
//        BarOptions[i].bBreak := ReadBool('Bars', 'Break'+IntToStr(i), False)
//      else
//        BarOptions[i].bBreak := ReadBool('Bars', 'Break'+IntToStr(i), True);
//      BarOptions[i].nWidth := ReadInteger('Bars', 'Width'+IntToStr(i), 0);
//    end;
  except
    Result := False;
  end;
end;

function TIniOperater.SaveSettings: Boolean;
var
  i: Integer;
begin 
  Result := True;
  try
    WriteString(C_sINI_SectionDef, 'LastDir', LastDir);
    WriteBool(C_sINI_SectionDef, 'AutoShowError', AutoShowError);
    WriteBool(C_sINI_SectionDef, 'CombineSql', CombineSql);
    WriteBool(C_sINI_SectionDef, 'LoadLastScript', LoadLastScript);  
    WriteBool(C_sINI_SectionDef, 'SingleInstance', SingleInstance);
    WriteBool(C_sINI_SectionDef, 'ConsoleModeWithView', ConsoleModeWithView);
    WriteBool(C_sINI_SectionDef, 'SortOnColumnClick', SortOnColumnClick);
    WriteString(C_sINI_SectionDef, 'LastScriptFile', LastScriptFile);

    WriteInteger(C_sINI_SectionDef, 'MaxRecord', MaxRecord);
    WriteInteger(C_sINI_SectionDef, 'MaxFitColCount', MaxFitColCount);
    WriteInteger(C_sINI_SectionDef, 'RecentDBCount', RecentDBCount);
    
    WriteString(C_sINI_SectionDB, 'DBType', DBTypeToStr(DBconfig.DBType, False));
    WriteString(C_sINI_SectionDB, 'DBEngineType',
      DBEngineTypeToStr(DBconfig.DBEngineType, False));

    WriteString(C_sINI_SectionAccess, 'DB', DBconfig.AcsDataSource);
    WriteString(C_sINI_SectionAccess, 'Secured', DBconfig.AcsSecuredDB);
    WriteString(C_sINI_SectionAccess, 'User', Base64Encode(DBconfig.AcsUser));
    WriteString(C_sINI_SectionAccess, 'Pwd', Base64Encode(DBconfig.AcsPwd));

    WriteString(C_sINI_SectionOracle, 'SID', DBconfig.SID);
    WriteString(C_sINI_SectionOracle, 'IP', DBconfig.Host);
    WriteString(C_sINI_SectionOracle, 'Protocol', DBconfig.Protocol);
    WriteString(C_sINI_SectionOracle, 'Port', DBconfig.Port);       
    WriteBool(C_sINI_SectionOracle, 'IsLocalTns', DBconfig.IsLocalTns);
    WriteString(C_sINI_SectionOracle, 'User', Base64Encode(DBconfig.OraUser));
    WriteString(C_sINI_SectionOracle, 'Pwd', Base64Encode(DBconfig.OraPwd));

    WriteString(C_sINI_SectionSybase, 'ServerName', DBconfig.SybServerName);
    WriteString(C_sINI_SectionSybase, 'DatabaseName', DBconfig.SybDatabaseName);
    WriteString(C_sINI_SectionSybase, 'SybPort', DBconfig.SybPort);
    WriteString(C_sINI_SectionSybase, 'User', Base64Encode(DBconfig.SybUser));
    WriteString(C_sINI_SectionSybase, 'Pwd', Base64Encode(DBconfig.SybPwd));
//    WriteInteger(C_sINI_SectionSybase, 'Type', SybType);

    WriteString(C_sINI_SectionMySQL, 'Host', DBconfig.mslHost);
    WriteString(C_sINI_SectionMySQL, 'DataBase', DBconfig.mslDataBase);
    WriteString(C_sINI_SectionMySQL, 'Charset', DBconfig.mslCharset);
    WriteString(C_sINI_SectionMySQL, 'User', Base64Encode(DBconfig.mslUser));
    WriteString(C_sINI_SectionMySQL, 'Pwd', Base64Encode(DBconfig.mslPwd));

    WriteString(C_sINI_SectionSQLite, 'DB', DBconfig.sltDB);

    WriteString(C_sINI_SectionDBF, 'DB', DBconfig.dbfDB);

    WriteString('UI', 'Skin', Skin);
//    WriteBool('UI', 'ShowMenu', ShowMenu);
//    WriteBool('UI', 'ShowBar', ShowBar);
//    WriteBool('UI', 'ShowTree', ShowTree);

    for i := Low(BarOptions) to High(BarOptions) do
    begin
      WriteBool('Bars', 'Break'+IntToStr(i), BarOptions[i].bBreak);
      WriteInteger('Bars', 'Width'+IntToStr(i), BarOptions[i].nWidth);
    end;

  except
    Result := False;
  end;
end;

initialization

finalization
  if Assigned(m_DBParams) then
    FreeAndNil(m_DBParams);


end.
