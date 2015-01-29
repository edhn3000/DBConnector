{**
 * <p>Title: U_DBCommand  </p>
 * <p>Copyright: Copyright (c) 2011/1/2</p>
 * @author fengyq
 * <p>Description: 数据库命令类 </p>
 * <p>Description: 针对TDBConnect的各种操作命令，从U_DBConnector单元脱离出来 </p>
 *}
unit U_DBCommand;

interface

uses
  ADODB, DB, Classes, Contnrs, Variants,
  U_DBConnect, U_DBEngineInterface, U_DBConnectInterface;

type
{ TDBCommandProc }
  TDBCommandProc = function (sCommandContent: string; var sError: string;
    var nEffectRows: Integer): Boolean of object;

{ TDBCommand }
  TDBCommand = class
  public
    CmdStr: string;
    Describe: string;
    Proc: TDBCommandProc;
    Error: string;
    EffectRows: Integer;
  public
    constructor Create(ACmd: string; ADescribe: string; AProc: TDBCommandProc);
    function Run(sCommandStr: string):Boolean;
  end;

{ TDBCommandList }
  TDBCommandList = class(TObjectList)
  private
    FFindResult: TDBCommand;
    function GetItem(index: Integer): TDBCommand;
    procedure SetItem(index: Integer; value: TDBCommand);
  public
    property Items[index: Integer]: TDBCommand read GetItem write SetItem;
    property FindResult: TDBCommand read FFindResult;
  public
    // 传入命令与所有命令对比,找到返回0,否则返回与最相似命令的差别的index,取最相似命令读取cmd属性
    function CompareAllCommand(sName: string): Integer;
  end;

  TDBCommandManager = class
  private
    FDBCommandList : TDBCommandList;
    FDBConnect: IDBConnect;
    function Bool2OnOff(b: Boolean): string;
    function OnOff2Bool(s: string): Boolean;             // 数据库命令列表，维护全部命令
    function GetParamByLevel(sFullCommand: string; nIndex: Integer): string;

    function GetDBConnect: IDBConnect;
    procedure SetDBConnect(value: IDBConnect);
  protected
    procedure LoadDBCommand;
    // 辅助类命令 
    function DBCommand_Help(sCommandContent: string; var sError: string;
      var nEffectRows: Integer):Boolean;virtual;         // 帮助
    function DBCommand_ShowSettings(sCommandContent: string; var sError: string;
       var nEffectRows: Integer):Boolean;virtual;        // 显示当前设置
    function DBCommand_SetDefault(sCommandContent: string; var sError: string;
      var nEffectRows: Integer):Boolean;virtual;         // 重置默认
    // 设置参数类命令
    function DBCommand_MaxRecords(sCommandContent: string; var sError: string;
      var nEffectRows: Integer):Boolean;virtual;         // 最大记录
    function DBCommand_DropNoError(sCommandContent: string; var sError: string;
      var nEffectRows: Integer):Boolean;virtual;         // Drop时不显示错误
    function DBCommand_PageIndex(sCommandContent: string; var sError: string;
      var nEffectRows: Integer):Boolean;virtual;         // 页码
    function DBCommand_Variable(sCommandContent: string; var sError: string;
      var nEffectRows: Integer):Boolean;virtual;         // 设置变量
    // 动作类命令
    function DBCommand_Connect(sCommandContent: string; var sError: string;
      var nEffectRows: Integer):Boolean;virtual;          // 连接数据库
    function DBCommand_Export(sCommandContent: string; var sError: string;
      var nEffectRows: Integer): Boolean;                 // 导出
    function DBCommand_SQL(sCommandContent: string; var sError: string;
      var nEffectRows: Integer):Boolean;virtual;          // 执行特殊sql
    function DBCommand_Stop(sCommandContent: string; var sError: string;
      var nEffectRows: Integer):Boolean;virtual;          // 停止执行
    function DBCommand_Exit(sCommandContent: string; var sError: string;
      var nEffectRows: Integer):Boolean;virtual;          // 退出

  public
    constructor Create; virtual;
    destructor Destroy; override;

    function IsDBConnectCommand(const sSql: string):Boolean;
    function ProcessCommand(sCommandStr: string): Integer;

  public
      
    property DBConnect: IDBConnect read GetDBConnect write SetDBConnect;
  end;

  function BuildCommandPrefix(sCommand: string): string;

implementation

uses
  SysUtils, StrUtils, Forms, 
  U_fStrUtil, U_PlanarList, U_DBExport, U_SqlUtils;

function BuildCommandPrefix(sCommand: string): string;
begin
  Result := C_sHeadFlag_DBCommand + ' ' + sCommand;
end;

{ TDBCommand }

constructor TDBCommand.Create(ACmd: string; ADescribe: string; AProc: TDBCommandProc);
begin
  Self.CmdStr     := ACmd;
  Self.Describe   := ADescribe;
  Self.proc       := AProc;
  Self.EffectRows := 0;
end;

function TDBCommand.Run(sCommandStr: string):Boolean;
var
  nPos: Integer;
  sCmdContent: string;
begin
  nPos := Pos(' '+ UpperCase(CmdStr), UpperCase(sCommandStr));
  if nPos > 0 then
    sCmdContent := Trim(Copy(sCommandStr, nPos+Length(CmdStr)+1, MaxInt))
  else
    sCmdContent := Trim(sCommandStr);
  Error := '';
  Result := proc(sCmdContent, Error, EffectRows);
end;

{ TDBCommandList }

function TDBCommandList.CompareAllCommand(sName: string): Integer;
var
  i: Integer;
  nCompareResult, nSimilerIndex, nSimilerGrade: Integer;
  bFind: Boolean;
begin
  bFind := False;
  // 通过 nSimilerGrade 找到最相似的
  nSimilerGrade := 0;
  nSimilerIndex := 0;
  for i := 0 to Count - 1 do
  begin     
    nCompareResult := fStrUtil.CompareTextAtIndex(sName, Items[i].CmdStr);
    if nCompareResult = 0 then
    begin
      bFind := True;
      FFindResult := Items[i];
      Break;
    end;
    if nCompareResult > nSimilerGrade then
    begin
      nSimilerIndex := i;
      nSimilerGrade := nCompareResult;
    end;
  end;
  
  if bFind then
    Result := 0
  else
  begin
    Result := nSimilerGrade;
    FFindResult := Items[nSimilerIndex];
  end;
end;

function TDBCommandList.GetItem(index: Integer): TDBCommand;
begin
  Result := inherited GetItem(index) as TDBCommand;
end;

procedure TDBCommandList.SetItem(index: Integer; value: TDBCommand);
begin
  inherited SetItem(index, value);
end;

{ TDBCommandManager }

function TDBCommandManager.IsDBConnectCommand(const sSql: string): Boolean;
begin
  Result := AnsiStartsText(C_sHeadFlag_DBCommand , sSql);
end;     

function TDBCommandManager.Bool2OnOff(b: Boolean): string;
begin
  if b then
    Result := C_sDBOptionSwitch_On
  else
    Result := C_sDBOptionSwitch_Off;
end;

function TDBCommandManager.OnOff2Bool(s: string): Boolean;
begin
  if LowerCase(s) = C_sDBOptionSwitch_On then
    Result := True
  else
    Result := False;
end;

function TDBCommandManager.GetParamByLevel(sFullCommand: string; nIndex: Integer): string;
var
  slstComParts: TStrings;
begin
  slstComParts := TStringList.Create;
  try
    fStrUtil.SplitEscapeQuote(sFullCommand, ' ', '"', '\', slstComParts);
//    ExtractStrings([' '], [], PChar(sFullCommand), slstComParts);
    if nIndex < slstComParts.Count then
      Result := slstComParts[nIndex]
    else
      Result := '';
  finally
    slstComParts.Free;
  end;
end;

procedure TDBCommandManager.LoadDBCommand;
  procedure AddCmd(ACmd: string; ADescribe: string; AProc: TDBCommandProc);
  begin
    FDBCommandList.Add(TDBCommand.Create(ACmd, ADescribe, AProc));
  end;  
begin
  FDBCommandList.Clear;
  
  AddCmd(C_sDBCommand_Help, '命令帮助', DBCommand_Help);
  AddCmd(C_sDBCommand_ShowSettings, '显示当前设置',DBCommand_ShowSettings);
  AddCmd(C_sDBCommand_SetDefault, '所有设置恢复默认。', DBCommand_SetDefault);
  AddCmd(C_sDBCommand_MaxRecords, '设置最大记录数。默认不限制',DBCommand_MaxRecords);
  AddCmd(C_sDBCommand_PageIndex, '分页开启时设置当前页号，默认不分页',DBCommand_PageIndex);
  AddCmd(C_sDBCommand_Export, '导出数据',DBCommand_Export);
  AddCmd(C_sDBCommand_DropNoError, 'drop命令时忽略错误 on/off值',DBCommand_DropNoError);
  AddCmd(C_sDBCommand_Conn, '连接数据库，形式dbtype user/pass@dbname',DBCommand_Connect); 
  AddCmd(C_sDBCommand_Exit,'断开连接',DBCommand_Exit);
  AddCmd(C_sDBCommand_Var, '定义变量，定义形式name=value。使用形式$name$',DBCommand_Variable);
  AddCmd(C_sDBCommand_SQL, '专用sql，如可更新blob的sql，详见手册',DBCommand_SQL);
end;

function TDBCommandManager.DBCommand_Help(sCommandContent: string;
  var sError: string;var nEffectRows: Integer):Boolean;
var
  planarList: TPlanarStringList;
  cmd: TDBCommand;
  i: Integer;
begin
  FDBConnect.AddLog('Help 命令帮助');  
  FDBConnect.AddLog(Format('命令形式：%s command [params] 可在单行注释中，不区分大小写',
    [C_sHeadFlag_DBCommand]));
  FDBConnect.AddLog('==========================================');
  planarList := TPlanarStringList.Create;
  try
    for i := 0 to FDBCommandList.Count - 1 do
    begin
      cmd := FDBCommandList.Items[i];
      planarList.Strs[i,0] := cmd.CmdStr;
      planarList.Strs[i,1] := ' ' + cmd.Describe;
    end;

    planarList.JustifyMode := pjmBothLeft;
    planarList.FormatItemLengths;
    for i := 0 to planarList.Count - 1 do
    begin
      FDBConnect.AddLog(planarList.ItemStr[i]);
    end;   
  finally
    planarList.Free;
  end;              
  FDBConnect.AddLog('命令更详细的使用方法请参见手册！');
  FDBConnect.AddLog('==========================================');
  Result := True;
end;

function TDBCommandManager.DBCommand_ShowSettings(sCommandContent: string;
  var sError: string;var nEffectRows: Integer):Boolean;
var
//  planarList: TPlanarStringList;
  i: Integer;
begin
  for i := 0 to FDBCommandList.Count - 1 do
  begin

  end;
//  FDBConnect.AddLog('MaxRecords Info');                          
  FDBConnect.AddLog(' ');                     
  FDBConnect.AddLog('ShowSettings');
  FDBConnect.AddLog('==========================================');
  DBCommand_MaxRecords('', sError, nEffectRows);
  DBCommand_PageIndex('', sError, nEffectRows);
  DBCommand_DropNoError('', sError, nEffectRows);
  DBCommand_Variable('', sError, nEffectRows);
  
  Result := True;
end;

function TDBCommandManager.DBCommand_Export(sCommandContent: string;
  var sError: string;var nEffectRows: Integer):Boolean;
var
  params: TStrings;
  i: Integer;
  sTableName, sType, sOut: string;
  exp: TDBExport;
  function ReadNextParam(var index: Integer): string;
  begin
    if index + 1 < params.Count then
    begin
      Result := params[index+1];
      Inc(i);
    end
    else
      Result := ''
  end;  
begin
  params := TStringList.Create;  
  exp := TDBExport.Create;
  try
    fStrUtil.SplitEscapeQuote(sCommandContent, ' ', '''', '', params);

    i := -1;
    while i < params.Count-1 do
    begin
      Inc(i);
      if SameText('-table', params[i]) then
        sTableName := ReadNextParam(i)
      else if SameText('-type', params[i]) then
        sType := ReadNextParam(i)
      else if SameText('-out', params[i]) then
      begin
        sOut := ReadNextParam(i);
        if Pos(':', sOut) = 0 then
          sOut := ExtractFilePath(Application.ExeName) + sOut;
      end;
    end;
    exp.DBConnect := Self.DBConnect;
    Result := exp.ExportTableData(sTableName, StrToDBExportType(sType), sOut);    
  finally      
    exp.Free;
    params.Free;
  end;
end;

function TDBCommandManager.DBCommand_MaxRecords(sCommandContent: string;
  var sError: string;var nEffectRows: Integer):Boolean;
begin        
  if Trim(sCommandContent) <> '' then
  begin
    if FDBConnect.GetDBEngine <> nil then
      FDBConnect.GetDBEngine.CloseQuery;
    FDBConnect.MaxRecords := StrToIntDef(
      Trim(sCommandContent), C_nDefaultSetting_MaxRecords);
  end;
  FDBConnect.AddLog(C_sDBCommand_MaxRecords + ': ' + IntToStr(FDBConnect.MaxRecords));  
  Result := True;
end;    

function TDBCommandManager.DBCommand_PageIndex(sCommandContent: string;
  var sError: string;var nEffectRows: Integer): Boolean;
begin       
  if Trim(sCommandContent) <> '' then
  begin
    if fStrUtil.IsNumber(Trim(sCommandContent)) then
      FDBConnect.PageIndex := StrToInt(Trim(sCommandContent));
  end;
  FDBConnect.AddLog(C_sDBCommand_PageIndex + ': ' + IntToStr(FDBConnect.PageIndex));
  Result := True;
end;

function TDBCommandManager.DBCommand_SetDefault(sCommandContent: string;
  var sError: string;var nEffectRows: Integer):Boolean;
begin
  if FDBConnect.GetDBEngine <> nil then
    FDBConnect.GetDBEngine.CloseQuery;
  FDBConnect.ResetParams;
  FDBConnect.AddLog('所有设置重置为默认！');
  Result := True;
end;

function TDBCommandManager.DBCommand_DropNoError(sCommandContent: string;
  var sError: string; var nEffectRows: Integer): Boolean;
begin
  if Trim(sCommandContent) <> '' then
  begin
    FDBConnect.DropNoError := OnOff2Bool(sCommandContent);
  end;
  FDBConnect.AddLog(C_sDBCommand_DropNoError + ': ' + Bool2OnOff(FDBConnect.DropNoError));
  Result := True;
end; 

function TDBCommandManager.DBCommand_Stop(sCommandContent: string; var sError: string;
  var nEffectRows: Integer): Boolean;
begin     
  FDBConnect.StopExec();
  Result := True;
end;

function TDBCommandManager.DBCommand_Exit(sCommandContent: string; var sError: string;
  var nEffectRows: Integer): Boolean;
begin
  if FDBConnect.GetConnected then
    FDBConnect.CloseDB;
  Result := not FDBConnect.GetConnected;
end;

function TDBCommandManager.DBCommand_Connect(sCommandContent: string;
  var sError: string; var nEffectRows: Integer): Boolean;
var
  dbt: TDBType;
  sUserName: string;
  sPassword: string;
  sDB, sOtherDBParam: string;
  function GetValidFullPath(Str: string):string;
  begin
    if (Trim(Str) = '') or (Pos(':', Str) <> 0)
       or (Pos('\\', Str) = 1) then
      Result := Str
    else
      Result := ExtractFilePath(Application.ExeName)+Str;
  end;
begin
  if FDBConnect.GetConnected then
  begin
    FDBConnect.CloseDB;
    Application.ProcessMessages;
  end;
  
  if SplitConnectParams(sCommandContent, dbt, sUserName, sPassword, sDB,
    sOtherDBParam, sError) then
  begin
    if sOtherDBParam <> '' then
    case dbt of
      dbtUnKnown: ;
      dbtAccess, dbtAccess2007:
      begin
        sDB := BuildAccessDataSource(GetValidFullPath(sDB),
              GetValidFullPath(sOtherDBParam));
      end;
  //    dbtAccess2007: ;
      dbtOracle:
        sDB := sDB + C_sSEPARATOR_DATASOURCE + sOtherDBParam;
      dbtSyBase:
        sDB := sDB + C_sSEPARATOR_DATASOURCE + sOtherDBParam;
      dbtMySql:
        sDB := sDB + C_sSEPARATOR_DATASOURCE + sOtherDBParam;
    end;
    FDBConnect.OpenDB(sDB, sUserName, sPassword, dbt, dbetAuto);
  end
  else
  begin

  end;    
  
  Result := FDBConnect.GetConnected;
end;

function TDBCommandManager.DBCommand_Variable(sCommandContent: string;
  var sError: string;var nEffectRows: Integer):Boolean;
var
  i: Integer;
  nPos: Integer;
  sName, sValue: string;
  qry: TDataSet;
begin            
  Result := False;
  if Trim(sCommandContent) = '' then
  begin
    if FDBConnect.Variables.Count = 0 then
    begin
      FDBConnect.AddLog('Variables: no var');
    end
    else
    begin
      FDBConnect.AddLog('Variables: ');
      for i := 0 to FDBConnect.Variables.Count - 1 do
      begin
        FDBConnect.AddLog(Format('var %s=%s',[FDBConnect.Variables.Names[i],
          FDBConnect.Variables.Values[FDBConnect.Variables.Names[i]]]));
      end;
    end;
    Result := True;
  end
  else
  begin
    nPos := Pos('=', sCommandContent);
    if nPos = 0 then
    begin
      sError := '(DBCommand_Variable)变量表达式缺少等号';
      Exit;
    end;
    sName := Trim(Copy(sCommandContent, 1, nPos-1));
    sValue := Trim(Copy(sCommandContent, nPos+1, MaxInt));

    // 变量的值是个查询
    if AnsiStartsText(C_sHeadFlag_ForceQuery, sValue) then
    begin    
      if not FDBConnect.GetConnected then
      begin
        Result := False;
        sError := '(DBCommand_Variable)还未连接数据库，无法执行查询获得变量的值';
        Exit;
      end;
      qry := FDBConnect.GetNewQuery;
      try
        if FDBConnect.DBEngine.ExecQuery(qry,
           Copy(sValue, 1+Length(C_sHeadFlag_ForceQuery), MaxInt)) then
          if not qry.Eof then
            sValue := qry.Fields[0].AsString;
      finally
        qry.Free;
      end;
    end;

    if sValue <> '' then
    begin
      for i := 0 to FDBConnect.Variables.Count - 1 do
      begin
        if SameText(FDBConnect.Variables.Names[i], sName) then
        begin
          FDBConnect.Variables.Values[sName] := sValue;
          FDBConnect.AddLog(Format('var Replaced: %s=%s',[sName, sValue]));
          Result := True;
          Exit;
        end;
      end;
      FDBConnect.Variables.Add(sName+'='+sValue);
      FDBConnect.AddLog(Format('var Added: %s=%s',[sName, sValue]));
      Result := True;
    end
    else
    begin
      for i := FDBConnect.Variables.Count - 1 downto 0 do
      begin
        if SameText(FDBConnect.Variables.Names[i], sName) then
        begin
          FDBConnect.Variables.Delete(i);
          FDBConnect.AddLog(Format('var Deleted: %s',[sName]));
          Result := True;
          Exit;
        end;
      end;
      FDBConnect.AddLog(Format('no var been deleted',[sName]));
      Result := True;
    end;
  end;
end;  

function TDBCommandManager.DBCommand_SQL(sCommandContent: string;
  var sError: string;var nEffectRows: Integer): Boolean; 
var
  sSql: string;
  sCommand: string;
  sRelData: string;
begin
  Result := False;
  if not FDBConnect.Connected then
  begin
    nEffectRows := TDBC_ERROR_NOTCONNECT;
    Exit;
  end;
  
  sSql := Trim(sCommandContent); 
  if AnsiStartsText(C_sHeadFlag_Blob, sSql) then
  begin
    sSql := FDBConnect.RemoveHeadFlag(sSql);
    nEffectRows := FDBConnect.UpdateBlobs(sSql);
    sCommand := TSqlUtils.GetSqlCommand(sSql, sRelData);
    FDBConnect.AddSqlExecLog(nEffectRows, sCommand, sRelData);
  end
  else if AnsiStartsText(C_sHeadFlag_Clob, sSql) then
  begin
    sSql := FDBConnect.RemoveHeadFlag(sSql);
    nEffectRows := FDBConnect.UpdateClobs(sSql);
    sCommand := TSqlUtils.GetSqlCommand(sSql, sRelData);
    FDBConnect.AddSqlExecLog(nEffectRows, sCommand, sRelData);
  end
  else
  begin
    nEffectRows := FDBConnect.ExecOneSql(sCommandContent, FDBConnect.LastQry);
  end;
        
  Result := nEffectRows <> TDBC_ERROR_EXECFAIL;
end;

function TDBCommandManager.ProcessCommand(sCommandStr: string): Integer;
var
  nFind: Integer;
begin
  if Copy(sCommandStr, Length(sCommandStr), 1) = ';' then
  begin
    sCommandStr := Copy(sCommandStr, 1, Length(sCommandStr)-1);
  end;
  nFind := FDBCommandList.CompareAllCommand(GetParamByLevel(sCommandStr, 1));
  if nFind = 0 then                        // 0表示命令匹配成功
  begin
    if Pos(C_sDBCmdParam_Var_Start, sCommandStr) > 0 then
      DBConnect.SetVariableValue(sCommandStr);
    if not FDBCommandList.FindResult.Run(sCommandStr) then
    begin
      FDBConnect.AddErrorLog(FDBCommandList.FindResult.Error);
      Result := TDBC_ERROR_EXECFAIL;
    end else
      Result := FDBCommandList.FindResult.EffectRows;
  end
  else if nFind > 1 then begin             // 匹配失败但有相似命令
    FDBConnect.AddLog('命令不正确，是否' + FDBCommandList.FindResult.CmdStr);
    Result := TDBC_ERROR_EXECFAIL;
  end else begin                           // 匹配失败，无相似命令
    FDBConnect.AddLog('命令无法识别，键入'+C_sHeadFlag_DBCommand+' Help');
    Result := TDBC_ERROR_EXECFAIL;
  end;
end;

constructor TDBCommandManager.Create;
begin
  FDBCommandList := TDBCommandList.Create;
  LoadDBCommand;
end;

destructor TDBCommandManager.Destroy;
begin     
  if Assigned(FDBCommandList) then
    FreeAndNil(FDBCommandList);

  inherited;
end;

function TDBCommandManager.GetDBConnect: IDBConnect;
begin
  Result := FDBConnect;
end;

procedure TDBCommandManager.SetDBConnect(value: IDBConnect);
begin
  FDBConnect := value;
end;

end.
