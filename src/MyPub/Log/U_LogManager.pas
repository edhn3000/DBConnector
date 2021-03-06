{**
 * <p>Title: U_LogManager </p>
 * <p>Copyright: Copyright (c) 2007-11-20 </p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 日志单元 </p>
 * <p>Description: 按天自动创建日志；新文件头记录主程序的名称和版本；日志级别分类等</p>
 * <p>Description: 从呼叫中心中修改分离出来，可单独使用此单元完成所有功能，如有特殊日志形式需求建议继承 TLogManager 2007-11-20</p>
 * <p>Description: 增加即时写入模式 2008-5-6</p>
 *}
unit U_LogManager;

{$WARN UNIT_PLATFORM OFF}

interface
uses
  windows, sysutils,syncobjs,classes,dialogs, FileCtrl,db, Variants;

const
  TTC_MINIMIUM_REMAINING_DISK_SPACE = 314572800;
  TTC_START_DELETE_REMAINING_DISK_SPACE = 314572800;
  TTC_INTERVAL_TIME = 10;  // 写日志到文件的时间间隔
  TTC_MAXKEEPDAYS = 60;     // 日志最大保留天数
  TTC_MINKEEPDAYS = 7;     // 最小保留天数
  TTC_MAXKEEYDAY_ALWAYS = 0;

  C_nLogLevel_NoLog        = 0;
  C_nLogLevel_Debug        = 1;
  C_nLogLevel_Info         = 2;
  C_nLogLevel_Warn         = 3;
  C_nLogLevel_Error        = 4;
  C_nLogLevel_Exception    = 5;

  C_as_LogLevel : array[0..5] of string = ('', '调试', '消息', '警告', '错误',
    '异常');

type
  TLogManager = class;
  TlogEvent = class;
  
{ Tthdloghandle }
  Tthdloghandle = class(TThread)
  private
    { Private declarations }
    FLogMgr: TLogManager;
    FNewLogAddEvent: TlogEvent;
   // fmemolog:Tmemo;
  protected
    procedure Execute; override;
  public
    constructor Create(ALogMgr: TLogManager); //;memolog:Tmemo
    destructor Destroy; override;
  end;

{ TThdcheck }
  TThdcheck = class(TThread)
  private
    FLogMgr: TLogManager;
    FTimerEvent: TlogEvent;
    FTimeCount: integer;
    FCurrentDate: TDateTime;
    FIsWarned: boolean;
    FnMaxKeepDays: Integer;
    FnMinKeepDays: Integer;
    FnInterval: Integer;

    function IsAnotherDay(ADate: TDateTime): boolean;
    procedure TestFreeDisk;
  protected
    procedure Execute; override;
  public
    property MaxKeepDays: Integer read FnMaxKeepDays write FnMaxKeepDays;
    property MinKeepDays: Integer read FnMinKeepDays write FnMinKeepDays;    
    property Interval: Integer read FnInterval write FnInterval;
  public
    constructor Create(ALogMgr: TLogManager);
    destructor Destroy; override;
  end;
  
{ TLogMessage }
  TLogMessage = class
  private
    FLogMsg: string;
  public
    constructor Create(AStr: string);
    property LogMessage: string read FLogMsg;
  end;

{ TlogEvent }
  TlogEvent = class(TEvent)
  public
    constructor Create(const Name: string; Manual: Boolean);//事件的创建
    procedure Signal;//事件执行状态的标志
    procedure Reset;
    function Wait (TimeOut: Dword): Boolean;
  end;  

{ TLogOption }
  TLogOption = class
  public
    LogDir: string;
    // 多数参数有默认值
    LogLevel: Integer;
    MaxKeepDays: Integer;
    MinKeepDays: Integer;
    CanLogSQL: Boolean;
    CanMulti: Boolean;
    LogInterval: Integer;  // seconds
    FlushInTime: Boolean;
  public
    constructor Create;
  end;

  TOnLogManagerLog = procedure (sFullLog: string; nLevel: Integer) of object;
  TOnLogManagerFullLog = procedure (sLog, sParams, sEvent, sResult: string;
    nLevel: Integer) of object;
{ TLogManager }
  TLogManager = class
  private
    FLogLevel : integer;
    FbLogSQL: Boolean;
    FLogMsgQueue: TList;
    FLogFile: TextFile; 
    FLogFileName: string;
    FQueueLock: TCriticalSection;
    FFileLock: TCriticalSection;
    FNewLogAddEvent: TlogEvent;
    FFileStoredDir: string;
    FAppFileName: string;    // 日志中记录的文件名及获取文件版本的依据，默认使用Application.ExeName
    FOnLog: TOnLogManagerLog;
    FOnFullLog: TOnLogManagerFullLog;
    procedure LockQueue;
    procedure UnLockQueue;
    procedure LockFile;
    procedure UnLockFile;
    procedure NewFileOfCurrentDay(ADate: TDateTime);
    function GetFileVersion(fn: string):string;
  protected
                     
    procedure CreateByDir(ADir: string);
  public
    TimeIntervalEvent: TlogEvent;
    NewLogAddEvent: TlogEvent;
    ThdLogHandler: TThdLogHandle;
    Thdcheck: TThdcheck;
    FlushInTime: Boolean;
  public
    property LogLevel: Integer read FLogLevel write FLogLevel;
    property LogSQL: Boolean read FbLogSQL write FbLogSQL;
    property LogFileName: string read FLogFileName;
    property AppFileName: string read FAppFileName write FAppFileName;
    constructor Create(ALogOption: TLogOption);virtual;
    destructor Destroy; override; 

    // 基本的添加log方法 被其他所有增加日志方法调用
    function AddLog( AStr: string; nLevel:Integer=C_nLogLevel_Debug ):Integer;
    // 包括事件、参数、返回值的完整日志
    function AddFullLog(sPrefix, sEvent: string;
        aryParamAndValues: array of Variant; vResult: Variant; sLogText: string;
        nLevel:Integer=C_nLogLevel_Debug):Integer;

    // 针对各种类别的Log方法
    function AddSQLLog(sMethod: string; sSQL: string;
        aryParams: array of Variant):Integer;overload;
    function AddSQLLog(sMethod: string; sSQL: string;
        slParams: TStrings):Integer;overload;

    function AddDebugLog(sEvent: string; aryParamAndValues: array of Variant;
        vResult: Variant; sLogText: string):Integer;
    function AddInfoLog(sEvent: string; aryParams: array of Variant;
        vResult: Variant; sLogText: string):Integer;
    function AddWarnLog(sEvent: string; aryParams: array of Variant;
        vResult: Variant; sLogText: string):Integer;
    function AddErrorLog(sEvent: string; aryParams: array of Variant;
        vResult: Variant; sLogText: string):Integer;
    function AddExceptionLog(sEvent: string; aryParams: array of Variant;
        vResult: Variant; sLogText: string; ex: Exception):Integer;

    function HasWaitingLog:boolean;
    function GetFirstWaitingLog: string;
    procedure WriteLogToFile(AStr: string);
    procedure FlushLogsToFile;
    procedure MoveToNextDay(ADate: TDateTime);
    procedure DelPreviousLogFiles(ADate: TDateTime; ADaysAgo: integer);

    property OnLog: TOnLogManagerLog read FOnLog write FOnLog;
    property OnFullLog: TOnLogManagerFullLog read FOnFullLog write FOnFullLog;
  end;

  function GetLogLevelMean(nLevel: Integer): string;


var
  g_log: TLogManager;
implementation

uses
  Forms;

function GetLogLevelMean(nLevel: Integer): string;
begin
  if (nLevel >= 0) and (nLevel <= High(C_as_LogLevel)) then
    Result := C_as_LogLevel[nLevel]
  else
    Result := '未知';
end;

function BuildLogTime: string;
begin
  Result := '['+FormatDateTime('yyyy/mm/dd hh:nn:ss:zzz', now)+']';
end;  

function VarValueToStr(v: Variant): string;
begin
  if VarIsType(v, varBoolean) then
  begin
    if v then  Result := 'True'
    else Result := 'False';
  end else Result := Variants.VarToStr(v);
end;

procedure FindFiles(sPath, sFileExt: string; AStrs: TStrings;
  SubDir, ShowDir: Boolean);
  function GetValidPath(APath: string):string;
  begin
    if Copy(APath, Length(APath), 1) <> '\' then  Result := APath + '\'
    else Result := APath;
  end;
var
  FSrchRec{, DSrchRec}: TSearchRec;
  nFindResult: Integer;
  nSearchMode: Integer;
begin
  sPath := GetValidPath(sPath);
  nSearchMode := faAnyFile;

  nFindResult := FindFirst( sPath + sFileExt, nSearchMode, FSrchRec);
  try
    while nFindResult = 0 do
    begin
      if (FSrchRec.Name <> '.') and (FSrchRec.Name <> '..') then
      begin
        if (FSrchRec.Attr and faDirectory) = faDirectory then
        begin
          if ShowDir then     // 是否显示目录
            AStrs.Add(sPath + GetValidPath(FSrchRec.Name));
          if SubDir then      // 是否搜索子目录模式
            FindFiles(sPath + FSrchRec.Name, sFileExt, AStrs, SubDir, ShowDir);
        end
        else
          AStrs.Add(sPath + FSrchRec.Name);
      end;
      
      nFindResult := FindNext(FSrchRec);
    end;
  finally
    SysUtils.FindClose(FSrchRec);
  end;
end;

function IsFileInUse(AFileName: string): Boolean;
var
 HFileRes: HFILE;
begin
  HFileRes := CreateFile( Pchar( AFileName ), GENERIC_READ,
                          0, nil, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0 );
  try
    Result:= (HFileRes = INVALID_HANDLE_VALUE);
  finally
    CloseHandle(HFileRes);
  end;
end; 

function IsDirInUse(sDir: string): Boolean;
var
  i: Integer;
  slstFiles: TStrings;
begin
  Result := False;
  slstFiles := TStringList.Create;
  try
    FindFiles(sDir, '*.*', slstFiles, True, False);
    for i := 0 to slstFiles.Count - 1 do
    begin
      if IsFileInUse(slstFiles[i]) then
      begin
        Result := True;
        Break;
      end;  
    end;  
  finally
    slstFiles.Free;
  end;
end;

{ TLogOption }

constructor TLogOption.Create;
begin
  LogLevel := C_nLogLevel_Debug;
  MaxKeepDays := TTC_MAXKEEPDAYS;
  MinKeepDays := TTC_MINKEEPDAYS;
  CanLogSQL := False;
  CanMulti := False;
  LogInterval := TTC_INTERVAL_TIME;
  FlushInTime := False;
end;

{ Tthdloghandle }

constructor Tthdloghandle.create(ALogMgr: TLogManager);    //;memolog:Tmemo
begin
  inherited create(true);
  FLogMgr := ALogMgr;
  FNewLogAddEvent := TlogEvent.Create('HAS_NEW_LOG_ADDED', false);
  Priority := tpLower;
end;

destructor Tthdloghandle.destroy;
begin
  FNewLogAddEvent.Free;
  inherited destroy;
end;

procedure Tthdloghandle.Execute;
var
  sTmp: string;
begin
  { Place thread code here }
  while true and (not terminated) do begin
    if FLogMgr.HasWaitingLog then
    begin
      sTmp := FLogMgr.GetFirstWaitingLog;
      if sTmp <> '' then
      begin
        FLogMgr.WriteLogToFile(sTmp);
      end; 
    end
    else
      FNewLogAddEvent.Wait(INFINITE);
  end;
end;

{ TThdcheck }

constructor TThdcheck.Create(ALogMgr: TLogManager);
begin
 { Place thread code here }
  inherited Create(true);
  FLogMgr := ALogMgr;
  FTimerEvent := TlogEvent.Create('TEN_SECONDS_INTERVAL_UP', false);
  FTimeCount := 1;
  FCurrentDate := now;
  FIsWarned := false;
  Priority := tpNormal;
  FnMaxKeepDays := TTC_MAXKEEPDAYS;    // default
  FnMinKeepDays := TTC_MINKEEPDAYS;
  FnInterval := TTC_INTERVAL_TIME;
end;

destructor TThdcheck.destroy;
begin
   FTimerEvent.Free;
   inherited destroy;  
end;

procedure TThdcheck.Execute;
var
  NowDate: TDateTime;
begin
  while (not terminated) do begin
//    FTimeCount := (FTimeCount + FnInterval) mod 3600;
    if FTimeCount mod FnInterval = 0 then begin
      FTimeCount := 1;
      FLogMgr.FlushLogsToFile;
      TestFreeDisk;
    end
    else
      Inc(FTimeCount);

    NowDate := now;

    if IsAnotherDay(NowDate) then
    begin
      FLogMgr.MoveToNextDay(NowDate);
      if FnMaxKeepDays > 0 then
      begin
        FLogMgr.DelPreviousLogFiles(NowDate, FnMaxKeepDays);
//        FLogMgr.AddLog('The data has been transfered successfully!');
//        FLogMgr.AddLog('Log files of '+IntToStr(FnMaxKeepDays)+' days ago have been deleted');
      end;
    end;
    FTimerEvent.Wait(1000);
    Sleep(1000);
  end;
 end;

function TThdcheck.IsAnotherDay(ADate: TDateTime): boolean;
var
  tmpOldDate, tmpNewDate: string;
begin
  tmpOldDate := FormatDateTime('yyyy/mm/dd', FCurrentDate);
  tmpNewDate := FormatDateTime('yyyy/mm/dd', ADate);
  if tmpOldDate<>tmpNewDate then begin
    result := true;
    FCurrentDate := ADate;
  end else
    result := false;
end;

procedure TThdcheck.TestFreeDisk;
var
  AmtFree: int64;//Cardinal;
begin
  AmtFree := DiskFree(0);

  if FIsWarned then begin
    FIsWarned := false;
    if (AmtFree<TTC_START_DELETE_REMAINING_DISK_SPACE)  and (AmtFree>=0) then begin
      FLogMgr.DelPreviousLogFiles(FCurrentDate, FnMinKeepDays);
      FLogMgr.AddLog('All log files of '+intToStr(FnMinKeepDays)+' days ago have been deleted');
      AmtFree := DiskFree(0);
    end;
  end;

  if (AmtFree < TTC_MINIMIUM_REMAINING_DISK_SPACE) and (AmtFree >= 0) then begin
    FLogMgr.AddLog('The remaining space is only: '+
    IntToStr(AmtFree)+' bytes. Please remove previous Log Files. All log files of 2 days ago will be deleted in one hour !!!');
    FIsWarned := true;
  end;
end;

{ TlogEvent }   

constructor TlogEvent.Create(const Name: string; Manual: Boolean);
begin
  inherited create(nil, Manual, false, PChar(Name));
end;

procedure TlogEvent.Reset;
begin
  inherited ResetEvent;
end;

procedure TlogEvent.Signal;
begin
  inherited SetEvent;
end;

function TlogEvent.Wait(TimeOut: Dword): Boolean;
begin
  if WaitFor(TimeOut) = wrSignaled then
    result := true
  else
    result := false;
end;

{ TLogManager }

function TLogManager.AddSQLLog(sMethod, sSQL: string; slParams: TStrings): Integer;
var
  i: Integer;
  aryParams: array of Variant;
begin
  Result := -1;
  if not FbLogSQL then Exit;
  if (slParams <> nil) and (slParams.Count > 0) then
  begin
    SetLength(aryParams, slParams.Count);
    for i := 0 to slParams.Count-1 do
    begin
      aryParams[i] := slParams[i];
    end;
  end
  else
    SetLength(aryParams, 0);
  Result := AddSQLLog(sMethod, sSQL, aryParams);
end;

function TLogManager.AddSQLLog(sMethod: string;sSQL: string; aryParams: array of Variant): Integer;
begin
  Result := -1;
  if not FbLogSQL then Exit;
  Result := AddFullLog('SQL', sMethod, aryParams, '', sSQL, C_nLogLevel_Debug);
end; 

function TLogManager.AddDebugLog(sEvent: string;
  aryParamAndValues: array of Variant; vResult: Variant; sLogText: string): Integer;
begin
  Result := AddFullLog('Debug', sEvent, aryParamAndValues, vResult, sLogText, C_nLogLevel_Debug);
end;    

function TLogManager.AddLog( AStr: string; nLevel : integer ): Integer;
begin
  Result := -1;
  LockQueue;
  try
    if ((FLogLevel <> C_nLogLevel_NoLog) and (nLevel >= FLogLevel))
       or (nLevel >= C_nLogLevel_Error) or (nLevel = C_nLogLevel_Exception) then       // 错误一定要记
    begin
      if (nLevel >= C_nLogLevel_Error) or (nLevel = C_nLogLevel_Exception)
         or FlushInTime then  // 现错误 或者 即时写入模式出。这样控制即时写入模式可在记录log前修改
      begin
        WriteLogToFile(BuildLogTime + ', ' +AStr);
        FlushLogsToFile;
      end
      else
      begin
        Result := FLogMsgQueue.Add(TLogMessage.Create(AStr));
      end;    
      if Assigned(FOnLog) then
        FOnLog(AStr, nLevel);
    end;
  finally
    UnLockQueue;
  end;
  FNewLogAddEvent.Signal;
end;

function TLogManager.AddFullLog(sPrefix, sEvent: string;
    aryParamAndValues: array of Variant; vResult: Variant; sLogText: string;
    nLevel:Integer): Integer;
var
  sFullLog, sParams: string;
  i: Integer;
begin
  if sEvent <> '' then
    sFullLog := '(' + sEvent + ')';
  // 参数
  if Length(aryParamAndValues) <> 0 then
  begin
    sParams := '';
    for i := Low(aryParamAndValues) to High(aryParamAndValues) do
      if sParams = '' then
        sParams := VarValueToStr(aryParamAndValues[i])
      else
        sParams := sParams + ',' + VarValueToStr(aryParamAndValues[i]);
    sFullLog := sFullLog + ' [' + sParams + ']';
  end;
  // 返回值
  if VarToStr(vResult) <> '' then
    sFullLog := sFullLog + ' <' + VarValueToStr(vResult) + '>';
  sFullLog :=  sLogText + ' ' + sFullLog;
  // 前缀 可传空则不添加前缀
  if sPrefix <> '' then
    sPrefix := sPrefix + ': ';
  // 增加日志
  Result := AddLog(sPrefix + sFullLog, nLevel);
  if Assigned(FOnFullLog) then
    FOnFullLog(sLogText, sParams, sEvent, VarValueToStr(vResult), nLevel);
end;

constructor TLogManager.Create(ALogOption: TLogOption);
var
  i: Integer;
  sLogDir: string;
begin
  inherited Create();
  sLogDir := ALogOption.LogDir;
  if not ALogOption.CanMulti then
    CreateByDir(sLogDir)
  else
  begin
    i := 1;
    while IsDirInUse(sLogDir) do
    begin
      sLogDir := ALogOption.LogDir + IntToStr(i);
      if not DirectoryExists(sLogDir) then    // 目录不存在相当于没有被使用
        Break;
      Inc(i);
    end;
    CreateByDir(sLogDir);
  end;

  LogLevel := ALogOption.LogLevel;   // 初始化
  LogSQL := ALogOption.CanLogSQL;
    
  TimeIntervalEvent := TlogEvent.Create('TEN_SECONDS_INTERVAL_UP', false);
  NewLogAddEvent := TlogEvent.Create('HAS_NEW_LOG_ADDED', false);

  Thdcheck:=TThdcheck.Create(Self);
  Thdcheck.Interval := ALogOption.LogInterval;
  Thdcheck.FreeOnTerminate := True;
  Thdcheck.MaxKeepDays := ALogOption.MaxKeepDays;
  Thdcheck.MinKeepDays := ALogOption.MinKeepDays;

  ThdLogHandler := TThdLogHandle.create(Self);
  ThdLogHandler.FreeOnTerminate := True;

  // 开始
  Thdcheck.Resume;
  ThdLogHandler.Resume;
end;

procedure TLogManager.DelPreviousLogFiles(ADate: TDateTime;
  ADaysAgo: integer);
var
  TmpDate: TDateTime;
  FileName: string;
  IndDay: integer;
  nExistLogCount: Integer;
  function GenFileName: string;
  begin
    result := FFileStoredDir+'\'+FormatDateTime('yyyymmdd', TmpDate)+'.Log';
  end;
begin
  nExistLogCount := 1;
  TmpDate := ADate - ADaysAgo;
  for IndDay := 1 to 120 do begin
    FileName := GenFileName;
    if FileExists(FileName) then
    begin
      if nExistLogCount > Thdcheck.MaxKeepDays then
        DeleteFile(PChar(FileName))
      else
        Inc(nExistLogCount);
    end;
    TmpDate := TmpDate - 1;
  end;
end;

destructor TLogManager.Destroy;
var
  i: integer;
begin
  if Assigned(TimeIntervalEvent) then
  begin
    TimeIntervalEvent.Signal;
    TimeIntervalEvent.Free;
  end;
  if Assigned(NewLogAddEvent) then
  begin
    NewLogAddEvent.Signal;
    NewLogAddEvent.Free;
  end;

  if Assigned(ThdLogHandler) then
    ThdLogHandler.Terminate;
  if Assigned(Thdcheck) then
    Thdcheck.Terminate;
      
  try
    Flush(FLogFile);
    CloseFile(FLogFile);
    for i:= FLogMsgQueue.Count-1 downto 0 do
    begin
      TLogMessage(FLogMsgQueue.Items[i]).Free;
    end;
    FLogMsgQueue.Free;
    FQueueLock.Free;
    FFileLock.Free;
    FNewLogAddEvent.Free;
  except

  end;
  inherited;
end;

procedure TLogManager.FlushLogsToFile;
begin
  LockFile;
  try
    Flush(FLogFile);
  finally
    UnLockFile;
  end;
end;

function TLogManager.GetFirstWaitingLog: string;
begin
  result := '';
  LockQueue;
  try
    if FLogMsgQueue.Count>0 then begin
      result := TLogMessage(FLogMsgQueue.items[0]).LogMessage;
      TLogMessage(FLogMsgQueue.items[0]).Free;
      FLogMsgQueue.Delete(0);
    end;
  finally
    UnLockQueue;
  end;  
end;

function TLogManager.HasWaitingLog: boolean;
begin
  if FLogMsgQueue.Count > 0 then
    result := true
  else
    result := false;
end;

procedure TLogManager.LockFile;
begin
  FFileLock.Enter;
end;

procedure TLogManager.LockQueue;
begin
  FQueueLock.Enter;
end;

procedure TLogManager.MoveToNextDay(ADate: TDateTime);
begin
  LockFile;
  try
    Flush(FLogFile);
    CloseFile(FLogFile);
    NewFileOfCurrentDay(ADate);
  finally
    UnLockFile;
  end;
end;

function TLogManager.GetFileVersion(fn: string):string;//得到本程序的版本号
var
  buf, p: pChar;
  sver: ^VS_FIXEDFILEINFO ;
  i: LongWord;
  sRes:string;
  iMajor,iMinor,iRelease,iBuild:integer;
  bResult: Boolean;
begin
  sRes:='';
  iMajor := 0;
  iMinor := 0;
  iRelease := 0;
  iBuild := 0;

  i:= GetFileVersionInfoSize(pchar(fn), i);
  new(sver);
  p:= pchar(sver);
  GetMem(buf, i);
  ZeroMemory(buf, i);
  bResult:= false;
  if GetFileVersionInfo(pchar(fn), 0, 4096, pointer(buf)) then
    if VerQueryValue(buf, '\', pointer(sver), i) then begin
      iMajor:= sVer^.dwFileVersionMS shr 16;
      iMinor:= sver^.dwFileVersionMS and $0000ffff;
      iRelease:= sver^.dwFileVersionLS shr 16;
      iBuild:= sver^.dwFileVersionLS and $0000ffff;
      bResult:= true;
    end;
  Dispose(p);
  FreeMem(buf);

  if bResult then
  begin
    sRes:= IntToStr(iMajor)+'.'+IntToStr(iMinor)
        +'.'+IntToStr(iRelease)+'.'+IntToStr(iBuild);
  end;
  Result:=sRes;
end;

procedure TLogManager.NewFileOfCurrentDay(ADate: TDateTime);
var
  sTitle: string;
  Stream: TStream;
  sAppName, sAppVer: string;
begin
  FLogFileName := FFileStoredDir + '\'+FormatDateTime('yyyymmdd', ADate)+'.Log';
  // 日志中加入文件名， 文件版本
  if (FAppFileName = '') or not FileExists(FAppFileName) then
    FAppFileName := Application.ExeName;
  sAppVer := GetFileVersion(FAppFileName);
  sAppName := ExtractFileName(FAppFileName);
  
  sTitle := Format('Following is the Logs of %s Version %s:  ', [
    ChangeFileExt(sAppName, ''), sAppVer])
    + FormatDateTime('yyyy/mm/dd', ADate) + ':' + #10;
  if not FileExists(FLogFileName) then
  begin
    Stream := TFileStream.Create(FLogFileName, fmCreate);
    try
      Stream.WriteBuffer(Pointer(sTitle)^, Length(sTitle));
    finally
      Stream.Free;
    end;
  end;

  AssignFile(FLogFile, FLogFileName);
  Append(FLogFile);
end;


procedure TLogManager.UnLockFile;
begin
 FFileLock.Leave;
end;

procedure TLogManager.UnLockQueue;
begin
  FQueueLock.Leave;
end;

procedure TLogManager.WriteLogToFile(AStr: string);
begin
  LockFile;
  try
    Writeln(FLogFile, AStr);
  finally
    UnLockFile;
  end; 
end;

function TLogManager.AddErrorLog(sEvent: string;
  aryParams: array of Variant; vResult: Variant; sLogText: string): Integer;
begin
  Result := AddFullLog('Error', sEvent, aryParams, vResult, sLogText,
    C_nLogLevel_Error);
end;

function TLogManager.AddExceptionLog(sEvent: string;
  aryParams: array of Variant; vResult: Variant; sLogText: string; ex: Exception): Integer;
begin
  sLogText := sLogText + ' ('+ ex.ClassName + ')'+ ex.Message;
  Result := AddFullLog('Exception', sEvent, aryParams, vResult, sLogText,
    C_nLogLevel_Error);
end;

function TLogManager.AddWarnLog(sEvent: string;
  aryParams: array of Variant; vResult: Variant; sLogText: string): Integer;
begin
  Result := AddFullLog('Warn', sEvent, aryParams, vResult, sLogText,
    C_nLogLevel_Warn);
end;

function TLogManager.AddInfoLog(sEvent: string;
  aryParams: array of Variant; vResult: Variant; sLogText: string): Integer;
begin
  Result := AddFullLog('Info', sEvent, aryParams, vResult, sLogText,
    C_nLogLevel_Info);
end;

procedure TLogManager.CreateByDir(ADir: string);
begin
  FFileStoredDir := ADir;
  if not DirectoryExists(ADir) then
    ForceDirectories(ADir);
  NewFileOfCurrentDay(now);
  FLogMsgQueue := TList.Create;
  FQueueLock := TCriticalSection.Create;
  FFileLock := TCriticalSection.Create;
  FNewLogAddEvent := TlogEvent.Create('HAS_NEW_LOG_ADDED', false);
  //设置访问name为'HAS_NEW_LOG_ADDED'，使用自动复位
end;

{ TLogMessage }

constructor TLogMessage.Create(AStr: string);
begin
  inherited create;
  FLogMsg := BuildLogTIme+ ', ' + AStr;
end;

end.
