{**
 * <p>Title: U_ThreadUtil.pas </p>
 * <p>Copyright: Copyright (c) 2007-3-20 </p>
 * <p>Company: Thunisoft</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: �Զ����߳���</p>
 * <p>Description: �����߳�ʵ��ʱ�������ʵ����ʹ��resume��ʼ�߳�</p>
 * <p>Description: �����źŵơ���������� 2007-09-12</p>
 * <p>Description: TLockObjectList�࣬�ɼ������б� �����̰߳�ȫ�б���� 2007-11-22</p>
 * <p>Description: �߳��๦������, resume����ʹ�߳̿��ظ�ִ��,suspend�����������ʱ���� </p>
 *}
unit U_ThreadUtil;

interface

uses
  Classes, Windows, SyncObjs, Contnrs, IniFiles;

type
  TMutex = class;
  TSemaphore = class;

  {TMyLock ������ֻ�����һ��������μ�����ȴ���
  // 2011-5-2ʹ��TRTLCriticalSection������EnterCriticalSection(FLock)��ʱ��ֱ�ӿ�ס
  // ���ɼ򵥵Ĳ�������ʵ�֡�
  }
  TMyLock = class
  private
    FLock: TRTLCriticalSection;
//    FLock: Boolean;
  public
    function Lock(ATimeOut: Integer = 0): Boolean;
    procedure UnLock;

    constructor Create;virtual;
    destructor Destroy;override;
  end;
  
{ TLockObjectList }
  TLockObjectList = class(TObjectList)
  private
    FLock: TRTLCriticalSection;
  public
    property LockObj: TRTLCriticalSection read FLock;

    procedure Lock;
    procedure UnLock;
    
    constructor Create;virtual;
    destructor Destroy;override;
    // ��ԭ��������ǰ��lock ��ɺ�unlock 
    function LockExtractFirst: TObject;
    function LockRemove(AObject: TObject): Integer;
  end;

{ TLockStringList }
  TLockStringList = class(TStringList)
  private
    FLock: TRTLCriticalSection;
  public              
    property LockObj: TRTLCriticalSection read FLock;
      
    procedure Lock;
    procedure UnLock;
    
    constructor Create;virtual;
    destructor Destroy;override;
  end;

{ THashedStringList }
  TLockHashedStringList = class(THashedStringList)
  private
    FLock: TRTLCriticalSection;
  public                     
    property LockObj: TRTLCriticalSection read FLock;
        
    procedure Lock;
    procedure UnLock;

    constructor Create;virtual;
    destructor Destroy;override;
  end;

{ THandleObjectExΪ��������źŵ���ĸ��� }
 THandleObjectEx = class(THandleObject)
  private
    function getLastWaitResultStr: string;
    function getLastErrorStr: string;
  protected
    Fhnd: THandle;
    FSystemLastError: Integer;
    FLastWaitResult: TWaitResult;
  public
    property LastError:Integer read FSystemLastError;
    property LastErrorStr: string read getLastErrorStr;
    property Handle: THandle read Fhnd;
    property LastWaitResult: TWaitResult read FLastWaitResult;
    property LastWaitResultStr: string read getLastWaitResultStr;
  public
    destructor Destroy; override;
    procedure Release;override;
    function WaitFor(Timeout: DWORD): Boolean;
  end;

{ TMutex ������ ���ý���֮��Ļ�����Դ���� }
  TMutex = class(THandleObjectEx)
  public                                                                                              
    constructor Create(const Name:string);overload;
    constructor Create(MutexAttributes: PSecurityAttributes;
      InitialOwner: Boolean;const Name:string);overload;
    // �ͷŶԻ�����������Ȩ
    procedure Release; override;
  end;

{ TSemaphore �źŵ��� }
  TSemaphore = class(THandleObjectEx)
  public
    ReleaseCount: Integer;
    PreviousCount:Pointer;
  public
   // �ĸ������ֱ�Ϊ����ȫ���ԡ���ʼ�źŵƼ���������źŵƼ������źŵ�����
    constructor Create(SemaphoreAttributes: PSecurityAttributes;
      InitialCount:Integer;MaximumCount: integer; const Name: string);
    procedure Release;override;
  end;
  
{ TProcOfObj }
  TProcOfObj = procedure of object;
  TProc = procedure;

{ TCustomThreadProxy �̴߳�����}
  TCustomThreadProxy = class
  private
    FProc: TProc;
    FProcOfObj: TProcOfObj;
    FEvent: TNotifyEvent;
    FOwner: TThread;
    FHasRun: Boolean;
  public
    constructor Create(AProc: TProc);overload;
    constructor Create(AProcOfObj: TProcOfObj);overload; 
    constructor Create(AEvent: TNotifyEvent);overload;
    destructor  Destroy;override;

    function HasRun: Boolean;
    procedure RunProc;
  end;

{ TCustomThread }
  TCustomThread = class(TThread)
  private                           
    FLastError: string;
    FProxy: TCustomThreadProxy;
    FDoBefore: TCustomThreadProxy;
    FDoAfter: TCustomThreadProxy;
    FDoOnAbort: TCustomThreadProxy;
    FExecuting: Boolean;

    function GetExecuting: Boolean;

    procedure SetThreadProxy(value: TCustomThreadProxy);
    procedure SetDoBefore(value: TCustomThreadProxy);
    procedure SetDoAfter(value: TCustomThreadProxy);
    procedure SetOnAbort(value: TCustomThreadProxy);

  protected
    procedure RunBeforeExecute;virtual;
    procedure RunMainProxy;virtual;
    procedure Execute;override;
    procedure RunAfterExecute;virtual;
  public
    property Executing: Boolean read GetExecuting;       
    property ThreadProxy: TCustomThreadProxy read FProxy write SetThreadProxy;
    // ����Excute����ִ��ʵ�ʵķ���֮ǰ
    property DoBeforeExecute: TCustomThreadProxy read FDoBefore write SetDoBefore;
    // Excuteִ�еĽ�βʱ����,����DoTerminate֮ǰ
    property DoAfterExecute: TCustomThreadProxy read FDoAfter write SetDoAfter;
    property DoOnAbort: TCustomThreadProxy read FDoOnAbort write SetOnAbort;
  public
    constructor Create(ACustomThreadProxy: TCustomThreadProxy;
      FreeOnTerminate: Boolean = True);
    destructor  Destroy;override;
                                      
    procedure Resume;   
    procedure Suspend;
    procedure Abort(WaitTimeOut: Integer=0);
  end;

  // �����߳̾�������߳�
  function KillThreadByHandle(ThreadHandle: THandle): Boolean;
  function WaitResultToStr(wr: TWaitResult): string;
  function WaitForExpect(var bVar: Boolean; bExpect: Boolean;
    TimeOut: Cardinal):TWaitResult;
  function GetSystemLastErrorStr(nLastError: Cardinal): string;
//  function CycleCount:Int64;

implementation

uses
  SysUtils, TypInfo, Forms;

function CycleCount:Int64;
asm
  db $0f
  db $32
end;

function WaitForExpect(var bVar: Boolean; bExpect: Boolean;
  TimeOut: Cardinal):TWaitResult;
var
  nOldTickCount,nTickCount: Cardinal;
begin
  nOldTickCount := GetTickCount;
  nTickCount := nOldTickCount;

  // û����GetTickCountѭ�����������ʱ���� Cardinal���ֵ4294967295
  while not (bVar=bExpect)
        and (nTickCount-nOldTickCount<TimeOut) do
  begin
    nTickCount := GetTickCount;
  end;
  if bVar=bExpect then
    Result := wrSignaled
  else
    Result := wrTimeout;
end;  

function KillThreadByHandle(ThreadHandle: THandle): Boolean;
var
  exitcode: Cardinal;
begin 
  GetExitCodeThread(ThreadHandle, exitcode);
  Result := TerminateThread(ThreadHandle, exitcode); 
end;

function WaitResultToStr(wr: TWaitResult): string;
var
  ti: PTypeInfo;
  td: PTypeData;
  i: Integer;
begin
  ti := TypeInfo(TWaitResult);
  td := GetTypeData(ti);
  for i := td^.MinValue to td^.MaxValue do
    if Integer(wr) = i then
    begin
      Result := GetEnumName(ti, i);
      Break;
    end;
end;     

{ TLock }

constructor TMyLock.Create;
begin
//  FLock := False;        
  InitializeCriticalSection(FLock);
end;

destructor TMyLock.Destroy;
begin
  DeleteCriticalSection(FLock);
  inherited;
end;

function TMyLock.Lock(ATimeOut: Integer): Boolean;
//var
//  dTime: TDateTime;
begin
//  dTime := Now;
//  while FLock
//        and ((ATimeOut <= 0)
//            or ((Now - dTIme) * 24 * 3600 * 1000 < ATimeOut)) do
//  begin
//    Sleep(100);
////    Application.ProcessMessages;
//  end;
//
//  if FLock then
//    Result := False
//  else
//    Result := True;
//
//  FLock := True;   
  EnterCriticalSection(FLock);
  Result := True;
end;

procedure TMyLock.UnLock;
begin
//  FLock := False;   
  LeaveCriticalSection(FLock);
end;

{ TLockObjectList }

constructor TLockObjectList.Create;
begin
  InitializeCriticalSection(FLock);
end;

destructor TLockObjectList.Destroy;
begin   
  DeleteCriticalSection(FLock);
  inherited;
end;

procedure TLockObjectList.Lock;
begin
  EnterCriticalSection(FLock);
end;   

procedure TLockObjectList.UnLock;
begin
  LeaveCriticalSection(FLock);
end;  

function TLockObjectList.LockExtractFirst: TObject;
begin
  Lock;
  try
    if Self.Count > 0 then
      Result := inherited Extract(Items[0])
    else
      Result := nil;
  finally
    UnLock;
  end;
end;

function TLockObjectList.LockRemove(AObject: TObject): Integer;
begin     
  Lock;
  try
    Result := inherited Remove(AObject);
  finally
    UnLock;
  end;
end;

{ TLockStringList }

constructor TLockStringList.Create;
begin
  InitializeCriticalSection(FLock);
end;

destructor TLockStringList.Destroy;
begin
  DeleteCriticalSection(FLock);
  inherited;
end;

procedure TLockStringList.Lock;
begin  
  EnterCriticalSection(FLock);
end;

procedure TLockStringList.UnLock;
begin           
  LeaveCriticalSection(FLock);
end;  


{ TLockHashedStringList }

constructor TLockHashedStringList.Create;
begin   
  InitializeCriticalSection(FLock);
end;

destructor TLockHashedStringList.Destroy;
begin
  DeleteCriticalSection(FLock);
  inherited;
end;

procedure TLockHashedStringList.Lock;
begin
  EnterCriticalSection(FLock);
end;

procedure TLockHashedStringList.UnLock;
begin
  LeaveCriticalSection(FLock);
end;

{ TCustomThreadStart }

constructor TCustomThreadProxy.Create(AProc: TProc);
begin
  FProc := AProc;
  FHasRun := False;
end;

constructor TCustomThreadProxy.Create(AProcOfObj: TProcOfObj);
begin
  FProcOfObj := AProcOfObj; 
  FHasRun := False;
end;

constructor TCustomThreadProxy.Create(AEvent: TNotifyEvent);
begin
  FEvent := AEvent;
end;

destructor TCustomThreadProxy.Destroy;
begin

  inherited;
end;

{ TCustomThread }

procedure TCustomThread.Abort(WaitTimeOut: Integer=0);
begin
//  Terminate;
//  WaitForExpect(FExecuting, False, WaitTimeOut);
  if Executing then
  begin
    if KillThreadByHandle(Self.Handle) then
      FExecuting := False
    else
      FLastError := GetSystemLastErrorStr(Windows.GetLastError);
    if Assigned(FDoOnAbort) then
      FDoOnAbort.RunProc;
  end;
end;

constructor TCustomThread.Create(ACustomThreadProxy: TCustomThreadProxy;
  FreeOnTerminate: Boolean);
begin
  inherited Create(True);
  FExecuting := False;    // ��Ϊʹ��Create(True)������FExecuting��ʼ��ΪFalse
  SetThreadProxy(ACustomThreadProxy);
  Self.FreeOnTerminate := FreeOnTerminate;
end;

destructor TCustomThread.Destroy;
begin
  if Executing then
  begin
    Abort();
  end;  
  if Assigned(FProxy) then
    FreeAndNil(FProxy);
  if Assigned(FDoBefore) then
    FreeAndNil(FDoBefore);
  if Assigned(FDoAfter) then
    FreeAndNil(FDoAfter);
  inherited;
end;

procedure TCustomThread.Execute;
//var
//  hr: HRESULT;
begin
  FExecuting := True;
  try
    // �߳̿�ʼǰ�Ĳ���
    Synchronize(RunBeforeExecute);
    try
//      hr := CoInitialize(nil);
//      if (hr = S_OK) or (hr = S_FALSE) then    // ��ʼ������
//      try
        Synchronize(RunMainProxy);
//      finally
//        CoUninitialize;
//      end
    except
     on ex: Exception do
       FLastError := ex.Message;
    end;

    // �߳̽�����Ĳ�����ȷ�е�˵�Ǽ�������
    Synchronize(RunAfterExecute);
  finally
    FExecuting := False;
  end;
end;   

procedure TCustomThread.RunBeforeExecute;
begin
  if Assigned(FDoBefore) then
    FDoBefore.RunProc;
end;

procedure TCustomThread.RunAfterExecute;
begin
  if Assigned(FDoAfter) then
  begin
    FDoAfter.RunProc;
//    with TCustomThread.Create(DoAfterExecute, True) do
//    begin
//      FreeOnTerminate := True;
//      Resume;
//    end;
  end;
end;

function TCustomThread.GetExecuting: Boolean;
begin
  Result := FExecuting and (not Suspended);
end;

procedure TCustomThread.SetThreadProxy(value: TCustomThreadProxy);
begin
  if Assigned(FProxy) then
    FreeAndNil(FProxy);
  FProxy := value;
  FProxy.FOwner := Self;
end;  

procedure TCustomThread.SetDoAfter(value: TCustomThreadProxy);
begin
  if Assigned(FDoAfter) then
    FreeAndNil(FDoAfter);
  FDoAfter := value;
  FDoAfter.FOwner := Self;
end;

procedure TCustomThread.SetDoBefore(value: TCustomThreadProxy);
begin
  if Assigned(FDoBefore) then
    FreeAndNil(FDoBefore);
  FDoBefore := value;
  FDoBefore.FOwner := Self;
end;  

procedure TCustomThread.SetOnAbort(value: TCustomThreadProxy);
begin     
  if Assigned(FDoOnAbort) then
    FreeAndNil(FDoOnAbort);
  FDoOnAbort := value;
  FDoOnAbort.FOwner := Self;
end;

function TCustomThreadProxy.HasRun: Boolean;
begin
  Result := FHasRun
end;

procedure TCustomThreadProxy.RunProc;
begin
  if Assigned(FEvent) or Assigned(FProc)
     or Assigned(FProcOfObj) then
    FHasRun := True;
  if Assigned(FEvent) then
    FEvent(FOwner)   // ��thread����
  else if Assigned(FProc) then
    FProc
  else if Assigned(FProcOfObj) then
    FprocOfObj
  else
    raise Exception.Create('�̴߳�����Ч���޷�ִ��!');
end; 

procedure TCustomThread.Resume;
begin
  // ��û�б�ִ�й�,��ִ��
  if not FProxy.HasRun then
    inherited Resume
  else
  begin
    // ���ִ�й�����Ŀǰ����ִ��״̬,����ִ�������
    // ��ʱ�̲߳��߱�ִ������,��Ҫ���´���һ��
    if not FExecuting then
    begin
      inherited Create(True);
      inherited Resume
    end
    else if Suspended then
      // ִ���в��ҹ�����
      inherited Resume;
  end;  
end;

procedure TCustomThread.Suspend;
begin
  if FExecuting then
    inherited Suspend;
end;

procedure TCustomThread.RunMainProxy;
begin        
  if Assigned(FProxy) then   // ���ú���
    FProxy.RunProc;
end;

{ THandleObjectEx }
destructor THandleObjectEx.Destroy;
begin
  if WaitFor(1000) then
    Windows.CloseHandle(Fhnd)
  else
  begin
    try
      Windows.CloseHandle(Fhnd);
    except
      //
    end;
  end;  
  inherited
end;

function THandleObjectEx.getLastErrorStr: string;
begin
  Result := GetSystemLastErrorStr(LastError);
end;

function THandleObjectEx.getLastWaitResultStr: string;
begin
  Result := WaitResultToStr(LastWaitResult);
end;

function GetSystemLastErrorStr(nLastError: Cardinal): string;
var
  acError: array[0..500] of Char;
begin
  FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, nil, nLastError,
      LANG_ENGLISH, @acError, Length(acError),  nil); 
  Result := Format('(%d)%s',[nLastError, string(acError)]);
end;

procedure THandleObjectEx.Release;
begin
  inherited;
end;

function THandleObjectEx.WaitFor(Timeout: DWORD): Boolean;
begin
  case WaitForSingleObject(Handle, Timeout) of
    WAIT_ABANDONED: FLastWaitResult := wrAbandoned;  //���ź�
    WAIT_OBJECT_0: FLastWaitResult := wrSignaled;    //���ź�
    WAIT_TIMEOUT: FLastWaitResult := wrTimeout;      //��ʱ
    WAIT_FAILED://ʧ��
    begin
      FLastWaitResult := wrError;
      FSystemLastError := Windows.GetLastError;
    end;
     else
     FLastWaitResult := wrError;
  end;
  Result := FLastWaitResult = wrSignaled;
end;

{ TMutex }
constructor TMutex.Create(MutexAttributes: PSecurityAttributes;
  InitialOwner: Boolean; const Name: string);
begin
  Fhnd := CreateMutex(MutexAttributes, InitialOwner, PChar(Name));
end;

constructor TMutex.Create(const Name: string);
begin
  Create(nil, False, PChar(Name));
end;

procedure TMutex.Release;
begin
  Windows.ReleaseMutex(Fhnd);
  inherited;
end; 

{ TSemaphore }
constructor TSemaphore.Create(SemaphoreAttributes: PSecurityAttributes;
   InitialCount, MaximumCount: integer; const Name: string);
begin
  Fhnd := CreateSemaphore
  (SemaphoreAttributes,InitialCount, 
  MaximumCount,PChar(Name));
  ReleaseCount := 1;
  PreviousCount := nil;
end;

procedure TSemaphore.Release;
begin
  Windows.ReleaseSemaphore(Fhnd, ReleaseCount, PreviousCount);
end;

end.

