unit U_ThreadPool;

{ aPool.AddRequest(TMyRequest.Create(RequestParam1, RequestParam2, ...)); }

interface

uses
  Windows, Classes, Forms, Contnrs, Generics.Collections, Generics.Defaults;

type
  { TCriticalSectionLock }
  TCriticalSectionLock = class(TObject)
  protected
    FSection: TRTLCriticalSection;
    FLocked: Boolean;
  public
    property Locked: Boolean read FLocked;

    constructor Create;
    destructor Destroy; override;
    // �����ٽ���
    procedure Lock;
    // �뿪�ٽ���
    procedure UnLock;
    // ���Խ���
    function TryLock: Boolean;
  end;

type
  // �����������ݵĻ�����
  TPoolTask = class(TObject)
  private
    FName: String;
    FCreateTime: TDateTime;
    FStartTime: TDateTime;
    FFinishTime: TDateTime;
  public
    property Name: String read FName write FName;
    property CreateTime: TDateTime read FCreateTime write FCreateTime;
    property StartTime: TDateTime read FStartTime write FStartTime;
    property FinishTime: TDateTime read FFinishTime write FFinishTime;

    constructor Create(); overload;
    constructor Create(name: String); overload;

    // ִ������
    procedure Run();virtual; abstract;
    function Equals(Obj: TObject): Boolean; override;
    function ToString: String; override;
  end;

type
  // ��־��־����   Level = 0 - ������Ϣ, 10 - ��������
  TThreadLogWriteProc = procedure(const Str: string; LogID: Integer; Level: Integer) of object;

type
  TThreadPool = class;

  // �߳�״̬
//  TThreadState = (tcsInitializing{�½�|��ʼ��}, tcsWaiting{�ȴ�|����}, tcsGetting, tcsProcessing{����},
//    tcsProcessed, tcsTerminating, tcsCheckingDown);
  TThreadState = (tcsInitializing{�½�|��ʼ��}, tcsWaiting{�ȴ�|����}, tcsProcessing{����}, tcsProcessed,
    tcsTerminating{����}, tcsCheckingDown);

  { TPoolProcessorThread �����߳̽������̳߳��ڣ� ��Ҫֱ�Ӵ�������������}
  TPoolProcessorThread = class(TThread)
  private
    // �����߳�ʱ��ʱ��Event����, �����߳�ֱ����ʼ�����
    hInitFinished: THandle;
    // ��ʼ��������Ϣ
    sInitError: string;
    // ��¼��־
    procedure WriteLog(const Str: string; Level: Integer = 0);
  protected
    // �߳��ٽ���ͬ������
    ProcessingDataObjectLock: TCriticalSectionLock;
    // ƽ������ʱ��
    FAverageProcessing: Integer;
    // �ȴ������ƽ��ʱ��
    FAverageWaitingTime: Integer;
    // ���߳�ʵ��������״̬
    FCurState: TThreadState;
    // ���߳�ʵ�����������̳߳�
    FPool: TThreadPool;
    // ��ǰ��������ݶ���
    FProcessingDataObject: TPoolTask;
    // �߳�ֹͣ Event, TProcessorThread.Terminate �п��̵�
    hThreadTerminated: THandle;
    uProcessingStart: DWORD;
    // ��ʼ�ȴ���ʱ��, ͨ�� GetTickCount ȡ�á�
    uWaitingStart: DWORD;
    // ����ƽ������ʱ��
    function AverageProcessingTime: DWORD;
    // ����ƽ���ȴ�ʱ��
    function AverageWaitingTime: DWORD;
    // ִ������
    procedure Execute; override;
    // ת��ö�����͵��߳�״̬Ϊ�ִ�����
    function InfoText: string;
    // �߳��Ƿ�ʱ�䴦��ͬһ�����󣿣�����������
    function IsDead: Boolean;
    // �߳��Ƿ�����ɵ�������
    function isFinished: Boolean;
    // �߳��Ƿ��ڿ���״̬
    function isIdle: Boolean;
    // ƽ��ֵУ�����㡣
    function NewAverage(OldAvg, NewVal: Integer): Integer;
  public
    Tag: Integer;
    constructor Create(APool: TThreadPool);
    destructor Destroy; override;
    procedure Terminate;
  end;

  // �̳߳�ʼ��ʱ�������¼�
  TProcessorThreadInitializing = procedure(Sender: TThreadPool;
                                aThread: TPoolProcessorThread) of object;
  // �߳̽���ʱ�������¼�
  TProcessorThreadFinalizing = procedure(Sender: TThreadPool;
                               aThread: TPoolProcessorThread) of object;
  // �̴߳�������ʱ�������¼�
  TProcessRequest = procedure(Sender: TThreadPool; WorkItem: TPoolTask;
                    aThread: TPoolProcessorThread) of object;
  TEmptyKind = (ekQueueEmpty, // ����ȡ�պ�
                ekProcessingFinished // ���һ����������Ϻ�
               );
  // ������п�ʱ�������¼�������������������񣬻����񶼴������
  TQueueEmpty = procedure(Sender: TThreadPool; EmptyKind: TEmptyKind) of object;
  // ���񶼴������ʱ�ᴥ�����൱��EmptyKind=ekProcessingFinished��ע�⣬��Ҫ��TQueueFinishʱ�ͷ�Pool���󣬱�����������
  TQueueFinish = procedure(Sender: TThreadPool) of object;

  { TThreadsPool }
  TThreadPool = class(TComponent)
  private
    QueueManagmentLock: TCriticalSectionLock;
    ThreadManagmentLock: TCriticalSectionLock;
    FProcessRequest: TProcessRequest;
    FOnWriteLog: TThreadLogWriteProc;
//    FQueue: TList;
    FQueue: TQueue;
    FQueueCapacity: Integer; // ���г���
    FQueueEmpty: TQueueEmpty;
    FQueueFinish: TQueueFinish;
    // �����߳�
    FThreads: TList<TPoolProcessorThread>;
    // �̳߳�ʱ��ֵ,0Ϊ����ʱ
    FThreadDeadTimeout: DWORD;
    // �̳߳�ʼ���¼���ÿ���̳߳�ʼ��ʱ����
    FThreadInitializing: TProcessorThreadInitializing;
    // �߳̽����¼���ÿ���߳̽���ʱ����
    FThreadFinalizing: TProcessorThreadFinalizing;
    // ִ���� terminat �����˳�ָ��, ���ڽ������߳�.
    FThreadsKilling: TList<TPoolProcessorThread>;
    FThreadsMax: Integer;    // ����߳���
    FThreadsMin: Integer;    // ����߳���

    FAllFinishedHandle: THandle;
    // ��ƽ���ȴ�ʱ��
    function PoolAverageWaitingTime: Integer;
    procedure WriteLog(const Str: string; Level: Integer = 0);
  protected
//    FLastGetPoint: Integer;
    // Semaphore�� ͳ���������
    hSemRequestCount: THandle;
    // Waitable timer. ÿ30����һ�ε�ʱ����ͬ��
    hTimCheckPoolDown: THandle;
    // �̳߳�ͣ������鲢��������̺߳����̣߳�
    procedure CheckPoolDown;
    // ������̣߳������䲻��Ĺ����߳�
    procedure CheckThreadsForGrow;
    procedure DoProcessed;
//    procedure DoProcessRequest(aDataObj: TPoolWorkItem; aThread: TPoolProcessorThread); virtual;
    procedure DoQueueEmpty(EmptyKind: TEmptyKind); virtual;
    procedure DoThreadFinalizing(aThread: TPoolProcessorThread); virtual;
    // ִ���¼�
    procedure DoThreadInitializing(aThread: TPoolProcessorThread); virtual;
    // �ͷ� FThreadsKilling �б��е��߳�
    procedure FreeFinishedThreads;
    // ��������
    procedure GetRequest(out Request: TPoolTask);
    // ������߳�
    procedure KillDeadThreads;

    // Ŀǰ������
  public
    constructor Create(AOwner: TComponent); overload; override;
    constructor Create(AOwner: TComponent; minThread, maxThread: Integer; queueCapacity: Integer); overload;
    destructor Destroy; override;

    // �ͽ��������Ƿ��ظ��ļ��, ��鷢���ظ��ͷ��� False
    function AddRequest(aDataObject: TPoolTask): Boolean;

    function WaitForAllFinish(timeOut: Integer = 0): Boolean;
  published
    property OnWriteLog: TThreadLogWriteProc read FOnWriteLog write FOnWriteLog;
    // �̴߳�������ʱ�������¼�
    property OnProcessRequest: TProcessRequest read FProcessRequest write FProcessRequest;
    // �����б�Ϊ��ʱ�ⷢ���¼�
    property OnQueueEmpty: TQueueEmpty read FQueueEmpty write FQueueEmpty;
    property OnQueueFinish: TQueueFinish read FQueueFinish write FQueueFinish;
    // �̳߳�ʼ��ʱ�������¼�
    property OnThreadInitializing: TProcessorThreadInitializing read FThreadInitializing write FThreadInitializing;
    // �߳̽���ʱ�������¼�
    property OnThreadFinalizing: TProcessorThreadFinalizing read FThreadFinalizing write FThreadFinalizing;
    // �̳߳�ʱֵ������), �������ʱ������Ϊ���߳�
    property ThreadDeadTimeout: DWORD read FThreadDeadTimeout write FThreadDeadTimeout default 0;
    // ����߳���
    property ThreadsMax: Integer read FThreadsMax write FThreadsMax default 1;
    // ��С�߳���
    property ThreadsMin: Integer read FThreadsMin write FThreadsMin default 0;
  end;

implementation

uses
  SysUtils, Math, TypInfo;

const
  LOG_LEVEL_DEBUG = 1;
  LOG_LEVEL_INFO = 2;
  LOG_LEVEL_WARN = 3;
  LOG_LEVEL_ERROR = 4;
  LOG_LEVEL_EXCEPTION = 9;


function TThreadStateToStr(Adbt: TThreadState): string;
var
  ti: PTypeInfo;
  td: PTypeData;
  i: Integer;
begin
  ti := TypeInfo(TThreadState);
  td := GetTypeData(ti);
  for i := td^.MinValue to td^.MaxValue do
    if Integer(Adbt) = i then begin
      Result := GetEnumName(ti, i);
      System.Break;
    end;
end;

{
  ******************************* TCriticalSectionLock *******************************
}

constructor TCriticalSectionLock.Create;
begin
  InitializeCriticalSection(FSection);
end;

destructor TCriticalSectionLock.Destroy;
begin
  DeleteCriticalSection(FSection);
end;

procedure TCriticalSectionLock.Lock;
begin
  EnterCriticalSection(FSection);
  FLocked := True;
end;

procedure TCriticalSectionLock.UnLock;
begin
  FLocked := False;
  LeaveCriticalSection(FSection);
end;

function TCriticalSectionLock.TryLock: Boolean;
begin
  Result := TryEnterCriticalSection(FSection);
  if Result then
    FLocked := True;
end;


// �����������ݵĻ�����
{
  ********************************** TRunnable ***********************************
}

constructor TPoolTask.Create(name: String);
begin
  Self.Name := name;
  FCreateTime := Now;
end;

constructor TPoolTask.Create;
var
  id: TGUID;
begin
  CreateGUID(id);
  Name := GUIDToString(id);
  FCreateTime := Now;
end;

function TPoolTask.Equals(Obj: TObject): Boolean;
begin
  Result := False;
  if not Assigned(obj) then
    Exit;
  if not (obj is TPoolTask) then
    Exit;
  Result := Name = TPoolTask(Obj).Name;
end;

function TPoolTask.ToString: String;
begin
  Result := Name;
end;

{
  ********************************* TThreadsPool *********************************
}

constructor TThreadPool.Create(AOwner: TComponent);
begin
  Create(AOwner, 1, 2, 0);
end;

constructor TThreadPool.Create(AOwner: TComponent; minThread, maxThread,
  queueCapacity: Integer);
var
  DueTo: Int64;
begin
  inherited Create(AOwner);
  WriteLog('�����̳߳�', LOG_LEVEL_INFO);
  QueueManagmentLock := TCriticalSectionLock.Create;
//  FQueue := TList.Create;
  FQueue := TObjectQueue.Create;
  FQueueCapacity := queueCapacity;
  ThreadManagmentLock := TCriticalSectionLock.Create;
  FThreads := TList<TPoolProcessorThread>.Create;
  FThreadsKilling := TList<TPoolProcessorThread>.Create;
  FThreadsMin := minThread;
  FThreadsMax := maxThread;
  FThreadDeadTimeout := 0;
//  FLastGetPoint := 0;
  //
  FAllFinishedHandle := CreateEvent(nil, True, False, nil);
  hSemRequestCount := CreateSemaphore(nil, 0, $7FFFFFFF, nil);

  DueTo := -1;
  // �ɵȴ��Ķ�ʱ����ֻ����Window NT4����ߣ�
  hTimCheckPoolDown := CreateWaitableTimer(nil, False, nil);

  if hTimCheckPoolDown = 0 then // Win9x��֧��
    // In Win9x number of thread will be never decrised
    hTimCheckPoolDown := CreateEvent(nil, False, False, nil)
  else
    SetWaitableTimer(hTimCheckPoolDown, DueTo, 30000, nil, nil, False);
end;

destructor TThreadPool.Destroy;
var
  n, i: Integer;
  Handles: array of THandle;
  task: TPoolTask;
begin
  WriteLog('�̳߳�����', LOG_LEVEL_INFO);
  ThreadManagmentLock.Lock;
  try
    SetLength(Handles, FThreads.Count);
    n := 0;
    for i := 0 to FThreads.Count - 1 do
      if FThreads[i] <> nil then
      begin
        Handles[n] := FThreads[i].Handle;
        FThreads[i].Terminate;
        Inc(n);
      end;
    WaitForMultipleObjects(n, @Handles[0], True, 10000);

    for i := 0 to FThreads.Count - 1 do
      FThreads[i].Free;
    FThreads.Free;
    FThreadsKilling.Free;
  finally
    ThreadManagmentLock.UnLock;
    ThreadManagmentLock.Free;
  end;

  QueueManagmentLock.Lock;
  try
    GetRequest(task);
    while Assigned(task) do begin
      FreeAndNil(task);
      GetRequest(task);
    end;
    FQueue.Free;
  finally
    QueueManagmentLock.UnLock;
    QueueManagmentLock.Free;
  end;

  CloseHandle(hSemRequestCount);
  CloseHandle(hTimCheckPoolDown);
  inherited;
end;

function TThreadPool.AddRequest(aDataObject: TPoolTask): Boolean;
//var
//  i: Integer;
begin
  Result := False;
  QueueManagmentLock.Lock;
  try
    ThreadManagmentLock.Lock;
    try
      // ������̣߳������䲻��Ĺ����߳�
      CheckThreadsForGrow;
    finally
      ThreadManagmentLock.UnLock;
    end;

    if (FQueueCapacity > 0) and (FQueue.Count >= FQueueCapacity) then begin
      WriteLog(Format('���г�������������������(%s)��', [aDataObject.Name]), LOG_LEVEL_WARN);
//      raise Exception.Create(Format('���г�������������������(%s)��', [aDataObject.Name]));
      Exit;
    end;

    // ������������
    FQueue.Push(aDataObject);
//    FQueue.Add(aDataObject);
    // �ͷ�һ��ͬ���ź������õȴ��е��̵߳õ�ִ��
    ReleaseSemaphore(hSemRequestCount, 1, nil);
    WriteLog(Format('������һ������(%s)', [aDataObject.Name]), LOG_LEVEL_DEBUG);
//    WriteLog('�ͷ�һ��ͬ���ź���', LOG_LEVEL_DEBUG);
    Result := True;
  finally
    QueueManagmentLock.UnLock;
  end;
end;

function TThreadPool.WaitForAllFinish(timeOut: Integer): Boolean;
//var
//  Handles:TWOHandleArray;
//  i: Integer;
begin
  Result := False;
  if timeOut = 0 then
    timeOut := INFINITE;
//  while time < timeOut do begin
//    Forms.Application.ProcessMessages;
//    Sleep(200);
//    Inc(time, 200);
//  end;

  // �������߳������������̳߳ص��߳�Ҳ������why��
  Result := WAIT_OBJECT_0 = WaitForSingleObject(FAllFinishedHandle, timeOut);
end;

{
  �� �� ����TThreadsPool.CheckPoolDown
  �����������̳߳�ͣ������鲢��������̺߳����̣߳�
  �����������
  �� �� ֵ: ��
  �������ڣ�2006.10.22 11:31
  �޸����ڣ�2006.
  ��    �ߣ�Kook
  ����˵����
}

procedure TThreadPool.CheckPoolDown;
var
  i: Integer;
begin
  WriteLog('TThreadsPool.CheckPoolDown', LOG_LEVEL_DEBUG);
  ThreadManagmentLock.Lock;
  try
    // ������߳�
    KillDeadThreads;
    // �ͷ� FThreadsKilling �б��е��߳�
    FreeFinishedThreads;

    // ����߳̿��У�����ֹ��
    for i := FThreads.Count - 1 downto FThreadsMin do
      if FThreads[i].isIdle then begin
        WriteLog(Format('��������߳�(%d)', [FThreads[i].ThreadID]), LOG_LEVEL_DEBUG);
        // ������ֹ����
        FThreads[i].Terminate;
        // ������������
        FThreadsKilling.Add(FThreads[i]);
        // �ӹ��������г���
        FThreads.Delete(i);
      end;
  finally
    ThreadManagmentLock.UnLock;
  end;
end;

{
  �� �� ����TThreadsPool.CheckThreadsForGrow
  ����������������̣߳������䲻��Ĺ����߳�
  �����������
  �� �� ֵ: ��
  �������ڣ�2006.10.22 11:31
  �޸����ڣ�2006.
  ��    �ߣ�Kook
  ����˵����
}

procedure TThreadPool.CheckThreadsForGrow;
var
  AvgWait: Integer;
  incThdCount, increasement: Integer;
  i: Integer;
begin
  {
    New thread created if:
    �½��̵߳�������
    1. �����߳���С����С�߳���
    2. �����߳���С������߳��� and �̳߳�ƽ���ȴ�ʱ�� < 100ms(ϵͳæ)
    3. ������ڹ����߳�����4��
  }
  increasement := 2; // һ����������2���߳�
  KillDeadThreads;
  if FThreads.Count < FThreadsMin then begin
    incThdCount := FThreadsMin - FThreads.Count;
    WriteLog('��ʼ���߳�����MinThreadCount', LOG_LEVEL_INFO);
    for i := 1 to incThdCount do
    try
      FThreads.Add(TPoolProcessorThread.Create(Self));
    except
      on e: Exception do
        WriteLog('TProcessorThread.Create raise: ' + e.ClassName +
            #13#10#9'Message: ' + e.Message, LOG_LEVEL_EXCEPTION);
    end
  end else if (FThreads.Count < FThreadsMax) then begin
    AvgWait := PoolAverageWaitingTime;
    WriteLog(Format('ThreadCount(%d) < MaxThreadCount(%d), AvgWartTime=%d',
        [FThreads.Count, FThreadsMax, AvgWait]), LOG_LEVEL_DEBUG);
    try
      if AvgWait < 100 then begin
        WriteLog('�����̣߳�ԭ��(WorkThreadCount < MaxThreadCount) and (AvgWartTime < 100ms)', LOG_LEVEL_INFO);
        incThdCount := Math.Min(increasement, (FThreadsMax - FThreads.Count));
        for i := 1 to incThdCount do begin
          FThreads.Add(TPoolProcessorThread.Create(Self));
        end;
      end else if (FQueue.Count > FThreadsMax * 4) then begin
        WriteLog('�����̣߳�ԭ��(WorkThreadCount < MaxThreadCount) and (QueueSize > MaxThreadCoun * 4)', LOG_LEVEL_INFO);
        incThdCount := Math.Min(increasement, (FThreadsMax - FThreads.Count));
        for i := 1 to incThdCount do begin
          FThreads.Add(TPoolProcessorThread.Create(Self));
        end;
      end;
    except
      on e: Exception do
        WriteLog('TThreadPool.CheckThreadsForGrow error: ' + e.ClassName +
            #13#10#9'Message: ' + e.Message, LOG_LEVEL_EXCEPTION);
    end;
  end;
end;

procedure TThreadPool.DoProcessed;
var
  i: Integer;
begin
  WriteLog(Format('DoProcessed Queue.Count=%d', [FQueue.Count]));
  if FQueue.Count > 0 then
    Exit;
//  if (FLastGetPoint < FQueue.Count) then
//    Exit;
  // ���Ƿ����ڴ����е��߳�
  ThreadManagmentLock.Lock;
  try
    for i := 0 to FThreads.Count - 1 do
      if FThreads[i].FCurState in [tcsProcessing] then
        Exit;
  finally
    ThreadManagmentLock.UnLock;
  end;
  SetEvent(FAllFinishedHandle);
  DoQueueEmpty(ekProcessingFinished);
end;

procedure TThreadPool.DoQueueEmpty(EmptyKind: TEmptyKind);
begin
  WriteLog('QueueEmpty! EmptyKind=' + IntToStr(Integer(EmptyKind)), LOG_LEVEL_DEBUG);
  if Assigned(FQueueEmpty) then
    FQueueEmpty(Self, EmptyKind);
  if EmptyKind = ekProcessingFinished then begin
    if Assigned(FQueueFinish) then
      FQueueFinish(Self);
  end;
end;

procedure TThreadPool.DoThreadFinalizing(aThread: TPoolProcessorThread);
begin
  if Assigned(FThreadFinalizing) then
    FThreadFinalizing(Self, aThread);
end;

procedure TThreadPool.DoThreadInitializing(aThread: TPoolProcessorThread);
begin
  if Assigned(FThreadInitializing) then
    FThreadInitializing(Self, aThread);
end;

procedure TThreadPool.FreeFinishedThreads;
var
  i: Integer;
begin
  if ThreadManagmentLock.TryLock then
  try
    for i := FThreadsKilling.Count - 1 downto 0 do
      if FThreadsKilling[i].isFinished then begin
        FThreadsKilling[i].Free;
        FThreadsKilling.Delete(i);
      end;
  finally
    ThreadManagmentLock.UnLock
  end;
end;

{
  �� �� ����TThreadsPool.GetRequest
  ������������������
  ���������out Request: TRequestDataObject
  �� �� ֵ: ��
  �������ڣ�2006.10.22 11:34
  �޸����ڣ�2006.
  ��    �ߣ�Kook
  ����˵����
}

procedure TThreadPool.GetRequest(out Request: TPoolTask);
begin
  QueueManagmentLock.Lock;
  try
    if FQueue.Count > 0 then begin
      Request := FQueue.Pop;
    end else begin
      Request := nil;
      DoQueueEmpty(ekQueueEmpty);
    end;
//    // �����յĶ���Ԫ��
//    while (FLastGetPoint < FQueue.Count) and (FQueue[FLastGetPoint] = nil) do
//      Inc(FLastGetPoint);
//
//    if (FLastGetPoint >= FQueue.Count) then begin
//      Request := nil;
//      Exit;
//    end;
//
//    // ѹ�����У������Ԫ��
//    if (FQueue.Count > 127) and (FLastGetPoint >= (3 * FQueue.Count) div 4) then
//    begin
//      WriteLog('ѹ�����У������Ԫ��', LOG_LEVEL_DEBUG);
//      FQueue.Pack;
//      FLastGetPoint := 0;
//    end;
//
//    Request := TPoolTask(FQueue[FLastGetPoint]);
//    FQueue[FLastGetPoint] := nil;
//    Inc(FLastGetPoint);
//    if (FLastGetPoint > 0) and (FLastGetPoint = FQueue.Count) then // ���������������
//    begin
//      DoQueueEmpty(ekQueueEmpty);
//      FQueue.Clear;
//      FLastGetPoint := 0;
//    end;
  finally
    QueueManagmentLock.UnLock;
  end;
end;

{
  �� �� ����TThreadsPool.KillDeadThreads
  ����������������߳�
  �����������
  �� �� ֵ: ��
  �������ڣ�2006.10.22 11:32
  �޸����ڣ�2006.
  ��    �ߣ�Kook
  ����˵����
}

procedure TThreadPool.KillDeadThreads;
var
  i: Integer;
begin
  // Check for dead threads
  if ThreadManagmentLock.TryLock then
  try
    for i := 0 to FThreads.Count - 1 do
      if FThreads[i].IsDead then
      begin
        WriteLog(Format('Thread(%d) dead', [FThreads[i].ThreadId]), LOG_LEVEL_WARN);
        // Dead thread moverd to other list.
        // New thread created to replace dead one
        FThreads[i].Terminate;
        FThreadsKilling.Add(FThreads[i]);
        try
          WriteLog(Format('�ؽ��������߳�, ID=', [FThreads[i].ThreadID]), LOG_LEVEL_DEBUG);
          FThreads[i] := TPoolProcessorThread.Create(Self);
        except
          on e: Exception do
          begin
            FThreads[i] := nil;
            WriteLog('TThreadPool.KillDeadThreads error! ' + e.ClassName +
                #13#10#9'Message: ' + e.Message, LOG_LEVEL_EXCEPTION);
          end;
        end;
      end;
  finally
    ThreadManagmentLock.UnLock;
  end;
end;

function TThreadPool.PoolAverageWaitingTime: Integer;
var
  i: Integer;
begin
  Result := 0;
  if FThreads.Count > 0 then
  begin
    for i := 0 to FThreads.Count - 1 do
      Inc(Result, FThreads[i].AverageWaitingTime);
    Result := Result div FThreads.Count
  end
  else
    Result := 1;
end;

procedure TThreadPool.WriteLog(const Str: string; Level: Integer = 0);
begin
  if Assigned(FOnWriteLog) then
    FOnWriteLog(Str, GetCurrentThreadId, Level);
end;

// �����߳̽������̳߳��ڣ� ��Ҫֱ�Ӵ�������������
{
  ******************************* TProcessorThread *******************************
}

constructor TPoolProcessorThread.Create(APool: TThreadPool);
begin
  WriteLog('���������߳�', 5);
  inherited Create(True);
  FPool := APool;

  FAverageWaitingTime := 1000;
  FAverageProcessing := 3000;

  sInitError := '';
  {
    ���������������£�
    ����һ������ nil ���ɡ�
    ���������Ƿ�����ֶ������ƺš�
    ���������ƺŵ���ʼ״̬��False ��ʾ��ơ�
    �����ģ�Event ����, ����������ͬ�Ļ�����ָ��ͬһ������������Ҫ������Event���󣬱�Ҫ��������ͬ������(���������ַ�������.ΪNIL�Ļ�ϵͳÿ�λ��Լ�����һ����ͬ�����֣����Ǳ��δ����Ķ����µ�EVENT)��
    ����ֵ��Event handle��
    }
  hInitFinished := CreateEvent(nil, True, False, nil);
  hThreadTerminated := CreateEvent(nil, True, False, nil);
  ProcessingDataObjectLock := TCriticalSectionLock.Create;
  try
    WriteLog('TProcessorThread.Create::Resume', 3);
    Resume;
    // ����, �ȴ���ʼ�����
    WaitForSingleObject(hInitFinished, INFINITE);
    if sInitError <> '' then
      raise Exception.Create(sInitError);
  finally
    CloseHandle(hInitFinished);
  end;
  WriteLog('TProcessorThread.Create::Finished', 3);
end;

destructor TPoolProcessorThread.Destroy;
begin
  WriteLog('�����߳�����', 5);
  CloseHandle(hThreadTerminated);
  ProcessingDataObjectLock.Free;
  inherited;
end;

function TPoolProcessorThread.AverageProcessingTime: DWORD;
begin
  if (FCurState in [tcsProcessing]) then
    Result := NewAverage(FAverageProcessing, GetTickCount - uProcessingStart)
  else
    Result := FAverageProcessing
end;

function TPoolProcessorThread.AverageWaitingTime: DWORD;
begin
  if (FCurState in [tcsWaiting, tcsCheckingDown]) then
    Result := NewAverage(FAverageWaitingTime, GetTickCount - uWaitingStart)
  else
    Result := FAverageWaitingTime
end;

procedure TPoolProcessorThread.Execute;

type
  THandleID = (hidTerminateThread, hidRequest, hidCheckPoolDown);
var
  WaitedTime: Integer;
  Handles: array [THandleID] of THandle;

begin
  WriteLog('�����߳̽�������', 3);
  // ��ǰ״̬����ʼ��
  FCurState := tcsInitializing;
  try
    // ִ���ⲿ�¼�
    FPool.DoThreadInitializing(Self);
  except
    on e: Exception do begin
      sInitError := e.Message;
    end;
  end;

  // ��ʼ����ɣ���ʼ��Event�̵�
  SetEvent(hInitFinished);

  WriteLog('TProcessorThread.Execute::Initialized', 3);

  // �����̳߳ص�ͬ�� Event
  Handles[hidTerminateThread] := hThreadTerminated;
  Handles[hidRequest] := FPool.hSemRequestCount;
  Handles[hidCheckPoolDown] := FPool.hTimCheckPoolDown;

  // ʱ�����
  // todo: �������߳����� GetTickCount; �᲻����
  uWaitingStart := GetTickCount;
  // �����ÿ�
  FProcessingDataObject := nil;

  // ��Ѳ��
  while not terminated do begin
    // ��ǰ״̬���ȴ�
    FCurState := tcsWaiting;
    // �����̣߳�ʹ�߳�����
    case WaitForMultipleObjects(Length(Handles), @Handles, False, INFINITE) - WAIT_OBJECT_0 of

      WAIT_OBJECT_0 + ord(hidTerminateThread):
        begin
          WriteLog('TProcessorThread.Execute:: Terminate event signaled ', LOG_LEVEL_DEBUG);
          // ��ǰ״̬��������ֹ�߳�
          FCurState := tcsTerminating;
          // �˳���Ѳ��(�����߳�)
          System.Break;
        end;

      WAIT_OBJECT_0 + ord(hidRequest):
        begin
          WriteLog('TProcessorThread.Execute:: Request semaphore signaled ', LOG_LEVEL_DEBUG);
          // �ȴ���ʱ��
          WaitedTime := GetTickCount - uWaitingStart;
          // ���¼���ƽ���ȴ�ʱ��
          FAverageWaitingTime := NewAverage(FAverageWaitingTime, WaitedTime);
          // ��ǰ״̬����������
          FCurState := tcsProcessing;
          // ���̳߳ص���������еõ�����
          FPool.GetRequest(FProcessingDataObject);
          if FProcessingDataObject = nil then begin
            Sleep(50);
            uWaitingStart := GetTickCount;
            Continue;
          end;

          WriteLog(Format('�̻߳�ȡ������(%s)', [FProcessingDataObject.Name]), LOG_LEVEL_INFO);
          try
            // ��ʼ�����ʱ���
            uProcessingStart := GetTickCount;

            // ��ǰ״̬��ִ������
            try
              WriteLog(Format('�߳�ִ������(%s)', [FProcessingDataObject.Name]), LOG_LEVEL_DEBUG);
              // ִ������
              FProcessingDataObject.StartTime := Now;
              FProcessingDataObject.Run;
              FProcessingDataObject.FinishTime := Now;

              with FProcessingDataObject do
                WriteLog(Format('�߳�ִ������(%s)��ϣ���ʱ%d ms', [FProcessingDataObject.Name,
                  Round((FinishTime - StartTime) * 24 * 3600 * 1000)]), LOG_LEVEL_DEBUG);
  //            FPool.DoProcessRequest(FProcessingDataObject, Self);

              // ���¼��㣬����δ���쳣������¼���
              FAverageProcessing := NewAverage(FAverageProcessing,
                GetTickCount - uProcessingStart);
              WriteLog(Format('Thread(%d) InfoText: ' + InfoText, [ThreadId]), LOG_LEVEL_DEBUG);
            except
              on e: Exception do
                WriteLog('ProcessRequest error! ' + FProcessingDataObject.Name +
                    #13#10'raise Exception: ' + e.Message, LOG_LEVEL_EXCEPTION);
            end;
          finally
            // �ͷ��������
            FreeAndNil(FProcessingDataObject);
          end;

          // ִ���߳����¼�
          FCurState := tcsProcessed;
          uWaitingStart := GetTickCount;
          FPool.DoProcessed;
        end;
      WAIT_OBJECT_0 + ord(hidCheckPoolDown):
        begin
          // !!! Never called under Win9x
          WriteLog('TProcessorThread.Execute:: CheckPoolDown timer signaled ', LOG_LEVEL_DEBUG);
          // ��ǰ״̬���̳߳�ͣ������鲢��������̺߳����̣߳�
          FCurState := tcsCheckingDown;
          FPool.CheckPoolDown;
        end;
    end;
  end;
  FCurState := tcsTerminating;

  FPool.DoThreadFinalizing(Self);
end;

function TPoolProcessorThread.InfoText: string;
begin
  Result := Format('%5d: %15s, AverageWaitingTime=%6d, AverageProcessingTime=%6d',
    [ThreadID, TThreadStateToStr(FCurState), AverageWaitingTime, AverageProcessingTime]);
  case FCurState of
    tcsWaiting:
      Result := Result + Format(', WaitingTime=%d', [GetTickCount - uWaitingStart]);
    tcsProcessing: begin
      Result := Result + Format(', ProcessingTime=%d', [GetTickCount - uProcessingStart]);
      if FProcessingDataObject <> nil then
        Result := Result + ', Task=' + FProcessingDataObject.Name;
    end;
  end;
end;

function TPoolProcessorThread.IsDead: Boolean;
begin
  { terminated״̬
    ���� ��ThreadDeadTimeout����0ʱ�������̳߳�ʱʱ��ThreadDeadTimeout
  }
  Result := terminated or
    ((FCurState = tcsProcessing) and (FPool.ThreadDeadTimeout > 0)
     and (GetTickCount - uProcessingStart > FPool.ThreadDeadTimeout));
end;

function TPoolProcessorThread.isFinished: Boolean;
begin
  Result := WaitForSingleObject(Handle, 0) = WAIT_OBJECT_0;
end;

function TPoolProcessorThread.isIdle: Boolean;
begin
  Result := False;
  // ����߳�״̬�� tcsWaiting, tcsCheckingDown
  // �� �ռ�ʱ�� > 30s
  //    ���� �ռ�ʱ�� > 10s and ƽ���Ⱥ�����ʱ�����ƽ������ʱ��� 50%
  // ����Ϊ���С�
  if (FCurState in [tcsWaiting, tcsCheckingDown]) then
    Result :=  (AverageWaitingTime > 30 * 1000)  or
      ((AverageWaitingTime > 10 * 1000) and (AverageWaitingTime * 2 > AverageProcessingTime));
end;

function TPoolProcessorThread.NewAverage(OldAvg, NewVal: Integer): Integer;
begin
  Result := (OldAvg * 2 + NewVal) div 3;
end;

procedure TPoolProcessorThread.Terminate;
begin
  WriteLog('TProcessorThread.Terminate', 5);
  inherited Terminate;
  SetEvent(hThreadTerminated);
end;

procedure TPoolProcessorThread.WriteLog(const Str: string; Level: Integer = 0);
begin
  if Assigned(FPool) and Assigned(FPool.OnWriteLog) then
    FPool.OnWriteLog(Str, ThreadID, Level);
end;

initialization

finalization

end.
