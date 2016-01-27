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
    // 进入临界区
    procedure Lock;
    // 离开临界区
    procedure UnLock;
    // 尝试进入
    function TryLock: Boolean;
  end;

type
  // 储存请求数据的基本类
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

    // 执行任务
    procedure Run();virtual; abstract;
    function Equals(Obj: TObject): Boolean; override;
    function ToString: String; override;
  end;

type
  // 日志记志函数   Level = 0 - 跟踪信息, 10 - 致命错误
  TThreadLogWriteProc = procedure(const Str: string; LogID: Integer; Level: Integer) of object;

type
  TThreadPool = class;

  // 线程状态
//  TThreadState = (tcsInitializing{新建|初始化}, tcsWaiting{等待|就绪}, tcsGetting, tcsProcessing{运行},
//    tcsProcessed, tcsTerminating, tcsCheckingDown);
  TThreadState = (tcsInitializing{新建|初始化}, tcsWaiting{等待|就绪}, tcsProcessing{运行}, tcsProcessed,
    tcsTerminating{结束}, tcsCheckingDown);

  { TPoolProcessorThread 工作线程仅用于线程池内， 不要直接创建并调用它。}
  TPoolProcessorThread = class(TThread)
  private
    // 创建线程时临时的Event对象, 阻塞线程直到初始化完成
    hInitFinished: THandle;
    // 初始化出错信息
    sInitError: string;
    // 记录日志
    procedure WriteLog(const Str: string; Level: Integer = 0);
  protected
    // 线程临界区同步对像
    ProcessingDataObjectLock: TCriticalSectionLock;
    // 平均处理时间
    FAverageProcessing: Integer;
    // 等待请求的平均时间
    FAverageWaitingTime: Integer;
    // 本线程实例的运行状态
    FCurState: TThreadState;
    // 本线程实例所附属的线程池
    FPool: TThreadPool;
    // 当前处理的数据对像。
    FProcessingDataObject: TPoolTask;
    // 线程停止 Event, TProcessorThread.Terminate 中开绿灯
    hThreadTerminated: THandle;
    uProcessingStart: DWORD;
    // 开始等待的时间, 通过 GetTickCount 取得。
    uWaitingStart: DWORD;
    // 计算平均工作时间
    function AverageProcessingTime: DWORD;
    // 计算平均等待时间
    function AverageWaitingTime: DWORD;
    // 执行任务
    procedure Execute; override;
    // 转换枚举类型的线程状态为字串类型
    function InfoText: string;
    // 线程是否长时间处理同一个请求？（已死掉？）
    function IsDead: Boolean;
    // 线程是否已完成当成任务
    function isFinished: Boolean;
    // 线程是否处于空闲状态
    function isIdle: Boolean;
    // 平均值校正计算。
    function NewAverage(OldAvg, NewVal: Integer): Integer;
  public
    Tag: Integer;
    constructor Create(APool: TThreadPool);
    destructor Destroy; override;
    procedure Terminate;
  end;

  // 线程初始化时触发的事件
  TProcessorThreadInitializing = procedure(Sender: TThreadPool;
                                aThread: TPoolProcessorThread) of object;
  // 线程结束时触发的事件
  TProcessorThreadFinalizing = procedure(Sender: TThreadPool;
                               aThread: TPoolProcessorThread) of object;
  // 线程处理请求时触发的事件
  TProcessRequest = procedure(Sender: TThreadPool; WorkItem: TPoolTask;
                    aThread: TPoolProcessorThread) of object;
  TEmptyKind = (ekQueueEmpty, // 任务被取空后
                ekProcessingFinished // 最后一个任务处理完毕后
               );
  // 任务队列空时触发的事件，分两种情况，无任务，或任务都处理完毕
  TQueueEmpty = procedure(Sender: TThreadPool; EmptyKind: TEmptyKind) of object;
  // 任务都处理完毕时会触发，相当于EmptyKind=ekProcessingFinished，注意，不要在TQueueFinish时释放Pool对象，避免引起死锁
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
    FQueueCapacity: Integer; // 队列长度
    FQueueEmpty: TQueueEmpty;
    FQueueFinish: TQueueFinish;
    // 工作线程
    FThreads: TList<TPoolProcessorThread>;
    // 线程超时阀值,0为不超时
    FThreadDeadTimeout: DWORD;
    // 线程初始化事件，每个线程初始化时触发
    FThreadInitializing: TProcessorThreadInitializing;
    // 线程结束事件，每个线程结束时触发
    FThreadFinalizing: TProcessorThreadFinalizing;
    // 执行了 terminat 发送退出指令, 正在结束的线程.
    FThreadsKilling: TList<TPoolProcessorThread>;
    FThreadsMax: Integer;    // 最大线程数
    FThreadsMin: Integer;    // 最大线程数

    FAllFinishedHandle: THandle;
    // 池平均等待时间
    function PoolAverageWaitingTime: Integer;
    procedure WriteLog(const Str: string; Level: Integer = 0);
  protected
//    FLastGetPoint: Integer;
    // Semaphore， 统计任务队列
    hSemRequestCount: THandle;
    // Waitable timer. 每30触发一次的时间量同步
    hTimCheckPoolDown: THandle;
    // 线程池停机（检查并清除空闲线程和死线程）
    procedure CheckPoolDown;
    // 清除死线程，并补充不足的工作线程
    procedure CheckThreadsForGrow;
    procedure DoProcessed;
//    procedure DoProcessRequest(aDataObj: TPoolWorkItem; aThread: TPoolProcessorThread); virtual;
    procedure DoQueueEmpty(EmptyKind: TEmptyKind); virtual;
    procedure DoThreadFinalizing(aThread: TPoolProcessorThread); virtual;
    // 执行事件
    procedure DoThreadInitializing(aThread: TPoolProcessorThread); virtual;
    // 释放 FThreadsKilling 列表中的线程
    procedure FreeFinishedThreads;
    // 申请任务
    procedure GetRequest(out Request: TPoolTask);
    // 清除死线程
    procedure KillDeadThreads;

    // 目前不能用
  public
    constructor Create(AOwner: TComponent); overload; override;
    constructor Create(AOwner: TComponent; minThread, maxThread: Integer; queueCapacity: Integer); overload;
    destructor Destroy; override;

    // 就进行任务是否重复的检查, 检查发现重复就返回 False
    function AddRequest(aDataObject: TPoolTask): Boolean;

    function WaitForAllFinish(timeOut: Integer = 0): Boolean;
  published
    property OnWriteLog: TThreadLogWriteProc read FOnWriteLog write FOnWriteLog;
    // 线程处理任务时触发的事件
    property OnProcessRequest: TProcessRequest read FProcessRequest write FProcessRequest;
    // 任务列表为空时解发的事件
    property OnQueueEmpty: TQueueEmpty read FQueueEmpty write FQueueEmpty;
    property OnQueueFinish: TQueueFinish read FQueueFinish write FQueueFinish;
    // 线程初始化时触发的事件
    property OnThreadInitializing: TProcessorThreadInitializing read FThreadInitializing write FThreadInitializing;
    // 线程结束时触发的事件
    property OnThreadFinalizing: TProcessorThreadFinalizing read FThreadFinalizing write FThreadFinalizing;
    // 线程超时值（毫秒), 如果处理超时，将视为死线程
    property ThreadDeadTimeout: DWORD read FThreadDeadTimeout write FThreadDeadTimeout default 0;
    // 最大线程数
    property ThreadsMax: Integer read FThreadsMax write FThreadsMax default 1;
    // 最小线程数
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


// 储存请求数据的基本类
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
  WriteLog('创建线程池', LOG_LEVEL_INFO);
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
  // 可等待的定时器（只用于Window NT4或更高）
  hTimCheckPoolDown := CreateWaitableTimer(nil, False, nil);

  if hTimCheckPoolDown = 0 then // Win9x不支持
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
  WriteLog('线程池销毁', LOG_LEVEL_INFO);
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
      // 清除死线程，并补充不足的工作线程
      CheckThreadsForGrow;
    finally
      ThreadManagmentLock.UnLock;
    end;

    if (FQueueCapacity > 0) and (FQueue.Count >= FQueueCapacity) then begin
      WriteLog(Format('队列超过限制数，丢弃任务(%s)！', [aDataObject.Name]), LOG_LEVEL_WARN);
//      raise Exception.Create(Format('队列超过限制数，丢弃任务(%s)！', [aDataObject.Name]));
      Exit;
    end;

    // 将任务加入队列
    FQueue.Push(aDataObject);
//    FQueue.Add(aDataObject);
    // 释放一个同步信号量，让等待中的线程得到执行
    ReleaseSemaphore(hSemRequestCount, 1, nil);
    WriteLog(Format('增加了一个任务(%s)', [aDataObject.Name]), LOG_LEVEL_DEBUG);
//    WriteLog('释放一个同步信号量', LOG_LEVEL_DEBUG);
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

  // 除了主线程阻塞，会让线程池的线程也阻塞，why？
  Result := WAIT_OBJECT_0 = WaitForSingleObject(FAllFinishedHandle, timeOut);
end;

{
  函 数 名：TThreadsPool.CheckPoolDown
  功能描述：线程池停机（检查并清除空闲线程和死线程）
  输入参数：无
  返 回 值: 无
  创建日期：2006.10.22 11:31
  修改日期：2006.
  作    者：Kook
  附加说明：
}

procedure TThreadPool.CheckPoolDown;
var
  i: Integer;
begin
  WriteLog('TThreadsPool.CheckPoolDown', LOG_LEVEL_DEBUG);
  ThreadManagmentLock.Lock;
  try
    // 清除死线程
    KillDeadThreads;
    // 释放 FThreadsKilling 列表中的线程
    FreeFinishedThreads;

    // 如果线程空闲，就终止它
    for i := FThreads.Count - 1 downto FThreadsMin do
      if FThreads[i].isIdle then begin
        WriteLog(Format('清理空闲线程(%d)', [FThreads[i].ThreadID]), LOG_LEVEL_DEBUG);
        // 发出终止命令
        FThreads[i].Terminate;
        // 加入待清除队列
        FThreadsKilling.Add(FThreads[i]);
        // 从工作队列中除名
        FThreads.Delete(i);
      end;
  finally
    ThreadManagmentLock.UnLock;
  end;
end;

{
  函 数 名：TThreadsPool.CheckThreadsForGrow
  功能描述：清除死线程，并补充不足的工作线程
  输入参数：无
  返 回 值: 无
  创建日期：2006.10.22 11:31
  修改日期：2006.
  作    者：Kook
  附加说明：
}

procedure TThreadPool.CheckThreadsForGrow;
var
  AvgWait: Integer;
  incThdCount, increasement: Integer;
  i: Integer;
begin
  {
    New thread created if:
    新建线程的条件：
    1. 工作线程数小于最小线程数
    2. 工作线程数小于最大线程数 and 线程池平均等待时间 < 100ms(系统忙)
    3. 任务大于工作线程数的4倍
  }
  increasement := 2; // 一次最少增加2个线程
  KillDeadThreads;
  if FThreads.Count < FThreadsMin then begin
    incThdCount := FThreadsMin - FThreads.Count;
    WriteLog('初始化线程数到MinThreadCount', LOG_LEVEL_INFO);
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
        WriteLog('创建线程，原因：(WorkThreadCount < MaxThreadCount) and (AvgWartTime < 100ms)', LOG_LEVEL_INFO);
        incThdCount := Math.Min(increasement, (FThreadsMax - FThreads.Count));
        for i := 1 to incThdCount do begin
          FThreads.Add(TPoolProcessorThread.Create(Self));
        end;
      end else if (FQueue.Count > FThreadsMax * 4) then begin
        WriteLog('创建线程，原因：(WorkThreadCount < MaxThreadCount) and (QueueSize > MaxThreadCoun * 4)', LOG_LEVEL_INFO);
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
  // 看是否还有在处理中的线程
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
  函 数 名：TThreadsPool.GetRequest
  功能描述：申请任务
  输入参数：out Request: TRequestDataObject
  返 回 值: 无
  创建日期：2006.10.22 11:34
  修改日期：2006.
  作    者：Kook
  附加说明：
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
//    // 跳过空的队列元素
//    while (FLastGetPoint < FQueue.Count) and (FQueue[FLastGetPoint] = nil) do
//      Inc(FLastGetPoint);
//
//    if (FLastGetPoint >= FQueue.Count) then begin
//      Request := nil;
//      Exit;
//    end;
//
//    // 压缩队列，清除空元素
//    if (FQueue.Count > 127) and (FLastGetPoint >= (3 * FQueue.Count) div 4) then
//    begin
//      WriteLog('压缩队列，清除空元素', LOG_LEVEL_DEBUG);
//      FQueue.Pack;
//      FLastGetPoint := 0;
//    end;
//
//    Request := TPoolTask(FQueue[FLastGetPoint]);
//    FQueue[FLastGetPoint] := nil;
//    Inc(FLastGetPoint);
//    if (FLastGetPoint > 0) and (FLastGetPoint = FQueue.Count) then // 如果队列中无任务
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
  函 数 名：TThreadsPool.KillDeadThreads
  功能描述：清除死线程
  输入参数：无
  返 回 值: 无
  创建日期：2006.10.22 11:32
  修改日期：2006.
  作    者：Kook
  附加说明：
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
          WriteLog(Format('重建死掉的线程, ID=', [FThreads[i].ThreadID]), LOG_LEVEL_DEBUG);
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

// 工作线程仅用于线程池内， 不要直接创建并调用它。
{
  ******************************* TProcessorThread *******************************
}

constructor TPoolProcessorThread.Create(APool: TThreadPool);
begin
  WriteLog('创建工作线程', 5);
  inherited Create(True);
  FPool := APool;

  FAverageWaitingTime := 1000;
  FAverageProcessing := 3000;

  sInitError := '';
  {
    各参数的意义如下：
    参数一：填上 nil 即可。
    参数二：是否采用手动调整灯号。
    参数三：灯号的起始状态，False 表示红灯。
    参数四：Event 名称, 对象名称相同的话，会指向同一个对象，所以想要有两个Event对象，便要有两个不同的名称(这名称以字符串来存.为NIL的话系统每次会自己创建一个不同的名字，就是被次创建的都是新的EVENT)。
    传回值：Event handle。
    }
  hInitFinished := CreateEvent(nil, True, False, nil);
  hThreadTerminated := CreateEvent(nil, True, False, nil);
  ProcessingDataObjectLock := TCriticalSectionLock.Create;
  try
    WriteLog('TProcessorThread.Create::Resume', 3);
    Resume;
    // 阻塞, 等待初始化完成
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
  WriteLog('工作线程销毁', 5);
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
  WriteLog('工作线程进常运行', 3);
  // 当前状态：初始化
  FCurState := tcsInitializing;
  try
    // 执行外部事件
    FPool.DoThreadInitializing(Self);
  except
    on e: Exception do begin
      sInitError := e.Message;
    end;
  end;

  // 初始化完成，初始化Event绿灯
  SetEvent(hInitFinished);

  WriteLog('TProcessorThread.Execute::Initialized', 3);

  // 引用线程池的同步 Event
  Handles[hidTerminateThread] := hThreadTerminated;
  Handles[hidRequest] := FPool.hSemRequestCount;
  Handles[hidCheckPoolDown] := FPool.hTimCheckPoolDown;

  // 时间戳，
  // todo: 好像在线程中用 GetTickCount; 会不正常
  uWaitingStart := GetTickCount;
  // 任务置空
  FProcessingDataObject := nil;

  // 大巡环
  while not terminated do begin
    // 当前状态：等待
    FCurState := tcsWaiting;
    // 阻塞线程，使线程休眠
    case WaitForMultipleObjects(Length(Handles), @Handles, False, INFINITE) - WAIT_OBJECT_0 of

      WAIT_OBJECT_0 + ord(hidTerminateThread):
        begin
          WriteLog('TProcessorThread.Execute:: Terminate event signaled ', LOG_LEVEL_DEBUG);
          // 当前状态：正在终止线程
          FCurState := tcsTerminating;
          // 退出大巡环(结束线程)
          System.Break;
        end;

      WAIT_OBJECT_0 + ord(hidRequest):
        begin
          WriteLog('TProcessorThread.Execute:: Request semaphore signaled ', LOG_LEVEL_DEBUG);
          // 等待的时间
          WaitedTime := GetTickCount - uWaitingStart;
          // 重新计算平均等待时间
          FAverageWaitingTime := NewAverage(FAverageWaitingTime, WaitedTime);
          // 当前状态：申请任务
          FCurState := tcsProcessing;
          // 从线程池的任务队列中得到任务
          FPool.GetRequest(FProcessingDataObject);
          if FProcessingDataObject = nil then begin
            Sleep(50);
            uWaitingStart := GetTickCount;
            Continue;
          end;

          WriteLog(Format('线程获取到任务(%s)', [FProcessingDataObject.Name]), LOG_LEVEL_INFO);
          try
            // 开始处理的时间戳
            uProcessingStart := GetTickCount;

            // 当前状态：执行任务
            try
              WriteLog(Format('线程执行任务(%s)', [FProcessingDataObject.Name]), LOG_LEVEL_DEBUG);
              // 执行任务
              FProcessingDataObject.StartTime := Now;
              FProcessingDataObject.Run;
              FProcessingDataObject.FinishTime := Now;

              with FProcessingDataObject do
                WriteLog(Format('线程执行任务(%s)完毕！耗时%d ms', [FProcessingDataObject.Name,
                  Round((FinishTime - StartTime) * 24 * 3600 * 1000)]), LOG_LEVEL_DEBUG);
  //            FPool.DoProcessRequest(FProcessingDataObject, Self);

              // 重新计算，仅在未出异常的情况下计算
              FAverageProcessing := NewAverage(FAverageProcessing,
                GetTickCount - uProcessingStart);
              WriteLog(Format('Thread(%d) InfoText: ' + InfoText, [ThreadId]), LOG_LEVEL_DEBUG);
            except
              on e: Exception do
                WriteLog('ProcessRequest error! ' + FProcessingDataObject.Name +
                    #13#10'raise Exception: ' + e.Message, LOG_LEVEL_EXCEPTION);
            end;
          finally
            // 释放任务对象
            FreeAndNil(FProcessingDataObject);
          end;

          // 执行线程外事件
          FCurState := tcsProcessed;
          uWaitingStart := GetTickCount;
          FPool.DoProcessed;
        end;
      WAIT_OBJECT_0 + ord(hidCheckPoolDown):
        begin
          // !!! Never called under Win9x
          WriteLog('TProcessorThread.Execute:: CheckPoolDown timer signaled ', LOG_LEVEL_DEBUG);
          // 当前状态：线程池停机（检查并清除空闲线程和死线程）
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
  { terminated状态
    或者 当ThreadDeadTimeout大于0时，超过线程超时时间ThreadDeadTimeout
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
  // 如果线程状态是 tcsWaiting, tcsCheckingDown
  // 当 空间时间 > 30s
  //    或者 空间时间 > 10s and 平均等候任务时间大于平均工作时间的 50%
  // 则视为空闲。
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
