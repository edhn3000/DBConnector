unit U_MemoryCpuUtils;

interface

uses
  Windows, SysUtils;

type
  TCPUID = array[1..4] of Longint;
  TVendor = array[0..11] of Char;

  TMemCpuUtils = class
  public
    mem_PhysicsTotal  : DWORD;
    mem_VirtualTotal  : DWORD;
    mem_PageFileTotal : DWORD;

    mem_PhysicsAvail  : DWORD;
    mem_VirtualAvail  : DWORD;
    mem_PageFileAvail : DWORD;

    mem_Usage: Double;

    procedure ReadMemStatus;

    //获取CPU使用率
    function GetCPUUsage( Index: Integer ): Double;
  end;

  {CPU区}
  //刷新CPU数据
  procedure CollectCPUData;

  //获取CPU在系统中的总数
  function GetCPUCount: Integer;


  //
  procedure ReleaseCPUData;    

  function GetCPUID: TCPUID; assembler; register;
  function GetCPUVendor: TVendor; assembler; register;
  function GetCPUIDStr: string;

var
  memCpuUtils: TMemCpuUtils;
implementation

type
PInt64 = ^TInt64;
TInt64 = int64;

type
TPERF_DATA_BLOCK = record
    Signature:        array[0..4 - 1] of WCHAR;
    LittleEndian:     DWORD;
    Version:          DWORD;
    Revision:         DWORD;
    TotalByteLength: DWORD;
    HeaderLength:     DWORD;
    NumObjectTypes:   DWORD;
    DefaultObject:    longint;
    SystemTime:       TSystemTime;
    Reserved:         DWORD;
    PerfTime:         TInt64;
    PerfFreq:         TInt64;
    PerfTime100nSec: TInt64;
    SystemNameLength: DWORD;
    SystemNameOffset: DWORD;
end;
PPERF_DATA_BLOCK = ^TPERF_DATA_BLOCK;

TPERF_OBJECT_TYPE = record
    TotalByteLength: DWORD;
    DefinitionLength: DWORD;
    HeaderLength: DWORD;
    ObjectNameTitleIndex: DWORD;
    ObjectNameTitle: LPWSTR;
    ObjectHelpTitleIndex: DWORD;
    ObjectHelpTitle: LPWSTR;
    DetailLevel: DWORD;
    NumCounters: DWORD;
    DefaultCounter: longint;
    NumInstances: longint;
    CodePage: DWORD;
    PerfTime: TInt64;
    PerfFreq: TInt64;
end;
PPERF_OBJECT_TYPE = ^TPERF_OBJECT_TYPE;

type
TPERF_COUNTER_DEFINITION = record
    ByteLength:       DWORD;
    CounterNameTitleIndex: DWORD;
    CounterNameTitle: LPWSTR;
    CounterHelpTitleIndex: DWORD;
    CounterHelpTitle: LPWSTR;
    DefaultScale:     longint;
    DetailLevel:      DWORD;
    CounterType:      DWORD;
    CounterSize:      DWORD;
    CounterOffset:    DWORD;
end;
PPERF_COUNTER_DEFINITION = ^TPERF_COUNTER_DEFINITION;

TPERF_COUNTER_BLOCK = record
    ByteLength: DWORD;
end;
PPERF_COUNTER_BLOCK = ^TPERF_COUNTER_BLOCK;

TPERF_INSTANCE_DEFINITION = record
    ByteLength: DWORD;
    ParentObjectTitleIndex: DWORD;
    ParentObjectInstance: DWORD;
    UniqueID:   longint;
    NameOffset: DWORD;
    NameLength: DWORD;
end;
PPERF_INSTANCE_DEFINITION = ^TPERF_INSTANCE_DEFINITION;

type
TInt64F = TInt64;

type
FInt64 = TInt64F;
Int64D = TInt64;
//------------------------------------------------------------------------------
const
Processor_IDX_Str = '238';
Processor_IDX     = 238;
CPUUsageIDX       = 6;

type
AInt64F = array[0..$FFFF] of TInt64F;
PAInt64F = ^AInt64F;

var
_PerfData:      PPERF_DATA_BLOCK;
_BufferSize:    integer;
_POT:           PPERF_OBJECT_TYPE;
_PCD:           PPerf_Counter_Definition;
_ProcessorsCount: integer;
_Counters:      PAInt64F;
_PrevCounters: PAInt64F;
_SysTime:       TInt64F;
_PrevSysTime:   TInt64F;
_IsWinNT:       boolean;
_W9xCollecting: boolean;
_W9xCpuUsage:   DWORD;
_W9xCpuKey:     HKEY;
//------------------------------------------------------------------------------
function GetCPUCount: integer;
begin
if _IsWinNT then
begin
    if _ProcessorsCount < 0 then
      CollectCPUData;
    Result := _ProcessorsCount;
end else
begin
    Result := 1;
end;
end;
//------------------------------------------------------------------------------
procedure ReleaseCPUData;
var
H: HKEY;
R: DWORD;
dwDataSize, dwType: DWORD;
begin
if _IsWinNT then
    exit;
if not _W9xCollecting then
    exit;
_W9xCollecting := False;
RegCloseKey(_W9xCpuKey);
R := RegOpenKeyEx(HKEY_DYN_DATA, 'PerfStats\StopStat', 0, KEY_ALL_ACCESS, H);
if R <> ERROR_SUCCESS then
    exit;
dwDataSize := sizeof(DWORD);
RegQueryValueEx(H, 'KERNEL\CPUUsage', nil, @dwType, PBYTE(@_W9xCpuUsage),
    @dwDataSize);
RegCloseKey(H);
end;
//------------------------------------------------------------------------------
function TMemCpuUtils.GetCPUUsage(Index: integer): double;
begin
  if _IsWinNT then
  begin
      if _ProcessorsCount < 0 then
        CollectCPUData;
      if (Index >= _ProcessorsCount) or (Index < 0) then
        raise Exception.Create('CPU index out of bounds');
      if _PrevSysTime = _SysTime then
        Result := 0
      else
        Result := 1 - (_Counters[index] - _PrevCounters[index])
          / (_SysTime - _PrevSysTime);
  end else
  begin
      if Index <> 0 then
        raise Exception.Create('CPU index out of bounds');
      if not _W9xCollecting then
        CollectCPUData;
      Result := _W9xCpuUsage / 100;
  end;
end;

var
VI: TOSVERSIONINFO;
//------------------------------------------------------------------------------
procedure CollectCPUData;
var
BS: integer;
i: integer;
_PCB_Instance: PPERF_COUNTER_BLOCK;
_PID_Instance: PPERF_INSTANCE_DEFINITION;
ST: TFileTime;
var
H: HKEY;
R: DWORD;
dwDataSize, dwType: DWORD;
begin
if _IsWinNT then
begin
    BS := _BufferSize;
    while RegQueryValueEx(HKEY_PERFORMANCE_DATA, Processor_IDX_Str,
        nil, nil, PByte(_PerfData), @BS) = ERROR_MORE_DATA do
    begin
      // Get a buffer that is big enough.
      Inc(_BufferSize, $1000);
      BS := _BufferSize;
      ReallocMem(_PerfData, _BufferSize);
    end;
    // Locate the performance object
    _POT := PPERF_OBJECT_TYPE(DWORD(_PerfData) + _PerfData.HeaderLength);
    for i := 1 to _PerfData.NumObjectTypes do
    begin
      if _POT.ObjectNameTitleIndex = Processor_IDX then
        Break;
      _POT := PPERF_OBJECT_TYPE(DWORD(_POT) + _POT.TotalByteLength);
    end;
    // Check for success
    if _POT.ObjectNameTitleIndex <> Processor_IDX then
      raise Exception.Create(
        'Unable to locate the "Processor" performance object');
    if _ProcessorsCount < 0 then
    begin
      _ProcessorsCount := _POT.NumInstances;
      GetMem(_Counters, _ProcessorsCount * SizeOf(TInt64));
      GetMem(_PrevCounters, _ProcessorsCount * SizeOf(TInt64));
    end;
    // Locate the "% CPU usage" counter definition
    _PCD := PPERF_Counter_DEFINITION(DWORD(_POT) + _POT.HeaderLength);
    for i := 1 to _POT.NumCounters do
    begin
      if _PCD.CounterNameTitleIndex = CPUUsageIDX then
        break;
      _PCD := PPERF_COUNTER_DEFINITION(DWORD(_PCD) + _PCD.ByteLength);
    end;
    // Check for success
    if _PCD.CounterNameTitleIndex <> CPUUsageIDX then
      raise Exception.Create(
        'Unable to locate the "% of CPU usage" performance counter');
    // Collecting coutners
    _PID_Instance := PPERF_INSTANCE_DEFINITION(DWORD(_POT) + _POT.DefinitionLength);
    for i := 0 to _ProcessorsCount - 1 do
    begin
      _PCB_Instance    := PPERF_COUNTER_BLOCK(DWORD(_PID_Instance) +
        _PID_Instance.ByteLength);
      _PrevCounters[i] := _Counters[i];
      _Counters[i]     := FInt64(PInt64(DWORD(_PCB_Instance) + _PCD.CounterOffset)^);
      _PID_Instance    := PPERF_INSTANCE_DEFINITION(DWORD(_PCB_Instance) +
        _PCB_Instance.ByteLength);
    end;
    _PrevSysTime := _SysTime;
    SystemTimeToFileTime(_PerfData.SystemTime, ST);
    _SysTime := FInt64(TInt64(ST));
end else
begin
    if not _W9xCollecting then
    begin
      R := RegOpenKeyEx(HKEY_DYN_DATA, 'PerfStats\StartStat',
        0, KEY_ALL_ACCESS, H);
      if R <> ERROR_SUCCESS then
        raise Exception.Create('Unable to start performance monitoring');
      dwDataSize := sizeof(DWORD);
      RegQueryValueEx(H, 'KERNEL\CPUUsage', nil, @dwType,
        PBYTE(@_W9xCpuUsage), @dwDataSize);
      RegCloseKey(H);
      R := RegOpenKeyEx(HKEY_DYN_DATA, 'PerfStats\StatData',
        0, KEY_READ, _W9xCpuKey);
      if R <> ERROR_SUCCESS then
        raise Exception.Create('Unable to read performance data');
      _W9xCollecting := True;
    end;
    dwDataSize := sizeof(DWORD);
    RegQueryValueEx(_W9xCpuKey, 'KERNEL\CPUUsage', nil, @dwType,
      PBYTE(@_W9xCpuUsage), @dwDataSize);
end;
end;


Function GetCPUID: TCPUID; assembler; register; 
asm
  PUSH    EBX         {Save affected register}
  PUSH    EDI
  MOV     EDI,EAX     {@Resukt}
  MOV     EAX,1
  DW      $A20F       {CPUID Command}
  STOSD                {CPUID[1]}
  MOV     EAX,EBX
  STOSD               {CPUID[2]} 
  MOV     EAX,ECX 
  STOSD               {CPUID[3]}
  MOV     EAX,EDX 
  STOSD               {CPUID[4]} 
  POP     EDI          {Restore registers} 
  POP     EBX
end;

Function GetCPUVendor: TVendor; assembler; register;
asm 
  PUSH    EBX          {Save affected register} 
  PUSH    EDI
  MOV     EDI,EAX      {@Result (TVendor)}
  MOV     EAX,0 
  DW      $A20F        {CPUID Command} 
  MOV     EAX,EBX 
  XCHG    EBX,ECX     {save ECX result}
  MOV      ECX,4
@1: 
  STOSB
  SHR     EAX,8
  LOOP    @1
  MOV     EAX,EDX 
  MOV      ECX,4
@2: 
  STOSB
  SHR     EAX,8
  LOOP    @2 
  MOV     EAX,EBX 
  MOV      ECX,4
@3: 
  STOSB
  SHR     EAX,8 
  LOOP    @3 
  POP     EDI          {Restore registers} 
  POP     EBX
end;

Function GetCPUIDStr: string; 
var 
  CPUID: TCPUID; 
  I: Integer;
begin
  for I := Low(CPUID) to High(CPUID) do
       CPUID[i]:=-1;
  CPUID := GetCPUID;
  Result :=  IntToHex(CPUID[1], 8) + IntToHex(CPUID[2], 8)
              + IntToHex(CPUID[3], 8)
              + IntToHex(CPUID[4], 8);
end;


{ TMemCpuUtils }

procedure TMemCpuUtils.ReadMemStatus;
var
  msMemory: TMemoryStatus;
begin
  msMemory.dwLength := SizeOf( msMemory );
  GlobalMemoryStatus( msMemory );

  mem_PhysicsTotal  := msMemory.dwTotalPhys;
  mem_VirtualTotal  := msMemory.dwTotalVirtual;
  mem_PageFileTotal := msMemory.dwTotalPageFile;

  mem_PhysicsAvail  := msMemory.dwAvailPhys;
  mem_VirtualAvail  := msMemory.dwAvailVirtual;
  mem_PageFileAvail := msMemory.dwAvailPageFile;

  mem_Usage := msMemory.dwMemoryLoad;
end;

initialization
  _ProcessorsCount := -1;
  _BufferSize := $2000;
  _PerfData := AllocMem( _BufferSize );
  VI.dwOSVersionInfoSize := SizeOf( VI );
  if not GetVersionEx( VI ) then
    raise Exception.Create( 'Can''t get the Windows version' );
  _IsWinNT := VI.dwPlatformId = VER_PLATFORM_WIN32_NT;
  
  memCpuUtils := TMemCpuUtils.Create;

finalization
  ReleaseCPUData;
  FreeMem( _PerfData );
  
  memCpuUtils.Free;

end.
