{**
 * <p>Title: Process  </p>
 * <p>Copyright: Copyright (c) 2007/3/5</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 进程处理单元 </P>
 * <p>Description: </P>
 *}
// todo: 进程监视：内存使用、CPU占用、文件使用等；进程屏蔽：屏蔽列表、
unit U_Process;

interface

uses
  ComCtrls,Controls, Windows, Classes, ShellAPI, Contnrs, Registry;

type
  TProcessInfo = class;
  TProcessInfoList = class;
  TWinNTInfo = class;

  { TWNDInfo }   
  PWNDInfo = ^TWNDInfo;
  TWNDInfo = record
   dwProcessId:DWORD;
   hWnd:HWND;
  end;

{ TProcessInfo 单个进程信息 }
  TProcessInfo = class(TCollectionItem)
  private
    FImageName: string;
    FImagePath: string;
    FProcessID: Integer;
    FHandle: THandle;
    FIcon: HICON;
    FCPUUse: Integer;
    FMemUse: Integer;

    FCount: Integer;

    FDestroyed: Boolean;       // 是否已注销，如已注销应从列表移除
    FCanDestroy: Boolean;      // 是否允许用户删除

    function GetHandleHex: string;    
  public         
    constructor Create(Collection: TCollection);override;
    destructor Destroy;override;
    procedure Init(AName: string; AID: Integer; AHandle: THandle; AIcon: HICON);
    procedure Assigned(processInfo: TProcessInfo);
    function Clone: TProcessInfo;

  published    // 可被 GetPropValue
    property ImageName: string read FImageName write FImageName;          // 映象名称
    property ImagePath: string read FImagePath write FImagePath;          // 映象路径
    property ProcessID: Integer read FProcessID write FProcessID;
    property Handle: THandle read FHandle write FHandle;
    property HandleHex: string read GetHandleHex;
    property Icon: HICON read FIcon write FIcon;
    property CPUUse: Integer read FCPUUse;
    property MemUse: Integer read FMemUse;

    property Count: Integer read FCount write FCount;
    property Destroyed: Boolean read FDestroyed write FDestroyed;
    property CanDestroy: Boolean read FCanDestroy write FCanDestroy;
  end;

{ TProcessInfoList 进程列表 }
  TProcessInfoList = class(TCollection)
  private
    FDelimiter: string;

    function GetText: string;
  protected
    function GetItem(index: Integer):TProcessInfo;
    function GetProcessCount(sProcessName: string): Integer;
  public
    property Items[Index: Integer]: TProcessInfo read GetItem;
    property Delimiter: string read FDelimiter write FDelimiter;
    property Text: string read GetText;
  public
    constructor Create;

    function Add: TProcessInfo;overload;
    function Add(p: TProcessInfo): Integer;overload;

    function FindProcess(ProcessID: Integer): TProcessInfo;overload;
    function FindProcess(ImagePath: string): TProcessInfo;overload;
    function IndexOfProcess(ProcessID: Integer): Integer;overload;  
    function IndexOfProcess(ImagePath: string): Integer;overload;

    function CountProcess(ImagePath: string):Integer;

    function KillProcess(ProcessID: Integer): Boolean;overload;
    function KillProcess(ImagePath: string): Boolean;overload;

    function KillProcessList(proclist: TProcessInfoList): Boolean;

    procedure FillRepeatProcessList(nRepeatTimes: Integer; RepeatPrecessList:TProcessInfoList);
    procedure SetAllProcessDestroyed;             // 重新设置进程是否已注销的标志，认为所有进程都已注销
    procedure RemoveDestroyedProcess;             // 移除已注销的进程
  end;
            
{ TWinInfo NT和95的父类 }
  TWinInfo = class
  public                   
    procedure Refresh;virtual;abstract;
    procedure FillProcessList(ProcessList: TProcessInfoList);virtual;abstract;
  end;

{ TWinNTInfoFunc 用于winnt }
  TWinNTInfo = class(TWinInfo)
  private
    SysImgList: TImageList;

  public
    ProcAry: array of DWORD;
    DrvAry: array of DWORD;
  public
    constructor Create;
    destructor Destroy;override;

    procedure Refresh;override;
    procedure FillProcessList(ProcessList: TProcessInfoList);override;
    procedure FillProcessLvw(var lvw: TListView; var iml: TImageList);
    procedure FillDrives(var lvw: TListView; var iml: TImageList);
  end;
  
{ TWin95Info 用于win95,98 }
  TWin95Info = class(TWinInfo)
  private
    FhndSnap: THandle;
    FProcList: TList;
  public
    constructor Create;
    destructor Destroy;override;

    procedure Refresh;override;
    procedure FillProcessList(ProcessList: TProcessInfoList);override;
    procedure FillProcessLvw(ALvw: TListView; AIml: TImageList);
  end;

  function WinInfo: TWinInfo;
  function WinNTInfo: TWinNTInfo; 
  function Win95Info: TWin95Info;

  // 获得Shell32.dll中未知程序的图标
  function GetWinIcon: HICON;
  // 根据进程句柄结束进程
  function KillProcessByHandle(ProcessHandle: THandle): Boolean;  
  function KillProcessByPID(ProcessID: Integer): Boolean;

  function getProcessHNDByID(dwProcessId:DWORD):HWND;

implementation

uses
  Forms, SysUtils, IntfInfo, PsAPI, TlHelp32, CommCtrl, Variants;

type
  { windoes版本枚举 }
  TWinVer = (wvUnKnown,
    wvWin95, wvWin98, wvWinMe,
    wvWin2k, wvWin2kServ,
    wvWinXp,
    wvWin2003,
    wvWinNT3, wvWinNT3Serv,
    wvWinNT4, wvWinNT4Serv,
    wvWinCE);

var                        
  FWinInfo: TWinInfo;
  FWinNTInfo: TWinNTInfo;
  FWin95Info: TWin95Info;
  m_WinIcon: HICON;

function GetWinVersionInfo(var SProduct, SVersion, SServicePack: string):
  TWinVer;overload;
type
  _OSVERSIONINFOEXA = record
    dwOSVersionInfoSize: DWORD;
    dwMajorVersion: DWORD;
    dwMinorVersion: DWORD;
    dwBuildNumber: DWORD;
    dwPlatformId: DWORD;
    szCSDVersion: array[0..127] of AnsiChar;
    wServicePackMajor: WORD;
    wServicePackMinor: WORD;
    wSuiteMask: Word;
    wProductType: Byte;
    wReserved: Byte;
  end;
  _OSVERSIONINFOEXW = record
    dwOSVersionInfoSize: DWORD;
    dwMajorVersion: DWORD;
    dwMinorVersion: DWORD;
    dwBuildNumber: DWORD;
    dwPlatformId: DWORD;
    szCSDVersion: array[0..127] of WideChar;
    wServicePackMajor: WORD;
    wServicePackMinor: WORD;
    wSuiteMask: Word;
    wProductType: Byte;
    wReserved: Byte;
  end;
  { this record only support Windows 4.0 SP6 and latter , Windows 2000 ,XP, 2003 }
  OSVERSIONINFOEXA = _OSVERSIONINFOEXA;
  OSVERSIONINFOEXW = _OSVERSIONINFOEXW;
  OSVERSIONINFOEX = OSVERSIONINFOEXA;
const
  VER_PLATFORM_WIN32_CE = 3;
  { wProductType defines }
  VER_NT_WORKSTATION = 1;
  VER_NT_DOMAIN_CONTROLLER = 2;
  VER_NT_SERVER = 3;
  { wSuiteMask defines }
  VER_SUITE_SMALLBUSINESS = $0001;
  VER_SUITE_ENTERPRISE = $0002;
  VER_SUITE_BACKOFFICE = $0004;
  VER_SUITE_TERMINAL = $0010;
  VER_SUITE_SMALLBUSINESS_RESTRICTED = $0020;
  VER_SUITE_DATACENTER = $0080;
  VER_SUITE_PERSONAL = $0200;
  VER_SUITE_BLADE = $0400;
  VER_SUITE_SECURITY_APPLIANCE = $1000;
var
  Info: OSVERSIONINFOEX;
  bEx: BOOL;
begin
  Result := wvUnKnown;
  FillChar(Info, SizeOf(OSVERSIONINFOEX), 0);
  Info.dwOSVersionInfoSize := SizeOf(OSVERSIONINFOEX);
  bEx := GetVersionEx(POSVERSIONINFO(@Info)^);
  if not bEx then
  begin
    Info.dwOSVersionInfoSize := SizeOf(OSVERSIONINFO);
    if not GetVersionEx(POSVERSIONINFO(@Info)^) then Exit;
  end;
  with Info do
  begin
    SVersion := IntToStr(dwMajorVersion) + '.' + IntToStr(dwMinorVersion)
      + '.' + IntToStr(dwBuildNumber and $0000FFFF);
    SProduct := 'Microsoft Windows unknown';
    case Info.dwPlatformId of
      VER_PLATFORM_WIN32s: { Windows 3.1 and earliest }
        SProduct := 'Microsoft Win32s';
      VER_PLATFORM_WIN32_WINDOWS:
        case dwMajorVersion of
          4: { Windows95,98,ME }
            case dwMinorVersion of
              0:
              begin
                Result := wvWin95;
                if szCSDVersion[1] in ['B', 'C'] then
                begin
                  SProduct := 'Microsoft Windows 95 OSR2';
                  SVersion := SVersion + szCSDVersion[1];
                end
                else
                  SProduct := 'Microsoft Windows 95';
              end;
              10:
              begin
                Result := wvWin98;
                if szCSDVersion[1] = 'A' then
                begin
                  SProduct := 'Microsoft Windows 98 SE';
                  SVersion := SVersion + szCSDVersion[1];
                end
                else
                  SProduct := 'Microsoft Windows  98';
              end;
              90:
              begin
                Result := wvWinMe;
                SProduct := 'Microsoft Windows Millennium Edition';
              end;
            end;
        end;
      VER_PLATFORM_WIN32_NT:
        begin
          SServicePack := szCSDVersion;
          case dwMajorVersion of
            0..4:
              if bEx then
              begin
                case wProductType of
                  VER_NT_WORKSTATION:
                    begin
                      Result := wvWinNT4;
                      SProduct := 'Microsoft Windows NT Workstation 4.0';
                    end;
                  VER_NT_SERVER:
                    begin
                      Result := wvWinNT4Serv;
                      if wSuiteMask and VER_SUITE_ENTERPRISE <> 0 then
                        SProduct := 'Microsoft Windows NT Advanced Server 4.0'
                      else
                        SProduct := 'Microsoft Windows NT Server 4.0';
                    end;
                end;
                end
              else { NT351 and NT4.0 SP5 earliest}
                with TRegistry.Create do
                try
                  RootKey := HKEY_LOCAL_MACHINE;
                  if OpenKey('SYSTEMCurrentControlSetControlProductOptions',
                    False) then
                  begin
                    if ReadString('ProductType') = 'WINNT' then
                    begin
                      Result := wvWinNT3;
                      SProduct := 'Microsoft Windows NT Workstation ' + SVersion
                    end
                    else if ReadString('ProductType') = 'LANMANNT' then
                    begin
                      Result := wvWinNT3Serv;
                      SProduct := 'Microsoft Windows NT Server ' + SVersion
                    end
                    else if ReadString('ProductType') = 'LANMANNT' then
                    begin
                      Result := wvWinNT3Serv;
                      SProduct := 'Microsoft Windows NT Advanced Server ' +
                        SVersion;
                    end
                  end;
                finally
                  Free;
                end;
            5:
              case dwMinorVersion of
                0: { Windows 2000 }
                  case wProductType of
                    VER_NT_WORKSTATION:
                      begin
                        Result := wvWin2k;
                        SProduct := 'Microsoft Windows 2000 Professional';
                      end;
                    VER_NT_SERVER:
                      begin
                        Result := wvWin2kServ;
                        if wSuiteMask and VER_SUITE_DATACENTER <> 0 then
                          SProduct := 'Microsoft Windows 2000 Datacenter Server'
                        else if wSuiteMask and VER_SUITE_ENTERPRISE <> 0 then
                          SProduct := 'Microsoft Windows 2000 Advanced Server'
                        else
                          SProduct := 'Microsoft Windows 2000 Server';
                      end;
                  end;
                1: { Windows XP }
                  begin
                    Result := wvWinXp;
                    if wSuiteMask and VER_SUITE_PERSONAL <> 0 then
                      SProduct := 'Microsoft Windows XP Home Edition'
                    else
                      SProduct := 'Microsoft Windows XP Professional';
                  end;
                2: { Windows Server 2003 }
                  begin
                    Result := wvWin2003;
                    if wSuiteMask and VER_SUITE_DATACENTER <> 0 then
                      SProduct := 'Microsoft Windows Server 2003 Datacenter Edition'
                    else if wSuiteMask and VER_SUITE_ENTERPRISE <> 0 then
                      SProduct := 'Microsoft Windows Server 2003 Enterprise Edition'
                    else if wSuiteMask and VER_SUITE_BLADE <> 0 then
                      SProduct := 'Microsoft Windows Server 2003 Web Edition'
                    else
                      SProduct :=
                        'Microsoft Windows Server 2003 Standard Edition';
                  end;
              end;
          end;
        end;
      VER_PLATFORM_WIN32_CE: { Windows CE }
        begin
          Result := wvWinCE;
          SProduct := SProduct + ' CE';
        end;
    end;
  end;
end;

function GetWinVersionInfo:TWinVer;overload;
var
  a, b, c: string;
begin
  Result := GetWinVersionInfo(a,b,c);
end;

function OSIsNT: Boolean;
begin
  Result := GetWinVersionInfo in [wvWin2k, wvWin2kServ, wvWinXp, wvWin2003,
    wvWinNT3, wvWinNT3Serv, wvWinNT4, wvWinNT4Serv];
end;  

function WinNTInfo: TWinNTInfo;
begin
  if not Assigned(FWinNTInfo) then
    FWinNTInfo := TWinNTInfo.Create;
  Result := FWinNTInfo;
end;

function Win95Info: TWin95Info;
begin
  if not Assigned(FWin95Info) then
    FWin95Info := TWin95Info.Create;
  Result := FWin95Info;
end;

function WinInfo: TWinInfo;
begin
  if not Assigned(FWinInfo) then
  begin
    if OSIsNT then
      FWinInfo := TWinNTInfo.Create
    else
      FWinInfo := TWin95Info.Create;
  end;
  Result := FWinInfo;
end;

function KillProcessByHandle(ProcessHandle: THandle): Boolean;
var
  exitcode: Cardinal;
begin
  GetExitCodeProcess(ProcessHandle, exitcode);
  Result := TerminateProcess(ProcessHandle, exitcode);
end; 

function KillProcessByPID(ProcessID: Integer): Boolean;
var
  hnd: THandle;
begin
  Result := False;
  hnd := OpenProcess(PROCESS_QUERY_INFORMATION or PROCESS_TERMINATE, False, ProcessID);
  if hnd <> Null then
  try
    Result := KillProcessByHandle(hnd);
//      GetExitCodeProcess(hnd, exitcode);
//      Result := TerminateProcess(hnd, exitcode);
  finally
    CloseHandle(hnd);
  end;
end;  

function TheEnumProc( hWnd: HWND; lp: LPARAM ): Boolean;
var
  dwProcessId: DWORD;
//  pInfo:PWNDInfo;
  wndInfo: TWNDInfo;
begin
  GetWindowThreadProcessId(hWnd, @dwProcessId);
  wndInfo := TWNDInfo(Pointer(lp)^);
  if(dwProcessId = wndInfo.dwProcessId) then
  begin
    wndInfo.hWnd := hWnd;
    Result := False;
    Exit;
  end;
  Result := True;
end;

function getProcessHNDByID(dwProcessId:DWORD):HWND;
begin
  Result := OpenProcess(PROCESS_QUERY_INFORMATION or PROCESS_TERMINATE, False,
    dwProcessId);
  CloseHandle(Result);
end;

{ TWinNTInfo }

constructor TWinNTInfo.Create;
//var
//  Sfi: TSHFileInfo;
begin
//  SysImgList := TImageList.Create(nil);
//  SysImgList.Handle := SHGetFileInfo('c:\windows\system32\shell32.dll', 0,
//                       Sfi, SizeOf(TSHFileInfo),
//                       SHGFI_SYSICONINDEX or SHGFI_SMALLICON);
//  SysImgList.ShareImages := True;
//  FWinIcon := ImageList_GetIcon(SysImgList.Handle, 2, ILD_NORMAL	);
//  FWinIcon := ExtractIcon(HInstance, 'SHELL32.DLL', 2);
end;

destructor TWinNTInfo.Destroy;
begin
  if Assigned(SysImgList) then
    FreeAndNil(SysImgList);
  inherited;
end;

procedure TWinNTInfo.FillDrives(var lvw: TListView; var iml: TImageList);
var
  i: Integer;
  acDrvName: array[0..300] of Char;
begin
  for i := Low(ProcAry) to High(ProcAry) do
    if GetDeviceDriverFileName(DrvAry, acDrvName, SizeOf(acDrvName)) <> 0 then
      with lvw.Items.Add do
      begin
        Caption := acDrvName;
        SubItems.Add('$' + IntToHex(Integer(DrvAry[i]), 8));
      end;
end;

procedure TWinNTInfo.FillProcessLvw(var lvw: TListView; var iml: TImageList);
var
  i: Integer;
  dwCount: DWORD;
  hndProcHand: THandle;
  ModHand: HMODULE;
  HAppIcon: HICON;
  acModName: array[0..300] of Char;
  item: TListItem;

  function FindItem(ID: Integer): TListItem;
  var
    nIndex: Integer;
  begin
    Result := nil;
    for nIndex := 0 to lvw.Items.Count - 1 do
    begin
      // SubItems[1] 是ProcessID
      if TProcessInfo(lvw.Items[nIndex].Data).ProcessID = ID then
      begin
        Result := lvw.Items[nIndex];
        Break;
      end;  
    end;  
  end;

//// FillProcesses Begin
begin  
  Refresh;

  lvw.Items.BeginUpdate;
  for i := Low(ProcAry) to High(ProcAry) do
  begin
    // 取得进程句柄
    hndProcHand := OpenProcess(PROCESS_QUERY_INFORMATION or PROCESS_VM_READ, False, ProcAry[i]);
    if hndProcHand <> NULL then
      try
        // 列举进程的模块 第一模块是进程可执行文件名
        EnumProcessModules(hndProcHand, @ModHand, SizeOf(ModHand), dwCount);
        if GetModuleFileNameEx(hndProcHand, ModHand, acModName, SizeOf(acModName)) <> 0 then  
//        if GetModuleFileName(ModHand, acModName, SizeOf(acModName)) <> 0 then
        begin
          HAppIcon := ExtractIcon(HInstance, acModName, 0);
          try
            if HAppIcon = 0 then
              HAppIcon := GetWinIcon;

            item := FindItem(ProcAry[i]);
            if nil = item then
              item := lvw.Items.Add;
            with item do
            begin
              Data := Pointer(ProcAry[i]);                   // 把进程ID保存到Data

              if Caption <> ExtractFileName(acModName) then
                Caption := ExtractFileName(acModName);         // 文件名
                
              while SubItems.Count < 3 do
                SubItems.Add('');
              if SubItems[0] <> acModName then
                SubItems[0] := acModName;
              if SubItems[1] <> IntToStr(ProcAry[i]) then
                SubItems[1] := IntToStr(ProcAry[i]);            // 进程ID
              if SubItems[2] <> '$' + IntToHex(hndProcHand, 8) then
                SubItems[2] := '$' + IntToHex(hndProcHand, 8);  // 进程句柄

//              SubItems.Add(acModName);
//              SubItems.Add(IntToStr(ProcAry[i]));            // 进程ID
//              SubItems.Add('$' + IntToHex(hndProcHand, 8));  // 进程句柄
//              SubItems.Add(GetPriorityClassString(GetPriorityClass(hndProcHand)));

              ImageIndex := ImageList_AddIcon(iml.Handle, HAppIcon); 
            end;
          finally
            if HAppIcon <> GetWinIcon then
              DestroyIcon(HAppIcon);
          end;
        end;
      finally
        CloseHandle(hndProcHand);
      end;
  end;
  lvw.Items.EndUpdate;
end;

procedure TWinNTInfo.FillProcessList(ProcessList: TProcessInfoList);
  var
  i: Integer;
  dwCount: DWORD;
//  lwHnd: LongWord;
  hndProcHand: Cardinal;
  ModHand: HMODULE;
  HAppIcon: HICON;
  acModName: array[0..300] of Char;
  process: TProcessInfo;
  hndWindow: HWND;
////  FillProcesses Begin
begin
  Refresh;
  ProcessList.SetAllProcessDestroyed;

  for i := Low(ProcAry) to High(ProcAry) do
  begin
    // 取得进程句柄
//    lwHnd := LongWord(OpenProcess(PROCESS_QUERY_INFORMATION or PROCESS_VM_READ, False, ProcAry[i]));
    hndProcHand := OpenProcess(PROCESS_QUERY_INFORMATION or PROCESS_VM_READ, False, ProcAry[i]);
//    if hndProcHand <> 0 then
      try
        // 列举进程的模块 第一模块是进程可执行文件名
        EnumProcessModules(hndProcHand, @ModHand, SizeOf(ModHand), dwCount);
        if GetModuleFileNameEx(hndProcHand, ModHand, acModName, SizeOf(acModName)) <> 0 then
//        if GetModuleFileName(ModHand, acModName, SizeOf(acModName)) <> 0 then
        begin
          HAppIcon := ExtractIcon(HInstance, acModName, 0);
          try
            if HAppIcon = 0 then
              HAppIcon := GetWinIcon;

            process := ProcessList.FindProcess(ProcAry[i]);
            if nil = process then
              process := ProcessList.Add;
            hndWindow := getProcessHNDByID(ProcAry[i]);
            if hndWindow = Null then
              hndWindow := 0;
            process.Init(acModName, ProcAry[i], hndWindow, HAppIcon);
          finally
//            if HAppIcon <> FWinIcon then  DestroyIcon(HAppIcon);
          end;
        end;
      finally
        CloseHandle(hndProcHand);
      end;
  end;
  ProcessList.RemoveDestroyedProcess;
end;

function GetWinIcon: HICON;
begin
  if 0 = m_WinIcon then
    m_WinIcon := ExtractIcon(HInstance, 'SHELL32.DLL', 2);
  Result := m_WinIcon;
end;

procedure TWinNTInfo.Refresh;
var
  dwCount: DWORD;
  AryProcesses: array[0..$3FFF-1] of DWORD;
begin
  // 列举进程
  if not EnumProcesses(@AryProcesses, SizeOf(AryProcesses), dwCount) then
    Exit // raise exception
  else
  begin
    SetLength(ProcAry, dwCount div SizeOf(DWORD));
    Move(AryProcesses, ProcAry[0], dwCount);
  end;

  if not EnumDeviceDrivers(@AryProcesses, SizeOf(AryProcesses), dwCount) then
    Exit
  else
  begin
    SetLength(DrvAry, dwCount div SizeOf(DWORD) );
    Move(AryProcesses, DrvAry[0], dwCount);
  end;
end;

{ TProcessInfo }

procedure TProcessInfo.Assigned(processInfo: TProcessInfo);
begin
  Self.ImagePath := processInfo.ImagePath;
  Self.ImageName := processInfo.ImageName;
  Self.ProcessID := processInfo.ProcessID;
  Self.Handle    := processInfo.Handle;
  Self.Icon      := processInfo.Icon;

  Self.Destroyed := processInfo.Destroyed;
  Self.CanDestroy := processInfo.CanDestroy;
end;

function TProcessInfo.Clone: TProcessInfo;
begin
  Result := TProcessInfo.Create(nil);
  Result.FImageName := Self.FImageName ;
  Result.FImagePath := Self.FImagePath ;
  Result.FProcessID := Self.FProcessID ;
  Result.FHandle    := Self.FHandle    ;
  Result.FIcon      := Self.FIcon      ;
  Result.FCPUUse    := Self.FCPUUse    ;
  Result.FMemUse    := Self.FMemUse    ;
  Result.FCount     := Self.FCount     ;
  Result.FDestroyed := Self.FDestroyed ;
  Result.FCanDestroy:= Self.FCanDestroy;
end;

constructor TProcessInfo.Create(Collection: TCollection);
begin
  if Collection <> nil then
    inherited Create(Collection);
  FIcon  := 0;
  FCount := 1;
  FDestroyed := False;
  FCanDestroy := True;
end;

destructor TProcessInfo.Destroy;
begin
  if (0 <> Icon) and (GetWinIcon <> Icon) then
    DestroyIcon(Icon);
  inherited;
end;

function TProcessInfo.GetHandleHex: string;
begin
  Result := '$' + IntToHex(Handle, 8);
end;

procedure TProcessInfo.Init(AName: string; AID: Integer; AHandle: THandle;
  AIcon: HICON);
begin
  ImagePath := AName;
  ImageName := ExtractFileName(AName);
  ProcessID := AID;
  Handle   := AHandle;
  Icon     := AIcon;

  Destroyed := False;
end;

{ TProcessInfoList }    

procedure TProcessInfoList.FillRepeatProcessList(nRepeatTimes: Integer;
   RepeatPrecessList:TProcessInfoList);
var
  i:Integer;
  p: TProcessInfo;
begin
  // 添加大于指定数目的进程
  for i := 0 to Count - 1 do
  begin
    if GetProcessCount(Items[i].ImagePath) > nRepeatTimes then
      if nil = RepeatPrecessList.FindProcess(Items[i].ImagePath) then
      begin
        p := TProcessInfo.Create(RepeatPrecessList);
        p.Assigned(Items[i]);
        if p.CanDestroy then
          RepeatPrecessList.Add(p);
      end;
  end;
  // 去除已小于指定数目的进程
  for i := RepeatPrecessList.Count -1 downto 0 do
  begin
    if GetProcessCount(RepeatPrecessList.Items[i].ImagePath) < nRepeatTimes then
      RepeatPrecessList.Delete(i);
  end;
end;

function TProcessInfoList.GetItem(index: Integer): TProcessInfo;
begin
  Result := inherited GetItem(index) as TProcessInfo;
end;

function TProcessInfoList.GetProcessCount(sProcessName: string): Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
  begin
    if SameText(sProcessName, Items[i].ImagePath) then
      Inc(Result);
  end;  
end;

function TProcessInfoList.FindProcess(ProcessID: Integer): TProcessInfo;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Count - 1 do
  begin
    if Items[i].ProcessID = ProcessID then
    begin
      Result := Items[i];
      Break;
    end;
  end;
end;

function TProcessInfoList.FindProcess(ImagePath: string): TProcessInfo;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Count - 1 do
  begin
    if SameText(Items[i].ImagePath, ImagePath) then
    begin
      Result := Items[i];
      Break;
    end;
  end;
end;

procedure TProcessInfoList.SetAllProcessDestroyed;
var
  i: Integer;
begin
  for i := 0 to Count - 1 do
  begin
    Items[i].Destroyed := True;
  end;
end;

procedure TProcessInfoList.RemoveDestroyedProcess;
var
  i: Integer;
begin
  for i := Count - 1 downto 0 do
  begin
    if Items[i].Destroyed then
      Delete(i);
  end;
end;

constructor TProcessInfoList.Create;
begin
  inherited Create(TProcessInfo);
  FDelimiter := ';';
end;

function TProcessInfoList.GetText: string;
var
  i: Integer;
begin
  for i := 0 to Count - 1 do
  begin
    if '' = Result then
      Result := Items[i].ImageName
    else
      Result := Result + FDelimiter + Items[i].ImageName;
  end;
end;

function TProcessInfoList.KillProcess(ProcessID: Integer): Boolean;
var
  hnd: THandle;
  nIndex: Integer;
begin
  Result := False;
  nIndex := IndexOfProcess(ProcessID);
  hnd := OpenProcess(PROCESS_QUERY_INFORMATION or PROCESS_TERMINATE, False, ProcessID);
  if hnd <> Null then
  try
    Result := KillProcessByHandle(hnd);
//      GetExitCodeProcess(hnd, exitcode);
//      Result := TerminateProcess(hnd, exitcode);
    if Result and (nIndex <> -1) then
    begin
      Self.Delete(nIndex);
    end;
  finally
    CloseHandle(hnd);
  end;
end;

function TProcessInfoList.KillProcess(ImagePath: string): Boolean;
var
  p: TProcessInfo;
  hnd: THandle;
begin
  Result := False;
  while True do
  begin
    p := FindProcess(ImagePath);
    if nil = p then Break;

    hnd := OpenProcess(PROCESS_TERMINATE, False, p.ProcessID);
    if hnd <> Null then
      try
        Result := KillProcessByHandle(hnd);
        if Result then
          Delete(p.Index);
      finally
        CloseHandle(hnd);
      end;
  end;
end;

function TProcessInfoList.IndexOfProcess(ProcessID: Integer): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to Count - 1 do
  begin
    if Items[i].ProcessID = ProcessID then
    begin
      Result := i;
      Break;
    end;
  end;
end;

function TProcessInfoList.IndexOfProcess(ImagePath: string): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to Count - 1 do
  begin
    if SameText(Items[i].ImagePath, ImagePath) then
    begin
      Result := i;
      Break;
    end;
  end;
end;  

function TProcessInfoList.CountProcess(ImagePath: string): Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
  begin
    if SameText(Items[i].ImagePath, ImagePath) then
    begin
      Result := Result + 1;
    end;
  end;
end;

function TProcessInfoList.Add: TProcessInfo;
begin
  Result := inherited Add as TProcessInfo;
end;

function TProcessInfoList.Add(p: TProcessInfo): Integer;
begin
  with Add do
  begin
    Assigned(p);
    Result := Index;
  end;
end;

function TProcessInfoList.KillProcessList(
  proclist: TProcessInfoList): Boolean;
var
  i: Integer;
begin
  Result := True;
  for i := proclist.Count - 1 downto 0 do
  begin
    if not KillProcess(proclist.Items[i].ProcessID) then
      Result := False;
  end;
end;

{ TWin95Info }

constructor TWin95Info.Create;
begin
  FProcList := TList.Create;
end;

destructor TWin95Info.Destroy;
begin
  if Assigned(FProcList) then
    FreeAndNil(FProcList);
  inherited;
end;

procedure TWin95Info.FillProcessLvw(ALvw: TListView;
  AIml: TImageList);
var
  i: Integer;
  sExeFile: string;
  PE: PProcessEntry32;
  HAppIcon: HICON;
begin
  Refresh;
//  ALvw.Columns.Clear;
  ALvw.Items.BeginUpdate;
//  for i := Low()
  for i:= 0 to FProcList.Count - 1 do
  begin
    PE := PProcessEntry32(FProcList.Items[i]);
    HAppIcon := ExtractIcon(HInstance, PE.szExeFile, 0);
    try
      if HAppIcon = 0 then HAppIcon := GetWinIcon;
      sExeFile := PE.szExeFile;
//      if ALvw.ViewStyle = vslist then
//        sExeFile := ExtractFileName(sExeFile);
      with ALvw.Items.Add do
      begin
        Caption := ExtractFileName(sExeFile);
        SubItems.Add(sExeFile);
//        SubItems.Add(IntToStr(PE.cntThreads));              //启用的线程数
        SubItems.Add(IntToHex(PE.th32ProcessID, 8));        // 进程ID
//        SubItems.Add(IntToHex(PE.th32ParentProcessID, 8));  // 父进程ID
//        SubItems.Add('$' + IntToHex(hndProcHand, 8));       // 进程句柄

        Data := FProcList.Items[i];
        if AIml <> nil then
          ImageIndex := ImageList_AddIcon(AIml.Handle, HAppIcon);
      end;
    finally
      if HAppIcon <> GetWinIcon then DestroyIcon(HAppIcon);
    end;
  end;
  ALvw.Items.EndUpdate;
end;

procedure TWin95Info.FillProcessList(ProcessList: TProcessInfoList);
var
  i: Integer;
  sExeFile: string;
  PE: PProcessEntry32;
  HAppIcon: HICON;
  process: TProcessInfo;  
begin
  Refresh;
  for i:= 0 to FProcList.Count - 1 do
  begin
    PE := PProcessEntry32(FProcList.Items[i]);
    HAppIcon := ExtractIcon(HInstance, PE.szExeFile, 0);
    try
      if HAppIcon = 0 then HAppIcon := GetWinIcon;
        sExeFile := PE.szExeFile; 

        process := ProcessList.FindProcess(PE.th32ProcessID);
        if nil = process then
          process := ProcessList.Add;
        process.Init(sExeFile, PE.th32ProcessID, PE.th32ParentProcessID, HAppIcon);
    finally
//      if HAppIcon <> GetWinIcon then DestroyIcon(HAppIcon);
    end;
  end;
end;

procedure TWin95Info.Refresh;
var
  PE: TProcessEntry32;
  pPE: PPROCESSENTRY32;
begin
  FProcList.Clear;
  if FhndSnap >0 then
    CloseHandle(FhndSnap);
  FhndSnap := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0); // 进程快照句柄
  if FhndSnap = 0 then
    Exit;
  PE.dwSize := SizeOf(PE);
  if Process32First(FhndSnap, PE) then //第一个进程
    repeat
      New(PPE);
      PPE^ := PE;
      FProcList.Add(PPE);
    until not Process32Next(FhndSnap, PE);
end;

initialization
  m_WinIcon := 0;

finalization
  if Assigned(FWinNTInfo) then
    FreeAndNil(FWinNTInfo);
  if Assigned(FWin95Info) then
    FreeAndNil(FWin95Info);
  if 0 <> m_WinIcon then
    DestroyIcon(m_WinIcon);

end.

