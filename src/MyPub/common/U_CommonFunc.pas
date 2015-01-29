{**
 * <p>Title: CommonFunc </p>
 * <p>Copyright: Copyright (c) 2006/10/14</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 通用方法单元 </p>
 *}
unit U_CommonFunc;

interface

{$WARN SYMBOL_PLATFORM OFF}

uses
  Windows, Messages, Dialogs, Forms, Controls, Classes, Types,
  Graphics, WinSvc;

const
  C_nFormatMsgBox_AppHandle = 0;

type
{ TStringClass }
  TStringClass = class
    FStr: string;
  public
    property Str: string read FStr write FStr;
  public
    constructor Create(AStr: string);
  end;

  TFunc = function: Variant;stdcall;

type
{ TSysConfig 系统配置信息 }
  TSysConfig = class
  public
    // reg key
    WindowsPath    : string;		 	 // Windows path, eg: c:\Windows
    WinSysPath     : string;		   // Windows System path, eg: c:\Windows\System32
    ProgramFiles   : string;	     // Program Files path, eg: c:\Program Files
    CommonFiles    : string;		 	 // Common files, eg: c:\Program Files\Common Files
    ApplicationData: string;       // %USERPROFILE%\Application Data
    Desktop        : string;       // %USERPROFILE%\桌面
    StartMenu      : string;       // %USERPROFILE%\「开始」菜单
    ProgramsMenu   : string;       // %USERPROFILE%\「开始」菜单\程序
    StartupMenu    : string;       // %USERPROFILE%\「开始」菜单\程序\启动
    WinTemp        : string;       // %SystemRoot%\Temp
  public
    constructor Create;

    function GetPreDefinePath(sDef: string): string;
  end;

{ Messagebox样式 }
  TMsgBoxStyle = (mbsInfo, mbsError, mbsQuestion, mbsYesNoCancel, mbsWarn,
    mbsStop, mbsHand, mbsFail);

{ windoes版本枚举 }
  TWinVer = (wvUnKnown,
    wvWin95, wvWin98, wvWinMe,
    wvWin2k, wvWin2kServ,
    wvWinXp,
    wvWin2003,
    wvWinNT3, wvWinNT3Serv,
    wvWinNT4, wvWinNT4Serv,
    wvWinCE);

{ TCommonFunc 通用方法类 }
  TCommonFunc = class
  private

  protected

  public

  public
    constructor Create;
    destructor Destroy;override;

    { 系统参数 }
    function GetWinVersionInfo(var SProduct, SVersion, SServicePack: string):TWinVer;overload;
    function GetWinVersionInfo:TWinVer;overload;   // 获得windows版本
    function OSIsNT: Boolean;
    function IsChineseChar(c: WideChar): Boolean;
    function HasChineseChar(WS: WideString): Boolean;
    function IsMultiByteChar(c: Char): Boolean;
    function UsesSmallFont: boolean;
    function ParseRunTimeParam(param: string; indexBegin, indexEnd: Integer;
        var pos: Integer):Boolean;overload;
    function ParseRunTimeParam(param: string; indexBegin, indexEnd: Integer):Boolean;overload;
    function DateTimeToPassedTimeStr(dt: TDateTime): string;
    function PrecisionNumStr(S: string; Precision: Integer): string;
    function RunProgram(sExeFile: string; nShowCmd: Integer; WaitFor: Boolean): Boolean;overload;
    function RunProgram(sExeFile: string; nShowCmd: Integer;
      WaitFor: Boolean; var si: TStartupInfo; var pi: TProcessInformation): Boolean;overload;
    // 方法未完善
    function IsInvalidObj(obj: TObject): Boolean;
    function ReallyAssigned(const o: TObject):Boolean;

    { 内容处理 }

    // 字符串加密解密
    function EncryptStr(const Astr: string): string;
    function DecryptStr(const Astr: string): string;

    { 窗体、控件 }
    function FormShowModal(F: TFormClass;FP: TComponent): Integer;   
    procedure FormShowOnTop(F: TForm);
    procedure SetOnTopMost(Ahnd: THandle;bOnTop: Boolean);
    procedure SetNotOnTaskBar();
    procedure MinOnShowDesktop(AHnd: THandle; bMin: Boolean);
    procedure ResetIconCache;

    { 文件、目录操作 }
    procedure FileRemoveReadOnly(sFile: string);
    // 去掉扩展名的文件名
    function RemoveFileExt(s: string): string;
     
    function GetNewGUID32: string;
  end;

  // 格式的messagebox    标题，按钮和图标已经配套规定，通过传枚举选择用哪种
  function FormatMsgBox(vMsg:Variant; mbs: TMsgBoxStyle =
      mbsInfo;Ahnd: THandle = 0; sCaptionDef: string = ''): Integer;


  // 调用此函数无效，此函数仅用于查看处理代码
  function LoadFuncFromDll(DllName, FuncName: string; var AFunc: TFunc): THandle;

  // 系统参数
  function SysConfig: TSysConfig;

  // 获得系统服务状态
  function GetServiceStatus(sServiceName: string): Integer;
  function GetServiceStatusString(status: Integer): string;

var
  commonFunc: TCommonFunc;

implementation

uses
  SysUtils, Registry, Variants, ActiveX;

var
  m_SysConfig: TSysConfig;


function SysConfig: TSysConfig;
begin
  if not Assigned(m_SysConfig) then
    m_SysConfig := TSysConfig.Create;
  Result := m_SysConfig;
end;

procedure FormShowModal(F: TFormClass;FP: TComponent);
begin
  with F.Create(FP) do
  try
    ShowModal;
  finally
    Free;
  end;
end;

procedure SetOnTop(Ahnd: THandle;bOnTop: Boolean);
begin
  if bOnTop then
    SetWindowPos(Ahnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE)
  else
    SetWindowPos(Ahnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
end;

{ TCommonFunc }
procedure TCommonFunc.MinOnShowDesktop(AHnd: THandle; bMin: Boolean);
begin
//  SetWindowPos(Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
  if bMin then
    SetWindowLong(AHnd,GWL_HWNDPARENT, GetDesktopWindow);
end;

function TCommonFunc.GetWinVersionInfo(var SProduct, SVersion, SServicePack: string):
  TWinVer;
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

function TCommonFunc.GetWinVersionInfo:TWinVer;
var
  a, b, c: string;
begin
  Result := GetWinVersionInfo(a,b,c);
end;   

function TCommonFunc.OSIsNT: Boolean;
begin
  Result := GetWinVersionInfo in [wvWin2k, wvWin2kServ, wvWinXp, wvWin2003,
    wvWinNT3, wvWinNT3Serv, wvWinNT4, wvWinNT4Serv];
end;

function FormatMsgBox(vMsg:Variant; mbs: TMsgBoxStyle; Ahnd: THandle;
  sCaptionDef: string): Integer;
var
  sCaption, sMsg: string;
  nStyle: Integer;
begin
  case mbs of
    mbsInfo:
      begin
        sCaption := '提示';
        nStyle := MB_OK + MB_ICONINFORMATION;
      end;
    mbsError:
      begin
        sCaption := '错误';
        nStyle := MB_OK + MB_ICONERROR;
      end;
    mbsQuestion:
      begin
        sCaption := '确认';
        nStyle := MB_YESNO + MB_ICONQUESTION;
      end;
    mbsWarn:
      begin
        sCaption := '警告';
        nStyle := MB_OK + MB_ICONWARNING;
      end;
    mbsStop:
      begin
        sCaption := '错误';
        nStyle := MB_OK + MB_ICONSTOP;
      end;
    mbsHand:
      begin
        sCaption := '提示';
        nStyle := MB_OK + MB_ICONHAND;
      end;
    mbsFail:
      begin
        sCaption := '失败';
        nStyle := MB_OK + MB_ICONEXCLAMATION;
      end;
    mbsYesNoCancel:
      begin
        sCaption := '确认';
        nStyle := MB_YESNOCANCEL + MB_ICONQUESTION;
      end;
  else
    begin
      sCaption := '提示';
      nStyle := MB_OK + MB_ICONINFORMATION;
    end;
  end;
  if sCaptionDef <> '' then
    sCaption := sCaptionDef;
  sMsg := VarToStrDef(vMsg, '');
  if Ahnd <> C_nFormatMsgBox_AppHandle then
    Result := MessageBox(Ahnd, PChar(sMsg), PChar(sCaption), nStyle)
  else
    Result := Application.MessageBox(PChar(sMsg), PChar(sCaption), nStyle);
end;

function LoadFuncFromDll(DllName, FuncName: string; var AFunc: TFunc): THandle;
begin
  Result := LoadLibrary(PChar(DllName));
  if Result <> Null then
  begin
    @AFunc := GetProcAddress(Result, PChar(FuncName));
  end;
end;     


constructor TCommonFunc.Create;
begin

end;

destructor TCommonFunc.Destroy;
begin
  inherited;
end;

function TCommonFunc.UsesSmallFont: boolean;
var
  DC: HDC;
begin
  DC := GetDC(0);
  Result := (GetDeviceCaps(DC, LOGPIXELSX) = 96);
  ReleaseDC(0, DC);
end;

function TCommonFunc.ParseRunTimeParam(param: string; indexBegin, indexEnd: Integer;
    var pos: Integer):Boolean;
var
  i: Integer;
begin
  if indexEnd > ParamCount then
    indexEnd := ParamCount;
  Result := False;
  for i := indexBegin to indexEnd do
  begin
    if SameText(ParamStr(i), param) then
    begin
      Pos := i;
      Result := True;
      Break;
    end;
  end;
end;

function TCommonFunc.ParseRunTimeParam(param: string; indexBegin, indexEnd: Integer):Boolean;
var
  n: Integer;
begin
  Result := ParseRunTimeParam(param, indexBegin, indexEnd, n);
end;

//function LoadFuncStrFromDll( DllName, FuncName: string; var AFuncStr:TFuncStr; var sResult: string ):THandle;
//begin
//  Result := LoadLibrary(PChar(DllName));
//  if Result <> Null then
//  begin
//    @AFuncStr := GetProcAddress(Result, PChar(FuncName));
//  end;
//end;

function TCommonFunc.FormShowModal(F: TFormClass; FP: TComponent): Integer;
begin
  with F.Create(FP) do
  try
    Result := ShowModal;
  finally
    Free;
  end;
end;

procedure TCommonFunc.SetOnTopMost(Ahnd: THandle; bOnTop: Boolean);
begin
  if bOnTop then
    SetWindowPos(Ahnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE)
  else
    SetWindowPos(Ahnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
end;

function TCommonFunc.DecryptStr(const Astr: string): string;
begin
  Result := '';
//  Result := Base64Decode(Astr);
end;

function TCommonFunc.EncryptStr(const Astr: string): string;
begin
  Result := '';
//  Result := Base64Encode(Astr);
end;

function TCommonFunc.DateTimeToPassedTimeStr(dt: TDateTime): string;
var
  fractional: Double;
  fsResult: Double;
begin
  fractional := dt - Trunc(dt);
  fsResult := fractional * 24 * 3600;  // 现在单位是秒
  if fsResult < 1 then
    Result := PrecisionNumStr(FloatToStr(fsResult*1000), 3) + '毫秒'
  else
    Result := PrecisionNumStr(FloatToStr(fsResult), 3) + '秒';
end;

function TCommonFunc.PrecisionNumStr(S: string; Precision: Integer): string;
var
  nPosPoint: Integer;
begin
  nPosPoint := Pos('.', S);
  if nPosPoint = 0 then
    Result := S
  else
    Result := Copy(S, 1, nPosPoint + Precision);
end;

procedure TCommonFunc.ResetIconCache;
var
  Reg: TRegistry;
  sShellIconSize: string;
  dwResult: DWORD;
begin
  Reg := TRegistry.Create;
  with Reg do
  try
    RootKey := HKEY_CURRENT_USER;
    if OpenKey('Control Panel\Desktop\WindowMetrics', False) then
    begin
      sShellIconSize := ReadString('Shell Icon Size');
      WriteString('Shell Icon Size', IntToStr(StrToInt(sShellIconSize)-1));
//      SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, 0);
      SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, SPI_SETNONCLIENTMETRICS, 0, SMTO_ABORTIFHUNG, 10000, dwResult);

      WriteString('Shell Icon Size', sShellIconSize);
//      SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, 0);
      SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, SPI_SETNONCLIENTMETRICS, 0, SMTO_ABORTIFHUNG, 10000, dwResult);
    end;
  finally
    Reg.Free;
  end;
end;

procedure TCommonFunc.FileRemoveReadOnly(sFile: string);
begin
  if FileExists( sFile ) then
    FileSetAttr( sFile, FileGetAttr (sFile) and (not SysUtils.faReadOnly) );
end;

function TCommonFunc.RemoveFileExt(s: string): string;
var
  i: Integer;
begin
  Result := s;
  for i := Length(Result) downto 1 do
  begin
    if Result[i] = '.' then
      Result := Copy(Result, 1, i - 1);
    if Result[i] = '\' then
    begin
      Result := Copy(Result, i + 1, Length(Result));
      Break;
    end;
  end;
end;    

function TCommonFunc.GetNewGUID32: string;
var
  guid: TGUID;
  r: HRESULT;
begin
  r := CoCreateGuid(guid);
  if S_OK = r then
  begin
    Result := GUIDToString(guid);
    Result := StringReplace(Result, '-', '', [rfReplaceAll]);
    Result := StringReplace(Result, '{', '', [rfReplaceAll]);  
    Result := StringReplace(Result, '}', '', [rfReplaceAll]);
  end
  else
    raise Exception.Create('生成guid时出错,code=' + IntToStr(r));
end;

function TCommonFunc.RunProgram(sExeFile: string; nShowCmd: Integer;
  WaitFor: Boolean; var si: TStartupInfo; var pi: TProcessInformation): Boolean;
begin
  FillMemory(@si, sizeof(si), 0);

  with si do
  begin
    cb := sizeof(si);
    dwFlags := STARTF_USESHOWWINDOW;
    wShowWindow := nShowCmd;
  end;

  Result := Boolean(CreateProcess(nil, PChar(sExeFile), nil, nil, False,
                                  NORMAL_PRIORITY_CLASS, nil, nil, si, pi));
  if Result and WaitFor then
    WaitForSingleObject(pi.hProcess, INFINITE);
end;

function TCommonFunc.RunProgram(sExeFile: string; nShowCmd: Integer;
  WaitFor: Boolean): Boolean;
var
  pi: TProcessInformation;
  si: TStartupInfo;
begin
  FillMemory(@si, sizeof(si), 0);
  with si do
  begin
    cb := sizeof(si);
    dwFlags := STARTF_USESHOWWINDOW;
    wShowWindow := nShowCmd;
  end;
  Result := CreateProcess(nil, PChar(sExeFile), nil, nil, False,
            NORMAL_PRIORITY_CLASS, nil, nil, si, pi);
  if Result and WaitFor then
    WaitForSingleObject(pi.hProcess, INFINITE);

  CloseHandle(pi.hProcess);
  CloseHandle(pi.hThread);
end;

function TCommonFunc.IsChineseChar(c: WideChar): Boolean;
begin
  Result := (Ord(c) >= 19968) and (Ord(c) <= 40869);
end;

function TCommonFunc.IsMultiByteChar(c: Char): Boolean;
begin
  Result := (ByteType(c, 1) <> mbSingleByte);
end;

function TCommonFunc.HasChineseChar(WS: WideString): Boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 1 to Length(WS) do
  begin
    if IsChineseChar(WS[i]) then
    begin
      Result := True;
      Break;
    end;
  end;
end;

function TCommonFunc.IsInvalidObj(obj: TObject): Boolean;
var
      dwStart,   dwAddr,   dwEnd:   Longword;
begin    
  Result := False;
  //   用内嵌汇编取变量的起始和末尾地址
  asm
    mov   dwStart,   ebp
    lea   eax,   obj
    mov   dwAddr,   eax
    lea   eax,   dwEnd
    mov   dwEnd,   eax
  end;

  //   根据变量的地址来判断是否合法
  if   dwAddr   <   dwEnd   then
    Result := True
        //   变量一般为有问题
  else   if   dwAddr   <   dwStart   then
//    Result := True
        //   变量是一个本函数的局部变量
  else   if   dwAddr   =   dwStart   then 
    Result := True
           //变量肯定有问题,   因为此值为函数的开始处
//  else   if   dwAddr   <   dwStart   +   ParamSize   then
//        //   变量为传值参数
//  else
//        //   变量为本函数外的变量
end;

function TCommonFunc.ReallyAssigned(const o: TObject): Boolean;
begin
  Result := System.Assigned(o);
  if Result then
  try
    o.ClassName;  // 随便进行一个调用
  except
    Result := False;
  end;
end;

procedure TCommonFunc.FormShowOnTop(F: TForm);
begin
  if not F.Showing then
  begin
    F.FormStyle := fsStayOnTop;
    F.Show;
  end
  else
    F.BringToFront;
end;


procedure TCommonFunc.SetNotOnTaskBar;
begin        
  SetWindowLong(Application.Handle, GWL_EXSTYLE,
    GetWindowLong(Application.Handle, GWL_EXSTYLE)
    or WS_EX_TOOLWINDOW AND NOT WS_EX_APPWINDOW);
end;

{ TStringClass }

constructor TStringClass.Create(AStr: string);
begin
  FStr := AStr;
end;

{ TSysConfig }

constructor TSysConfig.Create;
var
  Path: array[0..MAX_PATH] of Char;
  Reg: TRegistry;
begin
  GetWindowsDirectory(Path, MAX_PATH);
  WindowsPath := StrPas(Path);
  WinTemp := WindowsPath + '\temp';

  GetSystemDirectory(Path, MAX_PATH);
  WinSysPath := StrPas(Path);

  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKeyReadOnly('Software\Microsoft\Windows\CurrentVersion') then
    begin
      ProgramFiles := Reg.ReadString('ProgramFilesDir');
      CommonFiles := Reg.ReadString('CommonFilesDir');
    end;

    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKeyReadOnly('Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') then
    begin
      ApplicationData := Reg.ReadString('AppData');
      Desktop         := Reg.ReadString('Desktop');
      StartMenu       := Reg.ReadString('Start Menu');
      ProgramsMenu    := Reg.ReadString('Programs');
      StartupMenu     := Reg.ReadString('Startup');
    end;
  finally
    Reg.Free;
  end;
end;

function TSysConfig.GetPreDefinePath(sDef: string): string;
begin
  if SameText(sDef, 'Windows') then
    Result := WindowsPath
  else if SameText(sDef, 'System') then
    Result := WinSysPath
  else if SameText(sDef, 'ProgramFiles') then
    Result := ProgramFiles
  else if SameText(sDef, 'CommonFiles') then
    Result := CommonFiles
  else if SameText(sDef, 'ApplicationData') then
    Result := ApplicationData
  else if SameText(sDef, 'Desktop') then
    Result := Desktop
  else if SameText(sDef, 'StartMenu') then
    Result := StartMenu
  else if SameText(sDef, 'ProgramsMenu') then
    Result := ProgramsMenu
  else if SameText(sDef, 'StartupMenu') then
    Result := StartupMenu
  else
    Result := '';

//  Result := ExcludeTrailingBackslash(Result);
end;
     
function GetServiceStatus(sServiceName: string): Integer;
var
 hService, hSCManager: SC_HANDLE;
 SS: TServiceStatus;
begin
  hSCManager := OpenSCManager(nil, SERVICES_ACTIVE_DATABASE, SC_MANAGER_CONNECT);
  if hSCManager = 0 then
  begin
   result := -1;
   exit;
  end;
  hService := OpenService(hSCManager, PChar(sServiceName), SERVICE_QUERY_STATUS);
  if hService = 0 then
  begin
   CloseServiceHandle(hSCManager);
   result := -2;
   exit;
  end;
  if not QueryServiceStatus(hService, SS) then
   result := -3
  else
  begin
    Result := SS.dwCurrentState;
  end;
  CloseServiceHandle(hSCManager);
  CloseServiceHandle(hService);
end;  

function GetServiceStatusString(status: Integer): string;
begin
  case status of
   -1:
     Result :='Can not open the service control manager';
   -2:
     Result := 'Can not open the service';
   -3:
     Result := 'Can not query the service status';
   SERVICE_CONTINUE_PENDING:
     result := 'The service continue is pending';
   SERVICE_PAUSE_PENDING:
     result := 'The service pause is pending.';
   SERVICE_PAUSED:
     result := 'The service is paused.';
   SERVICE_RUNNING:
     result := 'The service is running.';
   SERVICE_START_PENDING:
     result := 'The service is starting.';
   SERVICE_STOP_PENDING:
     result := 'The service is stopping.';
   SERVICE_STOPPED:
     result := 'The service is not running.';
  else
   result := 'Unknown Status';
  end;
end; 

initialization
  commonFunc := TCommonFunc.Create;

finalization
  if Assigned(commonFunc) then
    FreeAndNil(commonFunc);
  if Assigned(m_SysConfig) then
    FreeAndNil(m_SysConfig);

end.
