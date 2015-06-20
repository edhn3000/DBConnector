unit U_Log;

interface
uses
  Windows, SysUtils, Dialogs;

const
  C_LogLevel_DEBUG = 1;
  C_LogLevel_INFO = 2;
  C_LogLevel_WARN = 3;
  C_LogLevel_ERROR = 4;
  C_ARY_LOG_LEVEL : array[1..4] of string = ('调试', '信息', '警告', '错误');

type
  { TSysLog }
  TSysLog = class
  private
    FLogfile: TextFile;
    FLogOpened: Boolean;
  protected
    FOwner: TObject;
    FRelativePath: string;
    FBlockFile: Boolean;
    FLogFileName: string;
    FShowToViewer: Boolean;
    FMaxLogCount: Integer;

    function getModulePath: String;
    procedure OpenLogFile();
    procedure CloseLogFile();
    procedure NewLogFile(const sFilename: string);
    procedure LogToDebugViewer(const AMsg: string; nLevel: Integer); virtual;
  public
    property RelativePath: String read FRelativePath write FRelativePath;
    property BlockFile: Boolean read FBlockFile write FBlockFile;
    property LogFileName: string read FLogFileName write FLogFileName;
    property ShowToViewer: Boolean read FShowToViewer write FShowToViewer;
    property MaxLogCount: Integer read FMaxLogCount write FMaxLogCount;

    constructor Create(owner: TObject = nil; blockFile: Boolean = True);
    destructor Destroy; override;

    class function GetLogLevelCN(nLevel: Integer): String;

    procedure InitLogger(fileName: string); virtual;
    procedure CloseLogger; virtual;
    procedure Log(const AMsg: string; nLevel: Integer); virtual;
    procedure Debug(AMsg: string);
    procedure Info(AMsg: string);
    procedure Warn(AMsg: string);
    procedure Error(AMsg: string; e: Exception = nil);
  end;

var
  // 日志级别，加载配置前默认为Debug，加载配置后默认为INFO
  G_LogLevel: Integer = C_LogLevel_DEBUG;
  SysLog: TSysLog;

implementation

uses
  DateUtils;
var
  m_dllPath: String;

function GetFileModifiedTime(const FileName: string): TDateTime;
var
  Handle: THandle;
  FindData: TWin32FindData;
  LocalFileTime: TFileTime;
  DosDateTime: Integer;
begin
  Handle := FindFirstFile(PChar(FileName), FindData);
  if Handle <> INVALID_HANDLE_VALUE then
  begin
    Windows.FindClose(Handle);
    if (FindData.dwFileAttributes and FILE_ATTRIBUTE_DIRECTORY) = 0 then
    begin
      FileTimeToLocalFileTime(FindData.ftLastWriteTime, LocalFileTime);
      if FileTimeToDosDateTime(LocalFileTime, LongRec(DosDateTime).Hi,
        LongRec(DosDateTime).Lo) then
      begin
        Result := FileDateToDateTime(DosDateTime);
        Exit;
      end;
    end;
  end;
  Result := -1;
end;

{ TSysLog }

function TSysLog.getModulePath: String;
var
  DllPath: array[0..255] of WideChar;
begin
  if m_dllPath = '' then
  begin
    GetModuleFileName(HInstance, DllPath, 255);
    m_dllPath := ExtractFilePath(StrPas(DllPath));
  end;
  Result := m_dllPath;
end;

procedure TSysLog.CloseLogger;
begin
  if FBlockFile then begin
    CloseLogFile;
  end;
//  Info('===========================结束记录=================================');
end;

constructor TSysLog.Create(owner: TObject; blockFile: Boolean);
begin
  FOwner := owner;
  FBlockFile := blockFile;
  FShowToViewer := True;
  FLogOpened := False;
  FMaxLogCount := 7;
end;

procedure TSysLog.OpenLogFile;
begin
  if not FLogOpened then begin
    AssignFile(FLogfile,FLogFileName);
    Append(FLogfile);
    FLogOpened := True;
  end;
end;

procedure TSysLog.CloseLogFile;
begin
  if FLogOpened then begin
    Flush(FLogfile);
    CloseFile(FLogfile);
    FLogOpened := False;
  end;
end;

procedure TSysLog.NewLogFile(const sFilename: string);
var
  sNewFileName, sOldLogFile: string;
  fileDate: TDateTime;
  i: Integer;
begin
  if FileExists(sFilename) then
  begin
    fileDate := GetFileModifiedTime(sFilename);
    if not SameDate(fileDate, Now) then
    begin
      sNewFileName := sFilename + '.' + FormatDateTime('yyyymmdd', fileDate);
      if FileExists(sNewFileName) then
        DeleteFile(sNewFileName);
      RenameFile(sFilename, sNewFileName);
    end;
  end;

  // 删除过多文件
  if FMaxLogCount > 0 then begin
    fileDate := Now - FMaxLogCount;
    sOldLogFile := sFilename + '.' + FormatDateTime('yyyymmdd', fileDate);
    repeat
      if FileExists(sOldLogFile) then
        DeleteFile(sOldLogFile)
      else
        Break;
      fileDate := fileDate - 1;
      sOldLogFile := sFilename + '.' + FormatDateTime('yyyymmdd', fileDate);
    until not FileExists(sOldLogFile);
  end;

  if not FileExists(sFilename) then
    FileClose(FileCreate(sFilename));
end;

procedure TSysLog.InitLogger(fileName: string);
var
  sDir: String;
begin
  if FRelativePath = '' then
    sDir := getModulePath
  else
    sDir := FRelativePath;
  FLogFileName := IncludeTrailingPathDelimiter(sDir) + fileName;
  sDir := ExtractFilePath(FLogFileName);
  if not DirectoryExists(sDir) then
    ForceDirectories(sDir);

  NewLogFile(FLogFileName);

  if FBlockFile then begin
    OpenLogFile;
  end;

//  Info('===========================开始记录=================================');
end;

procedure TSysLog.Log(const AMsg: string; nLevel: Integer);
var
  className: String;
begin
  if (nLevel >= G_LogLevel) or (nLevel = C_LogLevel_ERROR) then
  begin
    if ShowToViewer then
      LogToDebugViewer(AMsg, nLevel);

    if not FBlockFile then begin
      OpenLogFile;
    end;
    try
      if Assigned(FOwner) then
        className := '[' + FOwner.ClassName + ']-';

      Writeln(FLogfile, Format('%s %s[%s] [ThreadId=%d] %s',[
        FormatDateTime('yyyy-mm-dd hh:nn:ss.zzz ',Now), className,
        TSysLog.GetLogLevelCN(nLevel), GetCurrentThreadId, AMsg]));
      Flush(FLogfile);
    finally
      if not FBlockFile then begin
        CloseLogFile;
      end;
    end;
  end;
end;

procedure TSysLog.LogToDebugViewer(const AMsg: string; nLevel: Integer);
begin
  Windows.OutputDebugString(PWideChar(Format('校对[%s]:%s',
    [GetLogLevelCN(nLevel), AMsg])));
end;

destructor TSysLog.Destroy;
begin
  CloseLogger;
  inherited;
end;

procedure TSysLog.Debug(AMsg: string);
begin
  Log(AMsg, C_LogLevel_DEBUG);
end;

procedure TSysLog.Info(AMsg: string);
begin
  Log(AMsg, C_LogLevel_INFO);
end;

procedure TSysLog.Warn(AMsg: string);
begin
  Log(AMsg, C_LogLevel_WARN);
end;

procedure TSysLog.Error(AMsg: string; e: Exception);
begin
  if Assigned(e) then begin
    Log(Format(AMsg + ' Exception:Class=%s, Message=%s.'
      ,[e.ClassName,e.ToString]), C_LogLevel_ERROR);
  end else begin
    Log(AMsg, C_LogLevel_ERROR);
  end;
end;

class function TSysLog.GetLogLevelCN(nLevel: Integer): String;
begin
  if (nLevel >=Low(C_ARY_LOG_LEVEL))
    and (nLevel <= High(C_ARY_LOG_LEVEL)) then begin
    Result := C_ARY_LOG_LEVEL[nLevel];
  end else begin
    Result := 'UnknownLevel';
  end;
end;

initialization

finalization
  if Assigned(SysLog) then
    FreeAndNil(SysLog);

end.

