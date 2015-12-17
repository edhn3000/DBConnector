unit U_Log;

interface
uses
  Windows, SysUtils, Dialogs, Classes;

const
  C_LogLevel_DEBUG = 1;
  C_LogLevel_INFO = 2;
  C_LogLevel_WARN = 3;
  C_LogLevel_ERROR = 4;
  C_ARY_LOG_LEVEL : array[1..4] of string = ('����', '��Ϣ', '����', '����');

type
  { TAbstractLog }
  TAbstractLog = class
  private
    FLogLevel: Integer;
  public
    // Ĭ��Ϊ0��ʹ��ȫ�ֱ������粻Ϊ0��Ӧʹ�ô˱�������ʵ��ÿ��Log����ͬLevel
    property LogLevel: Integer read FLogLevel write FLogLevel default 0;

    procedure Log(const AMsg: string; nLevel: Integer); virtual; abstract;
    procedure Debug(AMsg: string);
    procedure Info(AMsg: string);
    procedure Warn(AMsg: string);
    procedure Error(AMsg: string; e: Exception = nil);

    class function GetLogLevelCN(nLevel: Integer): String;
  end;

  { TConsoleLog }
  TConsoleLog = class(TAbstractLog)
  public
    procedure Log(const AMsg: string; nLevel: Integer); override;

  end;

  { TTextFileLog }
  TTextFileLog = class(TAbstractLog)
  private
    FLogfile: TextFile;
    FLogOpened: Boolean;
    FLogLock: TRTLCriticalSection;
  protected
    FOwner: TObject;
    FRelativePath: string;
    FBlockFile: Boolean;
    FLogFileName: string;
    FShowToViewer: Boolean;
    FMaxLogCount: Integer;

    function getModulePath: String;
    function IsEmptyFile(fileName: String): Boolean;
    procedure OpenLogFile();
    procedure CloseLogFile();
    procedure DeleteOldFiles(logFileName: string; reserveCount: Integer);
    procedure NewLogFile(const sFilename: string);
  public
    property RelativePath: String read FRelativePath write FRelativePath;
    property BlockFile: Boolean read FBlockFile write FBlockFile;
    property LogFileName: string read FLogFileName write FLogFileName;
    property ShowToViewer: Boolean read FShowToViewer write FShowToViewer;
    property MaxLogCount: Integer read FMaxLogCount write FMaxLogCount;

    constructor Create(owner: TObject = nil; blockFile: Boolean = True);
    destructor Destroy; override;

    procedure InitLogger(fileName: string); virtual;
    procedure CloseLogger; virtual;
    procedure Log(const AMsg: string; nLevel: Integer); override;

    procedure ClearFile;
  end;

var
  // ��־���𣬼�������ǰĬ��ΪDebug���������ú�Ĭ��ΪINFO
  G_LogLevel: Integer = C_LogLevel_DEBUG;
  SysLog: TTextFileLog;
  ConsoleLog: TConsoleLog;

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

procedure FindOldLogFiles(sPath: string; fileShrotName: string; outFileList: TStrings);
var
  FSrchRec: TSearchRec;
  nFindResult: Integer;
  nSearchMode: Integer;
begin
  sPath := IncludeTrailingPathDelimiter(sPath);
  nSearchMode := faAnyFile;

  nFindResult := FindFirst(sPath + fileShrotName + '.*', nSearchMode, FSrchRec);
  try
    while nFindResult = 0 do
    begin
      if (FSrchRec.Name <> '.') and (FSrchRec.Name <> '..') then
      begin
        if (FSrchRec.Attr and faDirectory) = faDirectory then
        begin
          FindOldLogFiles(FSrchRec.Name, fileShrotName, outFileList);
        end else if fileShrotName <> FSrchRec.Name then begin
          outFileList.Add(sPath + FSrchRec.Name);
        end;
      end;

      nFindResult := FindNext(FSrchRec);
    end;
  finally
    FindClose(FSrchRec);
  end;
end;

{ TAbstractLog }

procedure TAbstractLog.Debug(AMsg: string);
begin
  Log(AMsg, C_LogLevel_DEBUG);
end;

procedure TAbstractLog.Info(AMsg: string);
begin
  Log(AMsg, C_LogLevel_INFO);
end;

procedure TAbstractLog.Warn(AMsg: string);
begin
  Log(AMsg, C_LogLevel_WARN);
end;

procedure TAbstractLog.Error(AMsg: string; e: Exception);
begin
  if Assigned(e) then begin
    Log(Format(AMsg + ' Exception:Class=%s, Message=%s.'
      ,[e.ClassName,e.ToString]), C_LogLevel_ERROR);
  end else begin
    Log(AMsg, C_LogLevel_ERROR);
  end;
end;

class function TAbstractLog.GetLogLevelCN(nLevel: Integer): String;
begin
  if (nLevel >=Low(C_ARY_LOG_LEVEL))
    and (nLevel <= High(C_ARY_LOG_LEVEL)) then begin
    Result := C_ARY_LOG_LEVEL[nLevel];
  end else begin
    Result := 'UnknownLevel_' + IntToStr(nLevel);
  end;
end;

{ TTextFileLog }

function TTextFileLog.getModulePath: String;
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

procedure TTextFileLog.CloseLogger;
begin
  if FBlockFile then begin
    CloseLogFile;
  end;
//  Info('===========================������¼=================================');
end;

constructor TTextFileLog.Create(owner: TObject; blockFile: Boolean);
begin
  InitializeCriticalSection(FLogLock);
  FOwner := owner;
  FBlockFile := blockFile;
  FShowToViewer := True;
  FLogOpened := False;
  FMaxLogCount := 7;
end;

destructor TTextFileLog.Destroy;
begin
  CloseLogger;
  DeleteCriticalSection(FLogLock);
  inherited;
end;

function TTextFileLog.IsEmptyFile(fileName: String): Boolean;
var
  f: TextFile;
begin
  AssignFile(f, fileName);
  try
    Reset(f);
    Result := Eof(f);
  finally
    CloseFile(f);
  end;
end;

procedure TTextFileLog.OpenLogFile;
begin
  if not FLogOpened then begin
    // ������ģʽʱlog�ļ����������й����б�ɾ��
    if not FBlockFile and not FileExists(FLogFileName) then begin
      NewLogFile(FLogFileName);
    end;
    AssignFile(FLogfile,FLogFileName);
    Append(FLogfile);
    FLogOpened := True;
  end;
end;

procedure TTextFileLog.ClearFile;
begin
  EnterCriticalSection(FLogLock);
  try
    if not FBlockFile then begin
      OpenLogFile;
    end;
    try
      System.Rewrite(FLogFile);
      Flush(FLogfile);
    finally
      if not FBlockFile then begin
        CloseLogFile;
      end;
    end;
  finally
    LeaveCriticalSection(FLogLock);
  end;
end;

procedure TTextFileLog.CloseLogFile;
begin
  if FLogOpened then begin
    Flush(FLogfile);
    CloseFile(FLogfile);
    FLogOpened := False;
  end;
end;

procedure TTextFileLog.DeleteOldFiles(logFileName: String; reserveCount: Integer);
var
  i: Integer;
  sOldLogFile: String;
  fileList: TStringList;
begin
  // ɾ�������ļ�
  fileList := TStringList.Create;
  try
    FindOldLogFiles(ExtractFilePath(logFileName), ExtractFileName(logFileName), fileList);
    // ������ǰ����ɾ������������˳�򿿺�ģ����ǽ��µģ���Ϊ�ļ����������
    fileList.Sort;
    for i := 0 to fileList.Count - 1 - reserveCount do begin
      sOldLogFile := fileList[i];
      DeleteFile(sOldLogFile);
    end;
  finally
    fileList.Free;
  end;
end;

procedure TTextFileLog.NewLogFile(const sFilename: string);
var
  sNewFileName: string;
  fileDate: TDateTime;
begin
  if FileExists(sFilename) then
  begin
    fileDate := GetFileModifiedTime(sFilename);
    if not SameDate(fileDate, Now) then
    begin
      sNewFileName := sFilename + '.' + FormatDateTime('yyyymmdd', fileDate);
      if FileExists(sNewFileName) then
        DeleteFile(sNewFileName);
      if IsEmptyFile(sFilename) then
        DeleteFile(sFilename)
      else
        RenameFile(sFilename, sNewFileName);
    end;
  end;

  DeleteOldFiles(sFilename, FMaxLogCount);

  if not FileExists(sFilename) then
    FileClose(FileCreate(sFilename));
end;

procedure TTextFileLog.InitLogger(fileName: string);
var
  sDir, sExt, sName: String;
begin
  if FRelativePath = '' then
    sDir := getModulePath
  else
    sDir := FRelativePath;
  FLogFileName := IncludeTrailingPathDelimiter(sDir) + fileName;
  sDir := ExtractFilePath(FLogFileName);

  EnterCriticalSection(FLogLock);
  try
    if not DirectoryExists(sDir) then
      ForceDirectories(sDir);

    NewLogFile(FLogFileName);

    if FBlockFile then begin
      try
        OpenLogFile;
      except
        sExt := ExtractFileExt(FLogFileName);
        if sExt <> '' then begin
          sName := Copy(FLogFileName, 1, Length(FLogFileName) - Length(sExt));
        end else
          sName := FLogFileName;
        FLogFileName := Format(sName + '_Thd%d' + sExt, [GetCurrentThreadId]);
        NewLogFile(FLogFileName);

        OpenLogFile;
      end;
    end;
  finally
    LeaveCriticalSection(FLogLock);
  end;

//  Info('===========================��ʼ��¼=================================');
end;

procedure TTextFileLog.Log(const AMsg: string; nLevel: Integer);
var
  className: String;
  level: Integer;
begin
  if Self.LogLevel = 0 then
    level := G_LogLevel
  else
    level := Self.LogLevel;

  if (nLevel >= level) or (nLevel >= C_LogLevel_ERROR) then
  begin
    if ShowToViewer then
      ConsoleLog.Log(AMsg, nLevel);

    EnterCriticalSection(FLogLock);
    try
      if not FBlockFile then begin
        OpenLogFile;
      end;
      try
        if Assigned(FOwner) then
          className := '[' + FOwner.ClassName + ']-';

        Writeln(FLogfile, Format('%s %s[%s] [ThreadId=%d] %s',[
          FormatDateTime('yyyy-mm-dd hh:nn:ss.zzz ',Now), className,
          TTextFileLog.GetLogLevelCN(nLevel), GetCurrentThreadId, AMsg]));
        Flush(FLogfile);
      finally
        if not FBlockFile then begin
          CloseLogFile;
        end;
      end;
    finally
      LeaveCriticalSection(FLogLock);
    end;
  end;
end;

{ TConsoleLog }

procedure TConsoleLog.Log(const AMsg: string; nLevel: Integer);
begin
  Windows.OutputDebugString(PWideChar(Format('Console[%s][ThreadId=%d]:%s',
    [GetLogLevelCN(nLevel), GetCurrentThreadId, AMsg])));
end;

initialization
  ConsoleLog := TConsoleLog.Create;

finalization
  if Assigned(ConsoleLog) then
    FreeAndNil(ConsoleLog);

end.

