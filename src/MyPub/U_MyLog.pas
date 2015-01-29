{**
 * <p>Title: TLog </p>
 * <p>Copyright: Copyright (c) 2006/10/14</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: ��־��  </p>
 *}
unit U_MyLog;

interface

type
  { ��¼��־���� }
  TLog = class
  private
    Ftxtlog: TextFile;
    FsFileName: string;
    FbFirstWrite: Boolean;  //�Ƿ��һ��д��
    FbReWriteFile: Boolean;
    FbFileOpend: Boolean;

    // ������
    FnInfoCount: Integer;
    FnWarnCount: Integer;
    FnErrorCount: Integer;

    function GetLogsCount: Integer;
    procedure SetFsFileName(value: string);

  protected
    procedure WriteLog(sLog: string);
  public
    property FileName: string read FsFileName write SetFsFileName;
    property ReWriteFile: Boolean read FbReWriteFile write FbReWriteFile;
    property InfoCount: Integer read FnInfoCount;
    property WarnCount: Integer read FnWarnCount;
    property ErrorCount: Integer read FnErrorCount;
    property LogsCount: Integer read GetLogsCount;

  public
    constructor Create(sLogFileFullName: string; bReWriteFile: Boolean = False);
    destructor Destroy; override;

    procedure Info(sInfo: string; bLogTime: Boolean = True);
    procedure Warn(sWarn: string; bLogTime: Boolean = True);
    procedure Error(sError: string; bLogTime: Boolean = True);

    procedure ClearInfoCount;
    procedure ClearWarnCount;
    procedure ClearErrorCount;
    procedure ClearLogsCount;
  end;

  function GlobalLog: TLog;

implementation

uses
  SysUtils, Forms;

const
  C_sLogFile = 'log.txt';

var
  m_Log: TLog;

function GlobalLog: TLog;
begin
  if not Assigned(m_log) then
  begin
    m_Log := TLog.Create(ExtractFilePath(Application.ExeName) + C_sLogFile);
  end;
  Result := m_Log;
end;  

{ TLog }

procedure TLog.WriteLog(sLog: string);  
var
  nHandle: Integer;
begin
  // ��һ��д����־ʱ����ļ��Ƿ���ڣ������д�ļ��ı�־����ռ�������
  if FbFirstWrite then
  begin 
    if not FileExists(FsFileName) then
    begin
      nHandle := FileCreate(FsFileName);
      FileClose(nHandle);
    end;
    if FbFileOpend then
      CloseFile(Ftxtlog);
    AssignFile(Ftxtlog, FsFileName);
    if not ReWriteFile then
      Append(Ftxtlog)
    else
      ReWrite(Ftxtlog);
    FbFileOpend := True;
    Writeln(Ftxtlog, '======================================');
    Writeln(Ftxtlog, '====��ʼ��¼��־'+ FormatDateTime('yyyy-mm-dd hh:mm:ss' ,Now) +'===');
    Writeln(Ftxtlog, '======================================');
    ClearLogsCount;
    FbFirstWrite := False;
  end;
  Writeln(Ftxtlog, sLog);
end;

constructor TLog.Create(sLogFileFullName: string; bReWriteFile: Boolean);
begin
  FsFileName := sLogFileFullName;
  FbFirstWrite := True;
  FbReWriteFile := bReWriteFile;
  FbFileOpend := False;
end;

destructor TLog.Destroy;
begin
  Writeln(Ftxtlog, '======================================');
  Writeln(Ftxtlog, '====������¼��־'+ FormatDateTime('yyyy-mm-dd hh:mm:ss' ,Now) +'===');
  Writeln(Ftxtlog, '======================================');
  CloseFile(Ftxtlog);
  inherited;
end;

procedure TLog.Error(sError: string; bLogTime: Boolean = True);
var
  sPrefix: string;
begin
  if bLogTime then
    sPrefix := FormatDateTime('yyyy-mm-dd hh:mm:ss' ,Now)
  else
    sPrefix := '';
  WriteLog( sPrefix + ' Error: ' + sError);
  Inc(FnErrorCount);
end; 

procedure TLog.Warn(sWarn: string; bLogTime: Boolean);
var
  sPrefix: string;
begin
  if bLogTime then
    sPrefix := FormatDateTime('yyyy-mm-dd hh:mm:ss' ,Now)
  else
    sPrefix := '';
  WriteLog( sPrefix + ' Warn: ' + sWarn);
  Inc(FnWarnCount);
end;

procedure TLog.Info(sInfo: string; bLogTime: Boolean = True);
var
  sPrefix: string;
begin 
  if bLogTime then
    sPrefix := FormatDateTime('yyyy-mm-dd hh:mm:ss' ,Now)
  else
    sPrefix := '';
  WriteLog(sPrefix + ' Info: ' + sInfo);
  Inc(FnInfoCount);
end;

procedure TLog.ClearErrorCount;
begin
  FnErrorCount := 0;
end;

procedure TLog.ClearInfoCount;
begin
  FnInfoCount := 0;
end;

procedure TLog.ClearLogsCount;
begin
  FnInfoCount := 0; 
  FnWarnCount := 0; 
  FnErrorCount := 0;
end;

procedure TLog.ClearWarnCount;
begin
  FnWarnCount := 0;
end;

function TLog.GetLogsCount: Integer;
begin
  Result := FnInfoCount + FnWarnCount + FnErrorCount;
end;

procedure TLog.SetFsFileName(value: string);
begin
  FsFileName := value;
  FbFirstWrite := True;
end;

initialization

finalization
  if Assigned(m_Log) then
    FreeAndNil(m_Log);

end.
