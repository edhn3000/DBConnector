unit U_Log4d;

interface
uses
  Windows, SysUtils, Dialogs, Log4D, U_Log;

type
  { TSysLog4d }
  TSysLog4d = class(TSysLog)
  private
    FLogger: TLogLogger;
  public
    constructor Create(owner: TObject = nil; blockFile: Boolean = True);
    destructor Destroy; override;

    procedure InitLogger(configFile: string); override;
    procedure CloseLogger; override;
    procedure Log(const AMsg: string; nLevel: Integer); override;
  end;

implementation


{ TSysLog4d }

procedure TSysLog4d.CloseLogger;
begin

//  Info('===========================结束记录=================================');
end;

constructor TSysLog4d.Create(owner: TObject; blockFile: Boolean);
begin
  FBlockFile := blockFile;
  FOwner := owner;
  FShowToViewer := True;
end;

destructor TSysLog4d.Destroy;
begin

  inherited;
end;

procedure TSysLog4d.InitLogger(configFile: string);
var
  sDir, configFileName: String;
begin
  if FRelativePath = '' then
    sDir := getModulePath
  else
    sDir := FRelativePath;
  Log4D.relativePath := sDir;
  configFileName := sDir + configFile;
  TLogPropertyConfigurator.Configure(configFileName);
  FLogger := TLogLogger.GetLogger('logger');
  FLogFileName := '';

//  Info('===========================开始记录=================================');
end;

procedure TSysLog4d.Log(const AMsg: string; nLevel: Integer);
begin
  if (nLevel >= G_LogLevel) or (nLevel = C_LogLevel_ERROR) then
  begin
    if ShowToViewer then
      LogToDebugViewer(AMsg, nLevel);

    case nLevel of
      C_LogLevel_DEBUG: begin
        FLogger.Log(Log4D.Debug, AMsg);
      end;
      C_LogLevel_INFO: begin
        FLogger.Log(Log4D.Info, AMsg);
      end;
      C_LogLevel_WARN: begin
        FLogger.Log(Log4D.Warn, AMsg);
      end;
      C_LogLevel_ERROR: begin
        FLogger.Log(Log4D.Error, AMsg);
      end;
      else begin
        // unknown level
      end;
    end;
  end;
end;

end.
