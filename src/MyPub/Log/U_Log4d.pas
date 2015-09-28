unit U_Log4d;

interface
uses
  Windows, SysUtils, Dialogs, Log4D, U_Log;

type
  { TTextLog4d }
  TTextLog4d = class(TTextFileLog)
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

procedure TTextLog4d.CloseLogger;
begin

//  Info('===========================结束记录=================================');
end;

constructor TTextLog4d.Create(owner: TObject; blockFile: Boolean);
begin
  FBlockFile := blockFile;
  FOwner := owner;
  FShowToViewer := True;
end;

destructor TTextLog4d.Destroy;
begin

  inherited;
end;

procedure TTextLog4d.InitLogger(configFile: string);
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

procedure TTextLog4d.Log(const AMsg: string; nLevel: Integer);
begin
  if (nLevel >= G_LogLevel) or (nLevel = C_LogLevel_ERROR) then
  begin
    if ShowToViewer then
      ConsoleLog.Log(AMsg, nLevel);

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
