unit TDailyRollingFileAppenderUnit;

interface

uses
   Classes, DateUtils, Windows, Log4D;

const
  RetentDaysOpt = 'retentDays';

type
{*----------------------------------------------------------------------------
  每天保存一个日志文件，删除给定时间之前的日志 
  ----------------------------------------------------------------------------}
  TDailyRollingFileAppender = class (TLogFileAppender)
  private
    procedure DeleteOldFiles();
  protected
    FShortName : string;
    FRetentDays : Integer;
    FCurrDate : TDateTime;
  protected
    procedure SetOption(const Name, Value: string); override;
  public
    constructor Create(const Name, FileName: string;
      const Layout: ILogLayout = nil; const Append: Boolean = True);
      reintroduce; overload; virtual;

    procedure Init; override;
    procedure SetRetentDays(const ARetentDays : Integer); Virtual;
    procedure DoAppend(const AEvent : TLogEvent); Override;
    procedure RollOver(); Virtual;
    function GetDayOut() : Int64; Virtual;
    procedure SetLogFile(const AFileName: string); override;
  end;

implementation

uses
  SysUtils;

const
  C_FILE_DATEFORMAT = 'yyyy-mm-dd';

function GetDateFileName( AFileName: string; LogDate: TDateTime): string;
begin
  Result := AFileName + '.' + FormatDateTime( C_FILE_DATEFORMAT, LogDate );
end;  

constructor TDailyRollingFileAppender.Create(const Name, FileName: string;
      const Layout: ILogLayout = nil; const Append: Boolean = True);
begin
  inherited;
  SetOption(RetentDaysOpt, '30');
  FCurrDate := Now;
end;

procedure TDailyRollingFileAppender.Init;
begin
  inherited;
  SetOption(RetentDaysOpt, '30');
  FCurrDate := Now;
end;

function FileTimeToTime(Fd: _FileTime): TDateTime;
var
  Tct  : _SystemTime;
  Temp : _FileTime;
begin
  FileTimeToLocalFileTime( Fd , Temp );
  FileTimeToSystemTime( Temp , Tct );
  Result :=SystemTimeToDateTime( Tct );
end;

function GetFileCreateTime( FileName: string ):TDateTime;
var
  FileOp : TOFStruct ;
  FHandle : THandle ;
  FInfo : TByHandleFileInformation;

begin
  FHandle:= OpenFile(PAnsiChar(AnsiString(FileName)),FileOp ,OF_READ);
  try
    GetFileInformationByHandle(FHandle,FInfo);
    Result := FileTimeToTime(FInfo.ftCreationTime);
  finally
    _lclose(FHandle);
  end;
end;

procedure TDailyRollingFileAppender.SetLogFile(const AFileName: string);
begin
  FShortName := AFileName;
  inherited SetLogFile(GetDateFileName(AFileName, FCurrDate));
end;

procedure TDailyRollingFileAppender.SetOption(const Name, Value: string);
begin
  inherited SetOption(Name, Value);
  EnterCriticalSection(FCriticalAppender);
  try
    if (Name = RetentDaysOpt) and (Value <> '') then
    begin
      FRetentDays := StrToIntDef(Value, FRetentDays);
    end;
  finally
    LeaveCriticalSection(FCriticalAppender);
  end;
end;

procedure TDailyRollingFileAppender.SetRetentDays(const ARetentDays : Integer);
begin
  Self.FRetentDays := ARetentDays;
  DeleteOldFiles();
end;

function IsDiffDay( d1, d2: TDateTime ): boolean;
begin
  Result := (YearOf(d1)<>Yearof(d2)) or (Monthof(d1)<>MonthOf(d2)) or 
            (Dayof(d1)<>Dayof(d2));
end;    

function CombinePath( sPath1, sPath2: WideString ): WideString;
begin
  if ( sPath1 = '' ) or (sPath2 = '' ) then
    Result := sPath1 + sPath2
  else
  begin   
    if (Copy(sPath1, Length(sPath1),1 ) <> '\') and
       (Copy(sPath2, 1, 1 ) <> '\') then
      Result := sPath1 + '\' + sPath2
    else
      Result := sPath1 + sPath2;
  end;
end;

procedure ListFiles(sPath, sExt: string; slstFile: TStringList);
var
  SR : TSearchRec;
  nFound : integer;
begin
  nFound := FindFirst( CombinePath( sPath, sExt ), faAnyFile, SR );
  while nFound = 0 do
  begin
    if ( SR.Attr and faDirectory ) <> faDirectory then
      slstFile.Add( CombinePath( sPath, SR.Name ) );

    nfound := FindNext( SR );
  end;

  SysUtils.FindClose( SR );
end;

procedure TDailyRollingFileAppender.DoAppend(const AEvent : TLogEvent);
begin
  if IsDiffDay( FCurrDate, Now ) then
    Self.RollOver;
    
  LogLog.debug('TRollingFileAppender#Append');
  inherited DoAppend(Aevent);
end;

procedure TDailyRollingFileAppender.DeleteOldFiles();
var
  slstFile: TStringList;
  OldDateFormat, AFileName, DateStr, ShortFileName: String;
  i: Integer;
  ALogDate: TDateTime;
begin  
  if (Self.FRetentDays <= 0)then
    Exit;

  slstFile := TStringList.Create;
  try
    ListFiles( ExtractFilePath(FileName), '*.*', slstFile );
    OldDateFormat := ShortDateFormat;
    ShortDateFormat := C_FILE_DATEFORMAT;
    ShortFileName := ExtractFileName(FileName);
    try
      for i := 0 to slstFile.Count - 1 do
      begin
        AFileName := ExtractFileName(slstFile[i]);
        if Pos(FileName, AFileName) > 0 then
        begin
          //去掉前缀，包括.
          DateStr := Copy( AFileName, Length(ShortFileName)+2, length(AFileName));
          if Length(DateStr) <> Length(ShortDateFormat) then
            Continue;              

          ALogDate := StrToDateDef( DateStr, Now );
          if DaysBetween( Now, ALogDate ) > FRetentDays then
            DeleteFile( AFileName ); 
        end;
      end;
    finally
      ShortDateFormat := OldDateFormat;
    end;
  finally
    slstFile.Free;
  end;  
end;  

procedure TDailyRollingFileAppender.RollOver();
begin
  LogLog.debug('TRollingFileAppender#RollOver: Start');

  LogLog.debug('TRollingFileAppender#RollOver: RetentDays > 0');
  // Delete the oldest file if it exists
  DeleteOldFiles();
  LogLog.debug('TRollingFileAppender#RollOver: deleted old file');
                           
  // Rename fileName to fileName.yyyy-mm-dd
//  Self.FileStream.Free;
//  Self.FFileStream := Nil;
  CloseLogFile;
  LogLog.debug('TRollingFileAppender#RollOver: freed the filestream');

  // reinitialize the file.
  FCurrDate := Now;
  SetLogFile(FShortName);
  LogLog.debug('TRollingFileAppender#RollOver: End');
end;

function TDailyRollingFileAppender.GetDayOut() : Int64;
begin
  Result := Self.FRetentDays;
end;

end.

