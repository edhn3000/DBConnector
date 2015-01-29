unit U_FtpTrans;

interface

uses
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdFTPCommon,
  IdFTP, SysUtils, Classes;

type
  TFtpTrans = class
  private
    FFTP: TIdFTP;
    FLastError: string;
    function GetConnected: Boolean;
//    procedure SetFFtp(value: TIdFTP);
  public
    property FTP: TIdFTP read FFTP;   
    property Connected: Boolean read GetConnected;
    property LastError: string read FLastError;

  public
    constructor Create(bPassive: Boolean = true);
    destructor Destroy;override;
    function Connect(sHost, sUser, sPass: string; sDir: string = '';
      nPort: Integer = 0):Boolean;
    function DisConnect: Boolean;
    function Test(sHost, sUser, sPass: string; sDir: string = '';
      nPort: Integer = 0):Boolean;
    function UpFile(sSourceFile: string; sDestFile: string = ''): Boolean;
//    function DeleteFile()
    function DownFile(sFtpFile, sLocalDir: string): Boolean;
    procedure ChangeDir(sDir: string);
  end;

implementation

uses
  Forms;

{ TFtpTrans }

constructor TFtpTrans.Create(bPassive: Boolean);
begin
  FFTP := TIdFTP.Create(nil);
  FFTP.Passive := bPassive;
end;

destructor TFtpTrans.Destroy;
begin
  FFTP.Free;
  inherited;
end;

//procedure TFtpTrans.SetFFtp(value: TIdFTP);
//begin
//  FFTP := value;
//end;  

function TFtpTrans.Connect(sHost, sUser, sPass: string; sDir: string;
  nPort: Integer): Boolean;
begin
  try
    if FFTP.Connected then
      FFTP.Disconnect;
    FFTP.Host := sHost;
    if nPort <> 0 then
      FFTP.Port := nPort;
    FFTP.Username := sUser;
    FFTP.Password := sPass;
    FFTP.Connect(True);
    if sDir <> '' then
      FFTP.ChangeDir(sDir);
    Result := True;
  except
    on ex: Exception do
    begin
      Result := False;
      FLastError := ex.Message;
    end;
  end;
end; 

function TFtpTrans.DisConnect: Boolean;
begin 
  try
    FFTP.Disconnect;  
    Result := True;
  except
    on ex: Exception do
    begin
      Result := False;
      FLastError := ex.Message;
    end;
  end;
end;

function TFtpTrans.UpFile(sSourceFile: string; sDestFile: string = ''): Boolean;
var
  sTrueDestFile: string;
begin
  try
    FFTP.TransferType := ftBinary;
    if sDestFile = '' then
      sTrueDestFile := ExtractFileName(sSourceFile)
    else   
      sTrueDestFile := sDestFile;
    FFTP.Put(sSourceFile, sTrueDestFile, False);
    Result := True;
  except
    on ex: Exception do
    begin
      Result := False;
      FLastError := ex.Message;
    end;
  end;
end;

function TFtpTrans.DownFile(sFtpFile, sLocalDir: string): Boolean;
begin
  try
    FFTP.TransferType := ftBinary;
    FFTP.Get(sFtpFile, IncludeTrailingPathDelimiter(sLocalDir)+sFtpFile, False);
    Result := True;
  except
    on ex: Exception do
    begin
      Result := False;
      FLastError := ex.Message;
    end;
  end;
end;

function TFtpTrans.GetConnected: Boolean;
begin
  Result := FFTP.Connected;
end;

procedure TFtpTrans.ChangeDir(sDir: string);
begin
  FFTP.ChangeDir(sDir);
end;

function TFtpTrans.Test(sHost, sUser, sPass, sDir: string;
  nPort: Integer): Boolean;
var
  slstTest: TStrings;
  sFile: string;
begin
  Result := False;
  if Connect(sHost, sUser, sPass, sDir, nPort) then
  begin
    sFile := ExtractFilePath(Application.ExeName) + 'testFtp';
    slstTest := TStringList.Create;
    try
      slstTest.Add('this file is for testing ftp');
      slstTest.SaveToFile(sFile);
    finally
      slstTest.Free;
    end;
    // 连接成功并可上传下载文件才算测试完成
    if UpFile(sFile) then
    begin
      FFTP.Delete('testFtp'); 
      Result := True;
    end;      
    ftp.DisConnect;
    DeleteFile(sFile);
  end
end;

initialization

finalization

end.
