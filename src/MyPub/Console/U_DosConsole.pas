{* U_DosConsole
 * 命令窗口控制单元
 *}
unit U_DosConsole;

interface
uses
  Windows, SysUtils, Dialogs, Classes, Forms;

type
  { 命令窗口控制类 实现未完全 }
  TDOSConsole = class
  private
    FReadIn, FReadOut: THandle;
    FWriteIn, FWriteOut: THandle;    
    FSecurity: TSecurityAttributes;
    Fstart: TStartUpInfo;
    FProcessInfo:TProcessInformation;
    FReadBuffer: Integer;
  public
    property WriteIn: THandle read FWriteIn;  
    property ReadOut: THandle read FReadOut;
  public
    constructor Create;
    procedure InitConsole;
    procedure FinalConsole;
    procedure WriteToPipe(Value: string);
    function ReadFromPipe(): string;
  end;

  // 执行命名，并将命令结果加载到ResultList中
  procedure RunDosCommand(const DosCmd: string; ResultList: TStrings);

implementation

{ TDOSConsole }

constructor TDOSConsole.Create;
begin
  FReadBuffer := 2000;
end;

procedure TDOSConsole.InitConsole;
begin 
  with FSecurity do begin
    nlength := SizeOf(TSecurityAttributes);
    binherithandle := true;
    lpsecuritydescriptor := nil;
  end;
  //
  Createpipe(FReadOut, FWriteOut, @FSecurity, 0);
  Createpipe(FReadIn, FWriteIn, @FSecurity, 0);

  with FSecurity do begin
    nlength := SizeOf(TSecurityAttributes); 
    binherithandle := true;
    lpsecuritydescriptor := nil;
  end; 

  FillChar(FStart, Sizeof(FStart), #0);
  Fstart.cb := SizeOf(Fstart);
  Fstart.hStdOutput := FWriteOut;
  Fstart.hStdInput := FReadIn;
  Fstart.hStdError := FWriteOut;
  Fstart.dwFlags := STARTF_USESTDHANDLES +
    STARTF_USESHOWWINDOW;
  Fstart.wShowWindow := SW_HIDE;
end; 

procedure TDOSConsole.FinalConsole;
begin
  CloseHandle(FWriteIn);
  CloseHandle(FWriteOut);
    
  CloseHandle(FReadIn);
  CloseHandle(FReadOut);

  CloseHandle(FProcessInfo.hProcess);
  CloseHandle(FProcessInfo.hThread);
end;

function TDOSConsole.ReadFromPipe(): string;
var 
  Buffer: PAnsiChar;
  BytesRead: DWord;
begin 
  Result := '';
  if GetFileSize(FReadOut, nil) = 0 then
  begin
//    ShowMessage('GetFileSize(FReadOut, nil) = 0');
    Exit;
  end;

  Buffer := AllocMem(FReadBuffer + 1);
  repeat
    BytesRead := 0;
    ReadFile(FReadOut, Buffer[0],
      FReadBuffer, BytesRead, nil);
    if BytesRead > 0 then begin
      Buffer[BytesRead] := #0;
      OemToAnsi(Buffer, Buffer);
      Result := string(Buffer);
    end;
  until (BytesRead < Cardinal(FReadBuffer));
  FreeMem(Buffer);
end;

procedure TDOSConsole.WriteToPipe(Value: string);
var 
  len: integer;
  BytesWrite: DWord;
  Buffer: PChar;
begin 
  CreateProcess(nil,
    PChar(Value),
    @FSecurity,
    @FSecurity,
    true,
    NORMAL_PRIORITY_CLASS,
    nil,
    nil,
    Fstart,
    FProcessInfo);

  len := Length(Value) + 1;
  Buffer := PChar(Value + #10);
  WriteFile(FWriteIn, Buffer[0], len, BytesWrite, nil);
  WaitForSingleObject(FProcessInfo.hProcess,INFINITE);
end;

procedure RunDosCommand(const DosCmd:string; ResultList: TStrings);
const
  ReadBuffer=2400;  //设置ReadBuffer的大小
var
  Security:TSecurityAttributes;
  ReadPipe,WritePipe:THandle;
  start:TStartUpInfo;
  ProcessInfo:TProcessInformation;
  Buffer:PAnsiChar;
  BytesRead:DWord;
  Buf:string;
begin
  with Security do
  begin
    nlength:=sizeof(TSecurityAttributes);
    binherithandle:=true;
    lpsecuritydescriptor:=nil;
  end;
  {创建一个命名管道来捕获Console的输出}
  if CreatePipe(ReadPipe,WritePipe,@Security,0) then
  begin
    Buffer:=AllocMem(ReadBuffer+1);
    FillChar(Start,Sizeof(Start),#0);
    {设置Console程序的启动属性}
    with Start do
    begin
      cb:=sizeof(start);
      start.lpReserved:=nil;
      lpDesktop:=nil;
      lpTitle:=nil;
      dwX:=0;
      dwY:=0;
      dwXSize:=0;
      dwYSize:=0;
      dwXCountChars:=0;
      dwYCountChars:=0;
      dwFillAttribute:=0;
      hStdOutput:=WritePipe; //将输出定向到建立的WritePipe上
      hStdInput:=ReadPipe;   //将输入定向到建立的ReadPipe上
      hStdError:=WritePipe;  //将错误输出定向到建立的WritePipe上
      dwFlags:=STARTF_USESTDHANDLES or STARTF_USESHOWWINDOW;
      wshowwindow:=SW_HIDE;  //设置窗口为hide
    end;

    
    try
      {创建一个子进程，运行Console}
      if CreateProcess(nil,PChar(DosCmd),@Security,@Security,true,
        NORMAL_PRIORITY_CLASS,
        nil,nil,start,ProcessInfo)then
      begin
        Application.ProcessMessages;
        {等待进程运行结束}
        WaitForSingleObject(ProcessInfo.hProcess,INFINITE);
        {关闭程序...开始没有关掉它，如果没有输出的话，程序死掉了。}
        closeHandle(WritePipe);
        buf:='';
        {读取Console的输出}
        repeat
          bytesRead:=0;
          ReadFile(ReadPipe,Buffer[0],ReadBuffer,BytesRead,nil);
          Buffer[BytesRead]:=#0;
          OemToAnsi(Buffer,Buffer);
          Buf:=Buf+String(Buffer);
          Application.ProcessMessages;
        until(bytesRead<ReadBuffer);
        //sendDebug(Buf);

        {按照换行符进行切割，并在Memo中显示出来}
        while pos(#10,buf)>0 do
        begin
          ResultList.Add(Copy(Buf,1,Pos(#10,buf)-1));
          Delete(Buf,1,Pos(#10,buf));
          Application.ProcessMessages;
        end;
      end;
    finally
      FreeMem(Buffer);
      CloseHandle(ProcessInfo.hProcess);
      CloseHandle(ProcessInfo.hThread);
      CloseHandle(ReadPipe);
    end;
  end;
end;

end.
