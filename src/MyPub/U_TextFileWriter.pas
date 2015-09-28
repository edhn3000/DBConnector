{**
 * <p>Title: U_TextFileWriter  </p>
 * <p>Copyright: Copyright (c) 2012/3/21</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 文本文件操作类 </p>
 * <p>Description: 封装TextFile实现，更面向对象 </p>  
 * <p>Description: 主要用于解决使用list操作大文本时内存占用大问题 </p> 
 *}
unit U_TextFileWriter;

interface

uses
  Windows;

type
  { TTextFileWriter 可读写文本，自适应读写模式，
    但只能有一种模式，读了就不能写，写了就不能读 }
  TTextFileWriter = class
  private
    FFileName: string;
    FAppend: Boolean;
    FReadTextFile: TextFile;
    FHasReadInit: Boolean;
    FMode: Integer;
  protected
    procedure InitForWrite;  
    procedure InitForRead;
  public
    constructor Create(sFileName: string; bAppend: boolean = False);
    destructor Destroy;override;

    // 读取一行，并自动切换为读模式
    function ReadLine(): string;
    // 当前行是否文件结尾
    function Eof(): Boolean;
    // 写入一行，并自动切换为写模式
    procedure WriteLine(s: string);

    function GetFileName(): string;
  end;


implementation

uses
  SysUtils;  

const
  TEXTFILEWRITER_MODE_NONE = 0;
  TEXTFILEWRITER_MODE_READ = 1;
  TEXTFILEWRITER_MODE_WRITE = 2;

{ TTextFileWriter }

constructor TTextFileWriter.Create(sFileName: string; bAppend: boolean);
begin
  FFileName := sFileName;
  FAppend := bAppend;
  FMode := TEXTFILEWRITER_MODE_NONE;
end;

destructor TTextFileWriter.Destroy;
begin
  if FMode = TEXTFILEWRITER_MODE_READ then
  begin
    if FHasReadInit then
      Closefile(FReadTextFile);
  end;
  inherited;
end;

function TTextFileWriter.Eof: Boolean;
begin
  Result := System.Eof(FReadTextFile);
end;

function TTextFileWriter.GetFileName: string;
begin
  Result := FFileName;
end;

procedure TTextFileWriter.InitForRead;
begin
  FMode := TEXTFILEWRITER_MODE_READ;
  FHasReadInit := False;
  if FileExists(FFileName) then
  begin
    AssignFile(FReadTextFile, FFileName);
    Reset(FReadTextFile);
    FHasReadInit := True;
  end;
end;

procedure TTextFileWriter.InitForWrite;
begin
  FMode := TEXTFILEWRITER_MODE_WRITE;
  if FileExists(FFileName) then
  begin
    if not FAppend then
    begin
      DeleteFile(FFileName);  // 不续写，就删除文件
    end;
  end;
end;

function TTextFileWriter.ReadLine: string;
var
  sLine: string;
begin      
  if FMode = TEXTFILEWRITER_MODE_NONE then
    InitForRead;
  if not FHasReadInit then
  begin
    Result := '';
    Exit;
  end;
  Readln(FReadTextFile, sLine);
  Result := sLine;
end;

procedure TTextFileWriter.WriteLine(s: string);
var
  F: TextFile;
begin
  if FMode = TEXTFILEWRITER_MODE_NONE then
    InitForWrite;
  AssignFile(F, FFileName);
  // 写文件内容时，对存在的文件添加内容。
  if FileExists(FFileName) then
    Append(F)
  else
    ReWrite(F);
  try
    Writeln(F, s);
    Flush(F);
  finally
    CloseFile(F);
  end;
end;


end.
