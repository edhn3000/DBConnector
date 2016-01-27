{**
 * <p>Title: U_TextFileWriter  </p>
 * <p>Copyright: Copyright (c) 2012/3/21</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: �ı��ļ������� </p>
 * <p>Description: ��װTextFileʵ�֣���������� </p>  
 * <p>Description: ��Ҫ���ڽ��ʹ��list�������ı�ʱ�ڴ�ռ�ô����� </p> 
 *}
unit U_TextFileWriter;

interface

uses
  Windows;

type
  { TTextFileWriter �ɶ�д�ı�������Ӧ��дģʽ��
    ��ֻ����һ��ģʽ�����˾Ͳ���д��д�˾Ͳ��ܶ� }
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

    // ��ȡһ�У����Զ��л�Ϊ��ģʽ
    function ReadLine(): string;
    // ��ǰ���Ƿ��ļ���β
    function Eof(): Boolean;
    // д��һ�У����Զ��л�Ϊдģʽ
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
      DeleteFile(FFileName);  // ����д����ɾ���ļ�
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
  // д�ļ�����ʱ���Դ��ڵ��ļ�������ݡ�
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
