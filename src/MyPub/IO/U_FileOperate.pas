unit U_FileOperate;

interface


uses
  Windows, Classes, SysUtils, Generics.Collections, Generics.Defaults;

type
  TFileReader = class
  private
    FFileName: String;
    FTxtFile: TextFile;

  public
    constructor Create(fileName: string);
    destructor Destroy;override;
//
    procedure ResetFile;
    function IsEnd: Boolean;
    function ReadLine(): String;
  end;

implementation

{ TFileReader }

constructor TFileReader.Create(fileName: string);
begin
  FFileName := fileName;
  if not FileExists(FFileName) then begin
    FileCreate(FFileName);
  end;
  AssignFile(FTxtFile, FFileName);
end;

destructor TFileReader.Destroy;
begin
  CloseFile(FTxtFile);
  inherited;
end;

function TFileReader.IsEnd: Boolean;
begin
  Result := Eof(FTxtFile);
end;

function TFileReader.ReadLine(): String;
var
  s: String;
begin
  s := '';
  if not Eof(FTxtFile) then
    Readln(FTxtFile, s);
  Result := s;
end;

procedure TFileReader.ResetFile;
begin
  Reset(FTxtFile);
end;

end.
