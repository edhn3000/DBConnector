{**
 * <p>Title: TIniOperator </p>
 * <p>Copyright: Copyright (c) 2006/9/7</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: ini读写类, 构造函数传入文件名和读取的段  </p>
 *}
unit U_Ini;

interface

uses
  IniFiles, Classes;

type
  { TIniOperator }
  TIniOperator = class(TIniFile)
  private

  protected

  public

    destructor Destroy; override; // 已在finalization调用

    function LoadSettings: Boolean;
    function SaveSettings: Boolean;
                                                                 
    function ReadBool(const Section, Ident: string; Default: Boolean): Boolean; override;
    procedure WriteBool(const Section, Ident: string; Value: Boolean); override;
    procedure ReadSectionFromIni(sIniFile, sSection: string; AList: TStrings);
  end;

  function ReadStringFromIni(sIniFile, sSection, sKey, sValueDef: string):string;
  procedure WriteStringToIni(sIniFile, sSection, sKey, sValueToWrite:string);
  procedure ReadSectionFromIni(sIniFile, sSection: string; AList: TStrings);
  procedure ReadSectionsFromIni(sIniFile: string; AList: TStrings);

  // 获得ini操作对象
  function IniOperator: TIniOperator;

implementation

uses
  SysUtils, Forms;

var
  FIniOperator: TIniOperator;

const
  C_sIniFileDef = 'Options.ini';

  // 真假设置
  C_sTrue     = 'True';
  C_sFalse    = 'False';
  C_sTrueNum  = '1';
  C_sFalseNum = '0';

function IniOperator: TIniOperator;
var
  sExeName: string;
begin
  sExeName := ExtractFileName(Application.ExeName);
  sExeName := Copy(sExeName, 1, Length(sExeName)-Length(ExtractFileExt(Application.ExeName)));

  if not Assigned(FIniOperator) then
  begin
    FIniOperator := TIniOperator.Create(ExtractFilePath(Application.ExeName) +
       sExeName + '.ini');
    FIniOperator.LoadSettings;
  end;
  Result := FIniOperator;
end;

function ReadStringFromIni(sIniFile, sSection, sKey, sValueDef: string):string;
begin
  with TIniFile.Create( sIniFile ) do
  begin
    try
      Result := ReadString(sSection,sKey,sValueDef);
    finally
      Free;
    end;
  end;
end;

procedure WriteStringToIni(sIniFile, sSection, sKey, sValueToWrite:string);
begin
  with TIniFile.Create( sIniFile ) do
  begin
    try
      WriteString(sSection,sKey,sValueToWrite);
    finally
      Free;
    end;
  end;
end;

procedure ReadSectionFromIni(sIniFile, sSection: string; AList: TStrings);
begin
  with TIniFile.Create( sIniFile ) do
  begin
    try
      ReadSection(sSection,AList);
    finally
      Free;
    end;
  end;
end;

procedure ReadSectionsFromIni(sIniFile: string; AList: TStrings);
begin
  with TIniFile.Create( sIniFile ) do
  begin
    try
      ReadSections(AList);
    finally
      Free;
    end;
  end;
end;


{ TIniOperator }

destructor TIniOperator.Destroy;
begin
  inherited;
end;

function GetBool(s: string): Boolean;
begin
  if s = C_sTrue then
    Result := True
  else
    Result := False;
end;

function GetStr(b: Boolean): string;
begin
  if b then
    Result := C_sTrue
  else
    Result := C_sFalse;
end;

function TIniOperator.ReadBool(const Section, Ident: string; Default: Boolean): Boolean;
begin
  if Default then
  begin
    Result := SameText(C_sTrue, ReadString(Section, Ident, C_sTrue)) or
              SameText(C_sTrueNum, ReadString(Section, Ident, C_sTrueNum));
  end
  else
  begin
    Result := not SameText(C_sFalse, ReadString(Section, Ident, C_sFalse));
  end;
end;

procedure TIniOperator.WriteBool(const Section, Ident: string; Value: Boolean);
begin
  if Value then
    WriteString(Section, Ident, C_sTrue)
  else
    WriteString(Section, Ident, C_sFalse);
end;

procedure TIniOperator.ReadSectionFromIni(sIniFile, sSection: string;
  AList: TStrings);
var
  i: Integer;
begin
  ReadSection(sSection, AList);
  for i := 0 to AList.Count - 1 do
  begin
    AList[i] := AList[i] + '=' + ReadString(sSection, AList[i], '');
  end;
end;

function TIniOperator.LoadSettings: Boolean;
begin
  Result := True;
  try

  except
    Result := False;
  end;
end;

function TIniOperator.SaveSettings: Boolean;
begin
  Result := True;
  try

  except
    Result := False;
  end;
end;

initialization

finalization
  if Assigned(FIniOperator) then
    FIniOperator.Free;

end.
