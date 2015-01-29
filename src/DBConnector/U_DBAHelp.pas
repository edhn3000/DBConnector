unit U_DBAHelp;

interface

uses
  Classes, Windows, SysUtils;

type
  TDBAHelpScriptOption = (dhsoNotRun, dhsoNotOpenTab);
  TDBAHelpScriptOptions = set of TDBAHelpScriptOption;
  TDBAHelpItem = class;
  TDBAHelpRunProc = procedure(dbaItem: TDBAHelpItem) of object;

  TDBAHelpItem = class(TCollectionItem)
  public
    FileName: string;
    Name: string;
    SQL: TStrings;
    Option: TDBAHelpScriptOptions;

  public
    RunProc: TDBAHelpRunProc;
    procedure Run;
    function Load(FileName: string): Boolean;

    constructor Create(Collection: TCollection);override;
    destructor Destroy;override;
  end;

  TDBAHelpList = class(TCollection)
  private
    function GetItem(index: Integer): TDBAHelpItem;
    procedure SetItem(index: Integer; value: TDBAHelpItem);
  public
    property Items[index: Integer]: TDBAHelpItem read GetItem write SetItem;
    procedure Load(sDir: string);
  end;

implementation

procedure FindFiles(sPath, sFileExt: string; AStrs: TStrings;
    SubDir: Boolean;ShowDir: Boolean);
var
  FSrchRec{, DSrchRec}: TSearchRec;
  nFindResult: Integer;
  nSearchMode: Integer;
  function GetValidPath(APath: string): string;
  begin
    Result := APath;  
    if Copy(Result, Length(Result), 1) <> '\' then
      Result := Result+'\';
  end;  
begin
  // 判断路径尾是否合法 路径尾不是\不会被识别为路径
  sPath := GetValidPath(sPath);
  nSearchMode := faAnyFile + faDirectory;

  nFindResult := FindFirst( sPath + sFileExt, nSearchMode, FSrchRec);
  try
    while nFindResult = 0 do
    begin
      if (FSrchRec.Name <> '.') and (FSrchRec.Name <> '..') then
      begin
        if (FSrchRec.Attr and faDirectory) = faDirectory then
        begin
          if ShowDir then     // 是否显示目录
//            AStrs.Add(sPath + GetValidPath(FSrchRec.Name));
            AStrs.Add(sPath + GetValidPath(FSrchRec.Name));
          if SubDir then      // 是否搜索子目录模式
            FindFiles(sPath + FSrchRec.Name, sFileExt, AStrs, SubDir, ShowDir);
        end
        else
//          AStrs.Add(sPath + FSrchRec.Name);
          AStrs.Add(sPath + FSrchRec.Name);
      end;

      nFindResult := FindNext(FSrchRec);
    end;
  finally
    FindClose(FSrchRec);
  end;
end;

{ TDBAHelpItem }

constructor TDBAHelpItem.Create(Collection: TCollection);
begin
  inherited;
  SQL := TStringList.Create;
end;

destructor TDBAHelpItem.Destroy;
begin
  SQL.Free;
  inherited;
end;

function TDBAHelpItem.Load(FileName: string): Boolean;
var
  F: TextFile;
  sLine, sTrimLine: string;
  sFlag: string;
  nLine: Integer;
begin
  Result := False;
  if not FileExists(FileName) then
    Exit;
  AssignFile(F, FileName);
  try
    SQL.Clear;         
    Reset(F);
    nLine := 1;
//    Readln(F, sLine);
//    sLine := Trim(sLine);
    option := [];
    while not Eof(F) do
    begin
      Readln(F, sLine);
      sTrimLine := Trim(sLine);
      if (Pos('--', sTrimLine) = 1) then
      begin
        // get file head params
        sFlag := Trim(Copy(sTrimLine, 3, Length(sTrimLine)));
        if nLine = 1 then
          Name := sFlag
        else if SameText(sFlag, 'NotRun') then
          option := option + [dhsoNotRun]
        else if SameText(sFlag, 'NotOpenTab') then
          Option := option + [dhsoNotOpenTab]
      end
      else
      begin
        // 非开关注释，即为有效语句，添加到SQL中  
        SQL.Add(sLine);
      end;
      Inc(nLine);
    end;
  finally
    CloseFile(F);
  end;
end;

procedure TDBAHelpItem.Run;
begin
  if Assigned(RunProc) then
    RunProc(Self);
end;

{ TDBAHelpList }

function TDBAHelpList.GetItem(index: Integer): TDBAHelpItem;
begin
  Result := inherited GetItem(index) as TDBAHelpItem;
end;

procedure TDBAHelpList.Load(sDir: string);
var
  slstFiles: TStrings;
  i: Integer;
  item: TDBAHelpItem;
begin
  Self.Clear;
  slstFiles := TStringList.Create;
  try
    FindFiles(sDir, '*.sql', slstFiles, False, False);
    (slstFiles as TStringList).Sort;
    for i := 0 to slstFiles.Count - 1 do
    begin
      item := Self.Add as TDBAHelpItem;
      item.Load(slstFiles[i]);
    end;
  finally
    slstFiles.Free;
  end;
end;

procedure TDBAHelpList.SetItem(index: Integer; value: TDBAHelpItem);
begin
  inherited SetItem(index, value);
end;

end.
