{**
 * <p>Title: U_Plugin  </p>
 * <p>Copyright: Copyright (c) 2008/10/20</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 插件单元 </p>
 *}
unit U_Plugin;

interface
uses
  Classes, Windows, SysUtils;

type
  TPluginItem = class(TCollectionItem)
  public
    FileName: string;
    PluginName: string;
    ShowCmd: Integer;
    WaitFor: Boolean;

    // run 之后赋值
    processInfo: TProcessInformation;
  public
    function Run:Boolean;
    function ClosePI: Boolean;

    function ReadPlugin(sDir: string): Boolean;

    class function IsPluginDir(sDir: string; var sFileName: string): Boolean;
  end;

  TPluginList = class(TCollection)
  private
    function GetItem(index: Integer): TPluginItem;
    procedure SetItem(index: Integer; value: TPluginItem);
  public
    property Items[index: Integer]: TPluginItem read GetItem write SetItem;
    procedure LoadPlugins(sDir: string);
    function FindPluginByName(sPlugName: string): TPluginItem;
  end;


implementation

{ TPluginItem }

uses
  U_fStrUtil, IniFiles;

procedure FindDirs(sPath: string; AStrs: TStrings; SubDir: Boolean = False);
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
  nSearchMode := faDirectory;

  nFindResult := FindFirst( sPath + '*.*', nSearchMode, FSrchRec);
  try
    while nFindResult = 0 do
    begin
      if (FSrchRec.Name <> '.') and (FSrchRec.Name <> '..') then
      begin
        if (FSrchRec.Attr and faDirectory) = faDirectory then
        begin
          AStrs.Add(sPath + GetValidPath(FSrchRec.Name));
          if SubDir then      // 是否搜索子目录模式
            FindDirs(sPath + FSrchRec.Name, AStrs, SubDir);
        end;
      end;

      nFindResult := FindNext(FSrchRec);
    end;
  finally
    FindClose(FSrchRec);
  end;
end;

function TPluginItem.ClosePI: Boolean;
begin
  Result := CloseHandle(processInfo.hProcess)
            and CloseHandle(processInfo.hThread);
end;

function TPluginItem.ReadPlugin(sDir: string): Boolean;
var
  ini: TIniFile;
  sPluginFile: string;
begin
  Result := False;
  if IsPluginDir(sDir, FileName) then
  begin
    Result := True;
    ShowCmd := SW_SHOWNORMAL;
    WaitFor := False;
    PluginName := ChangeFileExt(ExtractFileName(FileName), '');

    sPluginFile := ExtractFilePath(FileName) + 'plugin.ini';
    if FileExists(sPluginFile) then
    begin
      ini := TIniFile.Create(sPluginFile);
      try
        PluginName := ini.ReadString('Option', 'PluginName', PluginName);
        ShowCmd := ini.ReadInteger('Option', 'ShowCmd', ShowCmd);
        WaitFor := ini.ReadBool('Option', 'WaitFor', WaitFor);
      finally
        ini.Free;
      end;   
    end;  
  end;
end;

function TPluginItem.Run: Boolean;
var
  si: TStartupInfo;
begin  
  FillMemory(@si, sizeof(si), 0);

  with si do
  begin
    cb := sizeof(si);
    dwFlags := STARTF_USESHOWWINDOW;
    wShowWindow := ShowCmd;
  end;

  Result := Boolean(CreateProcess(nil, PChar(FileName), nil, nil, False,
                                  NORMAL_PRIORITY_CLASS, nil, nil, si, processInfo));
  if Result and WaitFor then
    WaitForSingleObject(processInfo.hProcess, INFINITE);
end;  

class function TPluginItem.IsPluginDir(sDir: string; var sFileName: string): Boolean;
var
  n: Integer;
  sDirName, sFullDir: string;
begin
  Result := False;  
  if sDir = '' then
    Exit;
  // 存在目录同名exe则是plugin
  sDirName := sDir;
  if Copy(sDirName, Length(sDirName), 1) = '\' then
    sDirName := Copy(sDirName, 1, Length(sDirName)-1);
  n := fStrUtil.PosLast('\', sDirName);
  if n <> 0 then
    sDirName := Copy(sDirName, n+1, MaxInt);

  sFullDir := sDir;
  if Copy(sFullDir, Length(sFullDir), 1) <> '\' then
    sFullDir := sFullDir + '\';
  if FileExists(sFullDir + sDirName + '.exe') then
  begin
    sFileName := sFullDir + sDirName + '.exe';
    Result := True;
  end;
end;

{ TPluginList }

function TPluginList.FindPluginByName(sPlugName: string): TPluginItem;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Count - 1 do
  begin
    if Items[i].PluginName = sPlugName then
    begin
      Result := Items[i];
      Break;
    end;  
  end;  
end;

function TPluginList.GetItem(index: Integer): TPluginItem;
begin
  Result := inherited GetItem(index) as TPluginItem;
end;

procedure TPluginList.LoadPlugins(sDir: string);
var
  slst: TStrings;
  i: Integer;
  plugin: TPluginItem;
begin
  slst := TStringList.Create;
  try
    FindDirs(sDir, slst, False);
    for i := 0 to slst.Count - 1 do
    begin
      plugin := Add as TPluginItem;
      if not plugin.ReadPlugin(slst[i]) then
        Delete(plugin.Index);
    end;  
  finally
    slst.Free;
  end;
end;

procedure TPluginList.SetItem(index: Integer; value: TPluginItem);
begin
  inherited SetItem(index, value);
end;

end.
