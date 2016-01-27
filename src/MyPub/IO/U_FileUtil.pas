{**
 * <p>Title: U_FileUtil  </p>
 * <p>Copyright: Copyright (c) 2006/9/19</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 可查找指定目录下指定扩展名的文件,多个通用方法 </p>
 * <p>update: 2008-06-10 FindFiles方法bug修改，改正传入非*.*扩展名导致无法查找子目录的错误，
             同时传入的扩展名格式也改变了，仅传入.后面的内容 </p>
 *}
unit U_FileUtil;

interface

{$WARN SYMBOL_PLATFORM OFF}

uses
  Windows, Classes, SysUtils, Generics.Collections, Generics.Defaults,
  ShellAPI, ShlObj, Forms, ActiveX;

type
  TTextFormat=(tfAnsi,tfUnicode,tfUnicodeBigEndian,tfUtf8);
const
  TextFormatFlag:array[tfAnsi..tfUtf8] of word=($0000,$FFFE,$FEFF,$EFBB);

type
{ TFileFullInfo }
  TFileFullInfo = record
    FileName: string;
    FileVersion: string;
    CompanyName: string;
    FileDescription: string;
    InternalName: string;
    LegalCopyright: string;
    LegalTradeMarks: string;
    OriginalFilename: string;
    ProductName: string;
    ProductVersion: string;
    Comments: string;
    FileSize: LongInt;
  end;

  { TFileItem }
  TFileItem = class
  private
    FSubFileItems: TStrings;
  public
    FileName: String;
    CreateTime: TDateTime;
    ModifyTime: TDateTime;
    AccessTime: TDateTime;

    IsFile: Boolean;

    property SubFileItems: TStrings read FSubFileItems;

    constructor Create;
    destructor Destroy;override;

    function HasSubFile(): Boolean;
  end;

  { TFindFileCallBack }
  TOnFileSearching = procedure (fileItem: TFileItem; var CanSearch: Boolean) of object;

  { TFileList }
  TFileList = class
  private
    FFileList: TObjectList<TFileItem>;
    FFileMap: TDictionary<String,TFileItem>;
    FSearchMode: Integer;
    FSearchExts: TDictionary<String,String>;
    FSearchSubDir: Boolean;
    FContainSelf: Boolean;
    FbSearching: Boolean;

    FExcludeFiles: TStrings;

    FOnSearching: TOnFileSearching;

  protected
    function AddFile(sFileName: String; SrchRec: TSearchRec): Integer;
    function MatchExt(sExt: String): Boolean;
    procedure DoFindFiles(sPath: string);
    function StrIsDir(s: string): Boolean;
    function IsExcludeFile(AFileName: string):Boolean;

  public
    { FindFiles使用的标志 }
    property SearchMode: Integer read FSearchMode;
    property SearchSubDir: Boolean read FSearchSubDir write FSearchSubDir;
    property ContainSelfApp: Boolean read FContainSelf write FContainSelf;
    property ExcludeFiles: TStrings read FExcludeFiles write FExcludeFiles;

    property OnSearching: TOnFileSearching read FOnSearching write FOnSearching;
    property Searching: Boolean read FbSearching;

    constructor Create;
    destructor Destroy;override;
{-------------------------------------------------------------------------------
  过程名:    CountDir
  参数:      sPath: string 一个路径
  返回值:    Integer 目录数量
-------------------------------------------------------------------------------}
    function CountDir(sPath: string): Integer;
{-------------------------------------------------------------------------------
  过程名:    FindFiles
  参数:      sPath: string 一个路径;  sFileExt: string 文件扩展名部分，如 .*
  返回值:    无，找到的文件添加到FFileList
-------------------------------------------------------------------------------}
    procedure FindFiles(sPath: string; sFileExt: string);
{-------------------------------------------------------------------------------
  过程名:    SortIt  名称顺序排列，且文件夹在前，文件在后
  参数:      无
  返回值:    无
-------------------------------------------------------------------------------}
    procedure SortIt;
    function Count(): Integer;
    procedure Clear();

    function GetFileName(index: Integer): string;
    function GetFileItem(index: Integer): TFileItem; overload;
    function GetFileItem(fileName: String): TFileItem; overload;
  end;

  { TFileUtil }
  TFileUtil = class
  public
    //// 一般处理
    // 连接多个目录，中间没有分隔符自动添加
    class function ConcatDir(asDirs: array of string): string;
    // 连接多个脚本中的元素
    class function ConcatScript(asDirs: array of string): string;
    // 去掉扩展名的文件名
    class function RemoveFileExt(s: string): string;
    class function ParseFileEncoding(const FileName: string): TEncoding;
    class function ParseFileEncoding2(const FileName: string): TEncoding;

    ////文件检索类
    // 查找指定目录下某扩展名的文件，默认步包含子目录，结果中不包含目录名
    class procedure FindFiles(sPath, sFileExt: string; AStrs: TStrings;
        SubDir: Boolean = False;ShowDir: Boolean = False);
    class procedure FindDirs(sPath: string; AStrs: TStrings; SubDir: Boolean = False);
    // 找到改动过的文件，未完成
    class procedure FindChangedFile(sPath: string; FileList: TFileList);

    // 拷贝文件列表到目的目录，根据bChangeDir决定拷贝完成后是否更改源文件目录为目的目录
    class function CopyListFiles(SourceFiles: TStrings;sDestDir: string;bChangeDir:
        Boolean): Boolean;
    class function CopyDirFiles(ASourceDir, ADestDir: string; CopySub: Boolean = True;
        FileExt: string = '*.*'):Boolean;

    //// 文件、目录状态检查和控制
    class function ForceDeleteDir(sDir: string): Integer;
    class function BrowseFolder(const Folder: string): string;
    // 参数2是否参数1的直接子目录
    class function IsDirectSubDir(sDirMain, sDirSub: string):Boolean;
    class function IsFileInUse(AFileName: string): Boolean;
    class function IsDirInUse(ADir: string): Boolean;

    //// 文件属性类的处理
    // 去掉只读属性
    class procedure FileRemoveReadOnly(sFile: string);
    // 获取文件的详细信息
    class function GetFileFullInfo(AFileName: string): TFileFullInfo;
    // 显示文件属性窗口
    class procedure ShowFilePropWindow(AFileName: string);
    class function GetFileMainVersion(fn: string):string;
    class function GetFileVersion(fn: string):string;//得到程序的版本号
  end;

{-------------------------------------------------------------------------------
  过程名:    SelectDirectory 选择目录窗口
  参数:      Caption: string; 显示的标题
             Root: WideString; 根目录
             Directory: string; 引用传递，可传入初始选择目录，执行后内容为用户选择的路径
             Owner: HWND;
             ShowNewButton: Boolean; 显示新建按钮
  返回值:    Boolean 选择确定返回true，选择取消返回false
-------------------------------------------------------------------------------}
  function SelectDirectory(const Caption: string; const Root: WideString;
    var Directory: string; Owner: HWND; ShowNewButton: Boolean = True): Boolean;

  // 旧式选择目录窗口
  function BrowseFolder(const Folder: string): string;

implementation

const
  C_sSLASH_SCRIPT = '/';

function ConcatStr(Separator: string; arystr: array of string): string;
var
  i: Integer;
  sConcat: string;
begin
  Result := '';
  if Length(arystr) = 0 then
    Exit;
  sConcat := '';
  Result := arystr[0];
  for i := 1 to High(arystr) do
  begin
    // 前后两个串都不是空 并且前串末尾、后串开头都不是分隔符，则加分隔符
    if ('' <> Trim(arystr[i])) and
       ('' <> Trim(Result)) and
       (Separator <> Copy(Result, Length(Result), 1)) and
       (Separator <> Copy(arystr[i], 1, 1)) then
      sConcat := Separator
    else
      sConcat := '';

    Result := Result + sConcat + arystr[i];
  end;
end;

function Split(sText, sSeparator: String; ssParts: TStrings):TStrings;
var
  nPos: Integer;
  sPart, sRemain: String;
begin
  Result := ssParts;
  if Result = nil then
    Result := TStringList.Create;
  Result.Clear;
  sRemain:= sText;
  nPos:= Pos(sSeparator, sRemain);

  while nPos > 0 do
  begin
    sPart:= Copy(sRemain, 1, nPos- 1);
    if Trim(sPart) <> '' then
      Result.Add(sPart);
    sRemain:= Copy(sRemain, nPos+ Length(sSeparator), Length(sRemain)- nPos+ 1);
    nPos:= Pos(sSeparator, sRemain);
  end;

  if sRemain <> '' then
    Result.Add(sRemain);
end;

function CovFileDate(Fd:_FileTime):TDateTime;
var
  Tct:_SystemTime;
  Temp:_FileTime;
begin
  FileTimeToLocalFileTime(Fd,Temp);
  FileTimeToSystemTime(Temp,Tct);
  Result := SystemTimeToDateTime(Tct);
end;

type
  TCompFile = class(TCustomComparer<TFileItem>)
  public
  end;

function BrowseCallbackProc(Wnd: HWND; uMsg: UINT; lParam, lpData: LPARAM): Integer stdcall;
begin
  case uMsg of
    BFFM_INITIALIZED: SendMessage(Wnd, BFFM_SETSELECTION, 1, lpData);
  end;
  Result := 0;
end;

function SelectDirCB(Wnd: HWND; uMsg: UINT; lParam, lpData: LPARAM): Integer stdcall;
begin
  if (uMsg = BFFM_INITIALIZED) and (lpData <> 0) then
    SendMessage(Wnd, BFFM_SETSELECTION, Integer(True), lpdata);
  Result := 0;
end;

function SelectDirectory(const Caption: string; const Root: WideString;
  var Directory: string; Owner: HWND; ShowNewButton: Boolean = True): Boolean;
var
  BrowseInfo: TBrowseInfo;
  Buffer: PChar;
  RootItemIDList, ItemIDList: PItemIDList;
  ShellMalloc: IMalloc;
  IDesktopFolder: IShellFolder;
  Eaten, Flags: LongWord;
begin
  Result := False;
  FillChar(BrowseInfo, SizeOf(BrowseInfo), 0);
  if (ShGetMalloc(ShellMalloc) = S_OK) and (ShellMalloc <> nil) then
  begin
    Buffer := ShellMalloc.Alloc(MAX_PATH);
    try
      SHGetDesktopFolder(IDesktopFolder);
      if Root = '' then
        RootItemIDList := nil
      else
        IDesktopFolder.ParseDisplayName(Application.Handle, nil,
          POleStr(Root), Eaten, RootItemIDList, Flags);
      with BrowseInfo do
      begin
        hwndOwner := Owner;
        pidlRoot := RootItemIDList;
        pszDisplayName := Buffer;
        lpszTitle := PChar(Caption);
        ulFlags := BIF_RETURNONLYFSDIRS;
        if ShowNewButton then
          ulFlags := ulFlags or $0040;
        lpfn := SelectDirCB;
        lparam := Integer(PChar(Directory));
      end;
      ItemIDList := SHBrowseForFolder(BrowseInfo);
      Result :=  ItemIDList <> nil;
      if Result then
      begin
        ShGetPathFromIDList(ItemIDList, Buffer);
        ShellMalloc.Free(ItemIDList);
        Directory := Buffer;
      end;
    finally
      ShellMalloc.Free(Buffer);
    end;
  end;
end;

function BrowseFolder(const Folder: string): string;
var
  TitleName: string;
  lpItemID: PItemIDList;
  BrowseInfo: TBrowseInfo;
  DisplayName: array[0..MAX_PATH] of char;
  TempPath: array[0..MAX_PATH] of char;
begin
  Result := Folder;
  FillChar(BrowseInfo, sizeof(TBrowseInfo), #0);
  BrowseInfo.hwndOwner := GetActiveWindow;
  BrowseInfo.pszDisplayName := @DisplayName;
  TitleName := '请选择一个目录';
  BrowseInfo.lpszTitle := PChar(TitleName);
  BrowseInfo.ulFlags := BIF_RETURNONLYFSDIRS or BIF_NEWDIALOGSTYLE;
  BrowseInfo.lpfn := BrowseCallbackProc;
  BrowseInfo.lParam := Integer(PChar(Folder));
  lpItemID := SHBrowseForFolder(BrowseInfo);
  // 选择了某目录
  if Assigned(lpItemId) then
  begin
    SHGetPathFromIDList(lpItemID, TempPath);
    GlobalFreePtr(lpItemID);
    Result := string(TempPath);
  end
  else
  // 选择了“取消 ”
    Result:=Folder;
end;

{ TFileUtil }

class function TFileUtil.ConcatDir(asDirs: array of string): string;
begin
  Result := ConcatStr(PathDelim, asDirs);
end;

class function TFileUtil.ConcatScript(asDirs: array of string): string;
begin
  Result := ConcatStr(C_sSLASH_SCRIPT, asDirs);
end;

class function TFileUtil.BrowseFolder(const Folder: string): string;
var
  TitleName: string;
  lpItemID: PItemIDList;
  BrowseInfo: TBrowseInfo;
  DisplayName: array[0..MAX_PATH] of char;
  TempPath: array[0..MAX_PATH] of char;
begin
  Result := Folder;
  FillChar(BrowseInfo, sizeof(TBrowseInfo), #0);
  BrowseInfo.hwndOwner := GetActiveWindow;
  BrowseInfo.pszDisplayName := @DisplayName;
  TitleName := '请选择一个目录';
  BrowseInfo.lpszTitle := PChar(TitleName);
  BrowseInfo.ulFlags := BIF_RETURNONLYFSDIRS or BIF_NEWDIALOGSTYLE;
  BrowseInfo.lpfn := BrowseCallbackProc;
  BrowseInfo.lParam := Integer(PChar(Folder));
  lpItemID := SHBrowseForFolder(BrowseInfo);
  // 选择了某目录
  if Assigned(lpItemId) then
  begin
    SHGetPathFromIDList(lpItemID, TempPath);
    GlobalFreePtr(lpItemID);
    Result := string(TempPath);
  end
  else
  // 选择了“取消 ”
    Result:=Folder;
end;

class function TFileUtil.ForceDeleteDir( sDir: String): Integer;
var
  T:TSHFileOpStruct;
begin
  Result := 0;

  if sDir = '' then
    Exit;
  if sDir[Length( sDir )] = '\' then
    sDir := Copy( sDir, 1, Length( sDir ) - 1 );

  if not DirectoryExists( sDir ) then
    Exit;

  With T do
  Begin
    Wnd := 0;
    wFunc := FO_DELETE;
    pFrom := Pchar( sDir+#0 );
    pTo := nil;
    //fFlags := FOF_ALLOWUNDO + FOF_NOCONFIRMATION + FOF_NOERRORUI; //标志表明允许恢复，无须确认并不显示出错信息]
    //标志表明允许恢复，无须确认并不显示出错信息
    fFlags := FOF_NOCONFIRMATION + FOF_NOERRORUI;
    hNameMappings := nil;
    lpszProgressTitle := '正在删除文件夹';
    fAnyOperationsAborted := False;
  End;
  Result := SHFileOperation( T );
end;

class function TFileUtil.GetFileMainVersion(fn: string):string;
var
  buf, p: pChar;
  sver: ^VS_FIXEDFILEINFO ;
  i: LongWord;
  sRes:string;
  iMajor,iMinor: Integer;
//  iRelease,iBuild:integer;
  bResult: Boolean;
begin
  sRes:='';
  iMajor := 0;
  iMinor := 0;
//  iRelease := 0;
//  iBuild := 0;

  i:= GetFileVersionInfoSize(pchar(fn), i);
  new(sver);
  p:= pchar(sver);
  GetMem(buf, i);
  ZeroMemory(buf, i);
  bResult:= false;
  if GetFileVersionInfo(pchar(fn), 0, 4096, pointer(buf)) then
    if VerQueryValue(buf, '\', pointer(sver), i) then begin
      iMajor:= sVer^.dwFileVersionMS shr 16;
      iMinor:= sver^.dwFileVersionMS and $0000ffff;
//      iRelease:= sver^.dwFileVersionLS shr 16;
//      iBuild:= sver^.dwFileVersionLS and $0000ffff;
      bResult:= true;
    end;
  Dispose(p);
  FreeMem(buf);

  if bResult then
  begin
    sRes:= IntToStr(iMajor)+'.'+IntToStr(iMinor);
//        +'.'+IntToStr(iRelease)+'.'+IntToStr(iBuild);
  end;
  Result:=sRes;
end;

class function TFileUtil.GetFileVersion(fn: string):string;
var
  buf, p: pChar;
  sver: ^VS_FIXEDFILEINFO ;
  i: LongWord;
  sRes:string;
  iMajor,iMinor,iRelease,iBuild:integer;
  bResult: Boolean;
begin
  sRes:='';
  iMajor := 0;
  iMinor := 0;
  iRelease := 0;
  iBuild := 0;

  i:= GetFileVersionInfoSize(pchar(fn), i);
  new(sver);
  p:= pchar(sver);
  GetMem(buf, i);
  ZeroMemory(buf, i);
  bResult:= false;
  if GetFileVersionInfo(pchar(fn), 0, 4096, pointer(buf)) then
    if VerQueryValue(buf, '\', pointer(sver), i) then begin
      iMajor:= sVer^.dwFileVersionMS shr 16;
      iMinor:= sver^.dwFileVersionMS and $0000ffff;
      iRelease:= sver^.dwFileVersionLS shr 16;
      iBuild:= sver^.dwFileVersionLS and $0000ffff;
      bResult:= true;
    end;
  Dispose(p);
  FreeMem(buf);

  if bResult then
  begin
    sRes:= IntToStr(iMajor)+'.'+IntToStr(iMinor)
        +'.'+IntToStr(iRelease)+'.'+IntToStr(iBuild);
  end;
  Result:=sRes;
end;

class procedure TFileUtil.FindFiles(sPath, sFileExt: string; AStrs: TStrings;
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

class procedure TFileUtil.FindDirs(sPath: string; AStrs: TStrings; SubDir: Boolean = False);
var
  FSrchRec{, DSrchRec}: TSearchRec;
  nFindResult: Integer;
  nSearchMode: Integer;
begin
  // 判断路径尾是否合法 路径尾不是\不会被识别为路径
  sPath := IncludeTrailingPathDelimiter(sPath);
  nSearchMode := faDirectory;

  nFindResult := FindFirst( sPath + '*.*', nSearchMode, FSrchRec);
  try
    while nFindResult = 0 do
    begin
      if (FSrchRec.Name <> '.') and (FSrchRec.Name <> '..') then
      begin
        if (FSrchRec.Attr and faDirectory) = faDirectory then
        begin
          AStrs.Add(sPath + IncludeTrailingPathDelimiter(FSrchRec.Name));
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

class procedure TFileUtil.FindChangedFile(sPath: string; FileList: TFileList);
var
  hndFindChange: THandle;
  dwWaitState: DWORD;
  info: BY_HANDLE_FILE_INFORMATION;
begin
  hndFindChange := FindFirstChangeNotification(PChar(sPath), True,
      FILE_NOTIFY_CHANGE_LAST_WRITE);
  if hndFindChange <> INVALID_HANDLE_VALUE then
  begin
    dwWaitState := WaitForSingleObject(hndFindChange, 1000);
    case dwWaitState of
      WAIT_TIMEOUT:
      begin
        FindCloseChangeNotification(hndFindChange);
      end;
      WAIT_OBJECT_0:
      begin
        GetFileInformationByHandle(hndFindChange, info);
        // todo: FildChangedFile方法未完成
//        info.
        if not FindNextChangeNotification(hndFindChange) then
        //
      end;
      else
      begin
        FindCloseChangeNotification(hndFindChange);
      end;
    end;
  end;
end;

class function TFileUtil.RemoveFileExt(s: string): string;
var
  i: Integer;
begin
  Result := s;
  for i := Length(Result) downto 1 do
  begin
    if Result[i] = '.' then begin
      Result := Copy(Result, 1, i - 1);
      Break;
    end;
  end;
end;

class procedure TFileUtil.FileRemoveReadOnly(sFile: string);
begin
  if FileExists( sFile ) then
    FileSetAttr( sFile, FileGetAttr (sFile) and (not SysUtils.faReadOnly) );
end;

class function TFileUtil.IsDirectSubDir(sDirMain, sDirSub: string):Boolean;
var
  sSub: string;
begin
  Result := False;
  if Length(sDirMain) > Length(sDirSub) then
    Exit;

  // 去掉结尾的\
  if Copy(sDirMain, Length(sDirMain), 1) <> PathDelim then
    sDirMain := sDirMain + PathDelim;
  if Copy(sDirSub, Length(sDirSub), 1) = PathDelim then
    sDirSub := Copy(sDirSub, 1, Length(sDirSub)-1);

  sSub := Copy(sDirSub, Length(sDirMain)+1, Length(sDirSub));
  if (sSub <> '') and (Pos(PathDelim, sSub) = 0) then
    Result := True;
end;

class function TFileUtil.CopyListFiles(SourceFiles: TStrings;sDestDir: string;bChangeDir:
  Boolean): Boolean;
var
  i: Integer;
  sDestPath: string;
begin
  Result := False;
  if sDestDir[Length(sDestDir)] <> '\' then
    sDestPath := sDestDir + '\'
  else
    sDestPath := sDestDir;

  // 源文件列表为空则失败
  if SourceFiles.Count = 0 then
    Exit;

  // 源目录和目的目录相同则不复制，但返回成功
  if SameText(ExtractFilePath(SourceFiles.Strings[0]), sDestPath) then
  begin
    Result := True;
    Exit;
  end;

  for i := 0 to SourceFiles.Count - 1 do
  begin
    Result := CopyFile(PChar(SourceFiles.Strings[i]),
      PChar(sDestPath + ExtractFileName(SourceFiles.Strings[i])),
      False);
    if not Result then
      Break;
    if bChangeDir then
      SourceFiles.Strings[i] := sDestPath +
        ExtractFileName(SourceFiles.Strings[i]);
  end;
end;

class function TFileUtil.CopyDirFiles(ASourceDir, ADestDir: string; CopySub: Boolean;
    FileExt: string):Boolean;
var
  i: Integer;
  strsFiles: TStrings;
  function ParseRelativeDir(dir1, dir2: string): string;
  begin
    Result := Copy(dir2, Length(dir1)+1, Length(dir2));
  end;
begin
  Result := False;
  if (not DirectoryExists(ASourceDir)) or (ADestDir = '') then
    Exit;
  if Copy(ADestDir, Length(ADestDir), 1 ) <> '\' then
    ADestDir := ADestDir + '\'
  else
    ADestDir := ADestDir;

  strsFiles := TStringList.Create;
  try
    strsFiles.Clear;
    FindFiles(ASourceDir, FileExt, strsFiles, CopySub, True);
    for i := 0 to strsFiles.Count - 1 do
    begin
      if Copy(strsFiles[i], Length(strsFiles[i]), 1) = '\' then
        CopyDirFiles(ASourceDir, ADestDir +
          ParseRelativeDir(ASourceDir, strsFiles[i]), CopySub, FileExt)
      else
        CopyListFiles(strsFiles, ADestDir, False);
    end;
  finally
    strsFiles.Free
  end;
end;

class function TFileUtil.GetFileFullInfo(AFileName: string): TFileFullInfo;
var
  nFileInfoSize: LongInt;
  Temp, Len: Cardinal;
  pInfoBuf: Pointer;
  TranslationLength: Cardinal;
  TranslationTable: Pointer;
  Value: string;
  CodePage, LanguageID, LookupString: string;
  HFileRes: HFILE;
begin
  with Result do
  begin
    FileName := AFileName;

    // file size
    HFileRes := CreateFile( Pchar( AFileName ), GENERIC_READ{ or GENERIC_WRITE},
                0, nil, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0 );
    try
      if HFileRes <> INVALID_HANDLE_VALUE then
        FileSize := GetFileSize(HFileRes, nil);
    finally
      CloseHandle(HFileRes);
    end;

    nFileInfoSize := GetFileVersionInfoSize(PChar(FileName), Temp);
    if nFileInfoSize > 0 then
    begin
      pInfoBuf := AllocMem(nFileInfoSize);
      try
        GetFileVersionInfo(PChar(FileName), 0, nFileInfoSize, pInfoBuf);

        if VerQueryValue( pInfoBuf, '\VarFileInfo\Translation', TranslationTable, TranslationLength ) then
        begin
          CodePage := Format( '%.4x', [ HiWord( PLongInt( TranslationTable )^ ) ] );
          LanguageID := Format( '%.4x', [ LoWord( PLongInt( TranslationTable )^ ) ] );
        end;
        LookupString := 'StringFileInfo\' + LanguageID + CodePage + '\';

        if VerQueryValue( pInfoBuf, PChar( LookupString + 'CompanyName' ), Pointer( Value ), Len ) then
          CompanyName := Value;
        if VerQueryValue( pInfoBuf, PChar( LookupString + 'FileDescription' ), Pointer( Value ), Len ) then
          FileDescription := Value;
        if VerQueryValue( pInfoBuf, PChar( LookupString + 'FileVersion' ), Pointer( Value ), Len ) then
          FileVersion := Value;
        if VerQueryValue( pInfoBuf, PChar( LookupString + 'InternalName' ), Pointer( Value ), Len ) then
          InternalName := Value;
        if VerQueryValue( pInfoBuf, PChar( LookupString + 'LegalCopyright' ), Pointer( Value ), Len ) then
          LegalCopyright := Value;
        if VerQueryValue( pInfoBuf, PChar( LookupString + 'LegalTrademarks' ), Pointer( Value ), Len ) then
          LegalTradeMarks := Value;
        if VerQueryValue( pInfoBuf, PChar( LookupString + 'OriginalFilename' ), Pointer( Value ), Len ) then
          OriginalFilename := Value;
        if VerQueryValue( pInfoBuf, PChar( LookupString + 'ProductName' ), Pointer( Value ), Len ) then
          ProductName := Value;
        if VerQueryValue( pInfoBuf, PChar( LookupString + 'ProductVersion' ), Pointer( Value ), Len ) then
          ProductVersion := Value;
        if VerQueryValue( pInfoBuf, PChar( LookupString + 'Comments' ), Pointer( Value ), Len ) then
          Comments := Value;
      finally
        FreeMem(pInfoBuf, nFileInfoSize);
      end;
    end;
  end;
end;

class procedure TFileUtil.ShowFilePropWindow(AFileName: string);
var
  sei: TShellExecuteInfo;
begin
  if FileExists(AFileName) then
  begin
    FillChar(sei,SizeOf(sei),#0);
    sei.cbSize:=SizeOf(sei);
    sei.lpFile:=PChar(AFileName);
    sei.lpVerb:='properties';
    sei.fMask:=SEE_MASK_INVOKEIDLIST;
    ShellExecuteEx(@sei);
  end;
end;

class function TFileUtil.IsFileInUse(AFileName: string): Boolean;
var
 HFileRes: HFILE;
begin
  HFileRes := CreateFile( Pchar( AFileName ), GENERIC_READ{ or GENERIC_WRITE},
                          0, nil, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0 );
  try
    Result:= (HFileRes = INVALID_HANDLE_VALUE);
  finally
    CloseHandle(HFileRes);
  end;
end;

class function TFileUtil.IsDirInUse(ADir: string): Boolean;
var
  slstFiles: TStrings;
  i: Integer;
begin
  Result := False;
  slstFiles := TStringList.Create;
  try
    FindFiles(ADir, '*.*', slstFiles, True, False);
    for i := 0 to slstFiles.Count - 1 do
    begin
      if IsFileInUse(slstFiles[i]) then
      begin
        Result := True;
        Break;
      end;
    end;
  finally
    slstFiles.Free;
  end;
end;

class function TFileUtil.ParseFileEncoding(const FileName: string): TEncoding;
var
  w:Word;
  b:Byte;
  function WordLoHiExchange(w:Word):Word;register;
  asm
    XCHG AL, AH
  end;
begin
  Result:= TEncoding.Default;
  with TFileStream.Create(FileName,fmOpenRead or fmShareDenyNone) do
  try
    Read(w,2);
    w:=WordLoHiExchange(w);//因为是以Word数据类型读取，故高低字节互换
    if w = TextFormatFlag[tfUnicode] then
      Result := TEncoding.Unicode
    else if w = TextFormatFlag[tfUnicodeBigEndian] then
      TEncoding.BigEndianUnicode
    else if w = TextFormatFlag[tfUtf8] then
    begin
      Read(b,1);// BOM是3个字节，因此UFT-8必须要跳过三个字节。
      Result:= TEncoding.UTF8;
    end else begin  // 无可判断的文件头，使用默认编码，暂未实现通过内容识别编码
      Result:= TEncoding.Default;
      Position:=0;
    end;
//    SetLength(sText,Size-Position);
//    ReadBuffer(sText[1],Size-Position);
  finally
    Free;
  end;
end;

class function TFileUtil.ParseFileEncoding2(const FileName: string): TEncoding;
var
  buff: TBytes;
//  readCount: Integer;
  encoding: TEncoding;
  stream: TFileStream;
begin
  stream := TFileStream.Create(FileName,fmOpenRead or fmShareDenyNone);
  try
    SetLength(buff, 1024);
    stream.Read(buff, Length(buff));
    encoding := nil;
    TEncoding.GetBufferEncoding(buff, encoding);
    Result := encoding;
  finally
    stream.Free;
  end;
end;

{ TFileList }
constructor TFileList.Create;
begin
  FFileList := TObjectList<TFileItem>.Create();
  FFileMap := TDictionary<String,TFileItem>.Create;
  FFileList.OwnsObjects := True;
  FSearchSubDir := True;     // 是否查找子文件夹
  FContainSelf := False;  // 是否包含自身
  FExcludeFiles := TStringList.Create;  // 排除的文件
  FSearchMode := faDirectory + faArchive + faReadOnly + faTemporary + faNormal;
  FSearchExts := TDictionary<String,String>.Create();
end;

destructor TFileList.Destroy;
begin
  FreeAndNil(FFileList);
  FreeAndNil(FFileMap);
  FreeAndNil(FExcludeFiles);

  FreeAndNil(FSearchExts);
  inherited;
end;

function TFileList.AddFile(sFileName: String; SrchRec: TSearchRec): Integer;
var
  fileItem, parentFileItem: TFileItem;
begin
  fileItem := TFileItem.Create;
  fileItem.FileName := sFileName;

  if (SrchRec.Attr and faDirectory) = faDirectory then
  begin
    fileItem.IsFile := False;
  end else begin
    fileItem.IsFile := True;
    fileItem.CreateTime := CovFileDate(SrchRec.FindData.ftCreationTime);
    fileItem.ModifyTime := CovFileDate(SrchRec.FindData.ftLastWriteTime);
    fileItem.AccessTime := CovFileDate(SrchRec.FindData.ftLastAccessTime);
  end;

  parentFileItem := GetFileItem(ExtractFilePath(sFileName));
  if Assigned(parentFileItem) then begin
    parentFileItem.SubFileItems.AddObject(fileItem.FileName, fileItem);
  end;

  if not FFileMap.ContainsKey(fileItem.FileName) then
    FFileMap.Add(fileItem.FileName, fileItem);

  Result := FFileList.Add(fileItem);
end;

procedure TFileList.DoFindFiles(sPath: string);
var
  FSrchRec: TSearchRec;
  nFindResult: Integer;
  nSearchMode: Integer;
  sFileName: String;
  fileIndex: Integer;
  fileItem: TFileItem;
begin
  sPath := IncludeTrailingPathDelimiter(sPath);
  nSearchMode := FSearchMode; // faAnyFile + faDirectory;

  nFindResult := FindFirst( sPath + '*.*', nSearchMode, FSrchRec);
  try
    while nFindResult = 0 do
    begin
      if not FbSearching then
        Break;
      if (FSrchRec.Name <> '.') and (FSrchRec.Name <> '..') then
      begin
        if (FSrchRec.Attr and faDirectory) = faDirectory then
        begin
          sFileName := sPath + IncludeTrailingPathDelimiter(FSrchRec.Name);
          fileIndex := AddFile(sFileName, FSrchRec);
          fileItem := FFileList.Items[fileIndex];
          if Assigned(FOnSearching) then
            FOnSearching(fileItem, FbSearching);

          if SearchSubDir then       // 是否搜索子目录模式
            Self.DoFindFiles(sPath + FSrchRec.Name);
        end
        else if not IsExcludeFile(FSrchRec.Name) and
                not((not ContainSelfApp) and
                    SameText(sPath + FSrchRec.Name, Application.ExeName)
                   ) then           // 根据属性决定是否添加自身程序名到列表
        begin
          if MatchExt(ExtractFileExt(FSrchRec.Name)) then
          begin
            sFileName := sPath + FSrchRec.Name;
            fileIndex := AddFile(sFileName, FSrchRec);
            fileItem := FFileList.Items[fileIndex];

            if Assigned(FOnSearching) then
              FOnSearching(fileItem, FbSearching);
          end;
        end;
      end;

      nFindResult := FindNext(FSrchRec);
    end;
  finally
    FindClose(FSrchRec);
  end;
end;

procedure TFileList.FindFiles(sPath, sFileExt: string);
var
  strs: TStrings;
  i: Integer;
begin
  FbSearching := True;
  FFileList.Clear;
  FFileMap.Clear;

  FSearchExts.Clear;
  strs := TStringList.Create;
  try
    Split(sFileExt, ';', strs);
    for i := 0 to strs.Count - 1 do
    begin
      FSearchExts.Add(strs[i], strs[i]);
    end;
  finally
    strs.Free;
  end;

  DoFindFiles(sPath);

  FbSearching := False;
end;

procedure TFileList.Clear;
begin
  FFileList.Clear;
  FFileMap.Clear;
end;

function TFileList.Count: Integer;
begin
  Result := FFileList.Count;
end;

function TFileList.CountDir(sPath: string): Integer;
var
  FSrchRec: TSearchRec;
  nFindResult, nSearchMode: Integer;
begin
  sPath := IncludeTrailingPathDelimiter(sPath);
  nSearchMode := faDirectory;
  Result := 0;
  nFindResult := FindFirst( sPath + '*.*', nSearchMode, FSrchRec);
  try
    while nFindResult = 0 do
    begin
      if (FSrchRec.Name <> '.') and (FSrchRec.Name <> '..') then
        Inc(Result);
      nFindResult := FindNext(FSrchRec);
    end;
  finally
    FindClose(FSrchRec);
  end;
end;

function TFileList.IsExcludeFile(AFileName: string): Boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to FExcludeFiles.Count - 1 do
  begin
    if SameText(ExtractFileName(FExcludeFiles[i]), AFileName) then
    begin
      Result := True;
      Break;
    end;
  end;
end;

function TFileList.MatchExt(sExt: String): Boolean;
begin
  Result := (FSearchExts.ContainsKey('.*'))
    or (FSearchExts.ContainsKey(sExt));
end;

procedure TFileList.SortIt;
var
  i: Integer;
  CurrDirFiles: TObjectList<TFileItem>;
begin
  CurrDirFiles := TObjectList<TFileItem>.Create();
  try
    for i := FFileList.Count - 1 downto 0 do
    begin
      if StrIsDir(FFileList[i].FileName) then
      begin
        CurrDirFiles.Add(FFileList[i]);
        FFileList.Delete(i);
      end;
    end;

    FFileList.Sort;
    CurrDirFiles.Sort;
    
    for i := 0 to CurrDirFiles.Count - 1 do
    begin
      FFileList.Add(CurrDirFiles[i]);
    end;
  finally
    CurrDirFiles.Free;
  end;
end;

function TFileList.StrIsDir(s: string): Boolean;
begin
  Result := False;
  if Copy(s, Length(s), 1) = PathDelim then
    Result := True;
end;

function TFileList.GetFileItem(index: Integer): TFileItem;
begin
  Result := TFileItem(FFileList[index]);
end;

function TFileList.GetFileItem(fileName: String): TFileItem;
begin
  Result := nil;
  if FFileMap.ContainsKey(fileName) then
    Result := FFileMap.Items[fileName];
end;

function TFileList.GetFileName(index: Integer): string;
begin
  Result := FFileList[index].FileName;
end;

{ TFileItem }

constructor TFileItem.Create;
begin
  FSubFileItems := TStringList.Create;
end;

destructor TFileItem.Destroy;
begin
  FSubFileItems.Free;
  inherited;
end;

function TFileItem.HasSubFile: Boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to FSubFileItems.Count - 1 do begin
    if TFileItem(FSubFileItems.Objects[i]).IsFile then begin
      Result := True;
      Break;
    end;
  end;
end;

end.
