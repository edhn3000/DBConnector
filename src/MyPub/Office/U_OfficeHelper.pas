unit U_OfficeHelper;
{
 @author  fengyq
 @comment Office助手单元
          主要依靠Variant做接口操作，很多操作是Word和Wps通用了
          各Office实现类（Word、Wps）须实现抽象方法，主要是变体的创建、个性操作
 @version 1.0
 @version 2015/11/09
}

interface

uses
  SysUtils, Variants, Classes, Controls, Forms, ComCtrls, ComObj, ActiveX,
  Office_TLB, Word_TLB, Windows;

type
  // 避免使用Selection.End造成单元语法提示器的错乱
  TSelectionPos = record
    StartPos: Integer;
    EndPos: Integer;
  public
    function ContainSelectionPos(selPos: TSelectionPos): Boolean;
  end;

const
  WORD_FIND_MAX_LEN = 255;   // 查找文字时的最大长度

type
  { TOfficeHelper }
  TOfficeHelper = class
  protected
    FUseActiveApp: Boolean;   // 是否使用已激活的Application对象
    FIsInstalled: Boolean;    // Office是否已安装 实现类初始化时对其赋值
    FApplication: OleVariant; // Office的Application
    FDocument: OleVariant;    // 当前文档
    FWindow: OleVariant;      // 当前窗口
    FOpend: Boolean;          // 是否已打开文档
    FCloseOnFree: Boolean;    // 释放对象时是否关闭文档和Application
    FFileName: String;        // 文档的文件名，外部传入

    function GetShowComment(): Boolean;
    procedure SetShowComment(value: Boolean);
  protected
    function PosFrom(sSub, S: String; nFrom: Integer;ignoreCase: Boolean): Integer;
    function GetFitFindText(text: String): string;

    function CreateApplication(useActiveApp: Boolean; className: String): Boolean; virtual;
    function CheckApplication(): Boolean;virtual;
  public
    property ShowComment: Boolean read GetShowComment write SetShowComment;
    property CloseOnFree: Boolean read FCloseOnFree write FCloseOnFree;
    property Application: OleVariant read FApplication;
    property Document: OleVariant read FDocument;
    property Window: OleVariant read FWindow;
    property Opend: Boolean read FOpend;

    // 传入已激活的Application对象
    constructor Create(ActiveApp: OleVariant); overload; virtual;
    constructor Create(doc: OleVariant; window: OleVariant); overload; virtual;
    destructor Destroy;override;

    // Office是否已安装
    function IsInstalled: Boolean; virtual;

    // 通过ScreenUpdating控制WordApp是否自动刷新UI
    procedure BeginUpdate;
    procedure EndUpdate;

{-------------------------------------------------------------------------------
    文档的创建保存关闭等操作
-------------------------------------------------------------------------------}
    // 创建新文档
    procedure NewFile(sFileName: String); virtual;abstract;
    // 打开文档
    procedure OpenFile(sFileName: String; readOnly: Boolean);overload; virtual;abstract;
    procedure OpenFile(sFileName: String);overload;
    // 保存文档
    procedure SaveFile; virtual;
    procedure SaveAsFile(newFileName: String); virtual;
    // 关闭文档
    procedure CloseFile(bSave: Boolean = false);virtual;abstract;
    // 显示文档，默认打开时不会显示文档
    procedure ShowDocument(); virtual;

{-------------------------------------------------------------------------------
    批注相关操作
-------------------------------------------------------------------------------}
    function AddComment(text: string; userName: string): OleVariant; virtual;
    procedure RemoveComments(userName: String); virtual;

{-------------------------------------------------------------------------------
    选区相关操作
-------------------------------------------------------------------------------}
    procedure SetSelectionText(s: string); virtual;
    // 写入文本，通过insertAfter写，写之后文本会被选中，便于追加格式
    procedure AppendSelectionText(s: string); virtual;
    // 相当于选区的Start=End，在AppendSelectionText后通常要调用
    procedure MoveToSelectionEnd;
    // 设置样式
    procedure SetSelectionStyle(styleName: string);

{-------------------------------------------------------------------------------
    书签相关操作
-------------------------------------------------------------------------------}
    function GetBookmark(name: String; var bk: OleVariant): Boolean; virtual;
    function SelectBookmark(name: String): Boolean; virtual;
    function ReplaceBookmark(key, value: String; reAddBookmark: boolean = False): Boolean; virtual;
    function ReplaceBookmarks(values: TStrings; reAddBookmark: boolean = False): Boolean; virtual;
    procedure ClearBookmarks(); virtual;

{-------------------------------------------------------------------------------
    文本处理
-------------------------------------------------------------------------------}
    // 查找并选中文本
    function SelectText(text: String; fromHead: Boolean = True): Boolean; overload; virtual;
    function SelectText(startPos, endPos: LongInt): Boolean; overload;virtual;
    function SelectTextWithTag(text: String; fromHead: Boolean = True): Boolean;virtual;
    function SelectTextInRange(text: String; iStart,iEnd:Integer): Boolean;virtual;
    // 查找并选中文本，对文本上背景色
    procedure SelectAndColorText(text: String; iPos, iLen: Integer; color: Integer; fromHead: Boolean = True);virtual;
    procedure SelectAndColorSubText(selStart, selEnd: Integer; subText: string; color: Integer);virtual;
    // 查找并替换文本内容
    function ReplaceText(text: String; replacement: String; fromHead: Boolean = True): Boolean; virtual;
    function ReplaceTextInRange(text: String; replacement: String; startPos, endPos: Integer): Boolean; virtual;


{-------------------------------------------------------------------------------
    表格操作
-------------------------------------------------------------------------------}
    function AddTable(rowCount: integer; colCount: Integer): OleVariant; virtual;
    procedure DeleteTableRows(table: OleVariant; startRow, endRow: Integer); virtual;

{-------------------------------------------------------------------------------
    段落操作
-------------------------------------------------------------------------------}
    // 从给定位置获取到所在段落号
    function GetParagraphNoByPos(posInDoc: Integer; beginPaNo, endPaNo: Integer): Integer; overload;
    function GetParagraphNoByPos(posInDoc: Integer): Integer; overload;

{-------------------------------------------------------------------------------
    控件按钮操作
-------------------------------------------------------------------------------}
    // 执行控件功能，比如点击工具栏按钮
    function ExecuteControl(controlType: TOleEnum; tag: String): Boolean;virtual;abstract;

{-------------------------------------------------------------------------------
    静态方法
-------------------------------------------------------------------------------}
    // 将换行符号替换为^p，使用Word查找功能时必须先对文本做此处理
    class function ReplaceWordNewLine(sText: String): string;

    // 改变选取开头，moveStep：移动步伐，start+moveStep
    class procedure MoveSelectionStart(selection: OleVariant; moveStep: Integer);
    // 改变选取结尾，moveStep：移动步伐，end+moveStep
    class procedure MoveSelectionEnd(selection: OleVariant; moveStep: Integer);

    // 获取Range或Selection，传入对象须自身具备start和end属性
    class function GetRangePos(range: OleVariant): TSelectionPos;
    // 设置Range或Selection的位置，传入对象须自身具备start和end属性
    class procedure SetRangePos(range: OleVariant; selPos: TSelectionPos); overload;
    class procedure SetRangePos(range: OleVariant; startPos, endPos: Integer); overload;

    // 获取SelectionOwner.Selection，必须传入选区的拥有者，如Window、Application
    class function GetSelectionPos(SelectionOwner: OleVariant): TSelectionPos;
    // 设置SelectionOwner.Selection的位置，必须传入选区的拥有者，如Window、Application
    class procedure SetSelectionPos(SelectionOwner: OleVariant; selPos: TSelectionPos); overload;
    class procedure SetSelectionPos(SelectionOwner: OleVariant; startPos, endPos: Integer); overload;

  end;

implementation

{ TSelectionPos }

function TSelectionPos.ContainSelectionPos(selPos: TSelectionPos): Boolean;
begin
  Result := (StartPos <= selPos.StartPos) and (EndPos >= selPos.EndPos);
end;

{ TOfficeHelper }

constructor TOfficeHelper.Create(ActiveApp: OleVariant);
begin
  FApplication := ActiveApp;
  FUseActiveApp := True;
  if FApplication.Documents.Count > 0 then begin
    FDocument := FApplication.ActiveDocument;
    FWindow := FApplication.ActiveWindow;
  end;
  FCloseOnFree := False;  // 外部传入的WordApp对象，默认不管关闭
end;

constructor TOfficeHelper.Create(doc: OleVariant; window: OleVariant);
begin
  FDocument := doc;
  FWindow := window;
  FApplication := doc.Application;
  FUseActiveApp := True;
  FCloseOnFree := False;
end;

destructor TOfficeHelper.Destroy;
begin

  inherited;
end;

function TOfficeHelper.CreateApplication(useActiveApp: Boolean; className: String): Boolean;
begin
  Result := False;
  try
    if useActiveApp then begin
      try
        FApplication := GetActiveOleObject(className);
        FUseActiveApp := True;
      except
        on e: EOleSysError do begin
          // 没有打开的Application，创建新的
          FApplication := CreateOleObject(className);
        end;
      end
    end else begin
      FApplication := CreateOleObject(className);
    end;
    Result := True;
    FIsInstalled := Result;
  except
    on e: Exception do begin
      OutputDebugString(PChar(Format('CreateApplication using (%s) fail! %s' ,[className, e.Message])));
    end;
  end;
end;

function TOfficeHelper.CheckApplication(): Boolean;
var
  S: string;
begin
  Result := not VarIsNull(FApplication);

  if Result then
  begin
    try
      S := FApplication.Version;
    except
      Result := False;
    end;
  end;
end;

function TOfficeHelper.IsInstalled: Boolean;
begin
  Result := FIsInstalled;
end;

procedure TOfficeHelper.BeginUpdate;
begin
  FApplication.ScreenUpdating := False;
end;

procedure TOfficeHelper.EndUpdate;
begin
  FApplication.ScreenUpdating := True;
end;

procedure TOfficeHelper.OpenFile(sFileName: String);
begin
  OpenFile(sFileName, False);
end;

procedure TOfficeHelper.SaveFile;
begin
  FDocument.Save;
end;

procedure TOfficeHelper.SaveAsFile(newFileName: String);
begin
  FDocument.SaveAs(newFileName);
end;

function TOfficeHelper.GetFitFindText(text: String): string;
var
  sFindText: String;
begin
  if Length(text) > WORD_FIND_MAX_LEN then begin
    sFindText := Copy(text, 1, WORD_FIND_MAX_LEN);
    if Copy(text, Length(sFindText), 2) = '^p' then    // 防止^p被分隔开
      sFindText := Copy(sFindText, 1, Length(sFindText) - 1);
  end else
    sFindText := text;
  Result := sFindText;
end;

function TOfficeHelper.PosFrom(sSub, S: String; nFrom: Integer;ignoreCase: Boolean): Integer;
var
  nIndexMain, nIndexSub: Integer;
  nIndexSave: Integer;
begin
  Result := 0;
  // 特殊情况
  if Length(sSub) > Length(S) then
    Exit;

  nIndexSub  := 1;
  nIndexMain := nFrom;  // 起始位置

  while nIndexSub <= Length(sSub) do
  begin
    if not ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
            or (sSub[nIndexSub]=S[nIndexMain])) then
    begin
      Inc(nIndexMain);
      if nIndexMain > Length(S) then
        System.Break;
    end
    else
    begin
      nIndexSave := nIndexMain;    // 保存原位置

      while (nIndexSub<= Length(sSub))
            and ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
                 or (sSub[nIndexSub]=S[nIndexMain])) do
      begin
        Inc(nIndexMain);
        Inc(nIndexSub);
      end;

      if nIndexSub = Length(sSub)+1 then      // 找到
      begin
        Result := nIndexSave;       // 保存下的位置就是结果
        System.Break;
      end
      else                          // 没找到
      begin
        nIndexMain := nIndexSave + 1;  // 向后搜索
        nIndexSub := 1;
      end;
    end;
  end;
end;

class function TOfficeHelper.ReplaceWordNewLine(sText: String): string;
var
  s: string;
begin
  // word中换行符替换为^p便于查找
  s := StringReplace(sText, #13#10, '^p', [rfReplaceAll]);
  s := StringReplace(s, #10, '^p', [rfReplaceAll]);
  s := StringReplace(s, #13, '^p', [rfReplaceAll]);
  Result := s;
end;

function TOfficeHelper.AddTable(rowCount: integer; colCount: Integer): OleVariant;
begin
  Result := FDocument.Tables.Add( FWindow.Selection.Range, rowCount, colCount, EmptyParam, EmptyParam );
end;

procedure TOfficeHelper.DeleteTableRows(table: OleVariant; startRow, endRow: Integer);
var
  i: Integer;
begin
  for i := endRow downto startRow do begin
    table.Rows.Item(i).Delete;
  end;
end;

function TOfficeHelper.GetParagraphNoByPos(posInDoc: Integer; beginPaNo, endPaNo: Integer): Integer;
var
  frontNo, endNo, midNo: Integer;
  selPos: TSelectionPos;
  bMatch: Boolean;
begin
  Result := 0;
  frontNo := beginPaNo;
  endNo := endPaNo;
  midNo := Trunc((frontNo + endNo) / 2);

  // 由于上一段结尾=下一段开头，应认为交界位置属于后一段
  // 如：段5.End=586,同时也会有段6.Start=586，应认为586是段6的
  // 特殊情况是结尾段落，只要位置>=开头就可以
  selPos := TOfficeHelper.GetRangePos(FDocument.Paragraphs.Item(midNo).Range);
  if midNo = endNo then begin
    bMatch := selPos.StartPos <= posInDoc;
  end else begin
    bMatch := (selPos.StartPos <= posInDoc) and (selPos.EndPos > posInDoc);
  end;
  if not bMatch then begin
    if posInDoc < selPos.StartPos  then begin
      endNo := midNo - 1
    end else begin
      frontNo := midNo + 1;
    end;
  end else begin
    Result := midNo;
    Exit;
  end;
  if frontNo > endNo then
    Exit;
  Result := GetParagraphNoByPos(posInDoc, frontNo, endNo);
end;

function TOfficeHelper.GetParagraphNoByPos(posInDoc: Integer): Integer;
var
  frontNo, endNo: Integer;
begin
  frontNo := 1;
  endNo := FDocument.Paragraphs.Count;
  Result := GetParagraphNoByPos(posInDoc, frontNo, endNo);
end;

procedure TOfficeHelper.ShowDocument;
begin
  FApplication.Visible := True;
  FApplication.Activate;
end;

function TOfficeHelper.GetShowComment(): Boolean;
begin
  Result := FWindow.View.ShowRevisionsAndComments;
end;

procedure TOfficeHelper.SetShowComment(value: Boolean);
begin
  FWindow.View.ShowRevisionsAndComments := value;
//  FWindow.View.RevisionsView := wdRevisionsViewFinal;
end;

function TOfficeHelper.AddComment(text: string; userName: string): OleVariant;
var
  range: OleVariant;
  cmt: OleVariant;
begin
  range := FWindow.Selection.Range;
  cmt := FDocument.Comments.Add(range, text);
  cmt.Initial := userName;
  Result := cmt;
end;

procedure TOfficeHelper.RemoveComments(userName: String);
var
  i: Integer;
  cmt: OleVariant;
begin
  // 批注索引是从1开始直到Count
  for i := FDocument.Comments.Count downto 1 do
  begin
    cmt := FDocument.Comments.Item(i);
    if cmt.Initial = userName then begin
      cmt.Delete;
    end;
  end;
end;

procedure TOfficeHelper.SetSelectionText(s: string);
begin
  FWindow.Selection.Text := s;
end;

procedure TOfficeHelper.AppendSelectionText(s: string);
begin
  FWindow.Selection.InsertAfter(s);
end;

procedure TOfficeHelper.MoveToSelectionEnd;
var
  direct: OleVariant;
begin
  direct := wdCollapseEnd;
  FWindow.Selection.Collapse(direct);
end;

procedure TOfficeHelper.SetSelectionStyle(styleName: string);
var
  style: OleVariant;
begin
  if styleName <> '' then begin
    style := FDocument.Styles.item(styleName);
    if not VarIsNull(style) then
      FWindow.Selection.Style := style;
  end;
end;

function TOfficeHelper.GetBookmark(name: String; var bk: OleVariant): Boolean;
var
  bkItem: OleVariant;
  i: Integer;
begin
  Result := False;
  for i := 1 to FDocument.Bookmarks.Count do begin
    bkItem := FDocument.Bookmarks.Item(i);
    if name = bkItem.Name then begin
      bk := bkItem;
      Result := True;
      System.Break;
    end;
  end;
end;

function TOfficeHelper.SelectBookmark(name: String): Boolean;
var
  bk: OleVariant;
begin
  Result := GetBookmark(name, bk);
  if Result then
    bk.Range.Select;
end;

function TOfficeHelper.ReplaceBookmark(key, value: String; reAddBookmark: boolean): Boolean;
var
  bkValue: String;
  bk: OleVariant;
  bkRange: OleVariant;
begin
  Result := False;
  if not GetBookmark(key, bk) then
    Exit;

  bkValue := value;
  bkRange := bk.Range;
  bkRange.Text := bkValue;
  Result := True;

  if reAddBookmark then
    FDocument.Bookmarks.Add(key, bkRange);
end;

function TOfficeHelper.ReplaceBookmarks(values: TStrings; reAddBookmark: boolean): Boolean;
var
  i, nIndex: Integer;
  bkName, bkValue: String;
  bk: OleVariant;
  bkRange: OleVariant;
begin
  Result := False;
  for i := FDocument.Bookmarks.Count downto 1 do begin
    bk := FDocument.Bookmarks.Item(i);
    bkName := bk.Name;
    // 不存在的表示不用替换
    nIndex := values.IndexOfName(bkName);
    if nIndex <= 0 then
      Continue;

    bkValue := values.Values[bkName];
    bkRange := bk.Range;
    bkRange.Text := bkValue;
    Result := True;

    if reAddBookmark then
      FDocument.Bookmarks.Add(bkName, bkRange);
  end;
end;

procedure TOfficeHelper.ClearBookmarks();
var
  i: Integer;
begin
  for i := FDocument.Bookmarks.Count downto 1 do begin
    FDocument.Bookmarks.Item(i).Delete;
  end;
end;

function TOfficeHelper.SelectText(text: String; fromHead: Boolean): Boolean;
var
  upline,nomove,oleUnit,oleExtend: OleVariant;
  sFindText: WideString;
begin
  Result := False;
  if text = '' then
    Exit;
  sFindText := GetFitFindText(text);

  oleUnit := wdStory;
  oleExtend := wdMove;
  //先定位到文档最前面
  if fromHead then begin
    FWindow.Selection.HomeKey(oleUnit, oleExtend);
    SetSelectionPos(FWindow, 0, 0);
  end;
  FWindow.Selection.Find.ClearFormatting;
  FWindow.Selection.Find.Text:=sFindText;
  FWindow.Selection.Find.Replacement.Text:=sFindText;
  FWindow.Selection.Find.Forward:=True;
  FWindow.Selection.Find.Wrap:=1;
  FWindow.Selection.Find.Format:=False;
  FWindow.Selection.Find.MatchCase:=False;
  FWindow.Selection.Find.MatchWholeWord:=False;
  // 因为word不区分全角半角的"和'，因此特殊处理
  if (sFindText = '"') or (sFindText = '''') then
  begin
    FWindow.Selection.Find.MatchByte:=False;
    FWindow.Selection.Find.MatchWildcards:=True;
  end else begin
    FWindow.Selection.Find.MatchByte:=True;
    FWindow.Selection.Find.MatchWildcards:=False;
  end;
  FWindow.Selection.Find.MatchSoundsLike:=False;
  FWindow.Selection.Find.MatchAllWordForms:=False;
  FWindow.Selection.Find.Execute;

  // Find时word会跳转到结果处，但结果显示在窗口最顶部，向上滚动8格易于查看
  upline := 8;
  nomove := 0;
  FWindow.SmallScroll(nomove,upline,nomove,nomove);
  Result := FWindow.Selection.Find.Found;

  // 如果之前缩短了查找文本，这里扩展选区
  if Result and (sFindText <> text) then begin
    TOfficeHelper.MoveSelectionEnd(FWindow.Selection, Length(text)-Length(sFindText));
  end;
end;

function TOfficeHelper.SelectText(startPos, endPos: LongInt): Boolean;
begin
  SetSelectionPos(FWindow, startPos, endPos);
  FWindow.ScrollIntoView(FWindow.Selection.Range);
  Result := True;
end;

function TOfficeHelper.SelectTextWithTag(text: String; fromHead: Boolean): Boolean;
var
  count: integer;
  selPos, firstMatchPos: TSelectionPos;
begin
  Result := False;
  if text = '' then
    Exit;
  text := StringReplace(text,#13#10,'^p',[rfReplaceAll]);
  text := StringReplace(text,#10,'^p',[rfReplaceAll]);
  SelectText(text, fromHead);
  count := 0;
  firstMatchPos.StartPos := 0;

  while FWindow.Selection.Find.Found do begin
    // 修正选区
    selPos := TOfficeHelper.GetSelectionPos(FWindow);
    if (selPos.EndPos - selPos.StartPos + 1) < Length(text) then begin
      selPos.EndPos := selPos.StartPos + Length(text) - 1;
      TOfficeHelper.SetSelectionPos(FWindow, selPos);
    end;
    if firstMatchPos.StartPos = 0 then begin
      firstMatchPos := TOfficeHelper.GetSelectionPos(FWindow);
    end;

    // 找到文字
    if (ReplaceWordNewLine(FWindow.Selection.Text) = text) then begin
      Result := True;
      // 检查是否高亮，优先选中高亮，选区内有多个颜色会返回9999999
      if (FWindow.Selection.Font.Shading.BackgroundPatternColor <> wdColorAutomatic)
        and (FWindow.Selection.Font.Shading.BackgroundPatternColor <> Integer(wdColorAutomatic)) then begin
        Exit;
      end;
    end;
    //为了防止死循环
    inc(count);
    if count > 50 then
      System.Break;
    FWindow.Selection.Find.Execute;
  end;

  // 如果没有找到带标记的，优先定位到最开头匹配的
  if firstMatchPos.StartPos > 0 then begin
    SelectText(firstMatchPos.StartPos, firstMatchPos.EndPos);
  end;
end;

function TOfficeHelper.SelectTextInRange(text: String; iStart,iEnd:Integer): Boolean;
var
  sFindText: String;
begin
  Result := False;
  if text = '' then
    Exit;
  sFindText := GetFitFindText(text);

  SetSelectionPos(FWindow, iStart, iEnd);
  FWindow.Selection.Find.ClearFormatting;
  FWindow.Selection.Find.Text := sFindText;
  FWindow.Selection.Find.Replacement.Text := sFindText;
  FWindow.Selection.Find.Forward := True;
  FWindow.Selection.Find.Wrap := wdFindStop;//wdFindContinue 1 到达搜索区域的开始或者结尾时，继续执行查找操作
                                                     //wdFindStop 0 到达搜索范围的开始或者结尾时，停止执行查找操作
  FWindow.Selection.Find.Format := False;
  FWindow.Selection.Find.MatchCase := False;
  FWindow.Selection.Find.MatchWholeWord :=False;
  // 因为word不区分全角半角的"和'，因此特殊处理
  if (sFindText = '"') or (sFindText = '''') then
  begin
    FWindow.Selection.Find.MatchByte:=False;
    FWindow.Selection.Find.MatchWildcards:=True;  //可以指定通配符及其他高级搜索条件
  end else begin
    FWindow.Selection.Find.MatchByte:=True;
    FWindow.Selection.Find.MatchWildcards:=False;
  end;
  FWindow.Selection.Find.MatchSoundsLike:=False;
  FWindow.Selection.Find.MatchAllWordForms:=False;
  FWindow.Selection.Find.Execute;

  if not FWindow.Selection.Find.Found then
  begin
    if (iStart = 0) and (iEnd = 0) then
      exit;

    if iStart > 100 then   //扩大查找范围
      SelectTextInRange(text, iStart - 100, iEnd + 100)
    else
      SelectTextInRange(text, 0, 0);
  end;
  Result := FWindow.Selection.Find.Found;
  // 如果之前缩短了查找文本，这里扩展选区
  if Result and (sFindText <> text)  then begin
    TOfficeHelper.MoveSelectionEnd(FWindow.Selection, Length(text)-Length(sFindText));
  end;
end;

procedure TOfficeHelper.SelectAndColorText(text: String; iPos, iLen: Integer;
  color: Integer; fromHead: Boolean);
var
  i: Integer;
  findpos: array[0..255] of Integer;
  findcount: Integer;
  selPos: TSelectionPos;
begin
  SelectText(text, fromHead);
  findcount := 0;
  while FWindow.Selection.Find.Found do
  begin
    findpos[findcount] := FWindow.Selection.start;
    Inc(findcount);
    if findcount > 255 then
      System.break;
    FWindow.Selection.Find.Execute;
  end;

  for i := 0 to findcount - 1 do
  begin
    selPos.StartPos := findpos[i] + iPos - 1;
    selPos.EndPos := findpos[i] + iPos - 1 + iLen;
    SetSelectionPos(FWindow, selPos);
    if (ReplaceWordNewLine(FWindow.Selection.Text) = text)  then begin
      FWindow.Selection.Font.Shading.BackgroundPatternColor := color;
    end;
  end;
end;

procedure TOfficeHelper.SelectAndColorSubText(selStart, selEnd: Integer;
  subText: string; color: Integer);
var
  nPos: Integer;
  text: String;
  selPos: TSelectionPos;
begin
  SetSelectionPos(FWindow, selStart, selEnd);
  text := FWindow.Selection.Text;
  nPos := 0;
  repeat
    // 每次从上次找到的位置+1继续找
    nPos := PosFrom(subText, text, nPos + 1, False);
    if nPos > 0 then
    begin
      selPos.StartPos := selStart + nPos - 1;
      selPos.EndPos := selStart + nPos + Length(subText) - 1;
      SetSelectionPos(FWindow, selPos);
      FWindow.Selection.Font.Shading.BackgroundPatternColor := color;
    end;
  until nPos = 0;
end;

function TOfficeHelper.ReplaceText(text: String; replacement: String; fromHead: Boolean): Boolean;
var
  oleUnit,oleExtend: OleVariant;
  sFindText: WideString;
begin
  Result := False;
  if text = '' then
    Exit;
  sFindText := text;

  oleUnit := wdStory;
  oleExtend := wdMove;
  //先定位到文档最前面
  if fromHead then begin
    FWindow.Selection.HomeKey(oleUnit, oleExtend);
    SetSelectionPos(FWindow, 0, 0);
  end;
  FWindow.Selection.Find.ClearFormatting;
  FWindow.Selection.Find.Replacement.ClearFormatting;
  FWindow.Selection.Find.Text:=sFindText;
  FWindow.Selection.Find.Replacement.Text:=replacement;
  FWindow.Selection.Find.Forward:=True;
  FWindow.Selection.Find.Wrap:=wdFindContinue;
  FWindow.Selection.Find.Format:=False;
  FWindow.Selection.Find.MatchCase:=False;
  FWindow.Selection.Find.MatchWholeWord:=False;
  // 因为word不区分全角半角的"和'，因此特殊处理
  if (sFindText = '"') or (sFindText = '''') then
  begin
    FWindow.Selection.Find.MatchByte:=False;
    FWindow.Selection.Find.MatchWildcards:=True;
  end else begin
    FWindow.Selection.Find.MatchByte:=True;
    FWindow.Selection.Find.MatchWildcards:=False;
  end;
  FWindow.Selection.Find.MatchSoundsLike:=False;
  FWindow.Selection.Find.MatchAllWordForms:=False;
  FWindow.Selection.Find.Execute(sFindText, False{MatchCase},
    False{MatchWholeWord}, False{MatchWildcards}, False{MatchSoundsLike},
    False{MatchAllWordForms}, True{Forward}, wdFindContinue{Wrap},
    False{Format}, replacement{ReplaceWith}, wdReplaceAll{Replace});
  Result := FWindow.Selection.Find.Found;
end;

function TOfficeHelper.ReplaceTextInRange(text: String; replacement: String; startPos, endPos: Integer): Boolean;
begin
  Result := False;
  SetSelectionPos(FWindow, startPos, startPos);
  while SelectText(text, False) do begin
    if FWindow.Selection.Start > endPos then
      System.Break;
    SetSelectionText(replacement);
    Result := True;
  end;
end;

//// static

class function TOfficeHelper.GetRangePos(range: OleVariant): TSelectionPos;
begin
  Result.StartPos := range.start;
  Result.EndPos := range.end;
end;

class procedure TOfficeHelper.SetRangePos(range: OleVariant;
  selPos: TSelectionPos);
begin
  SetRangePos(range, selPos.StartPos, selPos.EndPos);
end;

class function TOfficeHelper.GetSelectionPos(SelectionOwner: OleVariant): TSelectionPos;
begin
  Result.StartPos := SelectionOwner.Selection.start;
  Result.EndPos := SelectionOwner.Selection.end;
end;

class procedure TOfficeHelper.SetSelectionPos(SelectionOwner: OleVariant;
  selPos: TSelectionPos);
begin
  SetSelectionPos(SelectionOwner, selPos.StartPos, selPos.EndPos);
end;

class procedure TOfficeHelper.MoveSelectionStart(selection: OleVariant;
  moveStep: Integer);
begin
  if moveStep = 0 then
    Exit;
  selection.Start := selection.Start + moveStep;
end;

class procedure TOfficeHelper.MoveSelectionEnd(selection: OleVariant;
  moveStep: Integer);
var
  endPos: Integer;
begin
  if moveStep = 0 then
    Exit;
  endPos := selection.End;
  selection.End := endPos + moveStep;
end;

class procedure TOfficeHelper.SetRangePos(range: OleVariant; startPos, endPos: Integer);
begin
  range.end := EndPos;
  range.start := StartPos;
end;

class procedure TOfficeHelper.SetSelectionPos(SelectionOwner: OleVariant;
  startPos, endPos: Integer);
begin
  SelectionOwner.Selection.end := EndPos;
  SelectionOwner.Selection.start := StartPos;
end;

end.
