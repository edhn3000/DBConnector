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
  end;

  function InSelectionPos(selPosMain, selPosSub: TSelectionPos): Boolean;

const
  WORD_FIND_MAX_LEN = 255;   // 查找文字时的最大长度

type
  { TOfficeHelper }
  TOfficeHelper = class
  protected
    FUseActiveApp: Boolean;   // 是否使用已激活的Application对象
    FIsInstalled: Boolean;    // Office是否已安装 实现类初始化时对其赋值
    FApplicatioin: OleVariant;// Office的Application
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
    property Application: OleVariant read FApplicatioin;
    property Document: OleVariant read FDocument;
    property Window: OleVariant read FWindow;
    property Opend: Boolean read FOpend;

    // 传入已激活的Application对象
    constructor Create(ActiveApp: OleVariant); overload; virtual;
    destructor Destroy;override;

    // Office是否已安装
    function IsInstalled: Boolean; virtual;

    // 打开文档
    procedure NewFile(sFileName: String); virtual;abstract;
    procedure OpenFile(sFileName: String; readOnly: Boolean);overload; virtual;abstract;
    procedure OpenFile(sFileName: String);overload;
    // 关闭文档
    procedure CloseFile(bSave: Boolean = false);virtual;abstract;
    // 显示文档，默认打开时不会显示文档
    procedure ShowDocument(); virtual;
    // 查找内容，返回内容在resultList中
    procedure Find(sFind: string; resultList: TStringList); virtual;

{-------------------------------------------------------------------------------
  过程名:    AddComment 增加批注
  作者:      fengyq
  日期:      2014.10.17
  参数:      text: string; userName: string
  返回值:    OleVariant 批注对象，Comment类型
-------------------------------------------------------------------------------}
    function AddComment(text: string; userName: string): OleVariant; virtual;
    procedure RemoveComments(userName: String); virtual;

{-------------------------------------------------------------------------------
  过程名:    SelText, 选中Word中文本
  作者:      fengyq
  日期:      2014.09.27
  参数:      text: String
  返回值:    无
-------------------------------------------------------------------------------}
    function SelectText(text: String): Boolean; overload; virtual;
    function SelectText(startPos, endPos: LongInt): Boolean; overload;virtual;
    function SelectTextWithTag(text: String): Boolean;virtual;
    function SelectTextInRange(text: String; iStart,iEnd:Integer): Boolean;virtual;
{-------------------------------------------------------------------------------
  过程名:    SelAndColorText 选中Word中文本并用给定颜色标出，
             标出的部分是靠iPos, iLen指定
  作者:      fengyq
  日期:      2014.09.27
  参数:      text: String; iPos, iLen: Integer; color: Integer
  返回值:    无
-------------------------------------------------------------------------------}
    procedure SelectAndColorText(text: String; iPos, iLen: Integer; color: Integer);virtual;

    procedure SelectAndColorSubText(selStart, selEnd: Integer; subText: string; color: Integer);virtual;

    {将换行符号替换为^p}
    function ReplaceWordNewLine(sText: String): string;

{-------------------------------------------------------------------------------
  过程名:    ExecuteControl 执行控件功能，比如点击工具栏按钮
  作者:      fengyq
  日期:      2014.10.20
  参数:      controlType: TOleEnum; 类型如msoControlButton
             tag: String 控件Tag，用于查找
  返回值:    Boolean
-------------------------------------------------------------------------------}
    function ExecuteControl(controlType: TOleEnum; tag: String): Boolean;virtual;abstract;


    // 获取选区，传入对象须自身具备start和end属性
    class function GetRangePos(range: OleVariant): TSelectionPos;
    // 获取选区，传入对象须自身具备start和end属性
    class procedure SetRangePos(range: OleVariant; selPos: TSelectionPos);

    // 获取选区，必须传入选取的拥有者，如Window、Application
    class function GetSelectionPos(SelectionOwner: OleVariant): TSelectionPos;
    // 设置选区，必须传入选取的拥有者，如Window、Application
    class procedure SetSelectionPos(SelectionOwner: OleVariant; selPos: TSelectionPos);

    // 改变选取开头，moveStep：移动步伐，start+moveStep
    class procedure MoveSelectionStart(selection: OleVariant; moveStep: Integer);
    // 改变选取结尾，moveStep：移动步伐，end+moveStep
    class procedure MoveSelectionEnd(selection: OleVariant; moveStep: Integer);

  end;

implementation

function InSelectionPos(selPosMain, selPosSub: TSelectionPos): Boolean;
begin
  Result := (selPosMain.StartPos <= selPosSub.StartPos)
     and (selPosMain.EndPos >= selPosSub.EndPos);
end;

{ TOfficeHelper }

constructor TOfficeHelper.Create(ActiveApp: OleVariant);
begin
  FApplicatioin := ActiveApp;
  FUseActiveApp := True;
  if FApplicatioin.Documents.Count > 0 then begin
    FDocument := FApplicatioin.ActiveDocument;
    FWindow := FApplicatioin.ActiveWindow;
  end;
  FCloseOnFree := False;  // 外部传入的WordApp对象，默认不管关闭
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
        FApplicatioin := GetActiveOleObject(className);
        FUseActiveApp := True;
      except
        on e: EOleSysError do begin
          // 没有打开的Application，创建新的
          FApplicatioin := CreateOleObject(className);
        end;
      end
    end else begin
      FApplicatioin := CreateOleObject(className);
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
  Result := not VarIsNull(FApplicatioin);

  if Result then
  begin
    try
      S := FApplicatioin.Version;
    except
      Result := False;
    end;
  end;
end;

function TOfficeHelper.IsInstalled: Boolean;
begin
  Result := FIsInstalled;
end;

procedure TOfficeHelper.OpenFile(sFileName: String);
begin
  OpenFile(sFileName, False);
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

function TOfficeHelper.ReplaceWordNewLine(sText: String): string;
var
  s: string;
begin
  // word中换行符替换为^p便于查找
  s := StringReplace(sText, #13#10, '^p', [rfReplaceAll]);
  s := StringReplace(s, #10, '^p', [rfReplaceAll]);
  s := StringReplace(s, #13, '^p', [rfReplaceAll]);
  Result := s;
end;

procedure TOfficeHelper.ShowDocument;
begin
  FApplicatioin.Visible := True;
  FApplicatioin.Activate;
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
    if cmt.Initial = userName then
    begin
      cmt.Delete;
    end;
  end;
end;

function TOfficeHelper.SelectText(text: String): Boolean;
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
  FWindow.Selection.HomeKey(oleUnit, oleExtend);
  //再定位错误
  FWindow.Selection.Start := 0;
  FWindow.Selection.End := 0;
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
  end
  else
  begin
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

  // 扩展选区
//  if Result and (sFindText <> text) then begin
//    TOfficeHelper.MoveSelectionEnd(FWindow.Selection, Length(text)-Length(sFindText));
//  end;
end;

function TOfficeHelper.SelectText(startPos, endPos: LongInt): Boolean;
begin
  FWindow.Selection.Start := startPos;
  FWindow.Selection.End := endPos;
  Result := True;
end;

function TOfficeHelper.SelectTextWithTag(text: String): Boolean;
var
  count: integer;
  selPos: TSelectionPos;
begin
  Result := False;
  if text = '' then
    Exit;
  text := StringReplace(text,#13#10,'^p',[rfReplaceAll]);
  text := StringReplace(text,#10,'^p',[rfReplaceAll]);
  SelectText(text);
  count := 0;
  while FWindow.Selection.Find.Found do begin
    // 修正选区
    selPos := TOfficeHelper.GetSelectionPos(FWindow);
    if (selPos.EndPos - selPos.StartPos + 1) < Length(text) then begin
      selPos.EndPos := selPos.StartPos + Length(text) - 1;
      TOfficeHelper.SetSelectionPos(FWindow, selPos);
    end;

    // 找到文字
    if (ReplaceWordNewLine(FWindow.Selection.Text) = text) then begin
      Result := True;
      // 检查是否高亮，优先选中高亮
      if (FWindow.Selection.Font.Shading.BackgroundPatternColor <> wdColorAutomatic)
        and (FWindow.Selection.Font.Shading.BackgroundPatternColor <> Integer(wdColorAutomatic)) then begin
        Exit;
      end;
    end;
    //为了防止死循环
    inc(count);
    if count > 50 then
      Exit;
    FWindow.Selection.Find.Execute;
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

  FWindow.Selection.start := iStart;
  FWindow.Selection.end := iEnd;
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
  end
  else
  begin
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
  // 扩展选区
//  if Result and (sFindText <> text)  then begin
//    TOfficeHelper.MoveSelectionEnd(FWindow.Selection, Length(text)-Length(sFindText));
//  end;
end;

procedure TOfficeHelper.SelectAndColorText(text: String; iPos, iLen: Integer;
  color: Integer);
var
  i: Integer;
  findpos: array[0..255] of Integer;
  findcount: Integer;
begin
  SelectText(text);
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
    FWindow.Selection.start := findpos[i] + iPos - 1;
    FWindow.Selection.end := findpos[i] + iPos - 1 + iLen;
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
begin
  FWindow.Selection.start := selStart;
  FWindow.Selection.end := selEnd;
  text := FWindow.Selection.Text;
  nPos := 0;
  repeat
    // 每次从上次找到的位置+1继续找
    nPos := PosFrom(subText, text, nPos + 1, False);
    if nPos > 0 then
    begin
      FWindow.Selection.start := selStart + nPos - 1;
      FWindow.Selection.end := selStart + nPos + Length(subText) - 1;
      FWindow.Selection.Font.Shading.BackgroundPatternColor := color;
    end;
  until nPos = 0;
end;

procedure TOfficeHelper.Find(sFind: string; resultList: TStringList);
var
  nStart, nEnd, nVertical, nHorizontal, nPos1, nPos2: Integer;
  StrList: TStringList;
  Range: Variant;
begin
  if SameText('', sFind) then
  begin
    Exit;
  end;

  Screen.Cursor := crHourGlass;
  StrList := resultList;
  //记住原始滚动条位置
  nStart      := FWindow.Selection.Start;
  nEnd        := FWindow.Selection.End;
  nVertical   := FWindow.ActivePane.VerticalPercentScrolled;
  nHorizontal := FWindow.ActivePane.HorizontalPercentScrolled;
  try

    FWindow.Selection.Find.ClearFormatting;
    FWindow.Selection.Find.Text    := sFind;
    FWindow.Selection.Find.Forward := True;

    while FWindow.Selection.Find.Execute do
    begin
      Range := FWindow.Selection.Range;
      Range.Start := Range.Start - 20;
      Range.End   := Range.End + 20;

      nPos1 := FWindow.Selection.Start;
      nPos2 := FWindow.Selection.End;
      StrList.AddObject(Format('%d=%s', [nPos1, Range.Text]), TObject(nPos2));

      Application.ProcessMessages;
    end;

  finally
    //恢复原始位置
    if StrList.Count > 0 then
    begin
      FWindow.ActivePane.VerticalPercentScrolled   := nVertical;
      FWindow.ActivePane.HorizontalPercentScrolled := nHorizontal;
      FWindow.Selection.Start := nStart;
      FWindow.Selection.End   := nEnd;
    end;

    Screen.Cursor := crDefault;
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
  range.end := selPos.EndPos;
  range.start := selPos.StartPos;
end;

class function TOfficeHelper.GetSelectionPos(SelectionOwner: OleVariant): TSelectionPos;
begin
  Result.StartPos := SelectionOwner.Selection.start;
  Result.EndPos := SelectionOwner.Selection.end;
end;

class procedure TOfficeHelper.SetSelectionPos(SelectionOwner: OleVariant;
  selPos: TSelectionPos);
begin
  SelectionOwner.Selection.end := selPos.EndPos;
  SelectionOwner.Selection.start := selPos.StartPos;
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
begin
  if moveStep = 0 then
    Exit;
  selection.End := selection.End + moveStep;
end;

end.
