unit U_WordHelper;

interface

uses
  SysUtils, Variants, Classes, Controls, Forms,
  ComCtrls, ComObj, ActiveX, Word_TLB, Office_TLB;

const
  WORD_FIND_MAX_LEN = 255;

type
  TSelectionPos = record
    StartPos: Integer;
    EndPos: Integer;
  end;

  TWordHelper = class
  private
    FWordApp: OleVariant;
    FWordDoc: OleVariant; // active document
    FOpend: Boolean;
    FCloseOnFree: Boolean;
    FFileName: String;
    function GetShowComment(): Boolean;
    procedure SetShowComment(value: Boolean);
  protected
    function CheckWordApp(): Boolean;
  public
    property CloseOnFree: Boolean read FCloseOnFree write FCloseOnFree;
    property WordApp: OleVariant read FWordApp;
    property Opend: Boolean read FOpend;
    property ShowComment: Boolean read GetShowComment write SetShowComment;

    constructor Create;overload;
    constructor Create(useActiveApp: Boolean);overload;
    constructor Create(WordApp: OleVariant);overload;
    destructor Destroy;override;

    { 打开Wrod文档 }
    procedure OpenFile(sFileName: String);overload;
    procedure OpenFile(sFileName: String; readOnly: Boolean);overload;
    { 关闭Word，不保存更改，要保存先调用SaveFile}
    procedure CloseFile();
    { 保存Word文档 }
    procedure SaveFile();
    { 显示Word文档，默认打开时不会显示文档 }
    procedure ShowWord();
    { 查找内容，返回内容在resultList中 }
    procedure Find(sFind: string; resultList: TStringList);

{-------------------------------------------------------------------------------
  过程名:    AddComment 增加批注
  作者:      fengyq
  日期:      2014.10.17
  参数:      text: string; userName: string
  返回值:    OleVariant 批注对象，Comment类型
-------------------------------------------------------------------------------}
    function AddComment(text: string; userName: string): OleVariant;
    procedure RemoveComments(userName: String);

{-------------------------------------------------------------------------------
  过程名:    SelText, 选中Word中文本
  作者:      fengyq
  日期:      2014.09.27
  参数:      text: String
  返回值:    无
-------------------------------------------------------------------------------}
    procedure SelectText(text: String);
    procedure SelectTextWithTag(text: String);
{-------------------------------------------------------------------------------
  过程名:    SelAndColorText 选中Word中文本并用给定颜色标出，
             标出的部分是靠iPos, iLen指定
  作者:      fengyq
  日期:      2014.09.27
  参数:      text: String; iPos, iLen: Integer; color: Integer
  返回值:    无
-------------------------------------------------------------------------------}
    procedure SelectAndColorText(text: String; iPos, iLen: Integer; color: Integer);

    procedure SelectAndColorSubText(selStart, selEnd: Integer; subText: string; color: Integer);

{-------------------------------------------------------------------------------
  过程名:    ExecuteControl 执行控件功能，比如点击工具栏按钮
  作者:      fengyq
  日期:      2014.10.20
  参数:      controlType: TOleEnum; 类型如msoControlButton
             tag: String 控件Tag，用于查找
  返回值:    Boolean
-------------------------------------------------------------------------------}
    function ExecuteControl(controlType: TOleEnum; tag: String): Boolean;

    class function GetSelectionPos(wordApp: OleVariant): TSelectionPos;
    class procedure SetSelectionPos(wordApp: OleVariant; selPos: TSelectionPos);

  end;

implementation

{ TWordHelper }

uses
  Windows;

function PosFrom(sSub, S: String;
  nFrom: Integer;ignoreCase: Boolean): Integer;
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

function TWordHelper.CheckWordApp(): Boolean;
var
  S: string;
begin
  Result := not VarIsNull(FWordApp);

  if Result then
  begin
    try
      S := FWordApp.Version;
    except
      Result := False;
    end;
  end;
end;

function TWordHelper.GetShowComment(): Boolean;
begin
  Result := wordApp.ActiveWindow.View.ShowRevisionsAndComments;
end;

procedure TWordHelper.SetShowComment(value: Boolean);
begin
  wordApp.ActiveWindow.View.ShowRevisionsAndComments := value;
//  wordApp.ActiveWindow.View.RevisionsView := wdRevisionsViewFinal;
end;

constructor TWordHelper.Create;
begin
  Create(False);
end;

constructor TWordHelper.Create(useActiveApp: Boolean);
begin
  if useActiveApp then begin
    try
      FWordApp := GetActiveOleObject('Word.Application');
    except
      on e: EOleSysError do
        FWordApp := CreateOleObject('Word.Application');
    end
  end else begin
    FWordApp := CreateOleObject('Word.Application');
  end;
  // 不保存到Normal.dotm，避免文档被重复打开后在关闭时报错
  FWordApp.Options.SaveNormalPrompt := False;
//  if not CheckWordApp then
//  begin
//    //
//  end;
  FCloseOnFree := True;
end;

constructor TWordHelper.Create(WordApp: OleVariant);
begin
  FWordApp := WordApp;
  FCloseOnFree := False;  // 外部传入的WordApp对象，默认不管关闭
end;

destructor TWordHelper.Destroy;
begin
  if FCloseOnFree then
  begin
    try
      if FOpend then
        CloseFile;
      FWordApp.Quit(wdDoNotSaveChanges);
    except
      // 关闭过程中可能有异常
    end;
  end;

  FWordApp := NULL;
  FWordDoc := NULL;
  inherited;
end;

procedure TWordHelper.OpenFile(sFileName: String);
begin
  OpenFile(sFileName, False);
end;

procedure TWordHelper.OpenFile(sFileName: String; readOnly: Boolean);
begin
  FFileName := sFileName;
  if readOnly then
    FWordApp.Documents.Open(sFileName, False, True)
  else
    FWordApp.Documents.Open(sFileName);
  FWordApp.Options.SaveNormalPrompt := False;
  FWordDoc := FWordApp.ActiveDocument;
  FOpend := True;
end;

procedure TWordHelper.CloseFile();
begin
  FWordDoc.Close(wdDoNotSaveChanges);
  FOpend := False;
end;

procedure TWordHelper.SaveFile;
begin
  FWordDoc.Save;
end;

procedure TWordHelper.ShowWord;
begin
  FWordApp.Visible := True;
  FWordApp.Activate;
end;

function TWordHelper.AddComment(text: string; userName: string): OleVariant;
var
  range: OleVariant;
  cmt: OleVariant;
begin
  range := FWordApp.Selection.Range;
  cmt := FWordApp.ActiveDocument.Comments.Add(range, text);
  cmt.Initial := userName;
  Result := cmt;
end;

procedure TWordHelper.RemoveComments(userName: String);
var
  i: Integer;
  cmt: OleVariant;
begin
  // 批注索引是从1开始直到Count
  for i := FWordApp.ActiveDocument.Comments.Count downto 1 do
  begin
    cmt := FWordApp.ActiveDocument.Comments.Item(i);
    if cmt.Initial = userName then
    begin
      cmt.Delete;
    end;
  end;
end;

procedure TWordHelper.SelectText(text: String);
var
  upline,nomove,oleUnit,oleExtend: OleVariant;
  sFindText: WideString;
begin
  if Length(text) > WORD_FIND_MAX_LEN then
    sFindText := Copy(text, 1, WORD_FIND_MAX_LEN)
  else
    sFindText := text;
  oleUnit := wdStory;
  oleExtend := wdMove;
  //先定位到文档最前面
  wordApp.Selection.HomeKey(oleUnit, oleExtend);
  //再定位错误
  wordApp.Selection.Start := 0;
  wordApp.Selection.End := 0;
  wordApp.Selection.Find.ClearFormatting;
  wordApp.Selection.Find.Text:=sFindText;
  wordApp.Selection.Find.Replacement.Text:=sFindText;
  wordApp.Selection.Find.Forward:=True;
  wordApp.Selection.Find.Wrap:=1;
  wordApp.Selection.Find.Format:=False;
  wordApp.Selection.Find.MatchCase:=False;
  wordApp.Selection.Find.MatchWholeWord:=False;
  // 因为word不区分全角半角的"和'，因此特殊处理
  if (sFindText = '"') or (sFindText = '''') then
  begin
    wordApp.Selection.Find.MatchByte:=False;
    wordApp.Selection.Find.MatchWildcards:=True;
  end
  else
  begin
    wordApp.Selection.Find.MatchByte:=True;
    wordApp.Selection.Find.MatchWildcards:=False;
  end;
  wordApp.Selection.Find.MatchSoundsLike:=False;
  wordApp.Selection.Find.MatchAllWordForms:=False;
  wordApp.Selection.Find.Execute;
  // Find时word会跳转到结果处，但结果显示在窗口最顶部，向上滚动8格易于查看
  upline := 8;
  nomove := 0;
  wordApp.ActiveWindow.SmallScroll(nomove,upline,nomove,nomove);
end;

procedure TWordHelper.SelectTextWithTag(text: String);
begin
  SelectText(text);
  while wordApp.Selection.Find.Found do
  begin
    if (wordApp.Selection.Font.Shading.BackgroundPatternColor <> wdColorAutomatic)
      and (wordApp.Selection.Font.Shading.BackgroundPatternColor <> Integer(wdColorAutomatic)) then
    begin
      if Length(text) > WORD_FIND_MAX_LEN then
      begin
        wordApp.Selection.end := wordApp.Selection.start + Length(text);
      end;
      Exit;
    end;
    wordApp.Selection.Find.Execute;
  end;
end;

procedure TWordHelper.SelectAndColorText(text: String; iPos, iLen: Integer;
  color: Integer);
var
  i: Integer;
  findpos: array[0..255] of Integer;
  findcount: Integer;
begin
  SelectText(text);
  findcount := 0;
  while wordApp.Selection.Find.Found do
  begin
    findpos[findcount] := wordApp.Selection.start;
    Inc(findcount);
    if findcount > 255 then
      System.break;
    wordApp.Selection.Find.Execute;
  end;

  for i := 0 to findcount - 1 do
  begin
    wordApp.Selection.start := findpos[i] + iPos - 1;
    wordApp.Selection.end := findpos[i] + iPos - 1 + iLen;
    if wordApp.Selection.Text = text then
    begin
      wordApp.Selection.Font.Shading.BackgroundPatternColor := color;
    end;
  end;
end;

procedure TWordHelper.SelectAndColorSubText(selStart, selEnd: Integer;
  subText: string; color: Integer);
var
  nPos: Integer;
  text: String;
begin
  wordApp.Selection.start := selStart;
  wordApp.Selection.end := selEnd;
  text := FWordApp.Selection.Text;
  nPos := 0;
  repeat
    // 每次从上次找到的位置+1继续找
    nPos := PosFrom(subText, text, nPos + 1, False);
    if nPos > 0 then
    begin
      wordApp.Selection.start := selStart + nPos - 1;
      wordApp.Selection.end := selStart + nPos + Length(subText) - 1;
      wordApp.Selection.Font.Shading.BackgroundPatternColor := color;
    end;
  until nPos = 0;
end;

procedure TWordHelper.Find(sFind: string; resultList: TStringList);
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
  nStart      := FWordApp.Selection.Start;
  nEnd        := FWordApp.Selection.End;
  nVertical   := FWordApp.ActiveWindow.ActivePane.VerticalPercentScrolled;
  nHorizontal := FWordApp.ActiveWindow.ActivePane.HorizontalPercentScrolled;
  try

    FWordApp.Selection.Find.ClearFormatting;
    FWordApp.Selection.Find.Text    := sFind;
    FWordApp.Selection.Find.Forward := True;

    while FWordApp.Selection.Find.Execute do
    begin
      Range := FWordApp.Selection.Range;
      Range.Start := Range.Start - 20;
      Range.End   := Range.End + 20;

      nPos1 := FWordApp.Selection.Start;
      nPos2 := FWordApp.Selection.End;
      StrList.AddObject(Format('%d=%s', [nPos1, Range.Text]), TObject(nPos2));

      Application.ProcessMessages;
    end;

  finally
    //恢复原始位置
    if StrList.Count > 0 then
    begin
      FWordApp.ActiveWindow.ActivePane.VerticalPercentScrolled   := nVertical;
      FWordApp.ActiveWindow.ActivePane.HorizontalPercentScrolled := nHorizontal;
      FWordApp.Selection.Start := nStart;
      FWordApp.Selection.End   := nEnd;
    end;

    Screen.Cursor := crDefault;
  end;
end;

function TWordHelper.ExecuteControl(controlType: TOleEnum; tag: String): Boolean;
var
  cmdBar: OleVariant;
  cmdCtrl: OleVariant;
begin
  Result := False;
  cmdBar := FWordApp.CommandBars.Item['Standard'];
  cmdCtrl := cmdBar.FindControl(msoControlButton, 1, tag, 1, true);
  // TODO 这个VarisNull检测常常不好使
  if not VarisNull(cmdCtrl) then
  begin
    try
      cmdCtrl.Execute;
      Result := True;
    except
      Result := False;
    end;
  end;
end;

class function TWordHelper.getSelectionPos(wordApp: OleVariant): TSelectionPos;
begin
  Result.StartPos := wordApp.Selection.start;
  Result.EndPos := wordApp.Selection.end;
end;

class procedure TWordHelper.SetSelectionPos(wordApp: OleVariant;
  selPos: TSelectionPos);
begin
  wordApp.Selection.start := selPos.StartPos;
  wordApp.Selection.end := selPos.EndPos;
end;

end.
