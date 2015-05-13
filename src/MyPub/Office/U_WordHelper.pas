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
    FApplicatioin: OleVariant;
    FActiveDocument: OleVariant;
//    FWordApp: WordApplication;  // ��FApplicatioinת��Ϊ�ṹ���󣬷������
    FOpend: Boolean;
    FCloseOnFree: Boolean;
    FFileName: String;
    function GetShowComment(): Boolean;
    procedure SetShowComment(value: Boolean);
  protected
    function CheckWordApp(): Boolean;
  public
    property CloseOnFree: Boolean read FCloseOnFree write FCloseOnFree;
    property Application: OleVariant read FApplicatioin;
//    property WordApp: WordApplication read FWordApp;
    property Opend: Boolean read FOpend;
    property ShowComment: Boolean read GetShowComment write SetShowComment;

    constructor Create;overload;
    constructor Create(useActiveApp: Boolean);overload;
    constructor Create(WordApp: OleVariant);overload;
    destructor Destroy;override;

    { ��Wrod�ĵ� }
    procedure NewFile(sFileName: String);
    procedure OpenFile(sFileName: String);overload;
    procedure OpenFile(sFileName: String; readOnly: Boolean);overload;
    { �ر�Word}
    procedure CloseFile(bSave: Boolean = false);
    { ��ʾWord�ĵ���Ĭ�ϴ�ʱ������ʾ�ĵ� }
    procedure ShowWord();
    { �������ݣ�����������resultList�� }
    procedure Find(sFind: string; resultList: TStringList);

{-------------------------------------------------------------------------------
  ������:    AddComment ������ע
  ����:      fengyq
  ����:      2014.10.17
  ����:      text: string; userName: string
  ����ֵ:    OleVariant ��ע����Comment����
-------------------------------------------------------------------------------}
    function AddComment(text: string; userName: string): OleVariant;
    procedure RemoveComments(userName: String);

{-------------------------------------------------------------------------------
  ������:    SelText, ѡ��Word���ı�
  ����:      fengyq
  ����:      2014.09.27
  ����:      text: String
  ����ֵ:    ��
-------------------------------------------------------------------------------}
    function SelectText(text: String): OleVariant; overload;
    function SelectText(startPos, endPos: LongInt): OleVariant; overload;
    procedure SelectTextWithTag(text: String);
{-------------------------------------------------------------------------------
  ������:    SelAndColorText ѡ��Word���ı����ø�����ɫ�����
             ����Ĳ����ǿ�iPos, iLenָ��
  ����:      fengyq
  ����:      2014.09.27
  ����:      text: String; iPos, iLen: Integer; color: Integer
  ����ֵ:    ��
-------------------------------------------------------------------------------}
    procedure SelectAndColorText(text: String; iPos, iLen: Integer; color: Integer);

    procedure SelectAndColorSubText(selStart, selEnd: Integer; subText: string; color: Integer);

{-------------------------------------------------------------------------------
  ������:    ExecuteControl ִ�пؼ����ܣ���������������ť
  ����:      fengyq
  ����:      2014.10.20
  ����:      controlType: TOleEnum; ������msoControlButton
             tag: String �ؼ�Tag�����ڲ���
  ����ֵ:    Boolean
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
  // �������
  if Length(sSub) > Length(S) then
    Exit;

  nIndexSub  := 1;
  nIndexMain := nFrom;  // ��ʼλ��

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
      nIndexSave := nIndexMain;    // ����ԭλ��

      while (nIndexSub<= Length(sSub))
            and ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
                 or (sSub[nIndexSub]=S[nIndexMain])) do
      begin
        Inc(nIndexMain);
        Inc(nIndexSub);
      end;

      if nIndexSub = Length(sSub)+1 then      // �ҵ�
      begin
        Result := nIndexSave;       // �����µ�λ�þ��ǽ��
        System.Break;
      end
      else                          // û�ҵ�
      begin
        nIndexMain := nIndexSave + 1;  // �������
        nIndexSub := 1;
      end;
    end;
  end;
end;

function TWordHelper.CheckWordApp(): Boolean;
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

function TWordHelper.GetShowComment(): Boolean;
begin
  Result := FApplicatioin.ActiveWindow.View.ShowRevisionsAndComments;
end;

procedure TWordHelper.SetShowComment(value: Boolean);
begin
  FApplicatioin.ActiveWindow.View.ShowRevisionsAndComments := value;
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
      FApplicatioin := GetActiveOleObject('Word.Application');
    except
      on e: EOleSysError do
        FApplicatioin := CreateOleObject('Word.Application');
    end
  end else begin
    FApplicatioin := CreateOleObject('Word.Application');
  end;
  // �����浽Normal.dotm�������ĵ����ظ��򿪺��ڹر�ʱ����
  FApplicatioin.Options.SaveNormalPrompt := False;
//  FWordApp := WordApplication(IDispatch(FApplicatioin));
//  if not CheckWordApp then
//  begin
//    //
//  end;
  FCloseOnFree := True;
end;

constructor TWordHelper.Create(WordApp: OleVariant);
begin
  FApplicatioin := WordApp;
  FCloseOnFree := False;  // �ⲿ�����WordApp����Ĭ�ϲ��ܹر�
end;

destructor TWordHelper.Destroy;
begin
  if FCloseOnFree then
  begin
    try
      if FOpend then
        CloseFile(False);
      FApplicatioin.Quit(wdDoNotSaveChanges);
    except  // �رչ����п������쳣
      on e: Exception do
        OutputDebugString(PChar('TWordHelper.Free�ر�Word����' + e.Message));
    end;
  end;

  FApplicatioin := NULL;
  FActiveDocument := NULL;
  inherited;
end;

procedure TWordHelper.NewFile(sFileName: String);
begin
  FFileName := sFileName;
  FApplicatioin.Documents.Add;
//  FWordApp := WordApplication(IDispatch(FApplicatioin));
  FActiveDocument := FApplicatioin.ActiveDocument;
end;

procedure TWordHelper.OpenFile(sFileName: String);
begin
  OpenFile(sFileName, False);
end;

procedure TWordHelper.OpenFile(sFileName: String; readOnly: Boolean);
begin
  FFileName := sFileName;
  if readOnly then
    FApplicatioin.Documents.Open(sFileName, False, True)
  else
    FApplicatioin.Documents.Open(sFileName);
  FApplicatioin.Options.SaveNormalPrompt := False;
//  FWordApp := WordApplication(IDispatch(FApplicatioin));
  FActiveDocument := FApplicatioin.ActiveDocument;
  FOpend := True;
end;

procedure TWordHelper.CloseFile(bSave: Boolean);
begin
  if bSave then begin
    FActiveDocument.SaveAs(FFileName);
    FActiveDocument.Close(wdDoNotSaveChanges);
  end else begin
    FActiveDocument.Close(wdDoNotSaveChanges);
  end;
  FOpend := False;
end;

procedure TWordHelper.ShowWord;
begin
  FApplicatioin.Visible := True;
  FApplicatioin.Activate;
end;

function TWordHelper.AddComment(text: string; userName: string): OleVariant;
var
  range: OleVariant;
  cmt: OleVariant;
begin
  range := FApplicatioin.Selection.Range;
  cmt := FApplicatioin.ActiveDocument.Comments.Add(range, text);
  cmt.Initial := userName;
  Result := cmt;
end;

procedure TWordHelper.RemoveComments(userName: String);
var
  i: Integer;
  cmt: OleVariant;
begin
  // ��ע�����Ǵ�1��ʼֱ��Count
  for i := FApplicatioin.ActiveDocument.Comments.Count downto 1 do
  begin
    cmt := FApplicatioin.ActiveDocument.Comments.Item(i);
    if cmt.Initial = userName then
    begin
      cmt.Delete;
    end;
  end;
end;

function TWordHelper.SelectText(text: String): OleVariant;
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
  //�ȶ�λ���ĵ���ǰ��
  FApplicatioin.Selection.HomeKey(oleUnit, oleExtend);
  //�ٶ�λ����
  FApplicatioin.Selection.Start := 0;
  FApplicatioin.Selection.End := 0;
  FApplicatioin.Selection.Find.ClearFormatting;
  FApplicatioin.Selection.Find.Text:=sFindText;
  FApplicatioin.Selection.Find.Replacement.Text:=sFindText;
  FApplicatioin.Selection.Find.Forward:=True;
  FApplicatioin.Selection.Find.Wrap:=1;
  FApplicatioin.Selection.Find.Format:=False;
  FApplicatioin.Selection.Find.MatchCase:=False;
  FApplicatioin.Selection.Find.MatchWholeWord:=False;
  // ��Ϊword������ȫ�ǰ�ǵ�"��'��������⴦��
  if (sFindText = '"') or (sFindText = '''') then
  begin
    FApplicatioin.Selection.Find.MatchByte:=False;
    FApplicatioin.Selection.Find.MatchWildcards:=True;
  end
  else
  begin
    FApplicatioin.Selection.Find.MatchByte:=True;
    FApplicatioin.Selection.Find.MatchWildcards:=False;
  end;
  FApplicatioin.Selection.Find.MatchSoundsLike:=False;
  FApplicatioin.Selection.Find.MatchAllWordForms:=False;
  FApplicatioin.Selection.Find.Execute;
  // Findʱword����ת����������������ʾ�ڴ�����������Ϲ���8�����ڲ鿴
  upline := 8;
  nomove := 0;
  FApplicatioin.ActiveWindow.SmallScroll(nomove,upline,nomove,nomove);
  Result := FApplicatioin.Selection;
end;

function TWordHelper.SelectText(startPos, endPos: LongInt): OleVariant;
begin
  FApplicatioin.Selection.Start := startPos;
  FApplicatioin.Selection.End := endPos;
  Result := FApplicatioin.Selection;
end;

procedure TWordHelper.SelectTextWithTag(text: String);
begin
  SelectText(text);
  while FApplicatioin.Selection.Find.Found do
  begin
    if (FApplicatioin.Selection.Font.Shading.BackgroundPatternColor <> wdColorAutomatic)
      and (FApplicatioin.Selection.Font.Shading.BackgroundPatternColor <> Integer(wdColorAutomatic)) then
    begin
      if Length(text) > WORD_FIND_MAX_LEN then
      begin
        FApplicatioin.Selection.end := FApplicatioin.Selection.start + Length(text);
      end;
      Exit;
    end;
    FApplicatioin.Selection.Find.Execute;
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
  while FApplicatioin.Selection.Find.Found do
  begin
    findpos[findcount] := FApplicatioin.Selection.start;
    Inc(findcount);
    if findcount > 255 then
      System.break;
    FApplicatioin.Selection.Find.Execute;
  end;

  for i := 0 to findcount - 1 do
  begin
    FApplicatioin.Selection.start := findpos[i] + iPos - 1;
    FApplicatioin.Selection.end := findpos[i] + iPos - 1 + iLen;
    if FApplicatioin.Selection.Text = text then
    begin
      FApplicatioin.Selection.Font.Shading.BackgroundPatternColor := color;
    end;
  end;
end;

procedure TWordHelper.SelectAndColorSubText(selStart, selEnd: Integer;
  subText: string; color: Integer);
var
  nPos: Integer;
  text: String;
begin
  FApplicatioin.Selection.start := selStart;
  FApplicatioin.Selection.end := selEnd;
  text := FApplicatioin.Selection.Text;
  nPos := 0;
  repeat
    // ÿ�δ��ϴ��ҵ���λ��+1������
    nPos := PosFrom(subText, text, nPos + 1, False);
    if nPos > 0 then
    begin
      FApplicatioin.Selection.start := selStart + nPos - 1;
      FApplicatioin.Selection.end := selStart + nPos + Length(subText) - 1;
      FApplicatioin.Selection.Font.Shading.BackgroundPatternColor := color;
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
  //��סԭʼ������λ��
  nStart      := FApplicatioin.Selection.Start;
  nEnd        := FApplicatioin.Selection.End;
  nVertical   := FApplicatioin.ActiveWindow.ActivePane.VerticalPercentScrolled;
  nHorizontal := FApplicatioin.ActiveWindow.ActivePane.HorizontalPercentScrolled;
  try

    FApplicatioin.Selection.Find.ClearFormatting;
    FApplicatioin.Selection.Find.Text    := sFind;
    FApplicatioin.Selection.Find.Forward := True;

    while FApplicatioin.Selection.Find.Execute do
    begin
      Range := FApplicatioin.Selection.Range;
      Range.Start := Range.Start - 20;
      Range.End   := Range.End + 20;

      nPos1 := FApplicatioin.Selection.Start;
      nPos2 := FApplicatioin.Selection.End;
      StrList.AddObject(Format('%d=%s', [nPos1, Range.Text]), TObject(nPos2));

      Application.ProcessMessages;
    end;

  finally
    //�ָ�ԭʼλ��
    if StrList.Count > 0 then
    begin
      FApplicatioin.ActiveWindow.ActivePane.VerticalPercentScrolled   := nVertical;
      FApplicatioin.ActiveWindow.ActivePane.HorizontalPercentScrolled := nHorizontal;
      FApplicatioin.Selection.Start := nStart;
      FApplicatioin.Selection.End   := nEnd;
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
  cmdBar := FApplicatioin.CommandBars.Item['Standard'];
  cmdCtrl := cmdBar.FindControl(msoControlButton, 1, tag, 1, true);
  // TODO ���VarisNull��ⳣ������ʹ
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
