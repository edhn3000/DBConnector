unit U_OfficeHelper;
{
 @author  fengyq
 @comment Office���ֵ�Ԫ
          ��Ҫ����Variant���ӿڲ������ܶ������Word��Wpsͨ����
          ��Officeʵ���ࣨWord��Wps����ʵ�ֳ��󷽷�����Ҫ�Ǳ���Ĵ��������Բ���
 @version 1.0
 @version 2015/11/09
}

interface

uses
  SysUtils, Variants, Classes, Controls, Forms, ComCtrls, ComObj, ActiveX,
  Office_TLB, Word_TLB, Windows;

type
  // ����ʹ��Selection.End��ɵ�Ԫ�﷨��ʾ���Ĵ���
  TSelectionPos = record
    StartPos: Integer;
    EndPos: Integer;
  end;

  function InSelectionPos(selPosMain, selPosSub: TSelectionPos): Boolean;

const
  WORD_FIND_MAX_LEN = 255;   // ��������ʱ����󳤶�

type
  { TOfficeHelper }
  TOfficeHelper = class
  protected
    FUseActiveApp: Boolean;   // �Ƿ�ʹ���Ѽ����Application����
    FIsInstalled: Boolean;    // Office�Ƿ��Ѱ�װ ʵ�����ʼ��ʱ���丳ֵ
    FApplicatioin: OleVariant;// Office��Application
    FDocument: OleVariant;    // ��ǰ�ĵ�
    FWindow: OleVariant;      // ��ǰ����
    FOpend: Boolean;          // �Ƿ��Ѵ��ĵ�
    FCloseOnFree: Boolean;    // �ͷŶ���ʱ�Ƿ�ر��ĵ���Application
    FFileName: String;        // �ĵ����ļ������ⲿ����

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

    // �����Ѽ����Application����
    constructor Create(ActiveApp: OleVariant); overload; virtual;
    destructor Destroy;override;

    // Office�Ƿ��Ѱ�װ
    function IsInstalled: Boolean; virtual;

    // ���ĵ�
    procedure NewFile(sFileName: String); virtual;abstract;
    procedure OpenFile(sFileName: String; readOnly: Boolean);overload; virtual;abstract;
    procedure OpenFile(sFileName: String);overload;
    // �ر��ĵ�
    procedure CloseFile(bSave: Boolean = false);virtual;abstract;
    // ��ʾ�ĵ���Ĭ�ϴ�ʱ������ʾ�ĵ�
    procedure ShowDocument(); virtual;
    // �������ݣ�����������resultList��
    procedure Find(sFind: string; resultList: TStringList); virtual;

{-------------------------------------------------------------------------------
  ������:    AddComment ������ע
  ����:      fengyq
  ����:      2014.10.17
  ����:      text: string; userName: string
  ����ֵ:    OleVariant ��ע����Comment����
-------------------------------------------------------------------------------}
    function AddComment(text: string; userName: string): OleVariant; virtual;
    procedure RemoveComments(userName: String); virtual;

{-------------------------------------------------------------------------------
  ������:    SelText, ѡ��Word���ı�
  ����:      fengyq
  ����:      2014.09.27
  ����:      text: String
  ����ֵ:    ��
-------------------------------------------------------------------------------}
    function SelectText(text: String): Boolean; overload; virtual;
    function SelectText(startPos, endPos: LongInt): Boolean; overload;virtual;
    function SelectTextWithTag(text: String): Boolean;virtual;
    function SelectTextInRange(text: String; iStart,iEnd:Integer): Boolean;virtual;
{-------------------------------------------------------------------------------
  ������:    SelAndColorText ѡ��Word���ı����ø�����ɫ�����
             ����Ĳ����ǿ�iPos, iLenָ��
  ����:      fengyq
  ����:      2014.09.27
  ����:      text: String; iPos, iLen: Integer; color: Integer
  ����ֵ:    ��
-------------------------------------------------------------------------------}
    procedure SelectAndColorText(text: String; iPos, iLen: Integer; color: Integer);virtual;

    procedure SelectAndColorSubText(selStart, selEnd: Integer; subText: string; color: Integer);virtual;

    {�����з����滻Ϊ^p}
    function ReplaceWordNewLine(sText: String): string;

{-------------------------------------------------------------------------------
  ������:    ExecuteControl ִ�пؼ����ܣ���������������ť
  ����:      fengyq
  ����:      2014.10.20
  ����:      controlType: TOleEnum; ������msoControlButton
             tag: String �ؼ�Tag�����ڲ���
  ����ֵ:    Boolean
-------------------------------------------------------------------------------}
    function ExecuteControl(controlType: TOleEnum; tag: String): Boolean;virtual;abstract;


    // ��ȡѡ�����������������߱�start��end����
    class function GetRangePos(range: OleVariant): TSelectionPos;
    // ��ȡѡ�����������������߱�start��end����
    class procedure SetRangePos(range: OleVariant; selPos: TSelectionPos);

    // ��ȡѡ�������봫��ѡȡ��ӵ���ߣ���Window��Application
    class function GetSelectionPos(SelectionOwner: OleVariant): TSelectionPos;
    // ����ѡ�������봫��ѡȡ��ӵ���ߣ���Window��Application
    class procedure SetSelectionPos(SelectionOwner: OleVariant; selPos: TSelectionPos);

    // �ı�ѡȡ��ͷ��moveStep���ƶ�������start+moveStep
    class procedure MoveSelectionStart(selection: OleVariant; moveStep: Integer);
    // �ı�ѡȡ��β��moveStep���ƶ�������end+moveStep
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
  FCloseOnFree := False;  // �ⲿ�����WordApp����Ĭ�ϲ��ܹر�
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
          // û�д򿪵�Application�������µ�
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
    if Copy(text, Length(sFindText), 2) = '^p' then    // ��ֹ^p���ָ���
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

function TOfficeHelper.ReplaceWordNewLine(sText: String): string;
var
  s: string;
begin
  // word�л��з��滻Ϊ^p���ڲ���
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
  // ��ע�����Ǵ�1��ʼֱ��Count
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
  //�ȶ�λ���ĵ���ǰ��
  FWindow.Selection.HomeKey(oleUnit, oleExtend);
  //�ٶ�λ����
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
  // ��Ϊword������ȫ�ǰ�ǵ�"��'��������⴦��
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

  // Findʱword����ת����������������ʾ�ڴ�����������Ϲ���8�����ڲ鿴
  upline := 8;
  nomove := 0;
  FWindow.SmallScroll(nomove,upline,nomove,nomove);
  Result := FWindow.Selection.Find.Found;

  // ��չѡ��
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
    // ����ѡ��
    selPos := TOfficeHelper.GetSelectionPos(FWindow);
    if (selPos.EndPos - selPos.StartPos + 1) < Length(text) then begin
      selPos.EndPos := selPos.StartPos + Length(text) - 1;
      TOfficeHelper.SetSelectionPos(FWindow, selPos);
    end;

    // �ҵ�����
    if (ReplaceWordNewLine(FWindow.Selection.Text) = text) then begin
      Result := True;
      // ����Ƿ����������ѡ�и���
      if (FWindow.Selection.Font.Shading.BackgroundPatternColor <> wdColorAutomatic)
        and (FWindow.Selection.Font.Shading.BackgroundPatternColor <> Integer(wdColorAutomatic)) then begin
        Exit;
      end;
    end;
    //Ϊ�˷�ֹ��ѭ��
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
  FWindow.Selection.Find.Wrap := wdFindStop;//wdFindContinue 1 ������������Ŀ�ʼ���߽�βʱ������ִ�в��Ҳ���
                                                     //wdFindStop 0 ����������Χ�Ŀ�ʼ���߽�βʱ��ִֹͣ�в��Ҳ���
  FWindow.Selection.Find.Format := False;
  FWindow.Selection.Find.MatchCase := False;
  FWindow.Selection.Find.MatchWholeWord :=False;
  // ��Ϊword������ȫ�ǰ�ǵ�"��'��������⴦��
  if (sFindText = '"') or (sFindText = '''') then
  begin
    FWindow.Selection.Find.MatchByte:=False;
    FWindow.Selection.Find.MatchWildcards:=True;  //����ָ��ͨ����������߼���������
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

    if iStart > 100 then   //������ҷ�Χ
      SelectTextInRange(text, iStart - 100, iEnd + 100)
    else
      SelectTextInRange(text, 0, 0);
  end;
  Result := FWindow.Selection.Find.Found;
  // ��չѡ��
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
    // ÿ�δ��ϴ��ҵ���λ��+1������
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
  //��סԭʼ������λ��
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
    //�ָ�ԭʼλ��
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
