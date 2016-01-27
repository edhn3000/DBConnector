unit U_DocInfo;

interface

uses
  Classes, SysUtils, Forms, Windows, PerlRegEx, DateUtils, OleCtrls, StrUtils,
  Dialogs, Variants, Office_TLB, Word_TLB, U_OfficeHelper, U_WordHelper, Generics.Collections;

type
  TDocInfo = class
  private
    FDoc: OleVariant;
    FWindow: OleVariant;
    FOfficeHelper: TOfficeHelper;
    FExpectParagraphCount: Integer;
    FImagePositions: TStringList;   // 图片所在段落
    FTablePositions: TStringList;   // 表格所在段落的起止

    procedure SetDoc(value: OleVariant);

  protected

    procedure parseImageInfo();
    procedure parseTableInfo();

    function GetImageIndexInParagraph(paNo: Integer; beginImgIndex, endImgIndex: Integer): Integer; overload;
    function GetImageIndexInParagraph(paNo: Integer): Integer; overload;
  public
    property Doc: OleVariant read FDoc write SetDoc;
    property ImagePositions: TStringList read FImagePositions;
    property TablePositions: TStringList read FTablePositions;


    constructor Create();
    destructor Destroy; override;

    // 初始化，识别表格、图片等信息
    procedure init();
    // 当段落数改变时自动重新初始化，循环中可能改变段落数时调用此方法更新数据
    procedure refresh();

    // 文档所含图片个数
    function ImageCount: Integer;
    // 第N个图片的起止位置
    function GetImagePos(imageIndex: Integer): TSelectionPos;
    // 获取给定区域里除图片外的文本子区域
    procedure GetTextSubRanges(paragraphPos: TSelectionPos; subRangeList: TStringList);
    // 给定段落是否有图片
    function ParagraphHasImage(paNo: Integer): Boolean;
    function IsSingleLineImageParagraph(paNo: Integer): Boolean;
    // 给定区域是否有图片
    function RangeHasImage(startPos, endPos: Integer; var imagePos: TSelectionPos): Boolean;

    // 给定段落是否表格段落，返回表格段落的结束段, paNo必须是表格首个段落
    function ParagraphHasTable(paNo: Integer; var paCount: Integer): Boolean;
    // 段落是否在表格中
    function ParagraphInTable(paNo: Integer): Boolean;
  end;

implementation

uses
  IniFiles;

{ TDocInfo }

constructor TDocInfo.Create();
begin
  FImagePositions := THashedStringList.Create;
  FTablePositions := THashedStringList.Create;
  FExpectParagraphCount := 0;
end;

destructor TDocInfo.Destroy;
begin
  FImagePositions.Free;
  FTablePositions.Free;
  if Assigned(FOfficeHelper) then
    FreeAndNil(FOfficeHelper);
  inherited;
end;

procedure TDocInfo.SetDoc(value: OleVariant);
begin
  FDoc := value;
  FWindow := FDoc.ActiveWindow;
  if Assigned(FOfficeHelper) then
    FreeAndNil(FOfficeHelper);
  FOfficeHelper := TWordHelper.Create(FDoc, FWindow);
end;

procedure TDocInfo.init;
begin
  parseImageInfo;
  parseTableInfo;
  FExpectParagraphCount := doc.Paragraphs.Count;
end;

procedure TDocInfo.refresh;
begin
  if FExpectParagraphCount <> doc.Paragraphs.Count then begin
    init();
  end;
end;

function TDocInfo.IsSingleLineImageParagraph(paNo: Integer): Boolean;
var
  sValue: String;
begin
  Result := False;
  sValue := FImagePositions.Values[IntToStr(paNo)];
  if sValue = '1' then begin
    Result := True;
  end;
end;

function TDocInfo.ParagraphHasImage(paNo: Integer): Boolean;
begin
  Result := False;
  if paNo < 0 then
    Exit;
  Result := GetImageIndexInParagraph(paNo) > 0;
end;

function TDocInfo.ParagraphHasTable(paNo: Integer; var paCount: Integer): Boolean;
var
  endNoStr: string;
begin
  Result := False;
  if paNo < 0 then
    Exit;

  endNoStr := TablePositions.Values[IntToStr(paNo)];
  if endNoStr <> '' then begin
    paCount := StrToInt(endNoStr);
    Result := True;
  end;
end;

function TDocInfo.ParagraphInTable(paNo: Integer): Boolean;
var
  i: Integer;
  paTable, paCount: Integer;
begin
  Result := False;
  if paNo < 0 then
    Exit;

  for I := 0 to TablePositions.Count - 1 do begin
    paTable := StrToInt(TablePositions.Names[i]);
    paCount := StrToInt(TablePositions.ValueFromIndex[i]);
    if (paNo >= paTable) and (paNo <= paTable + paCount) then begin
      Result := True;
      System.Break;
    end;
  end;
end;

function TDocInfo.GetImageIndexInParagraph(paNo, beginImgIndex,
  endImgIndex: Integer): Integer;
var
  beginIndex, endIndex, midIndex, paNoResult: Integer;
  bMatch: Boolean;
  imagePos: TSelectionPos;
begin
  Result := 0;
  beginIndex := beginImgIndex;
  endIndex := endImgIndex;
  midIndex := Trunc((beginIndex + endIndex) / 2);

  imagePos := TOfficeHelper.GetRangePos(doc.InlineShapes.Item(midIndex).Range);

  paNoResult := FOfficeHelper.GetParagraphNoByPos(imagePos.StartPos);
  bMatch := paNoResult = paNo;
  if bMatch then begin
    Result := midIndex;
    Exit
  end else begin
    if doc.Paragraphs.Item(paNoResult).Range.Start > imagePos.StartPos then
      endIndex := midIndex - 1
    else
      beginIndex := midIndex + 1;
  end;
  if beginIndex > endIndex then
    Exit;
  Result := GetImageIndexInParagraph(paNo, beginIndex, endIndex);
end;

function TDocInfo.GetImageIndexInParagraph(paNo: Integer): Integer;
var
  beginIndex, endIndex: Integer;
begin
  beginIndex := 1;
  endIndex := doc.InlineShapes.Count;

  Result := GetImageIndexInParagraph(paNo, beginIndex, endIndex);
end;

function TDocInfo.GetImagePos(imageIndex: Integer): TSelectionPos;
begin
  Result := TOfficeHelper.GetRangePos(doc.InlineShapes.Item(imageIndex).Range);
end;

procedure TDocInfo.GetTextSubRanges(paragraphPos: TSelectionPos; subRangeList: TStringList);
var
  i: Integer;
  selPos, imagePos: TSelectionPos;
begin
  selPos := paragraphPos;
  for i := 1 to doc.InlineShapes.Count do begin
    imagePos := TOfficeHelper.GetRangePos(doc.InlineShapes.Item(i).Range);
    if (selPos.StartPos <= imagePos.StartPos)
       and (selPos.EndPos >= imagePos.EndPos) then begin
      subRangeList.Add(Format('%d=%d', [selpos.StartPos, imagePos.StartPos]));
      selPos.StartPos := imagePos.EndPos;
    end;
  end;
  if selPos.StartPos < selPos.EndPos then
    subRangeList.Add(Format('%d=%d', [selpos.StartPos, selPos.EndPos]));
  if subRangeList.Count > 0 then
end;

function TDocInfo.ImageCount: Integer;
begin
  Result := doc.InlineShapes.Count;
end;

procedure TDocInfo.parseImageInfo;
var
  i, j, m, picLines: Integer;
  selPos, imgPos, paPos, textPos: TSelectionPos;
  range: OleVariant;
  imgPosList: TStringList;
  subTextRanges: TStringList;
  singleImageParagraph: Boolean;
  docRange: OleVariant;
begin
  if doc.InlineShapes.Count = 0 then
    Exit;
  imgPosList := TStringList.Create;
  try
    // 全文所有图片的起止位置，由于文档稍有改变时位置信息就会变动，这里仅临时使用，不在内存中保留此信息
    for i := 1 to doc.InlineShapes.Count do begin
      range := doc.InlineShapes.Item(i).Range;
      selpos := TOfficeHelper.GetRangePos(range);
      imgPosList.Add(Format('%d=%d', [selpos.StartPos, selpos.EndPos]));
    end;

    // 记录包含图片的段落
    FImagePositions.Clear;
    for i := 1 to doc.Paragraphs.Count do begin
      paPos := TOfficeHelper.GetRangePos(doc.Paragraphs.Item(i).Range);
      for j := 0 to imgPosList.Count - 1 do begin
        imgPos.StartPos := StrToInt(imgPosList.Names[j]);
        imgPos.EndPos := StrToInt(imgPosList.ValueFromIndex[j]);
        if paPos.ContainSelectionPos(imgPos) then begin
          subTextRanges := TStringList.Create;
          try
            GetTextSubRanges(paPos, subTextRanges);
            singleImageParagraph := True;  // 是否仅有图片无文字
            for m := 0 to subTextRanges.Count - 1 do begin
              textPos.StartPos := StrToInt(subTextRanges.Names[m]);
              textPos.EndPos := StrToInt(subTextRanges.ValueFromIndex[m]);
              if textPos.StartPos = textPos.EndPos then
                Continue;
              docRange := doc.Range(textPos.StartPos, textPos.EndPos);
              if Trim(docRange.Text) = '' then
                Continue;
              singleImageParagraph := False;
            end;
          finally
            subTextRanges.Free;
          end;
          if singleImageParagraph then begin
            picLines := 1   // 单独成段的图片
          end else begin
            selPos.StartPos := paPos.EndPos - 1;
            selPos.EndPos := selPos.EndPos;
            TOfficeHelper.SetSelectionPos(FWindow, selPos);
            FWindow.Selection.HomeKey(wdLine, wdMove);
            if FWindow.Selection.Start = paPos.StartPos then
              picLines := 1   // 单独成段的图片，或图片所在段落只占一行
            else
              picLines := 2;  // 图片所在段落有多行
          end;

          FImagePositions.Add(Format('%d=%d', [i, picLines]));
          System.Break;
        end;
      end;
    end;
  finally
    imgPosList.Free;
  end;
end;

//procedure TDocInfo.parseTableInfo;
//var
//  s, content: string;
//  I: integer;
//  tabs: OleVariant;
//begin
//  if doc.Tables.Count = 0 then
//    Exit;
//
//  FWindow.Selection.WholeStory;
//  content := FWindow.Selection.text;
//  // 获取文档中所有的table
//  tabs := doc.Tables;
//
//  // 对有表格的文书进行格式化
//  s := '';
//  // content := StringReplace(content,#$B,#13#10,[rfReplaceAll]);//软回车
//
//  // 找出每个table在文档中的位置
//  for I := 1 to tabs.Count do
//  begin
//    s := s + getStrPosInfo(content, tabs.Item(I).Range.text);
//  end;
//
//  FTablePositions.text := s;
//
//  s := '';
//  // 除去重复表格值
//  for I := 0 to FTablePositions.Count - 1 do
//  begin
//    if Pos(FTablePositions.Strings[I], s) = 0 then
//      s := s + FTablePositions.Strings[I] + #13#10;
//  end;
//
//  FTablePositions.text := s;
//  FExpectParagraphCount := doc.Paragraphs.Count;
//end;

procedure TDocInfo.parseTableInfo;
var
  i, j, paCount: integer;
  lastTablePos: Integer;
  tablePos, selPos: TSelectionPos;
begin
  if doc.Tables.Count = 0 then
    Exit;

  FTablePositions.Clear;
  lastTablePos := 1;
  for i := 1 to doc.Tables.Count do begin
    tablePos := TOfficeHelper.GetRangePos(doc.Tables.Item(i).Range);
    paCount := doc.Tables.Item(i).Range.Paragraphs.Count;

    j := lastTablePos;
    while j <= doc.Paragraphs.Count do begin
      selPos := TOfficeHelper.GetRangePos(doc.Paragraphs.Item(j).Range);
      if selPos.StartPos = tablePos.StartPos then begin
        FTablePositions.Add(Format('%d=%d', [j, paCount]));
        lastTablePos := j + paCount;
        System.Break;
      end;
      Inc(j);
    end;
  end;
end;

function TDocInfo.RangeHasImage(startPos, endPos: Integer; var imagePos: TSelectionPos): Boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 1 to doc.InlineShapes.Count do begin
    imagePos := TOfficeHelper.GetRangePos(doc.InlineShapes.Item(i).Range);
    if (startPos <= imagePos.StartPos) and (endPos >= imagePos.EndPos) then begin
      Result := True;
      System.Break;
    end;
  end;
end;

end.
