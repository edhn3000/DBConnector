unit U_DocInfo;

interface

uses
  Classes, SysUtils, Forms, Windows, OleCtrls,
  Dialogs, Variants, Word_TLB,
  U_OfficeHelper;

type
  TGetDocObjectPos = function (objectIndex: Integer): TSelectionPos of object;
  { TDocInfo }
  TDocInfo = class
  private
    FOfficeHelper: TOfficeHelper;
    FExpectParagraphCount: Integer;
    FImagePositions: TStringList;   // 图片所在段落，key=图片段落号，value=段落是否单行（1为true）
    FTablePositions: TStringList;   // 表格所在段落的起止，key=表格段落号，value=表格段落数

  protected

    // 文档中二分法查找对象的函数
    function BinarySearchInDoc(paNo: Integer; beginIndex, endIndex: Integer; GetPosMethod: TGetDocObjectPos): Integer;
  public
    property OfficeHelper: TOfficeHelper read FOfficeHelper write FOfficeHelper;
//    property ImagePositions: TStringList read FImagePositions;
    property TablePositions: TStringList read FTablePositions;


    constructor Create();
    destructor Destroy; override;

    procedure parseImageInfo();
    procedure parseTableInfo();
    // 初始化，识别表格、图片等信息
    procedure init();
    // 当段落数改变时自动重新初始化，循环中可能改变段落数时调用此方法更新数据
    procedure refresh();

    // 文档所含图片个数
    function ImageCount: Integer;
    // 文档所含表格个数
    function TableCount: Integer;

    // 第N个图片的起止位置
    function GetImagePos(imageIndex: Integer): TSelectionPos;
    function GetTablePos(tableIndex: Integer): TSelectionPos;

    // 指定的段落中查找图片并返回图片的index
    function GetImageIndexInParagraph(paNo: Integer): Integer;
    // 指定的段落中查找表格并返回表格的index
    function GetTableIndexInParagraph(paNo: Integer): Integer;

    // 获取给定区域里除图片外的文本子区域
    procedure GetTextSubRanges(paragraphPos: TSelectionPos; subRangeList: TStringList);
    // 给定段落是否有图片
    function ParagraphHasImage(paNo: Integer): Boolean;
    // 是否单独成段的图片，相当于“段落只占一行&段落包含图片”
    function IsSingleLineImageParagraph(paNo: Integer): Boolean;

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
  inherited;
end;

procedure TDocInfo.init;
begin
  parseImageInfo;
  parseTableInfo;
  FExpectParagraphCount := OfficeHelper.Document.Paragraphs.Count;
end;

procedure TDocInfo.refresh;
begin
  if FExpectParagraphCount <> OfficeHelper.Document.Paragraphs.Count then begin
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
  if (paNo <= 0) or (paNo > OfficeHelper.Document.Paragraphs.Count) then
    Exit;
  Result := GetImageIndexInParagraph(paNo) > 0;
end;

function TDocInfo.ParagraphHasTable(paNo: Integer; var paCount: Integer): Boolean;
var
  tableIndex: Integer;
begin
  Result := False;
  if (paNo <= 0) or (paNo > OfficeHelper.Document.Paragraphs.Count) then
    Exit;
  if not ParagraphInTable(paNo) then
    Exit;

  tableIndex := GetTableIndexInParagraph(paNo);
  if tableIndex <= 0 then
    Exit;

  paCount := OfficeHelper.Document.Tables.Item(tableIndex).Range.Paragraphs.Count;
  Result := True;
end;

function TDocInfo.ParagraphInTable(paNo: Integer): Boolean;
begin
  Result := False;
  if (paNo <= 0) or (paNo > OfficeHelper.Document.Paragraphs.Count) then
    Exit;

  Result := OfficeHelper.Document.Paragraphs.Item(paNo).Range.Information[wdWithInTable];
end;

function TDocInfo.BinarySearchInDoc(paNo, beginIndex, endIndex: Integer;
  GetPosMethod: TGetDocObjectPos): Integer;
var
  midIndex, midPaNo: Integer;
  bMatch: Boolean;
  midObjectPos: TSelectionPos;
begin
  Result := 0;
  midIndex := Trunc((beginIndex + endIndex) / 2);
  if midIndex = 0 then
    Exit;

  // 获得midindex对象所在的段落
  midObjectPos := GetPosMethod(midIndex);
  midPaNo := FOfficeHelper.GetParagraphNoByPos(midObjectPos.StartPos);

  // 如果段落和传入段落相同的匹配成功
  bMatch := midPaNo = paNo;
  if bMatch then begin
    Result := midIndex;
    Exit;
  end else begin
//    if doc.Paragraphs.Item(paNo).Range.Start < midObjectPos.StartPos then
    if paNo < midPaNo then
      endIndex := midIndex - 1
    else
      beginIndex := midIndex + 1;
  end;
  if beginIndex > endIndex then
    Exit;
  Result := BinarySearchInDoc(paNo, beginIndex, endIndex, GetPosMethod);
end;

function TDocInfo.GetImageIndexInParagraph(paNo: Integer): Integer;
var
  beginIndex, endIndex: Integer;
begin
  beginIndex := 1;
  endIndex := OfficeHelper.Document.InlineShapes.Count;

  Result := BinarySearchInDoc(paNo, beginIndex, endIndex, GetImagePos);
end;

function TDocInfo.GetTableIndexInParagraph(paNo: Integer): Integer;
var
  beginIndex, endIndex: Integer;
begin
  beginIndex := 1;
  endIndex := OfficeHelper.Document.Tables.Count;

  Result := BinarySearchInDoc(paNo, beginIndex, endIndex, GetTablePos);
end;

function TDocInfo.GetTablePos(tableIndex: Integer): TSelectionPos;
begin
  if (tableIndex < 1) or (tableIndex > OfficeHelper.Document.Tables.Count) then
    raise Exception.Create(Format('tableIndex %d over range!', [tableIndex]));

  Result := TOfficeHelper.GetRangePos(OfficeHelper.Document.Tables.Item(tableIndex).Range);
end;
function TDocInfo.GetImagePos(imageIndex: Integer): TSelectionPos;
begin
  if (imageIndex < 1) or (imageIndex > OfficeHelper.Document.InlineShapes.Count) then
    raise Exception.Create(Format('imageIndex %d over range!', [imageIndex]));

  Result := TOfficeHelper.GetRangePos(OfficeHelper.Document.InlineShapes.Item(imageIndex).Range);
end;

procedure TDocInfo.GetTextSubRanges(paragraphPos: TSelectionPos; subRangeList: TStringList);
var
  i: Integer;
  selPos, imagePos: TSelectionPos;
begin
  selPos := paragraphPos;
  for i := 1 to OfficeHelper.Document.InlineShapes.Count do begin
    imagePos := TOfficeHelper.GetRangePos(OfficeHelper.Document.InlineShapes.Item(i).Range);
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
  Result := OfficeHelper.Document.InlineShapes.Count;
end;

function TDocInfo.TableCount: Integer;
begin
  Result := OfficeHelper.Document.Tables.Count;
end;

procedure TDocInfo.parseImageInfo;
var
  i, imageIndex, m, picLines: Integer;
  selPos, imgPos, paPos, textPos: TSelectionPos;
//  range: OleVariant;
//  imgPosList: TStringList;
  subTextRanges: TStringList;
  singleImageParagraph: Boolean;
  docRange: OleVariant;
begin
  if OfficeHelper.Document.InlineShapes.Count = 0 then
    Exit;
//  imgPosList := TStringList.Create;
  try
    // 全文所有图片的起止位置，由于文档稍有改变时位置信息就会变动，这里仅临时使用，不在内存中保留此信息
//    for i := 1 to doc.InlineShapes.Count do begin
//      range := doc.InlineShapes.Item(i).Range;
//      selpos := TOfficeHelper.GetRangePos(range);
//      imgPosList.Add(Format('%d=%d', [selpos.StartPos, selpos.EndPos]));
//    end;

    // 记录包含图片的段落
    FImagePositions.Clear;
    for i := 1 to OfficeHelper.Document.Paragraphs.Count do begin
      paPos := TOfficeHelper.GetRangePos(OfficeHelper.Document.Paragraphs.Item(i).Range);

      imageIndex := GetImageIndexInParagraph(i);
      if imageIndex <= 0 then
        Continue;
      imgPos := TOfficeHelper.GetRangePos(OfficeHelper.Document.InlineShapes.Item(imageIndex).Range);
//      for imageIndex := 0 to imgPosList.Count - 1 do begin
//        imgPos.StartPos := StrToInt(imgPosList.Names[imageIndex]);
//        imgPos.EndPos := StrToInt(imgPosList.ValueFromIndex[imageIndex]);
//        if paPos.ContainSelectionPos(imgPos) then begin
          subTextRanges := TStringList.Create;
          try
            GetTextSubRanges(paPos, subTextRanges);
            singleImageParagraph := True;  // 是否仅有图片无文字
            for m := 0 to subTextRanges.Count - 1 do begin
              textPos.StartPos := StrToInt(subTextRanges.Names[m]);
              textPos.EndPos := StrToInt(subTextRanges.ValueFromIndex[m]);
              if textPos.StartPos = textPos.EndPos then
                Continue;
              docRange := OfficeHelper.Document.Range(textPos.StartPos, textPos.EndPos);
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
            TOfficeHelper.SetSelectionPos(OfficeHelper.Window, selPos);
            OfficeHelper.Window.Selection.HomeKey(wdLine, wdMove);
            if OfficeHelper.Window.Selection.Start = paPos.StartPos then
              picLines := 1   // 单独成段的图片，或图片所在段落只占一行
            else
              picLines := 2;  // 图片所在段落有多行
          end;

          FImagePositions.Add(Format('%d=%d', [i, picLines]));
//          System.Break;
//        end;
//      end;
    end;
  finally
//    imgPosList.Free;
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
  i, paNo, paCount: integer;
//  lastTablePos: Integer;
  tablePos: TSelectionPos;
begin
  if OfficeHelper.Document.Tables.Count = 0 then
    Exit;

  FTablePositions.Clear;
//  lastTablePos := 1;
  for i := 1 to TableCount do begin
    tablePos := TOfficeHelper.GetRangePos(OfficeHelper.Document.Tables.Item(i).Range);
    paCount := OfficeHelper.Document.Tables.Item(i).Range.Paragraphs.Count;

    paNo := FOfficeHelper.GetParagraphNoByPos(tablePos.StartPos);
    FTablePositions.Add(Format('%d=%d', [paNo, paCount]));

//    paNo := lastTablePos;
//    while paNo <= doc.Paragraphs.Count do begin
//      selPos := TOfficeHelper.GetRangePos(doc.Paragraphs.Item(paNo).Range);
//      if selPos.StartPos = tablePos.StartPos then begin
//        FTablePositions.Add(Format('%d=%d', [paNo, paCount]));
//        lastTablePos := paNo + paCount;
//        System.Break;
//      end;
//      Inc(paNo);
//    end;
  end;
end;

end.
