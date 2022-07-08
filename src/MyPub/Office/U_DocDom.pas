unit U_DocDom;

interface

uses
  Classes, SysUtils, Forms, Windows, PerlRegEx, DateUtils, OleCtrls, StrUtils,
  Dialogs, Variants, Office_TLB, Word_TLB, XMLDoc, XMLIntf, msxml,
  U_OfficeHelper, U_WordHelper, Generics.Collections, U_RegexUtil;


const
  NODE_PATH_SEPARATOR = '/';

type
  TDocDom = class;
  { TDocNodeData }
  TDocNodeData = record
  private
    FNode: IXMLDOMNode;
    function GetChildCount: Integer;

  public
    Name: string; // 去掉编号的text
    Text: String; // 标题段落的内容
    Level: Integer;
    ParagraphNo: Integer;
    StartPos: Integer;
    Length: Integer;
    EndPos: Integer;

    property ChildCount: Integer read GetChildCount;

    procedure ParseNode(node: IXMLDOMNode);
    function GetChildNode(index: Integer): TDocNodeData;
    function HasParent: Boolean;
    function GetParentNode(): TDocNodeData;
    function GetPreSiblingNode(): TDocNodeData;
    function GetNextSiblingNode(): TDocNodeData;
    function GetFullName(separator: String = NODE_PATH_SEPARATOR): String;
    function GetAttribute(name: String): String;
  end;

  { TDocDom }
  TDocDom = class
  private
    FDoc: OleVariant;
    FOfficeHelper: TOfficeHelper;
    FXml: String;
    FDocumentDom: IXMLDomDocument2;
    FRootNode: TDocNodeData;

    function CreateNode(parent: IXMLNode; range: OleVariant): IXMLNode;
    function GetNodeLevel(node: IXMLNode): Integer;
    procedure UpdateNodePos(node: IXMLNode);


  protected
    procedure InitByParagraphs(parentNode: IXmlNode);
    procedure InitByGotoHeadings(parentNode: IXmlNode);

  public

    property XML: String read FXml;
    property DocumentDom: IXMLDomDocument2 read FDocumentDom;

    constructor Create();
    destructor Destroy; override;

    function GetRootNode: TDocNodeData;

    // 分析word文档
    procedure Init(doc: OleVariant);
    // 外部传入已分析好的xml
    procedure InitByXml(xml: String);

    procedure LoadFromFile(fileName: String);
    procedure SaveToFile(fileName: String);
  end;

implementation

uses
  U_Log;

{ TDocNodeData }

function TDocNodeData.GetChildCount: Integer;
begin
  Result := FNode.childNodes.length;
end;

procedure TDocNodeData.parseNode(node: IXMLDOMNode);
begin
  if node = nil then
    Exit;

  FNode := node;
  Name := GetAttribute('name');
  Text := GetAttribute('text');
  Level := StrToInt(GetAttribute('level'));
  StartPos := StrToInt(GetAttribute('start'));
  Length := StrToInt(GetAttribute('length'));
  EndPos := StrToInt(GetAttribute('end'));
  ParagraphNo := StrToInt(GetAttribute('paragraphno'))
end;

function TDocNodeData.GetChildNode(index: Integer): TDocNodeData;
var
  node: IXMLDOMNode;
begin
  if index >= FNode.childNodes.length then
    raise Exception.Create(Format('index %d out of range', [index]));

  node := FNode.childNodes.item[index];
  Result.parseNode(node);
end;

function TDocNodeData.GetFullName(separator: String): String;
var
  s: String;
  node: TDocNodeData;
begin
  s := Name;
  node := Self;
  while node.HasParent do begin
    node := node.GetParentNode;
    if node.Name = 'root' then
      System.Break;
    s := node.Name + separator + s;
  end;
  Result := s;
end;

function TDocNodeData.GetPreSiblingNode: TDocNodeData;
begin
  if FNode.previousSibling <> nil then begin
    Result.parseNode(FNode.previousSibling);
  end;
end;

function TDocNodeData.GetNextSiblingNode: TDocNodeData;
begin
  if FNode.nextSibling <> nil then begin
    Result.parseNode(FNode.nextSibling);
  end;
end;

function TDocNodeData.GetParentNode: TDocNodeData;
begin
  if FNode.parentNode <> nil then begin
    Result.parseNode(FNode.parentNode);
  end;
end;

function TDocNodeData.HasParent: Boolean;
begin
  Result := not ((FNode.parentNode = nil) or (FNode.nodeName = 'root'));
end;

function TDocNodeData.GetAttribute(name: String): String;
var
  node: IXMLDOMNode;
begin
  Result := '';
  node := FNode.selectSingleNode('@' + name);
  if Assigned(node) then
    Result := node.text;
end;


{ TDocDom }

constructor TDocDom.Create;
begin
  FDocumentDom := CoDOMDocument.create;
end;

destructor TDocDom.Destroy;
begin
//  FDocumentDom._Release;
  FDocumentDom := nil;

  if Assigned(FOfficeHelper) then
    FreeAndNil(FOfficeHelper);

  inherited;
end;

function TDocDom.GetRootNode: TDocNodeData;
begin
  Result := FRootNode;
end;


function TDocDom.GetNodeLevel(node: IXMLNode): Integer;
begin
  Result := StrToInt(node.Attributes['level']);
end;

// 创建node
function TDocDom.CreateNode(parent: IXMLNode; range: OleVariant): IXMLNode;
var
  selPos: TSelectionPos;
  sText: String;
begin
  Result := parent.AddChild('node');
  sText := Trim(range.Text);
  Result.Attributes['text'] := sText;
  sText := TRegexUtil.MatchFirstStr('(?:[\d.\x{0020}\x{3000}]*)?(.+)', 1, sText);
  Result.Attributes['name'] := sText;
  Result.Attributes['level'] := IntToStr(GetNodeLevel(parent) + 1);
  selPos := TOfficeHelper.GetRangePos(range);
  Result.Attributes['start'] := selPos.StartPos;
  Result.Attributes['length'] := Length(Range.Text);
  Result.Attributes['end'] := selPos.EndPos;
end;

// 更新endpos
procedure TDocDom.UpdateNodePos(node: IXMLNode);
var
  i: Integer;
  childNode,preNode: IXMLNode;
begin
  // 先设置最后一个节点的end，再逐层向前更新end
  if node.ChildNodes.Count > 0 then begin
    childNode := node.ChildNodes.Last;
    childNode.Attributes['end'] := node.Attributes['end'];

    // 向前更新
    preNode := childNode.PreviousSibling;
    while preNode <> nil do begin
      preNode.Attributes['end'] := IntToStr(StrToInt(preNode.NextSibling.Attributes['start'])-1);
      preNode := preNode.PreviousSibling;
    end;
  end;

  for i := node.ChildNodes.Count-1 downto 0 do begin
    childNode := node.ChildNodes[i];
    if childNode.ChildNodes.Count > 0 then
      UpdateNodePos(childNode);
  end;
end;

procedure TDocDom.initByParagraphs(parentNode: IXmlNode);
var
  childNode, lastNode: IXMLNode;
  i, nodeLevel, lastNodeLevel: Integer;
  range: OleVariant;
begin
  lastNode := parentNode;
  lastNodeLevel := 0;
  i := 1;
  while i <= FDoc.Paragraphs.Count do begin
    range := FDoc.Paragraphs.Item(i).Range;
    if range.Information[wdWithInTable] then begin
      Range.Select;
      i := i + FOfficeHelper.Selection.Tables.Item(1).Range.Paragraphs.Count;
      Continue;
    end;
    if Trim(Range.Text) = '' then begin
      Inc(i);
      Continue;
    end;
    nodeLevel := FDoc.Paragraphs.Item(i).OutlineLevel;
    if nodeLevel >= 10 then begin
      Inc(i);
      Continue;
    end;

    ConsoleLog.Debug(Format('text=%s, level=%d', [Trim(Range.Text), nodeLevel]));
    if nodeLevel = lastNodeLevel then begin // 与上一个节点同层级
      childNode := CreateNode(parentNode, Range);
    end else if nodeLevel < lastNodeLevel then begin // 上一个节点某层级父节点的兄弟
      parentNode := lastNode;
      while (GetNodeLevel(parentNode) >= nodeLevel) and (parentNode.ParentNode <> nil) do begin
        parentNode := parentNode.ParentNode;
      end;
      childNode := CreateNode(parentNode, Range);
    end else begin // 是上一个节点的子节点
      parentNode := lastNode;
      childNode := CreateNode(parentNode, Range);
    end;
    childNode.Attributes['paragraphno'] := IntToStr(i);

    lastNode := childNode;
    lastNodeLevel := nodeLevel;
    Inc(i);
  end;
end;

procedure TDocDom.initByGotoHeadings(parentNode: IXmlNode);
var
  childNode, lastNode: IXMLNode;
  i, nodeLevel, lastNodeLevel: Integer;
  range: OleVariant;
  selPos: TSelectionPos;
begin
  lastNode := parentNode;
  lastNodeLevel := 0;
  i := 1;
  while True do begin
    FOfficeHelper.Selection.Goto(wdGoToHeading, wdGoToFirst, i);
    selPos := TOfficeHelper.GetSelectionPos(FOfficeHelper.Window);
    if selPos.StartPos = 0 then
      System.Break;

    nodeLevel := FOfficeHelper.Selection.ParagraphFormat.OutlineLevel;

    FOfficeHelper.Selection.EndKey(wdLine, wdExtend);
    range := FOfficeHelper.Selection.Range;
    ConsoleLog.Debug(Format('text=%s, level=%d', [Trim(range.Text), nodeLevel]));
    if nodeLevel = lastNodeLevel then begin // 与上一个节点同层级
      childNode := CreateNode(parentNode, range);
    end else if nodeLevel < lastNodeLevel then begin // 上一个节点某层级父节点的兄弟
      parentNode := lastNode;
      while (GetNodeLevel(parentNode) >= nodeLevel) and (parentNode.ParentNode <> nil) do begin
        parentNode := parentNode.ParentNode;
      end;
      childNode := CreateNode(parentNode, range);
    end else begin // 是上一个节点的子节点
      parentNode := lastNode;
      childNode := CreateNode(parentNode, range);
    end;
    childNode.Attributes['paragraphno'] := IntToStr(FOfficeHelper.GetPageNoByPos(selPos.StartPos));

    lastNode := childNode;
    lastNodeLevel := nodeLevel;
    Inc(i);
  end;
end;

procedure TDocDom.initByXml(xml: String);
begin
  FXml := xml;
  FDocumentDom.LoadXml(xml);
  FDocumentDom.setProperty('SelectionLanguage','XPath');
  FRootNode.parseNode(FDocumentDom.documentElement);
end;

procedure TDocDom.LoadFromFile(fileName: String);
var
  xmlDoc: IXMLDocument;
begin
  XMLDoc := TXMLDocument.Create(nil);
  try
    XMLDoc.Active := True;
    XMLDoc.LoadFromFile(fileName);
    initByXml(XMLDoc.XML.Text);
  finally
    xmlDoc := nil;
  end;
end;

procedure TDocDom.SaveToFile(fileName: String);
var
  xmlDoc: IXMLDocument;
begin
  try
    XMLDoc := TXMLDocument.Create(nil);
    try
      XMLDoc.Active := True;
      XMLDoc.Encoding := 'UTF-8';
      XMLDoc.LoadFromXML(FXml);
      XMLDoc.SaveToFile(fileName);
    finally
      xmlDoc := nil;
    end;
  except
    on e: Exception do begin
      OutputDebugString(PChar('DocDom save '
      + fileName +' by xmldoc error! ' + e.Message
        + ', will save to ' + fileName + '.error.xml'));
      with TStringList.Create do begin
        Text := FXml;
        SaveToFile(fileName + '.error.xml', TEncoding.UTF8);
      end;
    end;
  end;
end;

procedure TDocDom.init(doc: OleVariant);
var
  xmlDoc: IXMLDocument;
  root: IXMLNode;
  sXml: String;
  stream: TStringStream;
  selPos: TSelectionPos;
begin
  FDoc := doc;
  if Assigned(FOfficeHelper) then
    FreeAndNil(FOfficeHelper);
  FOfficeHelper := TWordHelper.Create;
  FOfficeHelper.InitByDoc(FDoc);

  XMLDoc := TXMLDocument.Create(nil);
  try
    XMLDoc.Active := True;
    XMLDoc.Encoding := 'UTF-8';
    selPos := TOfficeHelper.GetRangePos(doc.Content);
    root := XMLDoc.AddChild('root');
    root.Attributes['name'] := 'root';
    root.Attributes['paragraphno'] := '0';
    root.Attributes['text'] := '';
    root.Attributes['level'] := '0';
    root.Attributes['start'] := IntToStr(selPos.StartPos);
    root.Attributes['length'] := '0';
    root.Attributes['end'] := IntToStr(selPos.EndPos);

    initByParagraphs(root);
//    initByGotoHeadings(root);
    UpdateNodePos(root);
    stream := TStringStream.Create(sXml, TEncoding.UTF8);
    try
      XMLDoc.SaveToStream(stream);
      initByXml(stream.DataString);
    finally
      stream.Free;
    end;
  finally
    xmlDoc := nil;
  end;
end;

end.
