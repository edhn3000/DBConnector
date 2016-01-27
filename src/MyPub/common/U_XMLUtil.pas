{**
 * <p>Title: XML处理 </p>
 * <p>Copyright: Copyright (c) 2006/11/30</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: XML处理单元 </p>
 *}

 unit U_XMLUtil;

interface

uses
  XMLDoc, XMLIntf, Classes, SysUtils;

type
  TCompareNodeFunc = function (s1, s2: string): Boolean;
  TXMLUtil = class(TXMLDocument)

  public
    constructor Create(sVersion: string = '1.0'; sEncoding: string = 'GBK');overload;
    destructor Destroy;override;

    // 是否含有非注释节点的子节点
    function HasNonCommentChild(AXmlNode: IXMLNode):Boolean;

    // 深度方式非递归遍历xml节点，找到指定节点
    function FindChildXMLNode(AParent:IXMLNode; ANodeName: string;
      cmp: TCompareNodeFunc; FindInChild: Boolean = False):IXMLNode;
    function FindParentXMLNode(Child:IXMLNode; ANodeName: string;
      cmp: TCompareNodeFunc):IXMLNode;

    // 找到节点的路径
    function RouteXmlNode(ANode: IXMLNode; sDelimiter: string='/'): string;
    function FindRouteNode(AParent: IXMLNode; sRoute: string;
      cmp: TCompareNodeFunc; sDelimiter: string = ','): IXMLNode;
    function GetNodeText(ANodeName: string): string;
  end;
  
implementation

{ TXMLUtil }

constructor TXMLUtil.Create(sVersion, sEncoding: string);
begin
  Create(nil);
  Active := True;
  if sVersion <> '' then
    Version := sVersion;
  if sEncoding <> '' then
    Encoding := sEncoding;
end;

destructor TXMLUtil.Destroy;
begin
//  Active := False;
  inherited;
end;

function TXMLUtil.FindChildXMLNode(AParent:IXMLNode; ANodeName: string;
    cmp: TCompareNodeFunc; FindInChild: Boolean = False):IXMLNode;
var
  CurrNode: IXMLNode;                  // 指向当前节点
  function CopmareNode(s1, s2: string): Boolean;
  begin
    if Assigned(cmp) then
      Result := cmp(s1, s2)
    else
      Result := s1=s2;
  end;
begin
  Result := nil;
  if not Assigned(AParent) then
    Exit;
  if CopmareNode(AParent.NodeName, ANodeName) then
  begin
    Result := AParent;
    Exit;
  end;
  if not AParent.HasChildNodes then
    Exit;

  CurrNode := AParent.ChildNodes[0];     // 取第一个子节点
  while CurrNode <> AParent do
  begin
    // 该节点就是要找的节点直接退出
    if CopmareNode(ANodeName, CurrNode.NodeName) then
    begin
      Result := CurrNode;
      Break;
    end;
    // 允许查找子节点 并且有子节点 进入子节点查找
    if FindInChild and CurrNode.HasChildNodes then
    begin
      CurrNode := CurrNode.ChildNodes[0];
      if CopmareNode(ANodeName, CurrNode.NodeName) then
      begin
        Result := CurrNode;
        Break;
      end;
    end
    // 没有子节点但有兄弟节点找兄弟
    else if CurrNode.NextSibling <> nil then
    begin
      CurrNode := CurrNode.NextSibling;
      if CopmareNode(ANodeName, CurrNode.NodeName) then
      begin
        Result := CurrNode;
        Break;
      end;
    end
    // 没有子节点也没有兄弟节点,找父节点的兄弟,如父节点是根则退出
    else if CurrNode.ParentNode <> nil then
    begin
      // 找到祖先节点中兄弟不为空的节点,直接父节点不能是根节点,否则结束
      while (CurrNode.ParentNode <> AParent) and (CurrNode.ParentNode.NextSibling = nil) do
        CurrNode := CurrNode.ParentNode;
      
      if CurrNode.ParentNode = AParent then
        Break
      else
        CurrNode := CurrNode.ParentNode.NextSibling;
    end
    else
      Break;
  end;
end;   

function TXMLUtil.FindRouteNode(AParent: IXMLNode; sRoute: string;
  cmp: TCompareNodeFunc; sDelimiter: string): IXMLNode;
var
  CurrNode: IXMLNode;
  sPath, sSubPath: string;
  nPos: Integer;
  function CopmareNode(s1, s2: string): Boolean;
  begin
    if Assigned(cmp) then
      Result := cmp(s1, s2)
    else
      Result := s1=s2;
  end;
  function GetChildNode(ParentNode: IXMLNode; sName: string): IXMLNode;
  var
    node: IXMLNode;
  begin
    Result := nil;
    node := ParentNode.ChildNodes.First;
    if node  = nil then
      Exit;
    if CopmareNode(sName, node.NodeName) then
    begin
      Result := node;
      Exit;
    end;  
    while node.NextSibling <> nil do
    begin
      node := node.NextSibling;
      if CopmareNode(sName, node.NodeName) then
      begin
        Result := node;
        Break;
      end;   
    end;  
  end;  
begin
  CurrNode := AParent;
  sPath := sRoute;
  sSubPath := '';
  nPos := Pos(sDelimiter, sPath);
  if nPos = 0 then
  begin
    sSubPath := sPath;
    sPath := '';
  end
  else
  begin
    sSubPath := Copy(sPath, 1, nPos-1);
    sPath := Copy(sPath, nPos+1, MaxInt);
  end;
  if not CopmareNode(sSubPath, CurrNode.NodeName) then
  begin
    Result := nil;
    Exit;
  end;

  while sPath <> '' do
  begin
    nPos := Pos(sDelimiter, sPath);
    if nPos = 0 then
    begin
      sSubPath := sPath;
      sPath := '';
    end
    else
    begin
      sSubPath := Copy(sPath, 1, nPos-1);
      sPath := Copy(sPath, nPos+1, MaxInt);
    end;
    Result := GetChildNode(CurrNode, sSubPath);
    if Result = nil then
      Break;
  end;
end;

function TXMLUtil.FindParentXMLNode(Child: IXMLNode; ANodeName: string;
  cmp: TCompareNodeFunc): IXMLNode;
  function CopmareNode(s1, s2: string): Boolean;
  begin
    if Assigned(cmp) then
      Result := cmp(s1, s2)
    else
      Result := s1=s2;
  end;
begin
  Result := nil;
  if CopmareNode(ANodeName, Child.NodeName) then
  begin
    Result := Child;
    Exit;
  end;
  
  while Assigned(Child.ParentNode) do
  begin
    Child := Child.ParentNode;
    if CopmareNode(ANodeName, Child.NodeName) then
    begin
      Result := Child;
      Break;
    end;  
  end;
end;

function TXMLUtil.HasNonCommentChild(AXmlNode: IXMLNode): Boolean;
var
  nChildNodeIndex: Integer;
begin
  Result := True;
  for nChildNodeIndex := 0 to AXmlNode.ChildNodes.Count - 1 do
  begin
    if AXmlNode.ChildNodes[nChildNodeIndex].NodeType <> ntComment then
      Exit;
  end;
  Result := False;
end;

function TXMLUtil.RouteXmlNode(ANode: IXMLNode; sDelimiter: string='/'): string;
begin
  Result := ANode.NodeName;
  while (ANode.ParentNode <> nil) and
        (Copy(ANode.ParentNode.NodeName,1,1) <> '#') do
  begin
    ANode := ANode.ParentNode;
    Result := ANode.NodeName + sDelimiter + Result;
  end;
end;

function TXMLUtil.GetNodeText(ANodeName: string): string;
var
  slst: TStrings;
  i: Integer;
  s1: string;
  node: IXMLNode;
  function GetChildNode(AParent: IXMLNode; sName: string): IXMLNode;
  var
    nd: IXMLNode;
  begin
    if AParent = nil then
    begin
      Result := DocumentElement;
      Exit;
    end;
    nd := AParent.ChildNodes.First;
    while nd <> nil do
    begin
      if SameText(nd.NodeName, sName) then
      begin
        Result := nd;
        Break;
      end;  
    end;  
  end;  
begin  
  slst := TStringList.Create;
  try
    ExtractStrings(['/'], [' '], PChar(ANodeName), slst);
    i := 0;
    node := nil;
    while i < slst.Count do
    begin
      s1 := slst[i];
      node := GetChildNode(node, s1);
      if node = nil then
        Break;
      Inc(i);
    end;
    if (i = slst.Count - 1) and (node <> nil) then
      Result := node.Text;
  finally
    slst.Free;
  end;
end;

initialization

finalization

end.
