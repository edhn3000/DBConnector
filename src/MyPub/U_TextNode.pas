{**
 * <p>Title: U_TextNode </p>
 * <p>Copyright: Copyright (c) 2008/8/31</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 文本节点解析 </p>
 *}
unit U_TextNode;

interface  
uses  
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ShellAPI, Contnrs;

type
  TTextNodeList = class;
  
  TTextNode = class(TCollectionItem)
  private          
    FLastError: string;   
    FNodeName: string;
    FNodeText: string;
    FParentNode: TTextNode;
    FChildNodes: TTextNodeList;
    function GetFirstNode(sText: string): string;     
    function ParseFullNodeText(S, sNodeName: string; var sNodeText: string;
      nFrom: Integer = 1): Boolean;
    function ParseNodeText(S, sNodeName: string; var sNodeText: string;
      nFrom: Integer = 1): Boolean;
    procedure Reset;
  public              
    property LastError: string read FLastError;
    property NodeName: string read FNodeName;
    property NodeText: string read FNodeText;
    property ParentNode: TTextNode read FParentNode write FParentNode;
    property ChildNodes: TTextNodeList read FChildNodes;
  public                      
    constructor Create(Collection: TCollection); override;
    destructor Destroy;override;

    function InitText(sText: string): Boolean;
    function ChildNodeText(sNodeName: string): string; 
    function NextSibling: TTextNode;
  end;

  TTextNodeList = class(TCollection)
  private
    function GetItem(index: Integer): TTextNode;
    procedure SetItem(index: Integer; value: TTextNode);
  public
    property Items[index: Integer]: TTextNode read GetItem write SetItem;
  public
    function FindNode(sNodeName: string): TTextNode;
    function Add: TTextNode;
    function First: TTextNode;
  end;

implementation   

{ TTextNode }

const
  C_sStartNodeStart = '<'; 
  C_sStartNodeEnd   = '>';
  C_sEndNodeStart   = '</';
  C_sEndNodeEnd     = C_sStartNodeEnd;

function TTextNode.ChildNodeText(sNodeName: string): string;
var
  node: TTextNode;
begin
  node := ChildNodes.FindNode(sNodeName);
  if node <> nil then
    Result := node.NodeText;
end;

constructor TTextNode.Create(Collection: TCollection);
begin
  inherited Create(Collection);
  FChildNodes := TTextNodeList.Create(TTextNode);
end;

destructor TTextNode.Destroy;
begin
  FChildNodes.Free;
  inherited;
end;

function TTextNode.GetFirstNode(sText: string): string;
var
  nPos1, nPos2: Integer;
begin
  nPos1 := Pos(C_sStartNodeStart, sText);
  nPos2 := Pos(C_sStartNodeEnd, sText);
  Result := Trim(Copy(sText, nPos1+Length(C_sStartNodeStart),
    nPos2-nPos1-Length(C_sStartNodeStart)))
end;

function TTextNode.InitText(sText: string): Boolean;
var
  sLeftText, sNodeName, sFullNodeText: string;
  node: TTextNode;
begin
  Reset;
  Result := False;
  FNodeName := GetFirstNode(sText);
  if FNodeName = '' then
  begin
    FLastError := '(TTextNode.InitText)未找到开始节点 '+ sText;
    Exit;
  end;
  if not ParseNodeText(sText, FNodeName, sLeftText) then
  begin
    FLastError := '(TTextNode.InitText)无法分析节点内容 '+ sText;
    Exit;
  end;  
  sNodeName := GetFirstNode(sLeftText);
  if sNodeName = '' then
    FNodeText := sLeftText
  else
  begin
    while sLeftText <> '' do
    begin        
      sNodeName := GetFirstNode(sLeftText);
      if sNodeName = '' then
        Break
      else
      begin
        if not ParseFullNodeText(sLeftText, sNodeName, sFullNodeText) then
        begin
          FLastError := '(TTextNode.InitText)无法分析节点完整内容 '+ sLeftText;
          Exit;
        end;
        node := ChildNodes.Add;
        node.InitText(sLeftText);
        node.ParentNode := Self;
      end;
      sLeftText := Copy(sLeftText, Length(sFullNodeText)+1, MaxInt);
    end;
  end;
  Result := True;
end;

function TTextNode.NextSibling: TTextNode;
//var
//  nIndex: Integer;
begin
  Result := nil;
  if Assigned(FParentNode) then
  begin
//    nIndex := FParentNode.ChildNodes.IndexOf(Self);
    if (index <> -1) and (index <> FParentNode.ChildNodes.Count-1) then
      Result := FParentNode.ChildNodes.Items[index+1];
  end;  
end;

function TTextNode.ParseFullNodeText(S, sNodeName: string; var sNodeText: string;
  nFrom: Integer): Boolean;
var
  nPos1, nPos2: Integer;
  sStart, sEnd: string;
  S2: string;
begin
  sStart := Format('%s%s%s',[C_sStartNodeStart, sNodeName, C_sStartNodeEnd]);
  sEnd   := Format('%s%s%s',[C_sEndNodeStart, sNodeName, C_sEndNodeEnd]);
  if nFrom <> 1 then
    S2 := Copy(S, nFrom, MaxInt)
  else
    S2 := S;
  nPos1 := Pos(sStart, S2);
  nPos2 := Pos(sEnd, S2);
  
  Result := nPos2 >= nPos1;
  sNodeText := Copy(S2, nPos1, nPos2-nPos1+Length(sEnd));
end;

function TTextNode.ParseNodeText(S, sNodeName: string; var sNodeText: string;
  nFrom: Integer): Boolean;
var
  nPos1, nPos2: Integer;
  sStart, sEnd: string;
  S2: string;
begin
  sStart := Format('%s%s%s',[C_sStartNodeStart, sNodeName, C_sStartNodeEnd]);
  sEnd   := Format('%s%s%s',[C_sEndNodeStart, sNodeName, C_sEndNodeEnd]);
  if nFrom <> 1 then
    S2 := Copy(S, nFrom, MaxInt)
  else
    S2 := S;
  nPos1 := Pos(sStart, S2);
  nPos2 := Pos(sEnd, S2);
                             
  Result := nPos2 >= nPos1;// nPos2小于nPos1的情况为 结束节点找不到;如果二者都找不到时都为0,认为正确
  sNodeText := Copy(S2, nPos1+Length(sStart), nPos2-nPos1-Length(sStart));
end;

procedure TTextNode.Reset;
begin
  FNodeName := '';
  FNodeText := '';
  ChildNodes.Clear;
end;

{ TTextNodeList }

function TTextNodeList.Add: TTextNode;
begin
  Result := inherited Add as TTextNode;
end;

function TTextNodeList.FindNode(sNodeName: string): TTextNode;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Count - 1 do
  begin
    if Items[i].NodeName = sNodeName then
    begin
      Result := Items[i];
      Exit;
    end;  
  end;  
end;

function TTextNodeList.First: TTextNode;
begin
  if Count = 0 then
    Result := nil
  else
    Result := Items[0];
end;

function TTextNodeList.GetItem(index: Integer): TTextNode;
begin
  Result := inherited GetItem(index) as TTextNode;
end;

procedure TTextNodeList.SetItem(index: Integer; value: TTextNode);
begin
  inherited SetItem(index, value);
end;

end.
