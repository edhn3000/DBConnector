{**
 * <p>Title: U_fCheckTree  </p>
 * <p>Copyright: Copyright (c) 2009/12/19</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 
 *}unit U_fCheckTree;

interface

uses
  ComCtrls, StdCtrls, Controls, Classes, U_TreeObjectList;
type
  TSelectTreeMode = (stmSingle, stmCheck);

  TfCheckTree = class(TTreeView)
  private
  
    FTreeRoot: TTreeObjectItem;
    FTreeList: TTreeObjectList;
    FSelectedItem: TTreeObjectItem;
    FSelectTreeMode: TSelectTreeMode;
    procedure AddItemToTreeView(treeItem: TTreeObjectItem; parentNode: TTreeNode);
    function FindNode(node: TTreeNode; item: TTreeObjectItem;
      compFunc: TFuncCompareTreeObjectItem): TTreeNode;
    procedure SetSelectTreeMode(value: TSelectTreeMode);
    // 设置tree节点是否被选中
    procedure SetNodeChecked(node: TTreeNode; value: Boolean);
    function GetNodeChecked(node: TTreeNode): Boolean;
    // 不仅选中传入节点，也检查子节点状态
    procedure DoCheckNode(node: TTreeNode; value: Boolean);
  protected
    procedure tvMainChanging(Sender: TObject; Node: TTreeNode;
      var AllowChange: Boolean);
    procedure tvMainMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  public
    function GetSelectedFullPath: string;
    { Public declarations }
    procedure Init(TreeRoot: TTreeObjectItem);
    procedure InitList(nodeList: TTreeObjectList);
    procedure SelectItem(itemToSelect:TTreeObjectItem;
      compFunc: TFuncCompareTreeObjectItem);
    procedure AutoExpanded;
  published
    property SelectedItem: TTreeObjectItem read FSelectedItem;
    property SelectTreeMode: TSelectTreeMode read FSelectTreeMode write SetSelectTreeMode;
  end;

implementation

procedure TfCheckTree.AddItemToTreeView(treeItem: TTreeObjectItem;
  parentNode: TTreeNode);
var
  i: Integer;
  node: TTreeNode;
begin
  node := Items.AddChildObject(parentNode, treeItem.Name, treeItem);
  if FSelectTreeMode = stmCheck then
  begin
    SetNodeChecked(node, treeItem.Checked);
  end;
  if treeItem.Childs = nil then
    Exit;
  for i := 0 to treeItem.Childs.Count - 1 do
  begin
    AddItemToTreeView(treeItem.Childs.Items[i], node);
  end;
end;

procedure TfCheckTree.Init(TreeRoot: TTreeObjectItem);
begin
  FTreeRoot := TreeRoot;
  Items.Clear;
  AddItemToTreeView(TreeRoot, nil);
end;

procedure TfCheckTree.InitList(nodeList: TTreeObjectList);
var
  i: Integer;
begin
  FTreeList := nodeList;
  Items.Clear;
  for i := 0 to FTreeList.Count - 1 do
    AddItemToTreeView(FTreeList.Items[i], nil);
end;

procedure TfCheckTree.SelectItem(itemToSelect: TTreeObjectItem;
  compFunc: TFuncCompareTreeObjectItem);
var
  node: TTreeNode;
begin
  node := Items.GetFirstNode;
  node := FindNode(node, itemToSelect, compFunc);
  if node <> nil then
    Selected := node;
end;

function TfCheckTree.FindNode(node: TTreeNode; item: TTreeObjectItem;
  compFunc: TFuncCompareTreeObjectItem): TTreeNode;
begin
  Result := nil;
  if node = nil then
    Exit;
  if compFunc(TTreeObjectItem(node.Data), item) then
  begin
    Result := node;
    Exit;
  end;
  if node.getFirstChild <> nil then
    Result := FindNode(node.getFirstChild, item, compFunc);
  if (Result = nil) and (node.getNextSibling <> nil) then
    Result := FindNode(node.getNextSibling, item, compFunc);
end;   

procedure TfCheckTree.tvMainChanging(Sender: TObject; Node: TTreeNode;
  var AllowChange: Boolean);
begin
  if Node <> nil then
    FSelectedItem := TTreeObjectItem(Node.Data)
  else
    FSelectedItem := nil;
end;

procedure TfCheckTree.tvMainMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  tndHit: TTreeNode;    
  theHT: THitTests;
begin
  tndHit := nil;  
  theHT := GetHitTestInfoAt(X, Y);
  if (htOnStateIcon in theHT) or (htOnLabel in theHT) or (htOnIcon in theHT)
     or (htOnItem in theHT) then
    tndHit := GetNodeAt(X, Y);
  if not Assigned(tndHit) then
    Exit;

  case FSelectTreeMode of
  stmSingle:
  begin
    if ssDouble in Shift then
//      btnSelect.Click;
  end;
  stmCheck:
  begin
    DoCheckNode(tndHit, not GetNodeChecked(tndHit));
  end;
  end;
end;

function TfCheckTree.GetSelectedFullPath: string;
var
  item: TTreeObjectItem;
begin
  item := FSelectedItem;
  while item <> nil do
  begin
    if Result = '' then
      Result := item.Name
    else
      Result := item.Name + '.' + Result;
    item := item.Parent;
  end;  
end;

procedure TfCheckTree.SetSelectTreeMode(value: TSelectTreeMode);
begin
  if FSelectTreeMode <> value then
  begin
    case value of
    stmSingle:
    begin
      Images := nil;
      btnSelect.Caption := '选择';
    end;
    stmCheck:
    begin
      Images := ilCheck;   
      btnSelect.Caption := '确定';
    end;
    else
    end;
    FSelectTreeMode := value;
  end;   
end;

procedure TfCheckTree.SetNodeChecked(node: TTreeNode; value: Boolean);
var
  treeItem: TTreeObjectItem;
begin
  if value then
  begin
    node.ImageIndex := 1;
    node.SelectedIndex := 1;
  end
  else
  begin
    node.ImageIndex := 0;
    node.SelectedIndex := 0;
  end;
  treeItem := node.Data;
  treeItem.Checked := value;
  Refresh;
end;

function TfCheckTree.GetNodeChecked(node: TTreeNode): Boolean;
begin
  Result := node.ImageIndex = 1;
end;

procedure TfCheckTree.DoCheckNode(node: TTreeNode; value: Boolean);
var
  childNode: TTreeNode;
  allSiblingChecked: Boolean;
  parentNode: TTreeNode;
  siblingNode: TTreeNode;
begin
  SetNodeChecked(node, value);

  childNode := node.getFirstChild;
  while childNode <> nil do
  begin
    DoCheckNode(childNode, value);
    childNode := childNode.getNextSibling;
  end;
  parentNode := node.Parent;
  if parentNode <> nil then
  begin
    allSiblingChecked := True;
    siblingNode := parentNode.getFirstChild;
    while siblingNode <> nil do
    begin
      if not GetNodeChecked(siblingNode) then
      begin
        allSiblingChecked := False;
        Break;
      end;
      siblingNode := siblingNode.getNextSibling;
    end;
    SetNodeChecked(parentNode, allSiblingChecked)
  end;
end;

procedure TfCheckTree.AutoExpanded;
begin
  AutoExpand := True;
end;

end.
