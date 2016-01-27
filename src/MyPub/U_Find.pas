{**
 * <p>Title: Find  </p>
 * <p>Copyright: Copyright (c) 2006/10/27</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 在控件中查找</p>

  使用如下方式显示查找窗体
  if not Assigned(gFindForm) then
    gFindForm := TF_Find.Create( Self );
  with gFindForm do
  begin
    PassControl( TWinControl( 控件对象 ) );
    FormStyle := fsStayOnTop;
    Show;
  end;
 *}
unit UF_Find;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls, DBGrids, DB;

type
  TFindCallBack = procedure(nLine, nIndex, nCount: Integer) of object;
  TTreeCompareMethod = function(Node:TTreeNode; findStr: String): Integer of object;
  TF_Find = class(TForm)
    lbl1: TLabel;
    edtKeyWord: TEdit;
    chkCaseSensitive: TCheckBox;
    chkWholeWordMatch: TCheckBox;
    bvl1: TBevel;
    btnFind: TButton;
    btnFindNext: TButton;
    chkMatchAtFirst: TCheckBox;
    procedure btnFindClick(Sender: TObject);
    procedure btnFindNextClick(Sender: TObject);
    procedure FormClose(Sender: TObject;var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure edtKeyWordKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
    //Fhnd: THandle;
    FControl: TWinControl;
    FStrs: TStrings;   
    FFindCallBack: TFindCallBack;
    FTreeCompareMethod: TTreeCompareMethod;
    // 查找，参数表示是否是查找下一个操作
    //procedure DoFindAtTree( mmo: TMemo; bNext: Boolean );
    function DoFind(bNext: Boolean): Boolean;
  public
    LastFindLine: Integer;
    LastFindIndex: Integer;
    property TreeCompareMethod: TTreeCompareMethod read FTreeCompareMethod write FTreeCompareMethod;

    { Public declarations }
    procedure PassControl(var ctrl: TWinControl;
      CaptionDef: string = '');
    procedure PassStrings(strs: TStrings; findcallback: TFindCallBack; CaptionDef: string = '');

    // 检查是否已找到
   // function FindIt(const sSub, S: string; CaseSensitive, WholeWordMatch: Boolean): Integer;
    function DoFindAtTreeNode(nodeStart: TTreeNode; sSearchText: string;
      bCaseSensitive, bWholeWordMatch, bMatchAtFirst, bNext: Boolean): TTreeNode;overload;
    function DoFindAtTreeNode(tv: TTreeView; nodeStart:
      TTreeNode;sSearchText: string):TTreeNode;overload;
    function DoFindTreePath(tv: TTreeView; path: String; separator: String = '/'): TTreeNode;
    function DoFindAtTree(tvw: TTreeView;sSearchText: string; bCaseSensitive,
        bWholeWordMatch, bMatchAtFirst, bNext: Boolean):Boolean;
    function DoFindAtList(lvw: TListView;sSearchText: string; bCaseSensitive,
        bWholeWordMatch, bMatchAtFirst, bNext: Boolean):Boolean;
    function DoFindAtGrid(dbgrd: TCustomDBGrid;sSearchText: string; bCaseSensitive,
        bWholeWordMatch, bMatchAtFirst, bNext: Boolean):Boolean;
//    function DoFindAtEditor(editor: TCustomEdit;sSearchText: string; bCaseSensitive,
//        bWholeWordMatch, bMatchAtFirst, bNext: Boolean):Boolean;
    function DoFindAtStrings(strs: TStrings;sSearchText: string; bCaseSensitive,
        bWholeWordMatch, bMatchAtFirst, bNext: Boolean; callback: TFindCallBack):Boolean;
  end;



var
  gFindForm: TF_Find;

implementation

{$R *.dfm}

uses
  U_CommonFunc, U_fStrUtil;

function FindIt(const sSub, S: string; nFrom: Integer; CaseSensitive,
  WholeWordMatch, MatchAtFirst: Boolean): Integer;
var
  nPos: Integer;
  nWildcardIndex: Integer;
  nMatchIndex: Integer;
begin
  // 大小写敏感处理
  nPos := fStrUtil.PosFrom(sSub, S, nFrom, not CaseSensitive);
  // 检查是否需要通配符比较
  nWildcardIndex := 0;
  if 0 < fStrUtil.PosArrayFrom(sSub, ['*', '?'], nMatchIndex, 1, not CaseSensitive) then
    nWildcardIndex := fStrUtil.WildcardCompareText(sSub,
      Copy(S, nFrom, Length(S)), not CaseSensitive);

  // 通配符匹配通过 通配符不受全字匹配选项的影响  并且就相当于从头匹配了
  if nWildcardIndex > 0 then
     nPos := nWildcardIndex + nFrom - 1
  // 全字匹配处理，前后都有空格
  else if (nPos > 0) and WholeWordMatch then
  begin
    if ((nPos = 1)         or (' ' = Copy(S, nPos-1, 1))) and     // 开头
       ((nPos+Length(sSub)-1 = Length(S)) or (' ' = Copy(S, nPos+1, 1)))         // 结尾
       then
      // right, needn't do anything
    else
      nPos := 0;
  end;

  Result := nPos;
  if (nPos <> 1) and MatchAtFirst then
    Result := 0;
end;

{ TF_Find }

procedure TF_Find.PassControl(var ctrl: TWinControl;
  CaptionDef: string);
begin
  FControl := ctrl;
  if not Assigned(FControl) then
    raise Exception.Create('未找到传递给查找窗体的控件！');
  if CaptionDef <> '' then
    Caption := CaptionDef
  else
    if FControl is TCustomTreeView then
      Caption := '在树中查找'
    else if FControl is TCustomListView then
      Caption := '在列表中查找'
    else if (FControl is TCustomMemo)
            or (FControl is TCustomEdit) then
      Caption := '在文本中查找'
    else if FControl is TCustomDBGrid then
      Caption := '在表格中查找';
end;

procedure TF_Find.PassStrings(strs: TStrings; findcallback: TFindCallBack; CaptionDef: string);
begin
  FStrs := strs;
  FFindCallBack := findcallback;
  if CaptionDef <> '' then
    Caption := CaptionDef
  else
    if FControl is TCustomTreeView then
      Caption := '在树中查找'
    else if FControl is TCustomListView then
      Caption := '在列表中查找'
    else if (FControl is TCustomMemo)
            or (FControl is TCustomEdit) then
      Caption := '在文本中查找'
    else if FControl is TCustomDBGrid then
      Caption := '在表格中查找';
end;

function TF_Find.DoFindAtTreeNode(nodeStart: TTreeNode; sSearchText: string;
  bCaseSensitive, bWholeWordMatch, bMatchAtFirst, bNext: Boolean): TTreeNode;
  function GetRootNode(node: TTreeNode): TTreeNode;
  begin
    while node.Parent <> nil do begin
      node := node.Parent;
    end;
    while node.getPrevSibling <> nil do begin
      node := node.getPrevSibling;
    end;
    Result := node;
  end;
  function ReachLastNode(node: TTreeNode): Boolean;
  begin
    Result := False;
    // 无父无子无邻居
    if (node.Parent = nil) and (node.getNextSibling = nil) and (node.getFirstChild = nil) then begin
      Result := True;
      Exit;
    end;

    // 根节点的下一层，自身无邻居，其父无邻居
    if (node.Parent <> nil) and (node.Parent.Parent = nil)
       and (node.getNextSibling = nil)
       and (node.getFirstChild = nil)
       and (node.Parent.getNextSibling = nil) then begin
      Result := True;
      Exit;
    end;
  end;
var
  sWholeText: string;
  bFind: Boolean;
  node, parentNode: TTreeNode;
  loopTree: Integer;
begin
  node := nodeStart;
  Result := nil;
  loopTree := 0;
  while Assigned(node) do
  begin
    Application.ProcessMessages;

    sWholeText := node.Text;
    if Assigned(FTreeCompareMethod) then
      bFind := FTreeCompareMethod(node, sSearchText) = 0
    else
      bFind := Pos(sSearchText, sWholeText) > 0;
//    bFind := FindIt(sSearchText, sWholeText, 1,
//                    bCaseSensitive, bWholeWordMatch, bMatchAtFirst) > 0;

    if bFind and
       not (bNext and node.Selected) then
    begin
      Result := node;
      Break;
    end;

    if (loopTree >= 1) and (node = nodeStart) then begin
      // 到结尾再次从头查找时又到了开始节点，未找到，结束
      Result := nil;
      Break;
    end;

    // 没找到时，查找一下个
    if node.getFirstChild <> nil then         // 先看子节点
      node := node.getFirstChild
    else if node.getNextSibling <> nil then   // 没有子节点，看下一个邻居
      node := node.getNextSibling
    else if node.Parent <> nil then begin     // 没有子节点和邻居，有父节点，看父节点邻居，可能要向上看多层
      // 找到第一个有邻居的父节点，取得之
      parentNode := node.Parent;
      node := nil;
      while parentNode <> nil do begin
        if parentNode.getNextSibling <> nil then begin
          node := parentNode.getNextSibling;
          Break;
        end;
        if (parentNode.Parent = nil) then
          Break;
        parentNode := parentNode.Parent;
      end;
      if (node = nil) then begin
        node := GetRootNode(parentNode);
        Inc(loopTree);
        Continue;
      end;

      // 看父节点邻居时，也可也能会一直看到最后的节点，有如下几种情况
      if ReachLastNode(node) then begin
         node := GetRootNode(node);
         Inc(loopTree);
      end;
    end else if ReachLastNode(node) then begin
      node := GetRootNode(node);
      Inc(loopTree);
    end;
  end;
end;

function TF_Find.DoFindAtTreeNode(tv: TTreeView; nodeStart: TTreeNode; sSearchText: string): TTreeNode;
var
  sWholeText, sSearchText_low: string;
  bFind: Boolean;
  node, parent: TTreeNode;
  loopTree: Integer;
begin
  node := nodeStart;
  Result := nil;
  loopTree := 0;
  sSearchText_low := LowerCase(sSearchText);
  while Assigned(node) do
  begin
    Application.ProcessMessages;

    sWholeText := LowerCase(node.Text);
    if Assigned(FTreeCompareMethod) then
      bFind := FTreeCompareMethod(node, sSearchText) = 0
    else
      bFind := Pos(sSearchText_low, sWholeText) > 0;

    if bFind then
    begin
      Result := node;
      Break;
    end;

    // 当前节点不匹配，则查找一下个,
    // 按照子节点->后续邻居节点->父节点的后续邻居节点的顺序，如果还找不到，再从头开始找
    if node.getFirstChild <> nil then
      node := node.getFirstChild
    else if node.getNextSibling <> nil then
      node := node.getNextSibling
    else
    begin
      //递归查找上级节点的后续邻居节点
      parent := node.Parent;
      node := nil;
      while parent <> nil do
      begin
        if parent.getNextSibling <> nil then
        begin
          node := parent.getNextSibling;
          Break;
        end;

        parent := parent.Parent;
      end;
    end;

    if (loopTree = 0) and (node = nil) then
    begin
      loopTree := 1;
      node := tv.Items.GetFirstNode;
    end;

    if (loopTree >= 1) and (node = nodeStart) then
    begin
      // 到结尾再次从头查找时又到了开始节点，未找到，结束
      Result := nil;
      Break;
    end;
  end;
end;

function TF_Find.DoFindTreePath(tv: TTreeView; path: String; separator: String): TTreeNode;
var
  pathList: TStrings;
  i: Integer;
  node, child: TTreeNode;
begin
  Result := nil;
  pathList := TStringList.Create;
  try
    fStrUtil.Split(path, separator, pathList);
    if pathList.Count = 0 then
      Exit;

    node := tv.Items.GetFirstNode;
    for i := 0 to pathList.Count - 1 do
    begin
      if (pathList[i] = tv.Items.GetFirstNode.Text) then
        Continue;

      child := node.getFirstChild;
      while child<>nil do
      begin
        if SameStr(child.Text,pathList[i]) then
        begin
          node := child;
          Break;
        end;
        child := child.getNextSibling;
      end;
      if child = nil then
      begin
        Break;
      end;
    end;
    Result := Node;
  finally
    pathList.Free;
  end;
end;

function TF_Find.DoFindAtTree(tvw: TTreeView;sSearchText: string; bCaseSensitive,
    bWholeWordMatch, bMatchAtFirst, bNext: Boolean):Boolean;
var
  sWholeText: string;
  bFind: Boolean;
  nodeStart: TTreeNode;
begin
  Result := False;
  if sSearchText = '' then Exit;
  nodeStart := nil;
  if bNext then
    nodeStart := tvw.Selected;
  if nodeStart = nil then
    nodeStart := tvw.Items.GetFirstNode;
  if not Assigned(nodeStart) then Exit;

  bFind := False;
  while True do
  begin
    Application.ProcessMessages;
    
    sWholeText := nodeStart.Text;

    bFind := FindIt(sSearchText, sWholeText, 1,
                    bCaseSensitive, bWholeWordMatch, bMatchAtFirst) > 0;

    {  一般查找，符合即返回；
      查找下一个时，从选中节点开始，找到其后匹配条件但没被选中的节点
    }
    if bFind and
       not (bNext and (tvw.Selected = nodeStart)) then
    begin
      if tvw.Selected <> nodeStart then  // 避免重复调用 DrawFocusRect
      begin
        tvw.Selected := nodeStart;
        tvw.Canvas.DrawFocusRect(tvw.Selected.DisplayRect(False));
      end;
      Break;
    end
    else if nodeStart.getFirstChild <> nil then
      nodeStart := nodeStart.getFirstChild
    else if nodeStart.getNextSibling <> nil then
      nodeStart := nodeStart.getNextSibling
    else if nodeStart.Parent <> nil then
    begin
      nodeStart := nodeStart.Parent;
      if nodeStart.getNextSibling <> nil then
        nodeStart := nodeStart.getNextSibling
      else
        nodeStart := nil;
    end
    else
      Break;
  end;
  Result := bFind;

  if bNext and (not bFind) then
    if FormatMsgBox( '已找到尾部，是否从头开始找', mbsQuestion) = IDYES then
    begin
      tvw.Selected := nil;
      Result := DoFindAtTree(tvw, sSearchText, bCaseSensitive, bWholeWordMatch,
        bMatchAtFirst, False);
    end;
end;

function TF_Find.DoFindAtList(lvw: TListView;sSearchText: string; bCaseSensitive,
    bWholeWordMatch, bMatchAtFirst, bNext: Boolean):Boolean;
var
  i, j, nStart: Integer;
  sWholeText: string;
  bFind: Boolean;
begin
  Result := False;
  if sSearchText = '' then Exit;
  nStart := 0; 
  bFind := False;
  // 是否是查找下一个操作
  if bNext then
    if Assigned(lvw.Selected) then
      nStart := lvw.Selected.Index + 1;
  for i := nStart to lvw.Items.Count - 1 do
  begin                     
    Application.ProcessMessages;
    
    sWholeText := lvw.Items[i].Caption;
    for j := 1 to lvw.Items[i].SubItems.Count - 1 do
      sWholeText := sWholeText + #13 + lvw.Items[i].SubItems[j];

    bFind := FindIt(sSearchText, sWholeText,1,
                    bCaseSensitive, bWholeWordMatch, bMatchAtFirst) > 0;

    if bFind then
    begin
      lvw.Selected := lvw.Items[i];
      lvw.SetFocus;
      Break;
    end;
  end;
  Result := bFind;

  if bNext and (not bFind) then
    if FormatMsgBox('已找到尾部，是否从头开始找', mbsQuestion) = IDYES
      then
      Result := DoFindAtList(lvw, sSearchText, bCaseSensitive, bWholeWordMatch,
        bMatchAtFirst, False);
end;

function TF_Find.DoFindAtGrid(dbgrd: TCustomDBGrid;sSearchText: string; bCaseSensitive,
    bWholeWordMatch, bMatchAtFirst, bNext: Boolean):Boolean;
var
  i, nStartRow, nStartCol: Integer;
  sWholeText: string;
  bFind: Boolean;
  qry: TDataSet;
begin
  Result := False;
  if sSearchText = '' then Exit;
  nStartRow := 0;
  nStartCol := 0;
  bFind := False;
  if bNext then
    if Assigned(dbgrd.SelectedField) then
    begin
      nStartCol := dbgrd.SelectedField.Index + 1;
      nStartRow := dbgrd.DataSource.DataSet.RecNo + 1;
    end;
  qry := dbgrd.DataSource.DataSet;
  qry.First;
  if nStartRow <> 0 then
    qry.MoveBy(nStartRow);
  while not qry.Eof do
  begin                
    Application.ProcessMessages;
    
    for i := nStartCol to qry.FieldCount - 1 do
    begin
      sWholeText := qry.Fields[i].AsString;

      bFind := FindIt(sSearchText, sWholeText, 1,
                      bCaseSensitive, bWholeWordMatch, bMatchAtFirst) > 0;

      if bFind then
      begin
        dbgrd.SelectedField := qry.Fields[i];
        dbgrd.SetFocus;
        Break;
      end;
    end;
    qry.Next;
  end;
  Result := bFind;
  if bNext and (not bFind) then
    if IDYES = FormatMsgBox('已找到尾部，是否从头开始找', mbsQuestion) then
      Result := DoFindAtGrid(dbgrd, sSearchText, bCaseSensitive, bWholeWordMatch,
        bMatchAtFirst, False);
end;

    
//function TF_Find.DoFindAtEditor(editor: TCustomEdit;sSearchText: string; bCaseSensitive,
//    bWholeWordMatch, bMatchAtFirst, bNext: Boolean):Boolean;
//begin
//  //TODO
//end;  

function TF_Find.DoFindAtStrings(strs: TStrings;sSearchText: string; bCaseSensitive,
  bWholeWordMatch, bMatchAtFirst, bNext: Boolean; callback: TFindCallBack):Boolean;
var
  i, nStartLine, nStartIndex: Integer;
  sWholeText: string;
  bFind: Boolean;
  nFindIndex: Integer;
begin
  Result := False;
  if sSearchText = '' then
    Exit;
  nStartLine := 0;             // 变量默认从头开始
  nStartIndex := 1;            // 变量默认从第一个索引开始
  bFind := False;
  if bNext then
  begin
    if LastFindLine > -1 then
      nStartLine := LastFindLine;    // 从上次查找的行开始
    if LastFindIndex > 0 then
      nStartIndex := LastFindIndex+Length(sSearchText); // 从上次查找的索引后面开始，如果上次没找到就从头开始
  end;

  i := nStartLine;
  while i < strs.Count do
  begin           
    Application.ProcessMessages;
    
    sWholeText := strs[i];
    // 在上次查找过的行查找的时候，就从指定索引开始，否则从第一个索引开始
    if bNext and (i=nStartLine) then
      nFindIndex := FindIt(sSearchText, strs[i], nStartIndex,
               bCaseSensitive, bWholeWordMatch, bMatchAtFirst)
    else
      nFindIndex := FindIt(sSearchText, strs[i], 1,
               bCaseSensitive, bWholeWordMatch, bMatchAtFirst);

    bFind := nFindIndex > 0;

    if bFind then
    begin
      LastFindLine := i;
      LastFindIndex := nFindIndex;
//      FindIndex := FindIndex + Length(sSearchText) + Pos(sSearchText,
//          Copy(strs[i], FindIndex + Length(sSearchText), Length(strs[i])));
      if Assigned(CallBack) then
        CallBack(LastFindLine, LastFindIndex, Length(sSearchText));
      Break;
    end;
    Inc(i);
  end;
  Result := bFind;
  if bNext and (not bFind) then
    if IDYES = FormatMsgBox('已找到尾部，是否从头开始找', mbsQuestion) then
      Result := DoFindAtStrings(strs, sSearchText, bCaseSensitive, bWholeWordMatch,
          bMatchAtFirst, False, callback);
end;

function TF_Find.DoFind(bNext: Boolean): Boolean;
begin
  Result := False;
  if '' = edtKeyWord.Text then
  begin
    MessageBox(Handle, '关键字不能为空！', '提示', MB_OK + MB_ICONINFORMATION);
    Exit;
  end;
  if FStrs <> nil then
    Result := DoFindAtStrings(FStrs, edtKeyWord.Text, chkCaseSensitive.Checked,
        chkWholeWordMatch.Checked, chkMatchAtFirst.Checked, bNext, FFindCallBack)  
//  else if FControl is TCustomEdit then
//    Result := DoFindAtEditor(TEdit(FControl), edtKeyWord.Text, chkCaseSensitive.Checked,
//        chkWholeWordMatch.Checked, chkMatchAtFirst.Checked, bNext)
  else if FControl is TTreeView then      // 在treeview中查找
    Result := DoFindAtTree(TTreeView(FControl), edtKeyWord.Text, chkCaseSensitive.Checked,
        chkWholeWordMatch.Checked, chkMatchAtFirst.Checked, bNext)
  else if FControl is TListView then     // 在listview中查找
    Result := DoFindAtList(TListView(FControl), edtKeyWord.Text, chkCaseSensitive.Checked,
        chkWholeWordMatch.Checked, chkMatchAtFirst.Checked, bNext)
  else if FControl is TDBGrid then       //
    Result := DoFindAtGrid(TDBGrid(FControl), edtKeyWord.Text, chkCaseSensitive.Checked,
        chkWholeWordMatch.Checked, chkMatchAtFirst.Checked, bNext);
end;

procedure TF_Find.btnFindClick(Sender: TObject);
begin
  if DoFind(False) then
    btnFindNext.Enabled := True;
end;

procedure TF_Find.btnFindNextClick(Sender: TObject);
begin
  DoFind(True);
end;

procedure TF_Find.FormClose(Sender: TObject;var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TF_Find.FormDestroy(Sender: TObject);
begin
  gFindForm := nil;
end;

procedure TF_Find.edtKeyWordKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    btnFindNext.Click;
end;

end.

