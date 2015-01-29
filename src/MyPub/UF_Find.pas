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
    // 查找，参数表示是否是查找下一个操作
    //procedure DoFindAtTree( mmo: TMemo; bNext: Boolean );
    function DoFind(bNext: Boolean): Boolean;
  public          
    LastFindLine: Integer;
    LastFindIndex: Integer;

    { Public declarations }
    procedure PassControl(var ctrl: TWinControl;
      CaptionDef: string = '');
    procedure PassStrings(strs: TStrings; findcallback: TFindCallBack; CaptionDef: string = '');

    // 检查是否已找到
   // function FindIt(const sSub, S: string; CaseSensitive, WholeWordMatch: Boolean): Integer;
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

