{**
 * <p>Title: Find  </p>
 * <p>Copyright: Copyright (c) 2006/10/27</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: �ڿؼ��в���</p>

  ʹ�����·�ʽ��ʾ���Ҵ���
  if not Assigned(gFindForm) then
    gFindForm := TF_Find.Create( Self );
  with gFindForm do
  begin
    PassControl( TWinControl( �ؼ����� ) );
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
    // ���ң�������ʾ�Ƿ��ǲ�����һ������
    //procedure DoFindAtTree( mmo: TMemo; bNext: Boolean );
    function DoFind(bNext: Boolean): Boolean;
  public          
    LastFindLine: Integer;
    LastFindIndex: Integer;

    { Public declarations }
    procedure PassControl(var ctrl: TWinControl;
      CaptionDef: string = '');
    procedure PassStrings(strs: TStrings; findcallback: TFindCallBack; CaptionDef: string = '');

    // ����Ƿ����ҵ�
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
  // ��Сд���д���
  nPos := fStrUtil.PosFrom(sSub, S, nFrom, not CaseSensitive);
  // ����Ƿ���Ҫͨ����Ƚ�
  nWildcardIndex := 0;
  if 0 < fStrUtil.PosArrayFrom(sSub, ['*', '?'], nMatchIndex, 1, not CaseSensitive) then
    nWildcardIndex := fStrUtil.WildcardCompareText(sSub,
      Copy(S, nFrom, Length(S)), not CaseSensitive);

  // ͨ���ƥ��ͨ�� ͨ�������ȫ��ƥ��ѡ���Ӱ��  ���Ҿ��൱�ڴ�ͷƥ����
  if nWildcardIndex > 0 then
     nPos := nWildcardIndex + nFrom - 1
  // ȫ��ƥ�䴦��ǰ���пո�
  else if (nPos > 0) and WholeWordMatch then
  begin
    if ((nPos = 1)         or (' ' = Copy(S, nPos-1, 1))) and     // ��ͷ
       ((nPos+Length(sSub)-1 = Length(S)) or (' ' = Copy(S, nPos+1, 1)))         // ��β
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
    raise Exception.Create('δ�ҵ����ݸ����Ҵ���Ŀؼ���');
  if CaptionDef <> '' then
    Caption := CaptionDef
  else
    if FControl is TCustomTreeView then
      Caption := '�����в���'
    else if FControl is TCustomListView then
      Caption := '���б��в���'
    else if (FControl is TCustomMemo)
            or (FControl is TCustomEdit) then
      Caption := '���ı��в���'
    else if FControl is TCustomDBGrid then
      Caption := '�ڱ���в���';
end;  

procedure TF_Find.PassStrings(strs: TStrings; findcallback: TFindCallBack; CaptionDef: string);
begin
  FStrs := strs;
  FFindCallBack := findcallback;
  if CaptionDef <> '' then
    Caption := CaptionDef
  else
    if FControl is TCustomTreeView then
      Caption := '�����в���'
    else if FControl is TCustomListView then
      Caption := '���б��в���'
    else if (FControl is TCustomMemo)
            or (FControl is TCustomEdit) then
      Caption := '���ı��в���'
    else if FControl is TCustomDBGrid then
      Caption := '�ڱ���в���';
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
      if tvw.Selected <> nodeStart then  // �����ظ����� DrawFocusRect
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
    if FormatMsgBox( '���ҵ�β�����Ƿ��ͷ��ʼ��', mbsQuestion) = IDYES then
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
  // �Ƿ��ǲ�����һ������
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
    if FormatMsgBox('���ҵ�β�����Ƿ��ͷ��ʼ��', mbsQuestion) = IDYES
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
    if IDYES = FormatMsgBox('���ҵ�β�����Ƿ��ͷ��ʼ��', mbsQuestion) then
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
  nStartLine := 0;             // ����Ĭ�ϴ�ͷ��ʼ
  nStartIndex := 1;            // ����Ĭ�ϴӵ�һ��������ʼ
  bFind := False;
  if bNext then
  begin
    if LastFindLine > -1 then
      nStartLine := LastFindLine;    // ���ϴβ��ҵ��п�ʼ
    if LastFindIndex > 0 then
      nStartIndex := LastFindIndex+Length(sSearchText); // ���ϴβ��ҵ��������濪ʼ������ϴ�û�ҵ��ʹ�ͷ��ʼ
  end;

  i := nStartLine;
  while i < strs.Count do
  begin           
    Application.ProcessMessages;
    
    sWholeText := strs[i];
    // ���ϴβ��ҹ����в��ҵ�ʱ�򣬾ʹ�ָ��������ʼ������ӵ�һ��������ʼ
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
    if IDYES = FormatMsgBox('���ҵ�β�����Ƿ��ͷ��ʼ��', mbsQuestion) then
      Result := DoFindAtStrings(strs, sSearchText, bCaseSensitive, bWholeWordMatch,
          bMatchAtFirst, False, callback);
end;

function TF_Find.DoFind(bNext: Boolean): Boolean;
begin
  Result := False;
  if '' = edtKeyWord.Text then
  begin
    MessageBox(Handle, '�ؼ��ֲ���Ϊ�գ�', '��ʾ', MB_OK + MB_ICONINFORMATION);
    Exit;
  end;
  if FStrs <> nil then
    Result := DoFindAtStrings(FStrs, edtKeyWord.Text, chkCaseSensitive.Checked,
        chkWholeWordMatch.Checked, chkMatchAtFirst.Checked, bNext, FFindCallBack)  
//  else if FControl is TCustomEdit then
//    Result := DoFindAtEditor(TEdit(FControl), edtKeyWord.Text, chkCaseSensitive.Checked,
//        chkWholeWordMatch.Checked, chkMatchAtFirst.Checked, bNext)
  else if FControl is TTreeView then      // ��treeview�в���
    Result := DoFindAtTree(TTreeView(FControl), edtKeyWord.Text, chkCaseSensitive.Checked,
        chkWholeWordMatch.Checked, chkMatchAtFirst.Checked, bNext)
  else if FControl is TListView then     // ��listview�в���
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

