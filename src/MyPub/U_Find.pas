{**
 * <p>Title: Find  </p>
 * <p>Copyright: Copyright (c) 2006/10/27</p>
 * <p>Company: Thunisoft</p>
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
unit U_Find;

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
    procedure DoFind(bNext: Boolean);
  public
    { Public declarations }
    procedure PassControl(var ctrl: TWinControl;
      CaptionDef: string = '');
    procedure PassStrings(strs: TStrings; findcallback: TFindCallBack; CaptionDef: string = '');
  end;
                

  // ����Ƿ����ҵ�
 // function FindIt(const sSub, S: string; CaseSensitive, WholeWordMatch: Boolean): Integer;
  procedure DoFindAtTree(tvw: TTreeView;sSearchText: string; bCaseSensitive,
      bWholeWordMatch, bNext: Boolean);
  procedure DoFindAtList(lvw: TListView;sSearchText: string; bCaseSensitive,
      bWholeWordMatch, bNext: Boolean);
  procedure DoFindAtGrid(dbgrd: TDBGrid;sSearchText: string; bCaseSensitive,
      bWholeWordMatch, bNext: Boolean);
  procedure DoFindAtStrings(strs: TStrings;sSearchText: string; bCaseSensitive,
      bWholeWordMatch, bNext: Boolean; callback: TFindCallBack);

var
  gFindForm: TF_Find;
  FindLine: Integer;
  FindIndex: Integer;

implementation

{$R *.dfm}

uses
  U_CommonFunc, Math, U_fStrUtil;

function FindIt(const sSub, S: string; nFrom: Integer; CaseSensitive,
  WholeWordMatch: Boolean): Integer;
var
  nPos: Integer;
begin
  // ��Сд���д���
  nPos := fStrUtil.PosFrom(sSub, S, nFrom, not CaseSensitive);

  // ȫ��ƥ�䴦��ǰ���пո�
  if (nPos > 0) and WholeWordMatch then
  begin
    if ((nPos = 1)         or (' ' = Copy(S, nPos-1, 1))) and     // ��ͷ
       ((nPos = Length(S)) or (' ' = Copy(S, nPos+1, 1)))         // ��β
       then
      // right, needn't do anything
    else
      nPos := 0;
  end;
  Result := nPos;
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
    if FControl is TTreeView then
      Caption := '�����в���'
    else if FControl is TListView then
      Caption := '���б��в���'
    else if FControl is TMemo then
      Caption := '���ı��в���'
    else if FControl is TDBGrid then
      Caption := '�ڱ���в���';
end;  

procedure TF_Find.PassStrings(strs: TStrings; findcallback: TFindCallBack; CaptionDef: string);
begin
  FStrs := strs;
  FFindCallBack := findcallback;
  if CaptionDef <> '' then
    Caption := CaptionDef
  else
    if FControl is TTreeView then
      Caption := '�����в���'
    else if FControl is TListView then
      Caption := '���б��в���'
    else if FControl is TMemo then
      Caption := '���ı��в���'
    else if FControl is TDBGrid then
      Caption := '�ڱ���в���';
end;

procedure DoFindAtTree(tvw: TTreeView;sSearchText: string; bCaseSensitive,
    bWholeWordMatch, bNext: Boolean);
var
  sWholeText: string;
  bFind: Boolean;
  nodeStart: TTreeNode;
begin
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
    sWholeText := nodeStart.Text;

    bFind := FindIt(sSearchText, sWholeText, 1,
                    bCaseSensitive, bWholeWordMatch) > 0;

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

  if bNext and (not bFind) then
    if FormatMsgBox( '���ҵ�β�����Ƿ��ͷ��ʼ��', mbsQuestion) = IDYES then
    begin
      tvw.Selected := nil;
      DoFindAtTree(tvw, sSearchText, bCaseSensitive, bWholeWordMatch, False);
    end;
end;

procedure DoFindAtList(lvw: TListView;sSearchText: string; bCaseSensitive,
    bWholeWordMatch, bNext: Boolean);
var
  i, j, nStart: Integer;
  sWholeText: string;
  bFind: Boolean;
begin      
  if sSearchText = '' then Exit;
  nStart := 0; 
  bFind := False;
  // �Ƿ��ǲ�����һ������
  if bNext then
    if Assigned(lvw.Selected) then
      nStart := lvw.Selected.Index + 1;
  for i := nStart to lvw.Items.Count - 1 do
  begin
    sWholeText := lvw.Items[i].Caption;
    for j := 1 to lvw.Items[i].SubItems.Count - 1 do
      sWholeText := sWholeText + #13 + lvw.Items[i].SubItems[j];

    bFind := FindIt(sSearchText, sWholeText,1,
                    bCaseSensitive, bWholeWordMatch) > 0;

    if bFind then
    begin
      lvw.Selected := lvw.Items[i];
      lvw.SetFocus;
      Break;
    end;
  end;  

  if bNext and (not bFind) then
    if FormatMsgBox('���ҵ�β�����Ƿ��ͷ��ʼ��', mbsQuestion) = IDYES
      then
      DoFindAtList(lvw, sSearchText, bCaseSensitive, bWholeWordMatch, False);
end;

procedure DoFindAtGrid(dbgrd: TDBGrid;sSearchText: string; bCaseSensitive,
    bWholeWordMatch, bNext: Boolean);
var
  i, nStartRow, nStartCol: Integer;
  sWholeText: string;
  bFind: Boolean;
  qry: TDataSet;
begin    
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
    for i := nStartCol to qry.FieldCount - 1 do
    begin
      sWholeText := qry.Fields[i].AsString;

      bFind := FindIt(sSearchText, sWholeText, 1,
                      bCaseSensitive, bWholeWordMatch) > 0;

      if bFind then
      begin
        dbgrd.SelectedField := qry.Fields[i];
        dbgrd.SetFocus;
        Exit;
      end;
    end;
    qry.Next;
  end;
  if bNext and (not bFind) then
    if IDYES = FormatMsgBox('���ҵ�β�����Ƿ��ͷ��ʼ��', mbsQuestion) then
      DoFindAtGrid(dbgrd, sSearchText, bCaseSensitive, bWholeWordMatch, False);
end; 

procedure DoFindAtStrings(strs: TStrings;sSearchText: string; bCaseSensitive, bWholeWordMatch, bNext: Boolean; callback: TFindCallBack);
var
  i, nStart: Integer;
  sWholeText: string;
  bFind: Boolean;
begin     
  if sSearchText = '' then Exit;
  nStart := 0;
  bFind := False;
  if bNext and (FindLine <> -1) then
    nStart := FindLine + 1;
  FindLine := -1;
//  if FindIndex <> 0 then
//  begin
//    i := nStart;
//    bFind := FindIt(sSearchText, Copy(strs[i], FindIndex + Length(sSearchText), Length(strs[i])),
//        chkCaseSensitive.Checked, chkWholeWordMatch.Checked) > 0;
//    if bFind then
//    begin
//      FFindLine := i;
//      FindIndex := FindIndex + Length(sSearchText) + Pos(sSearchText,
//          Copy(strs[i], FindIndex + Length(sSearchText)+1, Length(strs[i])));
//      if @FFindCallBack <> nil then
//        FFindCallBack(FFindLine, FindIndex, Length(sSearchText));
//      Exit;
//    end;
//  end;

  i := nStart;
  while i < strs.Count do
  begin
    sWholeText := strs[i];
    bFind := FindIt(sSearchText, strs[i], 1,
             bCaseSensitive, bWholeWordMatch) > 0;

    if bFind then
    begin
      FindLine := i;
      FindIndex := Pos(sSearchText,strs[i]);
//      FindIndex := FindIndex + Length(sSearchText) + Pos(sSearchText,
//          Copy(strs[i], FindIndex + Length(sSearchText), Length(strs[i])));
      if @CallBack <> nil then
        CallBack(FindLine, FindIndex, Length(sSearchText));
      Exit;
    end;
    Inc(i);
  end;
  if bNext and (not bFind) then
    if IDYES = FormatMsgBox('���ҵ�β�����Ƿ��ͷ��ʼ��', mbsQuestion) then
        DoFindAtStrings(strs, sSearchText, bCaseSensitive, bWholeWordMatch, False, callback);
end;

procedure TF_Find.DoFind(bNext: Boolean);
begin
  if '' = edtKeyWord.Text then
  begin
    MessageBox(Handle, '�ؼ��ֲ���Ϊ�գ�', '��ʾ', MB_OK + MB_ICONINFORMATION);
    Exit;
  end;
  if FStrs <> nil then
    DoFindAtStrings(FStrs, edtKeyWord.Text, chkCaseSensitive.Checked,
        chkWholeWordMatch.Checked, bNext, FFindCallBack)
  else if FControl is TTreeView then      // ��treeview�в���
    DoFindAtTree(TTreeView(FControl), edtKeyWord.Text, chkCaseSensitive.Checked,
        chkWholeWordMatch.Checked, bNext)
  else if FControl is TListView then     // ��listview�в���
    DoFindAtList(TListView(FControl), edtKeyWord.Text, chkCaseSensitive.Checked,
        chkWholeWordMatch.Checked, bNext)
  else if FControl is TDBGrid then       //
    DoFindAtGrid(TDBGrid(FControl), edtKeyWord.Text, chkCaseSensitive.Checked,
        chkWholeWordMatch.Checked, bNext);
end;

procedure TF_Find.btnFindClick(Sender: TObject);
begin
  DoFind(False);
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

