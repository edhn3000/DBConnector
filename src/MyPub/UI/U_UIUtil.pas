{**
 * <p>Title: U_UIUtil.pas </p>
 * <p>Copyright: Copyright (c) 2007-3-7 </p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 用户界面处理 </p>
 *}
unit U_UIUtil;

interface

uses
  Windows, Messages, Dialogs, Forms, ComCtrls, Controls, Classes, Types,
  ExtCtrls, Graphics, ADODB, StdCtrls, DBGrids, DB, Printers, OleServer,
  U_FieldInfo,
  U_fStrUtil;

const
  C_sSortFlag_Asc = '▲';
  C_sSortFlag_Desc = '▼';

type
  TAryInt = array of Integer;

  { TUIUtil }
  TUIUtil = class
  private
    FabColSortFlags: array of Boolean;

  protected

  public 
    procedure ExChangeListItem(lvw: TListView;nFst, nSnd: Integer; nTemp: Integer);
    // 排序listview，冒泡算法，对每次循环起止、列表是否beginupdate做过智能检查
    procedure DoSortList(Alvw: TListView;nSortedItem: Integer);

    function BuildStringByChar(c: Char; nLength: Integer): string;
  public
    constructor Create;
    destructor Destroy;override;

    { ListView处理 }
    procedure RefreshSortFlags(ALvw: TListView);
    procedure SortList(ALvw: TListView;nColIndex: Integer);       
    procedure ReverseSortList(Alvw: TListView);
    procedure ResortList(ALvw: TListView);

    function GetItemWidths(Alvw: TListView; ACanvas: TCanvas):TAryInt;
    function GetItemLengths(Alvw: TListView):TAryInt;
    procedure AjustListItemLengths(Aslst: TStrings; AItemSeparator: string;
        NewSeparator: string; aiLengths: TAryInt; nTitlePosition: Integer;
        cLineChar: Char = '-');  // 此方法不完善

    function FindListItem(Alvw: TListView; sKeyWord: string; ColumnIndex: Integer): TListItem;
    procedure ListViewAddColumns(Alvw: TListView; asCaption: array of string; aiWidth: array of Integer);
    function GetListViewHittedItem(lvw: TListView; X, Y: Integer; aHT:THitTests):TListItem;

    { ListBox }
    function GetListBoxSelected(lbx: TListBox):Integer;
    function GetListBoxHittedIndex(lbx: TListBox; X, Y: Integer):Integer;

    { 图象处理 }
    procedure ClearPhoto( AImg: TImage );
    procedure LoadPhoto( AImg: TImage; AImgFileName: string );
    // 写字符串， 水平布局：0中间，-1左侧，1右侧；垂直布局：0中间，-1上测，1下侧
    procedure DrawStrOnImage(AImg: TImage; AText: string;
      HorizontalAlign: Integer = 0; VerticalAlign: Integer = 0 ); 
      
    { DBGrid处理 }
    function GetFieldWidths(AQ: TADOQuery; ACanvas: TCanvas):TAryInt;
    function GetFieldLengths(AQ: TADOQuery; bAjustTitle: Boolean = True):TAryInt;
    procedure SortQuery(AQ: TADOQuery; nFieldIndex: Integer);
    procedure SortDBGrid(dbgrd: TDBGrid; nColumnIndex: Integer; bFitColWidth: Boolean = True);
    procedure PrintDBGrid(DbGrid:TDbGrid;Title:String);

    { TreeView }
    function GetTreeViewHittedNode(tvw: TTreeView; X, Y: Integer;
      bWithIndent: Boolean = False):TTreeNode;overload;
    function GetTreeViewHittedNode(tvw: TTreeView; X, Y: Integer;
      aHT:THitTests):TTreeNode;overload;

  end;

  function UIUtil:TUIUtil;

implementation

uses
  SysUtils, jpeg, Variants, U_PlanarList;

var
  FUIUtil: TUIUtil;

function UIUtil:TUIUtil;
begin
  if not Assigned(FUIUtil) then
    FUIUtil := TUIUtil.Create;
  Result := FUIUtil;
end;
       
{ TUIUtil }

constructor TUIUtil.Create;
begin

end;

destructor TUIUtil.Destroy;
begin

  inherited;
end;

procedure TUIUtil.RefreshSortFlags(ALvw: TListView);
var
  i: Integer;
begin
  SetLength(FabColSortFlags, ALvw.Columns.Count);
  for i := Low(FabColSortFlags) to High(FabColSortFlags) do
    FabColSortFlags[i] := False;

  for i := 0 to ALvw.Columns.Count - 1 do
  begin
    ALvw.Columns[i].Caption := StringReplace(
      Alvw.Columns[i].Caption,
      C_sSortFlag_Asc, '', [rfReplaceAll]);
    Alvw.Columns[i].Caption := StringReplace(
      Alvw.Columns[i].Caption,
      C_sSortFlag_Desc, '', [rfReplaceAll]);
  end;
end;

procedure TUIUtil.SortList(ALvw: TListView;nColIndex: Integer);
var
  i: Integer;
begin
  if Length(FabColSortFlags) < ALvw.Columns.Count then
  begin
    RefreshSortFlags(ALvw);
  end;

  // 当前列的排序方式先变成相反的方式
  FabColSortFlags[nColIndex] := not FabColSortFlags[nColIndex];

  if (Pos(C_sSortFlag_Asc, ALvw.Columns[nColIndex].Caption)>0) then   // 该列已经排列过，只需要反排一次
  begin
    ReverseSortList(ALvw);

    ALvw.Columns[nColIndex].Caption := StringReplace(
      Alvw.Columns[nColIndex].Caption,
      C_sSortFlag_Asc, '', [rfReplaceAll]);
    ALvw.Columns[nColIndex].Caption :=
      ALvw.Columns[nColIndex].DisplayName + C_sSortFlag_Desc;
  end
  else if (Pos(C_sSortFlag_Desc, ALvw.Columns[nColIndex].Caption)>0) then
  begin
    ReverseSortList(ALvw);

    Alvw.Columns[nColIndex].Caption := StringReplace(
      Alvw.Columns[nColIndex].Caption,
      C_sSortFlag_Desc, '', [rfReplaceAll]);
    ALvw.Columns[nColIndex].Caption :=
      ALvw.Columns[nColIndex].DisplayName + C_sSortFlag_Asc
  end
  else
  begin                                                    // 该列没有排列过
    DoSortList(ALvw, nColIndex);
    // 先把所有列的上或向下箭头去掉
    for i := 0 to ALvw.Columns.Count - 1 do
    begin
      ALvw.Columns[i].Caption := StringReplace(
        Alvw.Columns[i].Caption,
        C_sSortFlag_Asc, '', [rfReplaceAll]);
      Alvw.Columns[i].Caption := StringReplace(
        Alvw.Columns[i].Caption,
        C_sSortFlag_Desc, '', [rfReplaceAll]);
    end;

    // 为列标题添加箭头
    if FabColSortFlags[nColIndex] then
      ALvw.Columns[nColIndex].Caption :=
        ALvw.Columns[nColIndex].DisplayName + C_sSortFlag_Asc
    else
      ALvw.Columns[nColIndex].Caption :=
        ALvw.Columns[nColIndex].DisplayName + C_sSortFlag_Desc;
  end; 
end;

procedure TUIUtil.ExChangeListItem(lvw: TListView; nFst, nSnd,
  nTemp: Integer);
begin
  lvw.Items[nTemp] := lvw.Items[nFst];
  lvw.Items[nFst]  := lvw.Items[nSnd];
  lvw.Items[nSnd]  := lvw.Items[nTemp];
end;   

procedure TUIUtil.ReverseSortList(Alvw: TListView);
var
  nSortEnd, nOldCount: Integer;
  i: Integer;
begin
  nOldCount := Alvw.Items.Count;
  if nOldCount = 0 then
    Exit;
  nSortEnd := Trunc((nOldCount-1)/2); //索引一半取整
  Alvw.Items.Add;
  for i := 0 to nSortEnd do
  begin
    ExChangeListItem(Alvw, i, nOldCount-1-i, nOldCount);
  end;
  Alvw.Items.Delete(nOldCount);
end;

procedure TUIUtil.DoSortList(Alvw: TListView;nSortedItem: Integer);
var
  i, nTempForReplace, nCompareResult: Integer;
  nChanged, nBound: Integer;
  s1, s2: string;
  bBeginUpdate: Boolean;
  bAddedTemp: Boolean;
  bAsc: Boolean;
  nLoopStart, nLoopEnd: Integer;
  nAnother: Integer;
begin
  bBeginUpdate := False;
  bAddedTemp := False;
  nTempForReplace := Alvw.Items.Count;  // 在列表尾部添加一个元素作为临时交换变量，仅在需要时添加，后面有检测
  if nSortedItem = -1 then
    bAsc := True
  else
    bAsc := FabColSortFlags[nSortedItem];
  try
    if bBeginUpdate then
      nBound := Alvw.Items.Count-3  // 最多可使用到Count-2,下面要用i+1，所以这里要写Count-3
    else
      nBound := Alvw.Items.Count-2;
      
    nChanged := nBound;
    while nChanged <> 0 do begin
      nChanged := 0;
      nLoopStart := 0;
      nLoopEnd := nBound;

      i := nLoopStart;
      // 一次排序循环
      while i <= nLoopEnd do begin
        // 获得相邻两元素的值
        nAnother := i+1;
        if (nSortedItem = 0) or (nSortedItem = -1) then                // caption
        begin
          s1 := Alvw.Items[i].Caption;
          s2 := Alvw.Items[nAnother].Caption;
        end else begin
          s1 := Alvw.Items[i].SubItems[nSortedItem-1];
          s2 := Alvw.Items[nAnother].SubItems[nSortedItem-1];
        end;

        // 比较大小决定顺序
        nCompareResult := fStrUtil.CompareNumberText(s1, s2); // CompareText(s1, s2);
               
        // 检查相邻两元素是否需要调整顺序
        if (bAsc and (nCompareResult > 0))              //升序排列 且前者大于后者
           or ((not bAsc) and (nCompareResult < 0))     //降序     且后者大于前者
           then begin
          // 第一次进行元素交换时做BeginUpdate
          if not bBeginUpdate then begin
            bBeginUpdate := True;  
            Alvw.Items.BeginUpdate;
          end;
          if not bAddedTemp then begin
            bAddedTemp := True;
            Alvw.Items.Add;
          end;

          nChanged := i;   // 最后一次排序的索引，即之后是符合顺序的，之前是还需要排序的
          ExChangeListItem(Alvw, i, nAnother, nTempForReplace)
        end;
        
        Inc(i);
      end;
      nBound := nChanged; // 作为下次的最后索引界限
    end;
  finally
    if bAddedTemp then
      Alvw.Items[nTempForReplace].Delete;
    if bBeginUpdate then begin
      Alvw.Items.EndUpdate;
    end;
  end;
end;

procedure TUIUtil.ResortList(ALvw: TListView);
var
  nOldSortIndex: Integer;
//  nSelectedIndex: Integer;
  function GetOldSortIndex:Integer;
  var
    i: Integer;
  begin
    Result := -1;
    for i := 0 to ALvw.Columns.Count - 1 do
    begin
      if (Pos(C_sSortFlag_Asc, ALvw.Columns[i].Caption) > 0) or
         (Pos(C_sSortFlag_Desc, ALvw.Columns[i].Caption) > 0) then
      begin
        Result := i;
        Break;
      end;
    end;
  end;
begin
//  if Assigned(ALvw.Selected) then
//    nSelectedIndex := ALvw.Selected.Index
//  else
//    nSelectedIndex := -1;
  nOldSortIndex := GetOldSortIndex;
  if nOldSortIndex <> -1 then
    DoSortList(ALvw, nOldSortIndex);
//  if nSelectedIndex > 0 then
//    ALvw.Selected := ALvw.Items[nSelectedIndex];
end;

procedure TUIUtil.AjustListItemLengths(Aslst: TStrings; AItemSeparator: string;
    NewSeparator: string; aiLengths: TAryInt; nTitlePosition: Integer; cLineChar: Char);
var
  strs: TStrings;
  i, nStrIndex: Integer;
  sFormat, sLine, sNewStr: string;
begin
  strs := TStringList.Create;
  try
    for i := 0 to Aslst.Count - 1 do
    begin
      fStrUtil.Split(Aslst[i], AItemSeparator, strs); // 分解开再根据aiLengths逐个调整长度

//      while strs.Count < Length(aiLengths) do
//      begin
//      end;

//      Aslst[i] := '';
      sNewStr := '';
      for nStrIndex := 0 to strs.Count - 1 do
      begin
        sLine := strs[nStrIndex];
        if fStrUtil.IsNumber(sLine) then
          sFormat := '%' + IntToStr(aiLengths[nStrIndex]) + 's'      // 数字右对齐
        else
          sFormat := '%-' + IntToStr(aiLengths[nStrIndex]) + 's';    // 字符串左对齐

//        strs[nStrIndex] := Format(sFormat, [sLine]);
//        if nStrIndex < strs.Count - 1 then
        if sNewStr <> '' then
          sNewStr := sNewStr + NewSeparator + Format(sFormat, [sLine]) 
        else
          sNewStr := Format(sFormat, [sLine]);
      end;
      
      Aslst[i] := sNewStr;
    end;
    // 分隔线
    sLine := '';
    for i := Low(aiLengths) to High(aiLengths) do
    begin
      if sLine = '' then
        sLine := BuildStringByChar(cLineChar, aiLengths[i])
      else
        sLine := sLine + NewSeparator + BuildStringByChar(cLineChar, aiLengths[i]);
    end;
    Aslst.Insert(nTitlePosition, sLine);

  finally
    strs.Free;
  end;
end;  

function TUIUtil.GetItemLengths(Alvw: TListView): TAryInt;
var
  i, n, nLen: Integer;
  sitemText: string;
begin
  if not Assigned(Alvw) then
    Exit;
  with Alvw do
  begin
    SetLength(Result, Columns.Count);

    // 初时宽度为列名宽
    for i := 0 to Columns.Count - 1 do
    begin
      nLen := Length(Columns[i].DisplayName);
//      if Result[i] < nLen then
        Result[i] := nLen;
    end;

    for i := 0 to Items.Count - 1 do
      for n := 0 to Columns.Count - 1 do
      begin
        if n = 0 then
          sitemText := Items[i].Caption
        else
          sitemText := Items[i].SubItems[n-1];

        nLen := Length(sitemText);
        if Result[n] < nLen then
          Result[n] := nLen;
      end;
  end;
end;

function TUIUtil.GetItemWidths(Alvw: TListView;
  ACanvas: TCanvas): TAryInt;
var
  i, n, nLen: Integer;
  sitemText: string;
begin
  if not Assigned(Alvw) then
    Exit;
  with Alvw do
  begin
    SetLength(Result, Columns.Count);

    // 初时宽度为列名宽
    for i := 0 to Columns.Count - 1 do
    begin
      nLen := ACanvas.TextWidth(Columns[i].DisplayName);
//      if Result[i] < nLen then
        Result[i] := nLen;
    end;

    for i := 0 to Items.Count - 1 do
      for n := 0 to Columns.Count - 1 do
      begin
        if n = 0 then
          sitemText := Items[i].Caption
        else
          sitemText := Items[i].SubItems[n-1];

        nLen := ACanvas.TextWidth(sitemText);
        if Result[n] < nLen then
          Result[n] := nLen;
      end;
  end;
end;

procedure TUIUtil.ClearPhoto( AImg: TImage );
begin
  AImg.Picture.Graphic := nil;
  AImg.Refresh;
end;

procedure TUIUtil.LoadPhoto( AImg: TImage; AImgFileName: string );
var
  jpg: TJPEGImage;
  sExt: string;
begin
  if not FileExists(AImgFileName) then Exit;
  ClearPhoto(AImg);
  sExt := ExtractFileExt(AImgFileName);
  if SameText(sExt, 'jpg') or SameText(sExt, 'jpeg') then
  begin
    jpg := TJPEGImage.Create;
    try
      jpg.LoadFromFile(AImgFileName);
      AImg.Picture.Graphic := jpg;
    finally
      jpg.Free;
    end;
  end
  else
    AImg.Picture.LoadFromFile( AImgFileName );
end;

procedure TUIUtil.DrawStrOnImage(AImg: TImage;
    AText: string; HorizontalAlign: Integer; VerticalAlign: Integer );
var
  x, y: Integer;
begin
  with AImg do
  begin
    case HorizontalAlign of
      0:
        x := Trunc(Width/2) - Trunc(Canvas.TextWidth(AText)/2);
      -1:
        x := 5;
      1:
        x := Width-5;
      else
        Exit;
    end;
    case VerticalAlign of
      0:
        y := Trunc(Height/2) - Trunc(Canvas.TextHeight(AText)/2);
      -1:
        y := 5;
      1:
        y := Height-5;
      else
        Exit;
    end;
    Canvas.TextOut(x, y, AText);
  end;
end;


function TUIUtil.GetFieldLengths(AQ: TADOQuery; bAjustTitle: Boolean): TAryInt;
var
  i, nLen: Integer;
begin
  if not Assigned(AQ) then
    Exit;
  with AQ do
  begin 
    if not Active then            // 不是打开状态
      Exit;
    DisableControls;
    SetLength(Result,FieldCount );

    if bAjustTitle then
    begin
      // 初时宽度为字段名称宽度
      for i := 0 to FieldCount - 1 do
      begin
        nLen := Length(Fields[i].FieldName);
  //      if Result[i] < nLen then
          Result[i] := nLen;
      end;
    end;

    // 根据字段内容调整
    First;
    while not Eof do
    begin
      for i := 0 to FieldCount - 1 do
      begin
        nLen := Length(Fields[i].AsString);
        if Result[i] < nLen then
          Result[i] := nLen;
      end;
      Next;       
    end;
    
    // 再指向query对象第一条数据
    First;
    EnableControls;
  end;
end;

function TUIUtil.GetFieldWidths(AQ: TADOQuery;
  ACanvas: TCanvas): TAryInt;
var
  i, nLen: Integer;
begin
  if not Assigned(AQ) then
    Exit;
  with AQ do
  begin
    if not Active then            // 不是打开状态
      Exit;
    DisableControls;
    SetLength(Result,FieldCount );

    // 初始宽度为字段名称宽度
    for i := 0 to FieldCount - 1 do
    begin
      nLen := ACanvas.TextWidth(Fields[i].DisplayName);
//      if Result[i] < nLen then
        Result[i] := nLen;
    end;

    // 根据字段内容调整
    First;
    while not Eof do
    begin
      for i := 0 to FieldCount - 1 do
      begin
        nLen := ACanvas.TextWidth(Fields[i].DisplayText);
        if Result[i] < nLen then
          Result[i] := nLen;
      end;
      Next;    
    end;

    // 再指向query对象第一条数据
    First;
    EnableControls;
  end;
end;

function TUIUtil.FindListItem(Alvw: TListView; sKeyWord: string;
  ColumnIndex: Integer): TListItem;
var
  i: Integer;
  bFind: Boolean;
begin
  Result := nil;
  for i := 0 to Alvw.Items.Count - 1 do
  begin
    if 0 = ColumnIndex then
      bFind := Alvw.Items[i].Caption = sKeyWord
    else
      bFind := Alvw.Items[i].SubItems[ColumnIndex-1] = sKeyWord;

    if bFind then
    begin
      Result := Alvw.Items[i];
      Break;
    end;  
  end;
end;

function TUIUtil.GetListBoxSelected(lbx: TListBox): Integer;
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to lbx.Count - 1 do
  begin
    if lbx.Selected[i] then
    begin
      Result := i;
      Break;
    end;
  end;  
end;

function TUIUtil.GetListBoxHittedIndex(lbx: TListBox; X, Y: Integer):Integer;
begin
  Result := lbx.ItemAtPos(Point(X, Y), True);
end;  

function TUIUtil.GetTreeViewHittedNode(tvw: TTreeView; X, Y: Integer; bWithIndent: Boolean): TTreeNode;
var
  tndHit: TTreeNode;
  theHT: THitTests;
  bHitted: Boolean;
begin
  Result := nil;
  tndHit := tvw.GetNodeAt(X, Y);
  if Assigned(tndHit) then
  begin
    theHT := tvw.GetHitTestInfoAt(X, Y);
    bHitted := (htOnStateIcon in theHT) or (htOnLabel in theHT) or (htOnIcon in theHT)
               or (htOnItem in theHT);
    if bWithIndent then   // 节点前的+号，htOnButton表示有子节点时， htOnIndent表示无字节点但HasChildren为True
      bHitted := bHitted or (htOnButton in theHT) or  (htOnIndent in theHT);
    if bHitted then
      Result := tndHit;
  end;
end;

function TUIUtil.GetTreeViewHittedNode(tvw: TTreeView; X, Y: Integer;
  aHT: THitTests): TTreeNode;
var
  tndHit: TTreeNode;
  theHT: THitTests;
begin
  Result := nil;
  tndHit := tvw.GetNodeAt(X, Y);
  if Assigned(tndHit) then
  begin
    theHT := tvw.GetHitTestInfoAt(X, Y);
    if aHT * theHT <> [] then
      Result := tndHit;
  end;
end;

procedure TUIUtil.ListViewAddColumns(Alvw: TListView;
  asCaption: array of string; aiWidth: array of Integer);
var
  i: Integer;
begin
  Alvw.Columns.Clear;
  for i := Low(asCaption) to High(asCaption) do
  begin
    with Alvw.Columns.Add do
    begin
      Caption := asCaption[i];
      if i <= High(aiWidth) then
        Width := aiWidth[i]
      else
        Width := Alvw.Canvas.TextWidth(Caption);
    end;
  end;  
end;

function TUIUtil.GetListViewHittedItem(lvw: TListView; X, Y: Integer;
  aHT: THitTests): TListItem;
var
  item : TListItem;
  theHT: THitTests;
begin
  Result := nil;
  item := lvw.GetItemAt(X, Y);
  if Assigned(item) then
  begin
    TheHT := lvw.GetHitTestInfoAt(X, Y);
    if (aHT * theHT) <> [] then
    begin
      Result := item;
    end;
  end;
end;

function TUIUtil.BuildStringByChar(c: Char; nLength: Integer): string;
var
  s: string;
begin
  s := c;
  while Length(s) * 2 <= nLength do
    s := s + s;
  s := s + Copy(s, 1, nLength - Length(s));
  Result := s;
end;

procedure TUIUtil.SortQuery(AQ: TADOQuery; nFieldIndex: Integer);
var
  nPos: Integer;
  field: TField;
  sSort, sFieldName: string;
begin
  if AQ = nil then Exit;
  field := AQ.Fields[nFieldIndex];
  // 决定升序还是降序
  if Pos(C_sSortFlag_Asc, field.DisplayName) > 0 then
    sSort := ' desc'
  else if Pos(C_sSortFlag_Desc, field.DisplayName) > 0 then
    sSort := ' asc'
  else
    sSort := ' asc';
  nPos := Pos('order', LowerCase(AQ.SQL.Text));
  sFieldName := field.FieldName;

  // 重新生成sql语句
  AQ.Close;
  if nPos = 0 then
    AQ.SQL.Text := AQ.SQL.Text + 'order by ' + sFieldName + sSort
  else
  begin
    AQ.SQL.Text := Copy(AQ.SQL.Text, 1, nPos-1); 
    AQ.SQL.Text := AQ.SQL.Text + 'order by ' + sFieldName + sSort;
  end;
  AQ.Open;
end;

procedure TUIUtil.SortDBGrid(dbgrd: TDBGrid; nColumnIndex: Integer; bFitColWidth: Boolean);
var
  qry: TADOQuery;
  sSort: string;
  col: TColumn;
  i: Integer;
  anColWidths: TAryInt;
begin
  if dbgrd.DataSource.DataSet = nil then Exit;
  dbgrd.Columns.BeginUpdate;
  col := dbgrd.Columns[nColumnIndex];
  // 决定升序还是降序
  if Pos(C_sSortFlag_Asc, col.Title.Caption) > 0 then
    sSort := C_sSortFlag_Desc
  else if Pos(C_sSortFlag_Desc, col.Title.Caption) > 0 then
    sSort := C_sSortFlag_Asc
  else
    sSort := C_sSortFlag_Asc;
  qry := TADOQuery(dbgrd.DataSource.DataSet);
  dbgrd.DataSource.DataSet := nil;
  SortQuery(qry, nColumnIndex);
  if bFitColWidth then           // 是否要修整列宽度
  begin
    SetLength(anColWidths, 0);
    anColWidths := GetFieldWidths(qry, dbgrd.Canvas);
    anColWidths[nColumnIndex] := anColWidths[nColumnIndex] + dbgrd.Canvas.TextWidth(sSort);  // 要排序的列宽度增加标志的宽度
    dbgrd.DataSource.DataSet := qry;
    for i := 0 to dbgrd.Columns.Count -1 do
      dbgrd.Columns[i].Width := anColWidths[i] + 5;

    col := dbgrd.Columns[nColumnIndex];      //刚才的col已因qry操作而释放，需要再次获得
    col.Title.Caption := col.Title.Caption + sSort;
  end
  else
    dbgrd.DataSource.DataSet := qry;
  dbgrd.Columns.EndUpdate;
end;

procedure TUIUtil.PrintDBGrid(DbGrid: TDbGrid; Title: string);
var
  PixelPerCM_X, PixelPerCM_Y: integer;
  ScreenX: integer;
  i, lx, ly: integer;
  px1, py1, px2, py2: integer;
  RowPerPage, RowPrinted: integer;
  ScaleX: Real;
  LineHeight: integer;
  TitleWidth: integer;
  SumWidth: integer;
  PageCount: integer;
  SpaceX, SpaceY: integer;
  RowCount: integer;
  canvas: TCanvas;
  DataSet: TDataSet;
begin
  if not Assigned(DbGrid) or not Assigned(DbGrid.DataSource) or
     not Assigned(DbGrid.DataSource.DataSet) then
    Exit;
  PixelPerCM_X := Round(GetDeviceCaps(Printer.Handle, LOGPIXELSX) / 2.54);      // X方向象素每厘米 1英寸＝2.54厘米
  PixelPerCM_Y := Round(GetDeviceCaps(Printer.Handle, LOGPIXELSY) / 2.54);      // y方向象素每厘米
  ScreenX := Round(Screen.PixelsPerInch / 2.54);
  ScaleX := PixelPerCM_X / ScreenX;
  RowPrinted := 0;
  SumWidth := 0;
  canvas := Printer.Canvas;
  LineHeight := Round(canvas.TextHeight('我') * 1.5); // 设定每行高度为字符高的1.5倍
  SpaceY := Round(canvas.TextHeight('我') / 4);    // x方向间距
  SpaceX := Round(canvas.TextWidth('我') / 4);
  RowPerpage := Round((Printer.PageHeight - 5 * PixelPerCM_Y) / LineHeight);   // 每页行数

  DataSet := DbGrid.DataSource.DataSet;
  DataSet.DisableControls;
  Printer.BeginDoc;
  try
    //上下边缘各2厘米
    ly := 2 * PixelPerCM_Y;
    PageCount := 0;
    DataSet.First;
    while not DataSet.Eof do
    begin
      if (RowPrinted = RowPerPage) or (RowPrinted = 0) then
      begin
        if RowPrinted <> 0 then
          Printer.NewPage;
        RowPrinted := 0;
        PageCount := PageCount + 1;
        // 标题
        canvas.Font.Name := '宋体';
        canvas.Font.Size := 16;
        canvas.Font.Style := canvas.Font.Style + [fsBold];
        lx := Round((Printer.PageWidth - canvas.TextWidth(Title)) / 2);
        ly := 2 * PixelPerCM_Y;
        canvas.TextOut(lx, ly, Title);
        // 内容
        canvas.Font.Size := 9;
        canvas.Font.Style := canvas.Font.Style - [fsBold];
        lx := Printer.PageWidth - 5 * PixelPerCM_X;                  // x方向缩进5厘米
        ly := Round(2 * PixelPerCM_Y + 0.2 * PixelPerCM_Y);
        if RowPerPage * PageCount > DataSet.RecordCount then
          RowCount := DataSet.RecordCount                            // 最后一页的行数特殊处理
        else
          RowCount := RowPerPage * PageCount;
        canvas.TextOut(lx, ly, '第' + IntToStr(RowPerPage * (PageCount - 1) + 1) +
          '-' + IntToStr(RowCount) + '条，共' + IntToStr(DataSet.RecordCount) +
          '条');
        lx := 2 * PixelPerCM_X;
        ly := ly + LineHeight * 2;
        py1 := ly - SpaceY;
        if RowCount = DataSet.RecordCount then
          py2 := py1 + LineHeight * (RowCount - RowPerPage * (PageCount - 1) + 1)
        else
          py2 := py1 + LineHeight * (RowPerPage + 1);
        SumWidth := lx;
        for i := 0 to DBGrid.Columns.Count - 1 do
        begin
          px1 := SumWidth - SpaceX;
          px2 := SumWidth;
          canvas.MoveTo(px1, py1);
          canvas.LineTo(px2, py2);
          TitleWidth := canvas.TextWidth(DBGrid.Columns[i].Title.Caption);
          lx := Round(SumWidth + (DBGrid.Columns[i].width * scaleX - titleWidth
            ) / 2);
          canvas.TextOut(lx, ly, DBGrid.Columns[i].Title.Caption);
          SumWidth := Round(SumWidth + DBGrid.Columns[i].width * scaleX) +
            SpaceX * 2;
        end;
        px1 := SumWidth; //画最后一条竖线
        px2 := SumWidth;
        canvas.MoveTo(px1, py1);
        canvas.LineTo(px2, py2);
        px1 := 2 * PixelPerCM_X; //画第一条横线
        px2 := SumWidth;
        py1 := ly - SpaceY;
        py2 := ly - SpaceY;
        canvas.MoveTo(px1, py1);
        canvas.LineTo(px2, py2);
        py1 := py1 + LineHeight;
        py2 := py2 + LineHeight;
        canvas.MoveTo(px1, py1);
        canvas.LineTo(px2, py2);
      end;
      lx := 2 * PixelPerCM_X;
      ly := ly + LineHeight;
      px1 := lx;
      px2 := SumWidth;
      py1 := ly - SpaceY + LineHeight;
      py2 := ly - SpaceY + LineHeight;
      canvas.MoveTo(px1, py1);
      canvas.LineTo(px2, py2);
      for i := 0 to DBGrid.Columns.Count - 1 do
      begin
        // DataSet.FieldByname(DBGrid.Columns[i].Fieldname).AsString
        canvas.TextOut(lx, ly, DataSet.Fields[i].AsString);
        lx := Round(lx + DBGrid.Columns[i].width * ScaleX + SpaceX * 2);
      end;
      RowPrinted := RowPrinted + 1;

      DataSet.next;  
    end;
  finally
    DataSet.first;
    DataSet.EnableControls;
    Printer.EndDoc;
  end;
end;

initialization

finalization
  if Assigned(FUIUtil) then
    FreeAndNil(FUIUtil);

end.
