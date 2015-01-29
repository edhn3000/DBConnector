{**
 * <p>Title: fDBGrid  </p>
 * <p>Copyright: Copyright (c) 2006</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: fDBGrid </p>
 *}
unit fDBGrid;

interface

uses
  SysUtils, Classes, Controls, Grids, DBGrids, Types, Windows, Messages,ADODB,
  DB, Graphics, Forms, DBTables, Contnrs, SqlExpr;

type
  TAryInt = array of Integer;
  TFitColumnMode = (fcmTitle, fcmContent, fcmBoth);
  TSortFlag = (sfNone, sfAsc, sfDesc);
{ TSortedField 字段名和排序标志}
  TSortedField = class
  public
    Name: string;
    Flag: TSortFlag;
  end;

{ TfDBGrid }
  TfDBGrid = class(TCustomDBGrid)
  private
    FAutoFitColumnWidth: Boolean;
    FSortOnColumnClick: Boolean;
    FaiColWidths: TAryInt;
    FMaxFitColCount: Integer;
    FBrushColorOdd: TColor;
    FBrushColorEven: TColor;
    FFitColumnMode: TFitColumnMode;
    FolSortedFields: TObjectList;
    FRecordNo: Integer;   
    FLastSelectedField: TField;
    FOnSelectedField: TNotifyEvent;

    procedure WMMouseWheel(var Msg: TMessage); message WM_MOUSEWHEEL;
    procedure WMVScroll(var Message: TWMVScroll); message WM_VSCROLL;
    procedure WMKeyDown(var Msg: TWMKeyDown); message WM_KEYDOWN;
    procedure SetSortOnColumnClick(value: Boolean);
    procedure clearSortFields;
    function getSortStrBySortFlag(sf: TSortFlag): string;  
    function getSortSqlBySortFlag(sf: TSortFlag): string;
    function getColunmByFieldName(sFieldName: string): TColumn;
    function getSortFieldIndexByFieldName(sFieldName: string):Integer;
  protected
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState;
      X, Y: Integer); override;

    procedure AdjustFieldWidths(ADataSet: TDataSet; ACanvas: TCanvas);
    procedure SortQuery(AQ: TDataSet; nFieldIndex: Integer);
    procedure DrawCell(ACol, ARow: Longint; ARect: TRect; AState: TGridDrawState);override;
    
    procedure DrawColumnCell(const Rect: TRect; DataCol: Integer;
      Column: TColumn; State: TGridDrawState); override;
    procedure BrushOnDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    
    procedure TitleClick(Column: TColumn);override;
    procedure SortOnTitleClick(nColumnIndex: Integer);
  public
    constructor Create(AOwner:TComponent); override;
    destructor Destroy;override;

    procedure FitColumuWidth;
  published
    property AutoFitColumnWidth : Boolean read FAutoFitColumnWidth write FAutoFitColumnWidth;
    property FitColumnMode: TFitColumnMode read FFitColumnMode write FFitColumnMode;
    property MaxFitColCount: Integer read FMaxFitColCount write FMaxFitColCount;
    property BrushColorOdd: TColor read FBrushColorOdd write FBrushColorOdd;
    property BrushColorEven: TColor read FBrushColorEven write FBrushColorEven;
    property SortOnColumnClick: Boolean read FSortOnColumnClick write SetSortOnColumnClick;
//    property LastSelectedField: TField read FLastSelectedField write FLastSelectedField;
    property OnSelectedField: TNotifyEvent read FOnSelectedField write FOnSelectedField;

  // dbgrid
  public
    property Canvas;
    property SelectedRows;
  published
    property Align;
    property Anchors;
    property BiDiMode;
    property BorderStyle;
    property Color;
    property Columns stored False; //StoreColumns;
    property Constraints;
    property Ctl3D;
    property DataSource;
    property DefaultDrawing;
    property DragCursor;
    property DragKind;
    property DragMode;
    property Enabled;
    property FixedColor;
    property Font;
    property ImeMode;
    property ImeName;
    property Options;
    property ParentBiDiMode;
    property ParentColor;
    property ParentCtl3D;
    property ParentFont;
    property ParentShowHint;
    property PopupMenu;
    property ReadOnly;
    property ShowHint;
    property TabOrder;
    property TabStop;
    property TitleFont;
    property Visible;
    property OnCellClick;
    property OnColEnter;
    property OnColExit;
    property OnColumnMoved;
    property OnDrawDataCell;  { obsolete }
    property OnDrawColumnCell;
    property OnDblClick;
    property OnDragDrop;
    property OnDragOver;
    property OnEditButtonClick;
    property OnEndDock;
    property OnEndDrag;
    property OnEnter;
    property OnExit;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
//    property OnMouseActivate;
    property OnMouseDown;
//    property OnMouseEnter;
//    property OnMouseLeave;
    property OnMouseMove;
    property OnMouseUp;
    property OnMouseWheel;
    property OnMouseWheelDown;
    property OnMouseWheelUp;
    property OnStartDock;
    property OnStartDrag;
    property OnTitleClick;
  end;

procedure Register;

implementation

uses
  Dialogs;

const
  C_sSortFlag_Asc = '▲';
  C_sSortFlag_Desc = '▼';

  // 计算列宽度时使用逐行计算宽度并记录的方式，此常量指定最多检查多少行
  C_nMaxFitColCount = 20;

procedure Register;
begin
  RegisterComponents('DBTool', [TfDBGrid]);
end;

{ TfDBGrid }  

procedure TfDBGrid.SetSortOnColumnClick(value: Boolean);
begin
  FSortOnColumnClick := value;
end;   

procedure TfDBGrid.clearSortFields;
//var
//  i: Integer;
begin
  // 为什么会报错
//  for i := FolSortedFields.Count - 1 downto 0 do
//  begin
//    (FolSortedFields[i] as TSortedField).Free;
//    FolSortedFields[i] := nil;
//  end;
  FolSortedFields.Clear;
end;

function TfDBGrid.getSortStrBySortFlag(sf: TSortFlag): string;
begin
  case sf of
  sfAsc: Result := C_sSortFlag_Asc;
  sfDesc: Result := C_sSortFlag_Desc;
  else
    Result := '';
  end;
end;

function TfDBGrid.getSortSqlBySortFlag(sf: TSortFlag): string;
begin
  case sf of
  sfAsc: Result := ' asc';
  sfDesc: Result := ' desc';
  else
    Result := '';
  end;
end;

function TfDBGrid.getColunmByFieldName(sFieldName: string): TColumn;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Columns.Count - 1 do
  begin
    if Columns[i].Field.FieldName = sFieldName then
    begin
      Result := Columns[i];
      Break;
    end;  
  end;  
end;

function TfDBGrid.getSortFieldIndexByFieldName(
  sFieldName: string): Integer; 
var
  i: Integer;
begin
  Result := -1;
  for i := 0 to FolSortedFields.Count - 1 do
  begin
    if (FolSortedFields[i] as TSortedField).Name = sFieldName then
    begin
      Result := i;
      Break;
    end;  
  end;
end;        

procedure TfDBGrid.DrawCell(ACol, ARow: Integer; ARect: TRect;
  AState: TGridDrawState);
begin
  // 这里只能得到已显示的数据的行号，不是实际数据的行号
  FRecordNo := ARow;
  inherited;
//  if (ARow>=1) and (ACol=0) then
//     Canvas.TextRect(ARect,ARect.Left,ARect.Top,IntToStr(ARow));
end;

procedure TfDBGrid.BrushOnDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
var
  grid: TfDBGrid;
begin
  if (gdSelected in State) and Self.Focused then
    Exit;
  grid := (Sender as TfDBGrid);
  
//  if (grid.DataSource <> nil) and
//     (grid.DataSource.DataSet <> nil) then
//  begin
//    if grid.DataSource.DataSet.RecNo mod 2 = 0 then  // 偶数行
//    begin   
//      if Column.Field.AsString = '' then
//        grid.Canvas.Brush.Color := RGB(212,212,255)
//      else
//        grid.Canvas.Brush.Color := FBrushColorEven;
//    end
//    else
//    begin                                           // 奇数行
//      if Column.Field.AsString = '' then
//        grid.Canvas.Brush.Color := RGB(198,223,221)
//      else
//        grid.Canvas.Brush.Color := FBrushColorOdd;
//    end;  
//  end;

  if FRecordNo mod 2 = 0 then
  begin
    if Column.Field.AsString = '' then
      grid.Canvas.Brush.Color := RGB(212,212,255)
    else
      grid.Canvas.Brush.Color := FBrushColorEven;
  end
  else
  begin                                            // 奇数行
    if Column.Field.AsString = '' then
      grid.Canvas.Brush.Color := RGB(198,223,221)
    else
      grid.Canvas.Brush.Color := FBrushColorOdd;
  end;

  grid.DefaultDrawColumnCell(Rect,DataCol,Column,State);
end;


procedure TfDBGrid.AdjustFieldWidths(ADataSet:TDataSet;
  ACanvas: TCanvas);
var
  i, nLen, nLineCount, nOldRecordIndex: Integer;
begin
  if not Assigned(ADataSet) then
    Exit;

  if not ADataSet.Active then            // 不是打开状态
    Exit;
  ADataSet.DisableControls;
  try
    nOldRecordIndex := ADataSet.RecNo;   
    SetLength(FaiColWidths,ADataSet.FieldCount);

    // 根据标题调整列宽时
    if (FFitColumnMode = fcmTitle) or (FFitColumnMode = fcmBoth) then
    begin
      // 初始宽度为字段名称宽度
      for i := 0 to FieldCount - 1 do
      begin
          nLen := ACanvas.TextWidth(ADataSet.Fields[i].FieldName);
  //      if FaiColWidths[i] < nLen then
          FaiColWidths[i] := nLen;
      end;
    end;

    // 根据内容调整列宽时                  
    if (FFitColumnMode = fcmContent) or (FFitColumnMode = fcmBoth) then
    begin
      // 根据字段内容调整
      ADataSet.First;
      nLineCount := 1;
      while not ADataSet.Eof do
      begin
        for i := 0 to ADataSet.FieldCount - 1 do
        begin
          if (ADataSet.Fields[i].DataType in [
            ftBlob, ftOraBlob, ftOraClob, {ftWideMemo,} ftMemo]) then
            nLen := ACanvas.TextWidth('（WideMemo）') + 5
          else
            nLen := ACanvas.TextWidth(ADataSet.Fields[i].AsString);
          if FaiColWidths[i] < nLen then
            FaiColWidths[i] := nLen;
        end;
        ADataSet.Next;
        Inc(nLineCount);
        if (C_nMaxFitColCount <> 0) and (nLineCount > C_nMaxFitColCount) then
          Break;
      end;
    end;

    // 再指向query对象第一条数据
    ADataSet.First;
    ADataSet.MoveBy(nOldRecordIndex-1);
  finally
    ADataSet.EnableControls;
  end;
end;

procedure TfDBGrid.SortQuery(AQ: TDataSet; nFieldIndex: Integer);
const
  C_nIsADO = 1;
  C_nIsBDE = 2;
  C_nIsDBX = 3;
var
  nPos, i: Integer;
//  field: TField;
  sSql, sOrder: string;
  adoQry: TADOQuery;
  bdeQry: TQuery;
  dbxQry: TSQLDataSet;
  nQryClass: Integer;
begin
  if AQ = nil then Exit;
  adoQry := nil;
  bdeQry := nil;
  dbxQry := nil;
  if AQ is TADOQuery then
  begin
    nQryClass := C_nIsADO;
    if adoQry = nil then
      adoQry := TADOQuery(AQ);
    sSql := adoQry.SQL.Text
  end
  else if AQ is TQuery then
  begin
    nQryClass := C_nIsBDE;
    if bdeQry = nil then
      bdeQry := TQuery(AQ);
    sSql := bdeQry.SQL.Text;
  end
  else if AQ is TSQLDataSet then
  begin
    nQryClass := C_nIsDBX;
    if dbxQry = nil then
      dbxQry := TSQLDataSet(AQ);
    sSql := dbxQry.CommandText;
  end
  else
    Exit;

  sSql := StringReplace(sSql, #$D#$A, '', [rfReplaceAll, rfIgnoreCase]); 
  nPos := Pos('order', LowerCase(sSql));
  // 重新生成sql语句
  AQ.Close;

  // 去掉原来的order，根据排序列表中的字段再次生成一个order
  if nPos <> 0 then
    sSql := Copy(sSql, 1, nPos-1);
  sOrder := '';
  for i := 0 to FolSortedFields.Count - 1 do
  begin
    if sOrder = '' then
      sOrder := (FolSortedFields[i] as TSortedField).Name +
          getSortSqlBySortFlag((FolSortedFields[i] as TSortedField).Flag)
    else
      sOrder := sOrder + ',' + (FolSortedFields[i] as TSortedField).Name +
          getSortSqlBySortFlag((FolSortedFields[i] as TSortedField).Flag);
  end;
  if Trim(sOrder) <> '' then
    sSql := sSql + ' order by ' + sOrder;

  case nQryClass of
  C_nIsADO:
  begin
    adoQry.SQL.Text := sSql;
    adoQry.Open
  end;
  C_nIsBDE:
  begin
    bdeQry.SQL.Text := sSql;
    bdeQry.Open
  end;
  C_nIsDBX:
  begin
    dbxQry.CommandText := sSql;
    dbxQry.Open;
  end;  
  end;
end;

constructor TfDBGrid.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FRecordNo := 1;
  FAutoFitColumnWidth := True;
  FMaxFitColCount := C_nMaxFitColCount;
  FBrushColorOdd := clWindow;
//  FBrushColorEven := RGB(150,200,255); 
  FBrushColorEven := RGB(184,234,186);
  FFitColumnMode := fcmBoth;
  FolSortedFields := TObjectList.Create;
end;  

destructor TfDBGrid.Destroy;
begin
  FolSortedFields.Free;
  inherited;
end;

procedure TfDBGrid.DrawColumnCell(const Rect: TRect; DataCol: Integer;
  Column: TColumn; State: TGridDrawState);
begin
  inherited;
  BrushOnDrawColumnCell(Self, Rect, DataCol, Column, State);
end;

procedure TfDBGrid.SortOnTitleClick(nColumnIndex: Integer);
var
  qry: TDataSet;
  sSort: string;
  col: TColumn;
  i: Integer;
  sf: TSortFlag;
  nSortFieldIndex: Integer;
  function createSortFieldObj(ACol: TColumn; Asf: TSortFlag):TSortedField;
  var
    sfd: TSortedField;
  begin
    sfd := TSortedField.Create;
    sfd.Name := ACol.Field.FieldName;
    sfd.Flag := Asf;
    Result := sfd;
  end;
begin
  if DataSource.DataSet = nil then Exit;
  col := Columns[nColumnIndex];
  // 决定升序还是降序  无->降序->升序->无
  if Pos(C_sSortFlag_Asc, col.Title.Caption) > 0 then
    sf := sfNone
  else if Pos(C_sSortFlag_Desc, col.Title.Caption) > 0 then
    sf := sfAsc
  else 
    sf := sfDesc;
  // 按下shift可以子排序， 否则重新排序
  if not GetAsyncKeyState(VK_SHIFT) < 0 then
  begin
    clearSortFields;
    FolSortedFields.Add(createSortFieldObj(col, sf));
  end
  else
  begin
    nSortFieldIndex := getSortFieldIndexByFieldName(col.Field.FieldName);
    if nSortFieldIndex = -1 then
      FolSortedFields.Add(createSortFieldObj(col, sf))
    else
    begin
      if sf = sfNone then
        FolSortedFields.Delete(nSortFieldIndex)
      else
        (FolSortedFields[nSortFieldIndex] as TSortedField).Flag := sf;
    end;
  end;

  qry := DataSource.DataSet;
  DataSource.DataSet := nil;
  SortQuery(qry, nColumnIndex);
  if FAutoFitColumnWidth then           // 是否要修整列宽度
  begin
    AdjustFieldWidths(qry, Canvas);
    DataSource.DataSet := qry;

    Columns.BeginUpdate;
//    FaiColWidths[nColumnIndex] := FaiColWidths[nColumnIndex] +
//        Canvas.TextWidth(sSort);
    // 要排序的列宽度增加标志的宽度
    for i := 0 to FolSortedFields.Count - 1 do
    begin
      col := getColunmByFieldName((FolSortedFields[i] as TSortedField).Name);
      if col = nil then Continue;
      sSort := getSortStrBySortFlag((FolSortedFields[i] as TSortedField).Flag);
      if i = 0 then
      begin
        col.Title.Caption := col.Title.Caption + sSort;
        FaiColWidths[col.Index] := FaiColWidths[col.Index] + Canvas.TextWidth(sSort)
      end
      else
      begin
        if sSort <> '' then
          sSort := sSort + IntToStr(i+1);
        col.Title.Caption := col.Title.Caption + sSort;
        FaiColWidths[col.Index] := FaiColWidths[col.Index] + Canvas.TextWidth(sSort);
      end;
    end;  
    for i := 0 to Columns.Count -1 do
      Columns[i].Width := FaiColWidths[i] + 5;
//    col := Columns[nColumnIndex];  // 刚才的col已因qry操作而释放，需要再次获得
//    col.Title.Caption := col.Title.Caption + sSort;
    Columns.EndUpdate;
  end
  else
    DataSource.DataSet := qry;
end;

procedure TfDBGrid.TitleClick(Column: TColumn);
begin
  inherited;
  if FSortOnColumnClick then
    SortOnTitleClick(Column.Index);
end;

procedure TfDBGrid.WMKeyDown(var Msg: TWMKeyDown);
begin
  inherited;
  if Msg.CharCode in [37,38,39,40] then // left up right down
  begin
    if (Msg.CharCode in [37,38, 39,40])
        and ((SelectedField <> nil)) then
    begin
      FLastSelectedField := SelectedField;
      if Assigned(OnSelectedField) then
        OnSelectedField(Self);
    end;
    Paint;
  end;
end;    

procedure TfDBGrid.MouseDown(Button: TMouseButton; Shift: TShiftState; X,
  Y: Integer);
begin
  inherited;
  if (SelectedField <> nil) then
  begin
    FLastSelectedField := SelectedField;
    if Assigned(OnSelectedField) then
      OnSelectedField(Self);
  end;
end;

procedure TfDBGrid.WMMouseWheel(var Msg: TMessage);
var
  m: TMessage;
  pt: TPoint;
  nLeft, nTop: Integer;
  control: TWinControl;
begin
  // 检查鼠标是否在控件上，如果不在，即时控件有焦点也不响应鼠标滚轮事件
  pt := Mouse.CursorPos;

  nLeft := 0;
  nTop := 0;
  control := Self; 
  nLeft := nLeft + control.Left;
  nTop := nTop + control.Top;
  while (control.Parent <> nil) do
  begin       
    control := control.Parent;
    nLeft := nLeft + control.Left;
    nTop := nTop + control.Top;
  end;
  pt.X := pt.X- nLeft;
  pt.Y := pt.Y- nTop;
  
  if not PtInRect(Self.ClientRect,pt) then
    Exit;

  if Msg.WParam > 0 then // 向上滚动
    m.WParamLo:= SB_LINEUP
  else                   // 向下滚动
    m.WParamLo:= SB_LINEDOWN;
  m.Msg:=WM_VSCROLL;
  m.WParamHi:=1;
  m.LParam:=0;
  postmessage(Handle,m.Msg,m.WParam,m.LParam);
end;   

procedure TfDBGrid.WMVScroll(var Message: TWMVScroll);
begin
  inherited;
  Paint;
end;

procedure TfDBGrid.FitColumuWidth;
var
  i: Integer;
  qry: TDataSet;
begin
  qry := DataSource.DataSet;
  if (qry = nil) then Exit;
//  DataSource.DataSet := nil;
  try
    AdjustFieldWidths(qry, Canvas);
  //  DataSource.DataSet := qry;
    // 无数据时DBGrid显示一个无名字段，此时不需要调整宽度
    if (Columns.Count = 1) and (Columns[0].FieldName = '') then
      Exit;
    Columns.BeginUpdate;
    for i := 0 to Columns.Count -1 do
    begin
      if FaiColWidths[i] + 5 < 400 then
        Columns[i].Width := FaiColWidths[i] + 5
      else
        Columns[i].Width := 400;
    end;
    Columns.EndUpdate;
  except
    on ex: Exception do
      raise Exception.Create('(TfDBGrid.FitColumuWidth)无法调整Grid列宽 ' + ex.Message);
  end;
end;

end.
