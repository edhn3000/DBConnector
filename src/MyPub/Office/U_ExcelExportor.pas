{**
 * <p>Title: U_ExcelExportor </p>
 * <p>Copyright: Copyright (c) 2008-8-15 </p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 用户界面处理 </p>
 *}
unit U_ExcelExportor;

interface    

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, OleServer, ComObj, U_HashTable;

const
  C_CELL_Link = '=HYPERLINK("%s","%s")';

type           
  TExcelSheet = class;
  TExcelCell = class
  protected
    FCell: Variant;
    FRow: Integer;
    FCol: Integer;
    FOwnerSheet: TExcelSheet;

    function GetValue(): string;
    procedure SetValue(value: string);
  public      
    property Row: Integer read FRow write FRow;
    property Col: Integer read FCol write FCol;
    property Value: string read getValue write setValue;

    class function GetCellName(indexX, indexY: Integer): string; 

    function getName(): string;                                                                   
    constructor Create(OwnerSheet: TExcelSheet; rangeStr: string);overload;
    constructor Create(OwnerSheet: TExcelSheet; row, col: Integer);overload;
    destructor Destroy;override;
  end;

  TExcelRange = class
  private
    FStartCell: TExcelCell;
    FEndCell: TExcelCell;
    FOwnerSheet: TExcelSheet;
  public          
    property StartCell: TExcelCell read FStartCell;
    property EndCell: TExcelCell read FEndCell;

    constructor Create(OwnerSheet: TExcelSheet; rangeStr: string);
    destructor Destroy;override;

  end;
  
  TExcelSheet = class
  private 
    FSheet: Variant;
    FCellMap: THashTable;
    function GetManagedCell(row, col: Integer): TExcelCell;
    function GetName(): string;
    procedure SetName(value: string);                  
    function GetCell(indexX, indexY: Integer): TExcelCell;
    function GetCellValue(indexX, indexY: Integer): string;
    procedure SetCellValue(indexX, indexY: Integer; value: string);
    function getMergedCellLeftTop(row, col: Integer): TExcelCell;
  public
    property Name: string read GetName write SetName; 
    property Cell[indexX, indexY: Integer]: TExcelCell read GetCell;
    property CellValue[indexX, indexY: Integer]: string read GetCellValue write SetCellValue;
                           
    function GetColumnWidth(nColumnIndex: Integer): Integer;
    procedure SetColumnWidth(nColumnIndex, value: Integer);  
    
    procedure SetRowHeight(nRowIndex: Integer; value: Integer);
    function GetRowHeight(nRowIndex: Integer): Integer;
    procedure MergeRange(row1, col1, row2, col2: Integer);
    function isMergedCell(row, col: Integer): boolean;

    
    constructor Create;
    destructor Destroy;override;
  end;

  TExcelBook = class
  private
    FBook: Variant;
    FActiveSheet: TExcelSheet;
    function GetSheetCount: Integer;
    function GetSheet(index: Integer): TExcelSheet;
  public
    property SheetCount: Integer read GetSheetCount;
    property Sheets[index: Integer]: TExcelSheet read GetSheet;

    constructor Create;
    destructor Destroy;override;
  end;

  TExcelExportor = class
  private
    FExcelApp: Variant;
    FActiveBook: TExcelBook;
    function GetVisible(): Boolean;
    procedure SetVisible(value: Boolean);
    function GetActiveSheet(): TExcelSheet;
    function GetActiveWorkBook(): TExcelBook;
  public
    OnFreeCloseExcel: Boolean;
    property Visible: Boolean read GetVisible write SetVisible;
    property ExcelApp: Variant read FExcelApp write FExcelApp;
    property ActiveSheet: TExcelSheet read GetActiveSheet;      
    property ActiveWorkBook: TExcelBook read GetActiveWorkBook;
  public
    constructor Create;
    destructor Destroy;override;
    procedure Release;

    procedure NewFile(sFile: string);
    procedure LoadFile(sFile: string);
    procedure SaveFile(sFile: string);  
    procedure Close;

    procedure ShowExcel;
    procedure AddWorkBook;    
    procedure AddWorkSheet;        
    procedure ChangeActiveSheet(nSheetIndex: Integer);
  end;


implementation

{ T_ExcelExportor }

constructor TExcelExportor.Create;
begin
  FExcelApp := CreateOleObject( 'Excel.Application' );
  FActiveBook := TExcelBook.Create;
  OnFreeCloseExcel := True;
end;

destructor TExcelExportor.Destroy;
begin
  if OnFreeCloseExcel then
    Close;
  FActiveBook.Free;
  Release;
  inherited;
end;

procedure TExcelExportor.Release;
begin    
  FExcelApp := Unassigned;
end;

procedure TExcelExportor.Close;
begin
//  if FExcelApp <> Unassigned then
//  begin
    FExcelApp.WorkBooks.Close;
    FExcelApp.Quit;
//  end;
end;

procedure TExcelExportor.LoadFile(sFile: string);
begin
  FExcelApp.WorkBooks.Open(sFile);
end;

procedure TExcelExportor.ShowExcel;
begin
  FExcelApp.Visible := True;
end;

procedure TExcelExportor.ChangeActiveSheet(nSheetIndex: Integer);
begin
  FExcelApp.WorkSheets[nSheetIndex].Activate;
  FActiveBook.FActiveSheet.FSheet := FExcelApp.WorkSheets[nSheetIndex];
end;

procedure TExcelSheet.SetRowHeight(nRowIndex, value: Integer);
begin
  FSheet.Rows[nRowIndex].RowHeight := value;
end;

function TExcelSheet.GetRowHeight(nRowIndex: Integer): Integer;
begin
  Result := FSheet.Rows[nRowIndex].RowHeight;
end;

procedure TExcelExportor.SaveFile(sFile: string);
begin
  FExcelApp.ActiveWorkBook.SaveAs(sFile);
end;

procedure TExcelExportor.AddWorkBook;
begin
  FActiveBook.FBook := FExcelApp.WorkBooks.Add;
end;

procedure TExcelExportor.NewFile(sFile: string);
begin 
  FExcelApp.WorkBooks.Open(sFile);
end;

function TExcelExportor.GetVisible: Boolean;
begin
  Result := FExcelApp.Visible;
end;

procedure TExcelExportor.SetVisible(value: Boolean);
begin
  FExcelApp.Visible := value;
end;

procedure TExcelSheet.MergeRange(row1, col1, row2, col2: Integer);
var
  range: Variant;
  sRangeStr: string;
begin
  if (row1=row2) and (col1=col2) then
    Exit;
    
  sRangeStr := TExcelCell.GetCellName(row1,col1)
    + ':' + TExcelCell.GetCellName(row2,col2);
  range := FSheet.Range[sRangeStr];
  range.Select;
  range.Mergecells := True;
end;     

function TExcelSheet.isMergedCell(row, col: Integer): boolean;
var
  range: Variant;
  sRangeStr: string;
  addr1, addr2: string;
begin
  sRangeStr := TExcelCell.GetCellName(row,col);
  range := FSheet.Range[sRangeStr];
  addr1 := range.Address;  // like $A$1
  addr2 := range.MergeArea.Address; // like $A$1:$A$5
  Result := (addr1 <> addr2);
end;

function TExcelSheet.getMergedCellLeftTop(row, col: Integer): TExcelCell;
var
  range: Variant;
  sRangeStr: string;   
  addr1, addr2: string;
  rangeObj: TExcelRange;
begin
  sRangeStr := TExcelCell.GetCellName(row,col);
  range := FSheet.Range[sRangeStr];
  addr1 := range.Address;  // like $A$1
  addr2 := range.MergeArea.Address; // like $A$1:$A$5
  rangeObj := TExcelRange.Create(Self, addr2);
  try
    Result := GetManagedCell(rangeObj.StartCell.Row,
       rangeObj.StartCell.Col);
  finally
    rangeObj.Free;
  end;
end;

function TExcelExportor.GetActiveSheet: TExcelSheet;
begin
  FActiveBook.FActiveSheet.FSheet := FExcelApp.ActiveSheet;
  Result := FActiveBook.FActiveSheet;
end;  

procedure TExcelExportor.AddWorkSheet;
begin
  if FExcelApp.Workbooks.Count = 0 then
  begin
    FActiveBook.FBook := FExcelApp.WorkBooks.Add;
    FActiveBook.FActiveSheet.FSheet := FExcelApp.Sheets[1];
  end else
  begin
    FActiveBook.FBook  := FExcelApp.ActiveWorkbook;
    FActiveBook.FActiveSheet.FSheet := FExcelApp.Sheets.Add;
  end;
end;    

function TExcelExportor.GetActiveWorkBook: TExcelBook;
begin
  Result := FActiveBook;
end;

{ TExcelSheet }

constructor TExcelSheet.Create;
begin
  FCellMap := THashTable.Create();
end;

destructor TExcelSheet.Destroy;
var
  cell: TExcelCell;
begin
  FCellMap.ResetEnum;
  cell := TExcelCell(FCellMap.EnumNext);
  while Assigned(cell) do
  begin
    try
      FreeAndNil(cell);
    except
      // 释放cell，如果被free过可能会出现异常，这里忽略掉
    end;
    cell := TExcelCell(FCellMap.EnumNext);
  end;
  FCellMap.Free;
  inherited;
end;

function TExcelSheet.GetCellValue(indexX, indexY: Integer): string;
var
  row, col:Integer;
  cell: TExcelCell;
begin
  if (not isMergedCell(indexX, indexY)) then
  begin
    row := indexX;
    col := indexY;
  end
  else
  begin
    cell := getMergedCellLeftTop(indexX, indexY);
    row := cell.Row;
    col := cell.Col;
  end;   
  Result := FSheet.Cells[row, col].Value;
end;

function TExcelSheet.GetName: string;
begin
  Result := FSheet.Name;
end;

procedure TExcelSheet.SetCellValue(indexX, indexY: Integer; value: string);
begin  
  FSheet.Cells[indexX,indexY].Value := value;
end;

procedure TExcelSheet.SetColumnWidth(nColumnIndex, value: Integer);
begin
  if FSheet .Columns.Count <= nColumnIndex then
    FSheet.Columns.Add;
  FSheet.Columns[nColumnIndex].ColumnWidth := value;
end;

function TExcelSheet.GetColumnWidth(nColumnIndex: Integer): Integer;
begin
  Result := FSheet.Columns[nColumnIndex].ColumnWidth;
end;

procedure TExcelSheet.SetName(value: string);
begin
  FSheet.Name := value;
end;

function TExcelSheet.GetCell(indexX, indexY: Integer): TExcelCell;
var
  cell: TExcelCell;
begin
  cell := GetManagedCell(indexX, indexY);
  Result := cell;
end;

function TExcelSheet.GetManagedCell(row, col: Integer): TExcelCell;
var
  cell: TExcelCell;
  key: string;
begin
  key := TExcelCell.GetCellName(row, col);
  if FCellMap.Get(key) <> nil then
  begin
    cell := TExcelCell(FCellMap.Get(key));
    cell.Row := row;
    cell.Col := col;
  end
  else
  begin
    cell := TExcelCell.Create(Self, row, col);
    FCellMap.Put(key, cell); // cell 受管理，不用自动释放
  end;
  Result := cell;
end;

{ TExcelBook }

function TExcelBook.GetSheet(index: Integer): TExcelSheet;
begin
  FActiveSheet.FSheet := FBook.WorkSheets[index];
  Result := FActiveSheet;
end;

function TExcelBook.GetSheetCount: Integer;
begin
  Result := FBook.WorkSheets.Count;
end;

constructor TExcelBook.Create;
begin
  FActiveSheet := TExcelSheet.Create;
end;

destructor TExcelBook.Destroy;
begin
  FActiveSheet.Free;
  inherited;
end;

{ TCell }

constructor TExcelCell.Create(OwnerSheet: TExcelSheet; row, col: Integer);
begin
  Self.FOwnerSheet := OwnerSheet;
  Self.FRow := row;
  Self.FCol := col;
  Self.FCell := OwnerSheet.FSheet.Cells[FRow,FCol];
end;

constructor TExcelCell.Create(OwnerSheet: TExcelSheet; rangeStr: string);
var
  nPos: Integer;
  s: string;
  charTemp: Char;
begin
  Self.FOwnerSheet := OwnerSheet;
  nPos := Pos('$', rangeStr);
  if (nPos <> 0) then     // first $
  begin
    s := Copy(rangeStr, nPos + 1, MaxInt);
    nPos := Pos('$', s);    // second $
    if (nPos <> 0) then
    begin
      charTemp := Copy(s, 1, nPos-1)[1];  // TODO perhaps AA
      FCol := Byte(Ord(charTemp)) - 64;
      FRow := StrToInt(Copy(s, nPos + 1, MaxInt));
    end;
  end; 
  Self.FCell := OwnerSheet.FSheet.Cells[FRow,FCol];
end;

destructor TExcelCell.Destroy;
begin

  inherited;
end; 

function TExcelCell.getName: string;
begin
  Result := TExcelCell.GetCellName(FRow, FCol);
end;    

function TExcelCell.GetValue: string;
begin
  Result := FCell.Value;
end;

procedure TExcelCell.SetValue(value: string);
begin
  FCell.Value := value;
end;

class function TExcelCell.GetCellName(indexX, indexY: Integer): string;
begin          
  Result := Chr(Byte(indexY)+64) + IntToStr(indexX);
end;

{ TRange }

constructor TExcelRange.Create(OwnerSheet: TExcelSheet; rangeStr: string);
var
  nPos: Integer;
  cellStartStr, cellEndStr: string;
begin
  FOwnerSheet := OwnerSheet;
  nPos := Pos(':', rangeStr);
  if (nPos <> 0) then
  begin
    cellStartStr := Copy(rangeStr, 1, nPos - 1);
    cellEndStr := Copy(rangeStr, nPos + 1, MaxInt);
    FStartCell := TExcelCell.Create(FOwnerSheet, cellStartStr);
    FEndCell := TExcelCell.Create(FOwnerSheet, cellEndStr);
  end;  
end;

destructor TExcelRange.Destroy;
begin
  if Assigned(FStartCell) then
    FreeAndNil(FStartCell);
  if Assigned(FEndCell) then
    FreeAndNil(FEndCell);
  inherited;
end;

end.
