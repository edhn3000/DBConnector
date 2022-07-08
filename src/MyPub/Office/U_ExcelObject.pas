{**
 * <p>Title: U_ExcelObject </p>
 * <p>Copyright: Copyright (c) 2026-2-18 </p>
 * @author fengyq
 * @version 1.0
 * <p>Description: Excel对象的封装 </p>
 *}
unit U_ExcelObject;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, OleServer, ComObj, U_RegexUtil;

const
  C_CELL_Link = '=HYPERLINK("%s","%s")';

type           
  TExcelSheet = class;
  { TExcelCell }
  TExcelCell = class(TObject)
  private
    FCell: OleVariant;
    FOwnerSheet: TExcelSheet;
    FName: string;
    FRow: Integer;
    FCol: Integer;

    function GetName(): string;
    procedure SetName(value: String);
    function GetValue(): string;
    procedure SetValue(value: string);
  public
    property OwnerSheet: TExcelSheet read FOwnerSheet;
    property Cell: OleVariant read FCell write FCell;
    property Row: Integer read FRow write FRow;
    property Col: Integer read FCol write FCol;
    property Name: String read GetName write SetName;
    property Value: string read GetValue write SetValue;

    constructor Create(OwnerSheet: TExcelSheet; rangeName: string);overload;
    constructor Create(OwnerSheet: TExcelSheet; row, col: Integer);overload;
    destructor Destroy;override;

    function Equals(Obj: TObject): Boolean; override;

  end;

  { TExcelRange }
  TExcelRange = class
  private
    FRange: OleVariant;
    FOwnerSheet: TExcelSheet;
    FName: String;
    FStartCell: TExcelCell;
    FEndCell: TExcelCell;

    procedure SetName(value: String);
    function GetValue(): string;
    procedure SetValue(value: string);
  public
    property OwnerSheet: TExcelSheet read FOwnerSheet;
    property Range: OleVariant read FRange write FRange;
    property Name: String read FName write SetName;
    property Value: string read GetValue write SetValue;
    property StartCell: TExcelCell read FStartCell;
    property EndCell: TExcelCell read FEndCell;

    constructor Create(OwnerSheet: TExcelSheet; rangeName: string);
    destructor Destroy;override;

    function GetRowCount: Integer;
    function GetColCount: Integer;

    // 该是否合并区域
    function IsMerged: Boolean;

    procedure ExtendMergeRange;

  end;

  { TExcelSheet }
  TExcelSheet = class
  private
    FSheet: OleVariant;
    FUsedRange: TExcelRange;

    function GetName(): string;
    procedure SetName(value: string);                  
    function GetCell(indexX, indexY: Integer): TExcelCell;
    function GetCellValue(indexX, indexY: Integer): string;
    procedure SetCellValue(indexX, indexY: Integer; value: string);
    procedure SetSheet(sheet: OleVariant);
    function GetUsedRange: TExcelRange;
  public
    property Sheet: OleVariant read FSheet write SetSheet;
    property UsedRange: TExcelRange read GetUsedRange;
    property Name: string read GetName write SetName;
    property Cell[indexX, indexY: Integer]: TExcelCell read GetCell;
    property CellValue[indexX, indexY: Integer]: string read GetCellValue write SetCellValue;

    constructor Create();
    destructor Destroy;override;

    // 列宽
    function GetColumnWidth(nColumnIndex: Integer): Integer;
    procedure SetColumnWidth(nColumnIndex, value: Integer);

    // 行高
    function GetRowHeight(nRowIndex: Integer): Integer;
    procedure SetRowHeight(nRowIndex: Integer; value: Integer);

    // 合并单元格
    procedure MergeRange(row1, col1, row2, col2: Integer); overload;
    procedure MergeRange(cellName1, cellName2: String); overload;
    // 检查单元格是否属于合并单元格
    function IsMerged(row, col: Integer): Boolean;
  end;

  { TExcelWorkBook }
  TExcelWorkBook = class
  private
    FWorkBook: OleVariant;
    FActiveSheet: TExcelSheet;
    function GetSheetCount: Integer;
    function GetSheet(index: Integer): TExcelSheet;
    function GetActiveSheet(): TExcelSheet;
    procedure SetWorkBook(value: OleVariant);
  public
    property WorkBook: OleVariant read FWorkBook write SetWorkBook;
    property ActiveSheet: TExcelSheet read GetActiveSheet;
    property SheetCount: Integer read GetSheetCount;
    property Sheets[index: Integer]: TExcelSheet read GetSheet;

    constructor Create();
    destructor Destroy;override;
  end;

  { TExcelApplication }
  TExcelApplication = class
  private
    FApplication: OleVariant;
    FUseActiveApp: Boolean;
    FActiveWorkBook: TExcelWorkBook;
    FOpend: Boolean;
    FCloseOnFree: Boolean;
    FIsInstalled: Boolean;

    function GetVisible(): Boolean;
    procedure SetVisible(value: Boolean);
    function GetActiveSheet(): TExcelSheet;
    function GetActiveWorkBook(): TExcelWorkBook;
  protected
    function CreateApplication(useActiveApp: Boolean;
      className: String): Boolean;
  public
    property CloseOnFree: Boolean read FCloseOnFree write FCloseOnFree;
    property Visible: Boolean read GetVisible write SetVisible;
    property Opend: Boolean read FOpend;
    property Application: OleVariant read FApplication write FApplication;
    property ActiveSheet: TExcelSheet read GetActiveSheet;
    property ActiveWorkBook: TExcelWorkBook read GetActiveWorkBook;
  public
    constructor Create;overload;
    constructor Create(useActiveApp: Boolean);overload;
    constructor Create(ActiveApp: OleVariant);overload;
    destructor Destroy;override;

    procedure NewFile(sFile: string);
    procedure OpenFile(sFile: string);
    procedure SaveFile();
    procedure SaveAs(sFile: string);
    procedure CloseFile;

    procedure AddWorkBook;
    procedure AddWorkSheet;
    procedure ChangeActiveSheet(nSheetIndex: Integer);
  end;

  // 传入行列的数字，返回字符串形式单元格名，如A1
  function GetCellName(indexX, indexY: Integer): string;
  // 列数字转化为名称 如1转化为A
  function ExcelColIndex2Name(colIndex: Integer): String;
  // 列名称转化为数字 如A转化为1
  function ExcelColName2Index(colName: String): Integer;

implementation

const
  REGEX_CELL = '(\$?([A-Za-z]+)\$?(\d+))';

function ExcelColIndex2Name(colIndex: Integer): String;
var
  colName: String;
  num, bitNum: Integer;
begin
  num := colIndex;
  while num > 0 do begin
    bitNum := num mod 26;
    num := Trunc(num / 26);
    if bitNum = 0 then begin
      bitNum := 26;
      num := num - 1;
    end;
    colName := Chr(Byte(bitNum)+64) + colName;
  end;
  Result := colName;
end;

function ExcelColName2Index(colName: String): Integer;
var
  charTemp: Char;
  i: Integer;
  num: Byte;
begin
  Result := 0;
  colName := UpperCase(colName);
  for i := 1 to Length(colName) do begin
    charTemp := colName[i];
    if (Ord(charTemp) < 65) or (Ord(charTemp) > 90) then
      raise Exception.Create('ColName has non A-Za-z char');

    num := Byte(Ord(charTemp)) - 64;
    // 最后一位直接加
    if i = Length(colName) then
      Result := Result + num
    else // 26进制的计算
      Result := Result + num * 26 * (Length(colName) - i);
  end;
end;

function GetCellName(indexX, indexY: Integer): string;
begin
  Result := ExcelColIndex2Name(indexY) + IntToStr(indexX);
end;

{ TExcelApplication }

function TExcelApplication.CreateApplication(useActiveApp: Boolean; className: String): Boolean;
begin
  Result := False;
  try
    if useActiveApp then begin
      try
        FApplication := GetActiveOleObject(className);
        FUseActiveApp := True;
      except
        on e: EOleSysError do begin
          // 没有打开的Application，创建新的
          FApplication := CreateOleObject(className);
        end;
      end
    end else begin
      FApplication := CreateOleObject(className);
    end;
    Result := True;
    FIsInstalled := Result;
  except
    on e: Exception do begin
      OutputDebugString(PChar(Format('CreateApplication using (%s) fail! %s' ,[className, e.Message])));
    end;
  end;
end;

constructor TExcelApplication.Create;
begin
  FOpend := False;
  Create(False);
  FCloseOnFree := True;
end;

constructor TExcelApplication.Create(useActiveApp: Boolean);
begin
  CreateApplication(useActiveApp, 'Excel.Application');
  FActiveWorkBook := TExcelWorkBook.Create();
  if FApplication.WorkBooks.Count > 0 then begin
    FActiveWorkBook.WorkBook := FApplication.ActiveWorkbook;
  end;

  FCloseOnFree := not useActiveApp;
end;

constructor TExcelApplication.Create(ActiveApp: OleVariant);
begin
  FApplication := ActiveApp;
  FActiveWorkBook := TExcelWorkBook.Create();
  if FApplication.WorkBooks.Count > 0 then begin
    FActiveWorkBook.WorkBook := FApplication.ActiveWorkbook;
  end;
  FCloseOnFree := False;
end;

destructor TExcelApplication.Destroy;
begin
  if FCloseOnFree then begin
    CloseFile;
  end;
  FActiveWorkBook.Free;
  FApplication := Unassigned;
  inherited;
end;

procedure TExcelApplication.NewFile(sFile: string);
begin
  FActiveWorkBook.WorkBook := FApplication.WorkBooks.Add(sFile);
  FOpend := True;
end;

procedure TExcelApplication.OpenFile(sFile: string);
begin
  FActiveWorkBook.WorkBook := FApplication.WorkBooks.Open(sFile);
  FOpend := True;
end;

procedure TExcelApplication.SaveFile;
begin
  FApplication.ActiveWorkBook.Save;
end;

procedure TExcelApplication.SaveAs(sFile: string);
begin
  FApplication.ActiveWorkBook.SaveAs(sFile);
end;

procedure TExcelApplication.CloseFile;
begin
//  if FExcelApp <> Unassigned then
//  begin
    FApplication.WorkBooks.Close;
    FOpend := False;
    FApplication.Quit;
//  end;
end;

procedure TExcelApplication.ChangeActiveSheet(nSheetIndex: Integer);
begin
  FApplication.WorkSheets[nSheetIndex].Activate;
  FActiveWorkBook.FActiveSheet.FSheet := FApplication.WorkSheets[nSheetIndex];
end;

procedure TExcelApplication.AddWorkBook;
begin
  FActiveWorkBook.FWorkBook := FApplication.WorkBooks.Add;
end;

function TExcelApplication.GetVisible: Boolean;
begin
  Result := FApplication.Visible;
end;

procedure TExcelApplication.SetVisible(value: Boolean);
begin
  FApplication.Visible := value;
end;

function TExcelApplication.GetActiveSheet: TExcelSheet;
begin
  Result := FActiveWorkBook.ActiveSheet;
end;

procedure TExcelApplication.AddWorkSheet;
begin
  if FApplication.Workbooks.Count = 0 then begin
    FActiveWorkBook.FWorkBook := FApplication.WorkBooks.Add;
    FActiveWorkBook.FActiveSheet.FSheet := FApplication.Sheets[1];
  end else begin
    FActiveWorkBook.FWorkBook := FApplication.ActiveWorkbook;
    FActiveWorkBook.FActiveSheet.FSheet := FApplication.Sheets.Add;
  end;
end;

function TExcelApplication.GetActiveWorkBook: TExcelWorkBook;
begin
  Result := FActiveWorkBook;
end;

{ TExcelSheet }

constructor TExcelSheet.Create();
begin
  FUsedRange := TExcelRange.Create(Self, '');
end;

destructor TExcelSheet.Destroy;
begin
  if Assigned(FUsedRange) then
    FreeAndNil(FUsedRange);
  inherited;
end;

function TExcelSheet.GetCellValue(indexX, indexY: Integer): string;
var
  row, col:Integer;
  cell: TExcelCell;
  range: TExcelRange;
begin
  if (not isMerged(indexX, indexY)) then
  begin
    row := indexX;
    col := indexY;
  end else begin
    range := TExcelRange.Create(Self, GetCellName(indexY, indexY));
    try
      cell := range.StartCell;
      row := cell.Row;
      col := cell.Col;
    finally
      range.Free;
    end;
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

procedure TExcelSheet.MergeRange(row1, col1, row2, col2: Integer);
begin
  if (row1=row2) and (col1=col2) then
    Exit;

  MergeRange(GetCellName(row1,col1), GetCellName(row2,col2));
end;

procedure TExcelSheet.MergeRange(cellName1, cellName2: String);
var
  range: OleVariant;
  sRangeStr: string;
begin
  if cellName1 = cellName2 then
    Exit;

  sRangeStr := cellName1 + ':' + cellName2;
  range := FSheet.Range[sRangeStr];
  range.Select;
  range.Mergecells := True;
end;

function TExcelSheet.isMerged(row, col: Integer): boolean;
var
  sRangeStr: String;
  excelRange: TExcelRange;
begin
  sRangeStr := GetCellName(row,col);
  excelRange := TExcelRange.Create(Self, sRangeStr);
  try
    Result := excelRange.IsMerged;
  finally
    excelRange.Free;
  end;
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

procedure TExcelSheet.SetRowHeight(nRowIndex, value: Integer);
begin
  FSheet.Rows[nRowIndex].RowHeight := value;
end;

procedure TExcelSheet.SetSheet(sheet: OleVariant);
begin
  FSheet := sheet;
end;

function TExcelSheet.GetRowHeight(nRowIndex: Integer): Integer;
begin
  Result := FSheet.Rows[nRowIndex].RowHeight;
end;

function TExcelSheet.GetUsedRange: TExcelRange;
begin
  FUsedRange.Name := FSheet.UsedRange.Address;
  Result := FUsedRange;
end;

procedure TExcelSheet.SetName(value: string);
begin
  FSheet.Name := value;
end;

function TExcelSheet.GetCell(indexX, indexY: Integer): TExcelCell;
var
  cell: TExcelCell;
begin
  cell := TExcelCell.Create(Self, indexX, indexY);
  Result := cell;
end;

{ TExcelBook }

constructor TExcelWorkBook.Create();
begin
  FActiveSheet := TExcelSheet.Create();
end;

destructor TExcelWorkBook.Destroy;
begin
  FActiveSheet.Free;
  inherited;
end;

function TExcelWorkBook.GetActiveSheet: TExcelSheet;
begin
  FActiveSheet.Sheet := FWorkBook.ActiveSheet;
  Result := FActiveSheet;
end;

function TExcelWorkBook.GetSheet(index: Integer): TExcelSheet;
begin
  FActiveSheet.Sheet := FWorkBook.WorkSheets[index];
  Result := FActiveSheet;
end;

function TExcelWorkBook.GetSheetCount: Integer;
begin
  Result := FWorkBook.WorkSheets.Count;
end;

procedure TExcelWorkBook.SetWorkBook(value: OleVariant);
begin
  FWorkBook := value;
end;

{ TExcelCell }

constructor TExcelCell.Create(OwnerSheet: TExcelSheet; row, col: Integer);
begin
  Self.FOwnerSheet := OwnerSheet;
  if (row > 0) and (col > 0) then begin
    Self.FRow := row;
    Self.FCol := col;
    Self.FCell := OwnerSheet.FSheet.Cells[FRow,FCol];
  end;
end;

constructor TExcelCell.Create(OwnerSheet: TExcelSheet; rangeName: string);
begin
  Self.FOwnerSheet := OwnerSheet;
  Name := rangeName;
end;

destructor TExcelCell.Destroy;
begin

  inherited;
end;

function TExcelCell.Equals(Obj: TObject): Boolean;
var
  cell: TExcelCell;
begin
  Result := False;
  if not (obj is TExcelCell) then begin
    Exit;
  end;

  cell := TExcelCell(obj);

  Result := (cell.FOwnerSheet = Self.FOwnerSheet)
    and (cell.FRow = Self.Row) and (cell.FCol = Self.FCol);
end;

function TExcelCell.getName: string;
begin
  Result := GetCellName(FRow, FCol);
end;

function TExcelCell.GetValue: string;
begin
  Result := FCell.Value;
end;

procedure TExcelCell.SetName(value: String);
var
  mr: TMatchResult;
begin
  FName := value;
  if value = '' then
    Exit;
  mr := TRegexUtil.MatchFirst(REGEX_CELL, 0, value);
  if Assigned(mr) then
  try
    FCol := ExcelColName2Index(mr.Groups[2].MatchStr);
    FRow := StrToInt(mr.Groups[3].MatchStr);
    Self.Cell := FOwnerSheet.Sheet.Cells[FRow, FCol];
  finally
    mr.Free;
  end;
end;

procedure TExcelCell.SetValue(value: string);
begin
  FCell.Value := value;
end;

{ TExcelRange }

constructor TExcelRange.Create(OwnerSheet: TExcelSheet; rangeName: string);
begin
  FOwnerSheet := OwnerSheet;
  FStartCell := TExcelCell.Create(FOwnerSheet, '');
  FEndCell := TExcelCell.Create(FOwnerSheet, '');
  Name := rangeName;
end;

destructor TExcelRange.Destroy;
begin
  if Assigned(FEndCell) then
    FreeAndNil(FEndCell);
  if Assigned(FStartCell) then
    FreeAndNil(FStartCell);
  inherited;
end;

procedure TExcelRange.ExtendMergeRange;
var
  addr2: String;
begin
  addr2 := FRange.MergeArea.Address; // like $A$1:$A$5
  Name := addr2;
end;

function TExcelRange.IsMerged: Boolean;
var
  addr1, addr2: String;
begin
  Result := False;
  if FName = '' then
    Exit;

  addr1 := FRange.Address;  // like $A$1
  addr2 := FRange.MergeArea.Address; // like $A$1:$A$5
  Result := (addr1 <> addr2);
end;

procedure TExcelRange.SetName(value: String);
var
  nPos: Integer;
  cellStartStr, cellEndStr: string;
begin
  if value = '' then
    Exit;
  if FName = value then
    Exit;

  FName := value;
  FRange := FOwnerSheet.FSheet.Range[value];
  nPos := Pos(':', FName);
  if (nPos <> 0) then
  begin
    cellStartStr := Copy(FName, 1, nPos - 1);
    cellEndStr := Copy(FName, nPos + 1, MaxInt);
    FStartCell.Name := cellStartStr;
    FEndCell.Name := cellEndStr;
  end else begin
    FStartCell.Name := FName;
    FEndCell.Name := FName;
  end;
end;

function TExcelRange.GetColCount: Integer;
begin
  Result := EndCell.Col - StartCell.Col + 1;
end;

function TExcelRange.GetRowCount: Integer;
begin
  Result := EndCell.Row - StartCell.Row + 1;
end;

function TExcelRange.GetValue: string;
begin
  Result := FRange.Value;
end;

procedure TExcelRange.SetValue(value: string);
begin
  FRange.Value := value;
end;

end.
