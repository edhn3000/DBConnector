{**
 * <p>Title: U_ExcelHelper </p>
 * <p>Copyright: Copyright (c) 2008-8-15 </p>
 * @author fengyq
 * @version 1.0
 * <p>Description: Excel处理 </p>
 *}
unit U_ExcelHelper;

interface    

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, OleServer, ComObj, Excel_TLB;

const
  C_CELL_Link = '=HYPERLINK("%s","%s")';

  // 边框样式相关常量
//  xlNone = -4142;
//  xlContinuous = 1;
//  xlThin = 2;
//  xlThick = 4;
  xlDiagonalDown = 5;
  xlDiagonalUp = 6;
  xlEdgeLeft = 7;
  xlEdgeTop = 8;
  xlEdgeBottom = 9;
  xlEdgeRight = 10;
  xlInsideVertical = 11;
  xlInsideHorizontal = 12;

type
  TExcelCell = record
    Row: Integer;
    Col: Integer;
  end;

  { TExcelSheet Excel文档的一个也签 }
  TExcelSheet = class
  private
    FSheet: OleVariant;
    function GetName(): string;
    procedure SetName(value: string);
    function GetCellValue(row, col: Integer): string;
    procedure SetCellValue(row, col: Integer; value: string);
  public
    property Name: string read GetName write SetName;
    property CellValue[row, col: Integer]: string read GetCellValue write SetCellValue;

    constructor Create;
    destructor Destroy;override;

    function GetColumnWidth(nColumnIndex: Integer): Integer;
    procedure SetColumnWidth(nColumnIndex, value: Integer);
  end;

  { TExcelBook 一个Excel文档的管理 }
  TExcelBook = class
  private
    FBook: OleVariant;
    FActiveSheet: TExcelSheet;
    function GetSheetCount: Integer;
    function GetSheet(index: Integer): TExcelSheet;
    function GetActiveSheet(): TExcelSheet;
    procedure SetWorkbook(value: OleVariant);
  public
    property Workbook: OleVariant read FBook write SetWorkbook;
    property SheetCount: Integer read GetSheetCount;
    property Sheets[index: Integer]: TExcelSheet read GetSheet;
    property ActiveSheet: TExcelSheet read GetActiveSheet;

    constructor Create();
    destructor Destroy;override;
  end;

  { Excel助手类，管理ExcelApplication整个生命周期 }
  TExcelHelper = class
  private
    FFileName: String;
    FOpened: Boolean;
    FExcelApp: OleVariant;
    FActiveBook: TExcelBook;
    FCloseOnFree: Boolean;
    function GetVisible(): Boolean;
    procedure SetVisible(value: Boolean);
    function GetActiveSheet(): TExcelSheet;
    function GetActiveWorkbook(): TExcelBook;
  protected
    procedure Release;
  public
    property CloseOnFree: Boolean read FCloseOnFree write FCloseOnFree;
    property Visible: Boolean read GetVisible write SetVisible;
    property ExcelApp: OleVariant read FExcelApp;
    property ActiveSheet: TExcelSheet read GetActiveSheet;
    property ActiveWorkbook: TExcelBook read GetActiveWorkbook;
  public
    constructor Create;overload;
    constructor Create(ExcepApp: OleVariant);overload;
    destructor Destroy;override;

    procedure NewFile(sFileName: string);
    procedure OpenFile(sFileName: string);
    procedure SaveFile();
    procedure SaveAs(sFileName: string);
    procedure CloseFile();

    procedure AddWorkbook();
    procedure AddWorkSheet();
    { 修改激活页签，注意第1个页签索引为1 }
    procedure ChangeActiveSheet(nSheetIndex: Integer);

    procedure SetRowHeight(nRowIndex: Integer; value: Integer);
    function GetRowHeight(nRowIndex: Integer): Integer;
    procedure MergeRange(row1, col1, row2, col2: Integer);

    procedure SetBorderStyle(row1, col1, row2, col2: Integer; xlStyle: Integer);

    { 将坐标转为单元格名称，如3,1转为A3 }
    class function GetCellName(row, col: Integer): string;
    { 将单元格名称转为坐标，如A3转为3,1，暂不支持字母部分超过1位
      支持只传入字母，如A转为0,1}
    class function GetCellPos(cellName: String): TExcelCell;
  end;


implementation

{ T_ExcelExportor }

constructor TExcelHelper.Create;
begin
  FExcelApp := CreateOleObject( 'Excel.Application' );
  FActiveBook := TExcelBook.Create;
  FCloseOnFree := True;
end;

constructor TExcelHelper.Create(ExcepApp: OleVariant);
begin
  FExcelApp := ExcepApp;
  FActiveBook := TExcelBook.Create;
  FCloseOnFree := True;
end;

destructor TExcelHelper.Destroy;
begin
  if FCloseOnFree then
  begin
    try
      if FOpened then
        CloseFile;
      FExcelApp.Quit;
    except
      // 关闭过程中可能有异常
    end;
  end;
  FExcelApp := Null;
  FreeAndNil(FActiveBook);
  Release;
  inherited;
end;

procedure TExcelHelper.Release;
begin
  FExcelApp := Unassigned;
end;

procedure TExcelHelper.CloseFile;
begin
  if FOpened then
  begin
    FExcelApp.ActiveWorkbook.Close(False); // 不自动保存内容
    FOpened := False;
  end;
end;

procedure TExcelHelper.OpenFile(sFileName: string);
begin
  FFileName := sFileName;
  FExcelApp.Workbooks.Open(sFileName);
  FActiveBook.Workbook := FExcelApp.ActiveWorkbook;
  FOpened := True;
end;

procedure TExcelHelper.ChangeActiveSheet(nSheetIndex: Integer);
begin
  FExcelApp.WorkSheets[nSheetIndex].Activate;
  FActiveBook.FActiveSheet.FSheet := FExcelApp.WorkSheets[nSheetIndex];
end;

procedure TExcelHelper.SaveFile;
begin
  FExcelApp.ActiveWorkbook.Save;
end;

procedure TExcelHelper.SetRowHeight(nRowIndex, value: Integer);
begin
  FExcelApp.Rows[nRowIndex].RowHeight := value;
end;

function TExcelHelper.GetRowHeight(nRowIndex: Integer): Integer;
begin
  Result := FExcelApp.Rows[nRowIndex].RowHeight;
end;

procedure TExcelHelper.SaveAs(sFileName: string);
begin
  FExcelApp.ActiveWorkbook.SaveAs(sFileName);
end;

procedure TExcelHelper.AddWorkbook;
begin
  FActiveBook.Workbook := FExcelApp.Workbooks.Add;
end;

class function TExcelHelper.GetCellName(row, col: Integer): string;
begin
  Result := Chr(Byte(col) + 64) + IntToStr(row);
end;

class function TExcelHelper.GetCellPos(cellName: String): TExcelCell;
var
  pos: Integer;
  sCol, sRow: String;
  cel: TExcelCell;
begin
  pos := 1;
  while pos <= Length(cellName) do
  begin
    // 查找数字起始位置
    if (Ord(cellName[pos]) >=48) and (Ord(cellName[pos]) <= 58) then
    begin
      sCol := Copy(cellName, 1, pos - 1);
      sRow := Copy(cellName, pos, MaxInt);
      // TODO 字母超过1位，转换为数字要考虑26进制计算
      cel.Col := Ord(sCol[1]) - 64;
      cel.Row := StrToInt(sRow);
      Break;
    end;
    Inc(pos);
  end;

  Result := cel;
end;

procedure TExcelHelper.NewFile(sFileName: string);
begin
  //
end;

function TExcelHelper.GetVisible: Boolean;
begin
  Result := FExcelApp.Visible;
end;

procedure TExcelHelper.SetVisible(value: Boolean);
begin
  FExcelApp.Visible := value;
end;

procedure TExcelHelper.MergeRange(row1, col1, row2, col2: Integer);
var
  range: OleVariant;
  sRangeStr: string;
begin
  if (row1=row2) and (col1=col2) then
    Exit;

  sRangeStr := TExcelHelper.GetCellName(row1,col1)
    + ':' + TExcelHelper.GetCellName(row2,col2);
  range := FExcelApp.ActiveSheet.Range[sRangeStr];
  range.Select;
  range.Mergecells := True;
end;

procedure TExcelHelper.SetBorderStyle(row1, col1, row2, col2, xlStyle: Integer);
var
  range: OleVariant;
  sRangeStr: string;
  bd: OleVariant;
begin
  if (row1=row2) and (col1=col2) then
    Exit;

  sRangeStr := TExcelHelper.GetCellName(row1,col1)
    + ':' + TExcelHelper.GetCellName(row2,col2);
  range := FExcelApp.ActiveSheet.Range[sRangeStr];
  range.Select;
//  FExcelApp.Selection.Borders(xlDiagonalDown).LineStyle := xlNone;
//  FExcelApp.Selection.Borders(xlDiagonalUp).LineStyle := xlNone;
  bd := FExcelApp.Selection.Borders[xlEdgeLeft];
  bd.LineStyle := xlContinuous;
  bd.ColorIndex := 0;
  bd.TintAndShade := 0;
  bd.Weight := xlStyle;   // xlThin

  bd := FExcelApp.Selection.Borders[xlEdgeRight];
  bd.LineStyle := xlContinuous;
  bd.ColorIndex := 0;
  bd.TintAndShade := 0;
  bd.Weight := xlStyle;

  bd := FExcelApp.Selection.Borders[xlEdgeTop];
  bd.LineStyle := xlContinuous;
  bd.ColorIndex := 0;
  bd.TintAndShade := 0;
  bd.Weight := xlStyle;

  bd := FExcelApp.Selection.Borders[xlEdgeBottom];
  bd.LineStyle := xlContinuous;
  bd.ColorIndex := 0;
  bd.TintAndShade := 0;
  bd.Weight := xlStyle;

  bd := FExcelApp.Selection.Borders[xlInsideVertical];
  bd.LineStyle := xlContinuous;
  bd.ColorIndex := 0;
  bd.TintAndShade := 0;
  bd.Weight := xlStyle;

  bd := FExcelApp.Selection.Borders[xlInsideHorizontal];
  bd.LineStyle := xlContinuous;
  bd.ColorIndex := 0;
  bd.TintAndShade := 0;
  bd.Weight := xlStyle;
end;

function TExcelHelper.GetActiveSheet: TExcelSheet;
begin
  Result := FActiveBook.ActiveSheet;
end;

procedure TExcelHelper.AddWorkSheet;
begin
  if FExcelApp.Workbooks.Count = 0 then
  begin
    FActiveBook.Workbook := FExcelApp.Workbooks.Add;
//    FActiveBook.FActiveSheet.FSheet := FExcelApp.Sheets[1];
  end else
  begin
    FExcelApp.Sheets.Add;
    FActiveBook.Workbook  := FExcelApp.ActiveWorkbook;
//    FActiveBook.FActiveSheet.FSheet := FExcelApp.Sheets.Add;
  end;
end;

function TExcelHelper.GetActiveWorkbook: TExcelBook;
begin
  Result := FActiveBook;
end;

{ TExcelSheet }

constructor TExcelSheet.Create;
begin

end;

destructor TExcelSheet.Destroy;
begin

  inherited;
end;

function TExcelSheet.GetCellValue(row, col: Integer): string;
begin
  Result := FSheet.Cells[row,col].Value;
end;

function TExcelSheet.GetName: string;
begin
  Result := FSheet.Name;
end;

procedure TExcelSheet.SetCellValue(row, col: Integer; value: string);
begin
  FSheet.Cells[row,col].Value := value;
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

{ TExcelBook }

function TExcelBook.GetActiveSheet: TExcelSheet;
begin
  Result := FActiveSheet;
end;

function TExcelBook.GetSheet(index: Integer): TExcelSheet;
begin
  FActiveSheet.FSheet := FBook.WorkSheets[index];
  Result := FActiveSheet;
end;

function TExcelBook.GetSheetCount: Integer;
begin
  Result := FBook.WorkSheets.Count;
end;

procedure TExcelBook.SetWorkbook(value: OleVariant);
begin
  FBook := value;
  FActiveSheet.FSheet := FBook.ActiveSheet;
end;

constructor TExcelBook.Create();
begin
  FActiveSheet := TExcelSheet.Create;
end;

destructor TExcelBook.Destroy;
begin
  FActiveSheet.Free;
  inherited;
end;

end.
