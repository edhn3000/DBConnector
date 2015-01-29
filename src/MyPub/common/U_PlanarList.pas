{**
 * <p>Title: U_PlanarList  </p>
 * <p>Copyright: Copyright (c) 2008/12/7</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 二维数据列表 </p>
 *}

unit U_PlanarList;

interface          

uses
  Classes;

const
  C_PlanarItem_ReadSeperator = '|||';

type
  TPlanarJustifyMode = (pjmBothLeft, pjmBothRight, pjmNumLeftStrRight,
    pjmNumRightStrLeft);

  { 二维StringList 每个元素保存StringList 可规则输出二维表 }
  TPlanarStringList = class(TStringList)
  private                             
//    aryItemLengthToTal: array of Int64;
    aryItemMaxLength: array of Integer;
    aryIsNumbers: array of Boolean;
    FJustifyMode: TPlanarJustifyMode;
    FExcludeCols: TStringList;
    FExcludeRows: TStringList;
    function GetItem(indexRow: Integer): TStringList;
    procedure SetItem(indexRow: Integer; value: TStringList);
    function GetStr(indexRow, indexCol: Integer): string;
    procedure SetStr(indexRow, indexCol: Integer; s: string);
    function GetItemLengths(indexCol: Integer): Integer;
    function GetItemStr(indexRow: Integer): string;

    function IsNumber(S: string):Boolean;
    function Split(sText, sSeparator: String; ssParts: TStrings):TStrings;
  protected
  public
    property Items[indexRow: Integer]: TStringList read GetItem write Setitem;
    property ItemLengths[indexCol: Integer]: Integer read GetItemLengths;
    property Strs[indexRow, indexCol: Integer]: string read GetStr write SetStr;
    property ItemStr[indexRow: Integer]: string read GetItemStr;
    property ExcludeCols: TStringList read FExcludeCols;
    property ExcludeRows: TStringList read FExcludeRows;
    property JustifyMode: TPlanarJustifyMode read FJustifyMode write FJustifyMode;
  public
    constructor Create;
    destructor Destroy;override;
    procedure InsertPlanarItem(indexRow: Integer; slst: TStringList);
    function AddItem(slst: TStrings): Integer;overload;
    function AddItem(sSeperatedText: string; sSeperator: string =
      C_PlanarItem_ReadSeperator): Integer;overload;
    procedure AdjustItemLengths;
    procedure FormatItemLengths;

    function IsExcludeRow(nRow: Integer): Boolean;
    function IsExcludeCol(nCol: Integer): Boolean;  
    procedure AddExcludeRow(nRow: Integer);
    procedure AddExcludeCol(nCol: Integer);

    procedure Clear;override;
  end;

implementation

uses
  SysUtils, Math;

const
  C_Item_MaxLength = 100;

{ TPlanarStringList }

procedure TPlanarStringList.InsertPlanarItem(indexRow: Integer; slst: TStringList);
begin
  inherited InsertItem(indexRow, slst.Text, slst);
end;

function TPlanarStringList.AddItem(slst: TStrings): Integer;
begin
  Result := AddObject(slst.Text, slst);
end;

function TPlanarStringList.AddItem(sSeperatedText,
  sSeperator: string): Integer;
var
  slst: TStringList;
begin
  slst := TStringList.Create;
  Split(sSeperatedText, sSeperator, slst);
  Result := AddItem(slst);
end;

procedure TPlanarStringList.AdjustItemLengths;
var
  i, j: Integer;
  nLen: Integer;
  nCount: Integer;
  sValue: string;
  nPos: Integer;
  function CalcNewLen(nNew, nCurr: Integer): Integer;
  var
    nNewLen: Integer;
  begin
    if Abs(nNew - nCurr) < 5 then
      nNewLen := Math.Min(nNew, C_Item_MaxLength)
    else
    begin
      nNewLen := Round(Sqrt(nNew * Math.Max(nCurr, 1)));
      if nNewLen > C_Item_MaxLength then
        nNewLen := C_Item_MaxLength;
    end;
    Result := nNewLen;
  end;  
begin
  nCount := 0;
  for i := 0 to Count - 1 do
  begin
    if not IsExcludeRow(i) and (nCount < Items[i].Count) then
      nCount := Items[i].Count;
  end;                             
  SetLength(aryItemMaxLength, nCount);
  SetLength(aryIsNumbers, nCount);
  for i := Low(aryIsNumbers) to High(aryIsNumbers) do
  begin
    aryIsNumbers[i] := True;
  end;
  
  for i := 0 to Count - 1 do
  begin
    if not IsExcludeRow(i) then
    begin
      for j := 0 to Items[i].Count - 1 do
      begin
        sValue := GetStr(i, j);
        nPos := Pos(#$A, sValue);   // 如果有换行，通常因为太长，这里只统计换行前的长度
        if nPos <> 0 then
          nLen := nPos-1
        else
          nLen := Length(sValue);
        if not IsExcludeCol(j) then
        begin
          if (aryItemMaxLength[j] < nLen) then
            aryItemMaxLength[j] := CalcNewLen(nLen, aryItemMaxLength[j]);
        end;
        if aryIsNumbers[j] and not IsNumber(sValue) then
          aryIsNumbers[j] := False;
      end;
    end;
  end;
end;

function TPlanarStringList.GetItem(indexRow: Integer): TStringList;
begin
  Result := Objects[indexRow] as TStringList;
end;

function TPlanarStringList.GetStr(indexRow, indexCol: Integer): string;
begin
  if (Count > indexRow) and (GetItem(indexRow).Count > indexCol) then
    Result := GetItem(indexRow).Strings[indexCol]
  else
    Result := '';
end;  

procedure TPlanarStringList.SetStr(indexRow, indexCol: Integer; s: string);
begin
  while indexRow >= Count do
    AddItem(TStringList.Create);
  while indexCol >= Items[indexRow].Count do
    Items[indexRow].Add('');
  Items[indexRow].Strings[indexCol] := s;
  Self.Strings[indexRow] := Items[indexRow].Text;
end;

procedure TPlanarStringList.SetItem(indexRow: Integer; value: TStringList);
begin
  Objects[indexRow] := value;
end;    

constructor TPlanarStringList.Create;
begin
  inherited Create;
  FJustifyMode := pjmNumLeftStrRight;
  FExcludeCols := TStringList.Create;
  FExcludeRows := TStringList.Create;
end;

destructor TPlanarStringList.Destroy;
begin
  FExcludeCols.Free;
  FExcludeRows.Free;
  inherited;
end;

procedure TPlanarStringList.Clear;
begin
  inherited;
end;

function TPlanarStringList.GetItemLengths(indexCol: Integer): Integer;
begin
  Result := aryItemMaxLength[indexCol];
end;

procedure TPlanarStringList.FormatItemLengths;
var
  i, j: Integer;
  sFormat: string;
  sNumFormatHead, sStrFormatHead: string;
begin
  AdjustItemLengths;   
  case FJustifyMode of
  pjmBothLeft:
  begin
    sNumFormatHead := '%-';
    sStrFormatHead := '%-';
  end;
  pjmBothRight:
  begin
    sNumFormatHead := '%';
    sStrFormatHead := '%';
  end;
  pjmNumLeftStrRight:
  begin    
    sNumFormatHead := '%-';
    sStrFormatHead := '%';
  end;
  pjmNumRightStrLeft:
  begin
    sNumFormatHead := '%';
    sStrFormatHead := '%-';
  end;
  end;
  for i := 0 to Count - 1 do
  begin
    if not IsExcludeRow(i) then
    begin
      for j := 0 to Items[i].Count - 1 do
      begin
        if not IsExcludeCol(j) then
        begin
          if aryIsNumbers[j] then
            sFormat := sNumFormatHead + IntToStr(aryItemMaxLength[j]) + 's'      // 数字右对齐
          else
            sFormat := sStrFormatHead + IntToStr(aryItemMaxLength[j]) + 's';
          SetStr(i, j, Format(sFormat, [GetStr(i, j)]));
        end;
      end;
    end;
  end;
end;

function TPlanarStringList.GetItemStr(indexRow: Integer): string;
var
  i: Integer;
begin
  if IsExcludeRow(indexRow) then
    Result := Strings[indexRow]
  else
  begin
    for i := 0 to Items[indexRow].Count - 1 do
    begin
      if i = 0 then
        Result := Items[indexRow].Strings[i]
      else
        Result := Result + Items[indexRow].Strings[i];
    end;
  end;
end;

function TPlanarStringList.Split(sText, sSeparator: String;
  ssParts: TStrings): TStrings;
var
  nPos: Integer;
  sPart, sRemain: String;
begin
  Result := ssParts;
  if Result = nil then
    Result := TStringList.Create;
  Result.Clear;
  sRemain:= sText;
  nPos:= Pos(sSeparator, sRemain);

  while nPos > 0 do
  begin
    sPart:= Copy(sRemain, 1, nPos- 1);
    Result.Add(sPart);
    sRemain:= Copy(sRemain, nPos+ Length(sSeparator), Length(sRemain)- nPos+ 1);
    nPos:= Pos(sSeparator, sRemain);
  end;

  if sRemain <> '' then
    Result.Add(sRemain);
end;

function TPlanarStringList.IsNumber(S: string):Boolean;
var
  nCode: Integer;
  fdNum: Double;
begin
  Val(S, fdNum, nCode);
  Result := (nCode = 0) and (fdNum = fdNum);
end;    

procedure TPlanarStringList.AddExcludeCol(nCol: Integer);
begin
  ExcludeCols.Add(IntToStr(nCol));
end;

procedure TPlanarStringList.AddExcludeRow(nRow: Integer);
begin                             
  ExcludeRows.Add(IntToStr(nRow));
end;

function TPlanarStringList.IsExcludeCol(nCol: Integer): Boolean;
begin
  Result := -1 <> FExcludeCols.IndexOf(IntToStr(nCol));
end;

function TPlanarStringList.IsExcludeRow(nRow: Integer): Boolean;
begin                                                  
  Result := -1 <> FExcludeRows.IndexOf(IntToStr(nRow));
end;

end.
