{**
 * <p>Title: U_Statistic </p>
 * <p>Copyright: Copyright (c) 2006/8/30</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 统计实现单元 </p>
 * <p>Description: 对统计窗体生成的列表分析并显示 </p>
 * <p>Description: 显示方式包括横条图、柱图、饼图、列表 </p>
 *}
unit U_Statistic;

interface

uses
  DbChart, Series, Chart, TeEngine, Graphics, Classes, ComCtrls;

type
  { TDepartByTime }
  TDepartByTime = (eByNone, eByYear, eByMonth);

  { 一个TLabelValue对象包含在要显示的标签和对应的多个值 }
  TLabelValue = class
  private
    FsLabel: string;
    FslstValues: TStringList;

    function GetValue(index: Integer): Double;
    procedure SetValue(index: Integer;value: Double);
    function GetValuesCount: Integer;
    function GetValueSum: Double;
    function GetValidLabel: string;
  public
    constructor Create(sLabel: string;slstValues: TStringList = nil);
    destructor Destroy; override;

    procedure AddValue(value: string); overload;
    procedure AddValue(value: Integer); overload;
    procedure ClearValue;

    property sLabel: string read FsLabel;
    property ValidLabel: string read GetValidLabel;
    property Values[index: Integer]: Double read GetValue write SetValue;
    property ValuesCount: Integer read GetValuesCount;
    property ValueSum: Double read GetValueSum;
  end;

  // 清空dbchart
  procedure ClearDBChart(var Adbcht: TDbChart);

  // 显示横条, slstDatas的元素必须是TLabelValue对象
  procedure ShowHorizBars(var Adbcht: TDbChart;slstDatas: TStringList);

  // 显示饼图
  procedure ShowPieSeries(var Adbcht: TDbChart;slstDatas: TStringList);

  // 柱图
  procedure ShowPilarSeries(var Adbcht: TDbChart;slstDatas: TStringList);

  // 详细列表的列
  procedure AddListCols(var lvw: TListView;slstCols: TStringList);

  //
  procedure ShowDetailList(var lvw: TListView;slstDatas: TStringList);

implementation

uses
  SysUtils;

const
  aColors: array[0..9] of TColor = (clRed, clGreen, clBlue, clYellow,
    clWhite, clPurple, clSilver, clAqua, clGray, clBlack);

procedure ClearDBChart(var Adbcht: TDbChart);
begin
  Adbcht.SeriesList.Clear;
  Adbcht.Refresh;
end;

function GetValidIntInLst(slst: TStringList;I, J: Integer): Integer;
begin
  if I >= slst.Count then
    Result := 0
  else if not (slst.Objects[I] is TStringList) then
    Result := 0
  else if J >= (slst.Objects[I] as TStringList).Count then
    Result := 0
  else
  try
    Result := StrToInt((slst.Objects[I] as TStringList).Strings[J]);
  except
    Result := 0;
  end;
end;

function GetColor(nIndex: Integer): TColor;
begin
  Result := aColors[nIndex mod Length(aColors)];
end;

function GetBiggestCount(slst: TStringList): Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 0 to slst.Count - 1 do
  begin
    if slst.Objects[i] is TLabelValue then
      if (slst.Objects[i] as TLabelValue).ValuesCount > Result then
        Result := (slst.Objects[i] as TLabelValue).ValuesCount;
  end;
end;

procedure ShowHorizBars(var Adbcht: TDbChart;slstDatas: TStringList);
var
  ahobrSers: array of THorizBarSeries;
  i, j: Integer;
begin
  if (slstDatas.Count = 0) or (GetBiggestCount(slstDatas) = 0) then
  begin
    ClearDBChart(Adbcht);
    Exit;
  end;

  SetLength(ahobrSers, GetBiggestCount(slstDatas));

  Adbcht.SeriesList.Clear;

  for i := Low(ahobrSers) to High(ahobrSers) do
  begin
    ahobrSers[i] := THorizBarSeries.Create(nil);
    ahobrSers[i].ParentChart := Adbcht;
    ahobrSers[i].Marks.Style := smsValue;
  end;

  Adbcht.Legend.Alignment := laBottom; // 说明框在底部
  Adbcht.ClipPoints := True;
  Adbcht.AxisVisible := True; // 显示轴
  Adbcht.View3DWalls := True;
  Adbcht.Zoomed := False;
  Adbcht.AllowZoom := False;
  Adbcht.View3DOptions.Orthogonal := True;

  try
    for i := 0 to slstDatas.Count - 1 do
      for j := Low(ahobrSers) to High(ahobrSers) do
        Adbcht.Series[j].add((slstDatas.Objects[i] as TLabelValue).Values[j],
          (slstDatas.Objects[i] as TLabelValue).ValidLabel, GetColor(i));
  except

  end;
end;

procedure ShowPilarSeries(var Adbcht: TDbChart;slstDatas: TStringList);
var
  ahobrSers: array of TBarSeries;
  i, j: Integer;
begin
  if (slstDatas.Count = 0) or (GetBiggestCount(slstDatas) = 0) then
  begin
    ClearDBChart(Adbcht);
    Exit;
  end;

  SetLength(ahobrSers, GetBiggestCount(slstDatas));
  Adbcht.SeriesList.Clear;

  for i := Low(ahobrSers) to High(ahobrSers) do
  begin
    ahobrSers[i] := TBarSeries.Create(nil);
    ahobrSers[i].ParentChart := Adbcht;
    ahobrSers[i].Marks.Style := smsValue;
  end;

  Adbcht.Legend.Alignment := laBottom; // 说明框在底部
  Adbcht.ClipPoints := True;
  Adbcht.AxisVisible := True; // 显示轴
  Adbcht.View3DWalls := True;
  Adbcht.Zoomed := False;
  Adbcht.AllowZoom := False;
  Adbcht.View3DOptions.Orthogonal := True;

  try
    for i := 0 to slstDatas.Count - 1 do
      for j := Low(ahobrSers) to High(ahobrSers) do
        Adbcht.Series[j].add((slstDatas.Objects[i] as TLabelValue).Values[j],
          (slstDatas.Objects[i] as TLabelValue).ValidLabel, GetColor(i));
  except

  end;

end;

procedure ShowPieSeries(var Adbcht: TDbChart;slstDatas: TStringList);
var
  piSers1: TPieSeries;
  i, j: Integer;
  nSum: Double;
begin
  if (slstDatas.Count = 0) or (GetBiggestCount(slstDatas) = 0) then
  begin
    ClearDBChart(Adbcht);
    Exit;
  end;

  piSers1 := TPieSeries.Create(nil);

  Adbcht.SeriesList.Clear;
  piSers1.ParentChart := Adbcht;
  piSers1.Marks.Style := smsLabelPercent;

  Adbcht.Legend.Alignment := laBottom;
  Adbcht.Zoomed := False;
  Adbcht.AllowZoom := False;

  nSum := 0;
  try
    for i := 0 to slstDatas.Count - 1 do
      for j := 0 to (slstDatas.Objects[i] as TLabelValue).ValuesCount - 1 do
        nSum := nSum + (slstDatas.Objects[i] as TLabelValue).ValueSum;

    if nSum = 0 then
      Exit;

    for i := 0 to slstDatas.Count - 1 do
    begin
      Adbcht.Series[0].add((slstDatas.Objects[i] as TLabelValue).ValueSum /
        nSum,
        (slstDatas.Objects[i] as TLabelValue).ValidLabel,
        GetColor(i));
    end;
  except

  end;
end;

procedure AddListCols(var lvw: TListView;slstCols: TStringList);
var
  i: Integer;
begin
  lvw.Columns.Clear;
  for i := 0 to slstCols.Count - 1 do
    with lvw.Columns.Add do
    begin
      Caption := slstCols.Strings[i];
      Width := 100;
    end;
end;

procedure ShowDetailList(var lvw: TListView;slstDatas: TStringList);
var
  i, j: Integer;
  lblval: TLabelValue;
begin
  lvw.Items.Clear;
  for i := 0 to slstDatas.Count - 1 do
  begin
    with lvw.Items.Add do
    begin
      lblval := slstDatas.Objects[i] as TLabelValue;
      Caption := lblval.ValidLabel;
      for j := 0 to lblval.ValuesCount - 1 do
        SubItems.Add(FloatToStr(lblval.Values[j]));
    end;
  end;
end;

{ TLabelValue }

procedure TLabelValue.AddValue(value: string);
begin
  FslstValues.Add(value);
end;

procedure TLabelValue.AddValue(value: Integer);
begin
  AddValue(IntToStr(value));
end;

procedure TLabelValue.ClearValue;
begin
  FslstValues.Clear;
end;

constructor TLabelValue.Create(sLabel: string;slstValues: TStringList = nil);
begin
  FslstValues := slstValues;
  if not Assigned(FslstValues) then
    FslstValues := TStringList.Create;
  FsLabel := sLabel;
end;

destructor TLabelValue.Destroy;
begin
  if Assigned(FslstValues) then
    FslstValues.Free;
  inherited;
end;

function TLabelValue.GetValidLabel: string;
begin
  Result := sLabel;
  if Result = '' then
    Result := '未填写项';
end;

function TLabelValue.GetValue(index: Integer): Double;
begin
  if index >= FslstValues.Count then
    Result := 0
  else if FslstValues.Strings[index] = '' then
    Result := 0
  else
  try
    Result := StrToFloat(FslstValues.Strings[index]);
  except
    Result := 0;
  end;
end;

function TLabelValue.GetValuesCount: Integer;
begin
  Result := FslstValues.Count;
end;

function TLabelValue.GetValueSum: Double;
var
  i: Integer;
begin
  Result := 0;
  for i := 0 to FslstValues.Count - 1 do
    Result := Result + StrToFloatDef(FslstValues.Strings[i], 0);
end;

procedure TLabelValue.SetValue(index: Integer;value: Double);
begin
  FslstValues.Strings[index] := FloatToStr(value);
end;


end.
