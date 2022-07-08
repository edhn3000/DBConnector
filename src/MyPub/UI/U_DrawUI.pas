unit U_DrawUI;

{ 用于在界面绘制控件UI时的数据对象
  author: edhn
}

interface

uses
  Windows, Forms, ComCtrls, Controls, Classes,
  Types, Messages, Graphics, ExtCtrls, SysUtils, Buttons;

const
  BUTTON_DRAW_NORMAL = 1;
  BUTTON_DRAW_HOVER = 2;
  BUTTON_DRAW_CLICK = 3;

type
  { THighlightData }
  THighlightData = record// class(TObject)
  private
    FMainColor: TColor;
    FMainInverseColor: TColor;
    FKeyColor: TColor;
    FKeyInverseColor: TColor;
    FStr: string;
    FPos: Integer;
    FLen: Integer;
    procedure SetMainColor(value: TColor);
    procedure SetMainInverseColor(value: TColor);
    procedure SetKeyColor(value: TColor);
    procedure SetKeyInverseColor(value: TColor);
    procedure SetStr(value: String);
    procedure SetPos(value: Integer);
    procedure SetLen(value: Integer);
    function GetEndPos: Integer;
  public
    UseInverse: Boolean; // 是否使用反白色

    property MainColor: TColor read FMainColor write SetMainColor;
    property MainInverseColor: TColor read FMainInverseColor write SetMainInverseColor;
    property KeyColor: TColor read FKeyColor write SetKeyColor;
    property KeyInverseColor: TColor read FKeyInverseColor write SetKeyInverseColor;

    property Str: string read FStr write SetStr;          // 文本内容
    property Pos: Integer read FPos write SetPos;         // 关键字start
    property Len: Integer read FLen write SetLen;         // 关键字len
    property EndPos: Integer read GetEndPos;              // 关键字end

    procedure InitColor(mainColor, mainInverse, keyColor, keyInverse: TColor);
    function Clone(): THighlightData;
  end;

  TDrawUIUtil = class
  private
    // 输出高亮文本，只能传入一行文本
    class procedure HighlightTextOutLine(ACanvas: TCanvas; X, Y: Integer;
      const hlData: THighlightData);
  public
{-------------------------------------------------------------------------------
  过程名:    GetLineCount 字符串行数
  参数:      wrapChar: Char;s: WideString;
  返回值:    Integer
-------------------------------------------------------------------------------}
  class function GetLineCount( wrapChar: Char; s: WideString): Integer;
{-------------------------------------------------------------------------------
  过程名:    TrimStrRows 按行数缩短字符串，尾部会加...
  参数:      ACanvas: TCanvas;text: string;nWidth: Integer;maxLine: Integer;
  返回值:    String
-------------------------------------------------------------------------------}
  class function TrimStrRows(ACanvas: TCanvas; text:string; nWidth: Integer; maxLine: Integer): String;
{-------------------------------------------------------------------------------
  过程名:    WrapText 文本折行，折行部分添加换行符#$A
             用于在List、Hint的DrawItem中显示文字时折行计算
  参数:      ACanvas: TCanvas;error: string;
             nWidth: Integer 每行宽度
  返回值:    string
-------------------------------------------------------------------------------}
  class function WrapText(ACanvas: TCanvas; text:string; nWidth: Integer):string;overload;
  class function WrapText(ACanvas: TCanvas; text:string; nWidth: Integer; maxLine: Integer):string;overload;

{-------------------------------------------------------------------------------
  过程名:    HighlightTextOut 输出高亮文本，支持多行
  参数:      ACanvas: TCanvas;X: Integer;Y: Integer;hlData: THighlightData;
  返回值:    Integer 返回输出的行数
-------------------------------------------------------------------------------}
  class function HighlightTextOut(ACanvas: TCanvas; X, Y: Integer;
    const hlData: THighlightData): Integer;
  end;

  TDrawUIBaseControl = class
  private
    FOwner: TObject;
    FLeft: Integer;
    FTop: Integer;
    FWidth: Integer;
    FHeight: Integer;
    FColor: TColor;
    FHint: String;
    FEnabled: Boolean;
    FVisbile: Boolean;
    FOnMouseLeave: TNotifyEvent;
    FOnMouseEnter: TNotifyEvent;

    function GetBrushRect: TRect;
    function GetEdgeRect: TRect;
    procedure SetEdgeRect(value: TRect);
  protected
    procedure SetBounds(ALeft, ATop, AWidth, AHeight: Integer);
  public
    property Owner: TObject read FOwner write FOwner;
    property Left: Integer read FLeft write FLeft;
    property Top: Integer read FTop write FTop;
    property Width: Integer read FWidth write FWidth;
    property Height: Integer read FHeight write FHeight;
    property Color: TColor read FColor write FColor;
    property Hint: String read FHint write FHint;
    property Enabled: Boolean read FEnabled write FEnabled;
    property Visbile: Boolean read FVisbile write FVisbile;
    property BrushRect: TRect read GetBrushRect;
    property EdgeRect: TRect read GetEdgeRect write SetEdgeRect;

    // events
    property OnMouseLeave: TNotifyEvent read FOnMouseLeave write FOnMouseLeave;
    property OnMouseEnter: TNotifyEvent read FOnMouseEnter write FOnMouseEnter;

    constructor Create();overload; virtual;
    constructor Create(Owner: TObject);overload; virtual;
    destructor Destroy();override;

    procedure Draw(Canvas: TCanvas; drawParam: Integer);virtual; abstract;
  end;

  { TDrawUIButton }
  TDrawUIButton = class(TDrawUIBaseControl)
  private
    FCaption: String;
    FFont: TFont;
    FEnabled: Boolean;
    FNormalPicture: TPicture;
    FHoverPicture: TPicture;
    FClickPicture: TPicture;
    FDisablePicture: TPicture;
    FDrawState: TButtonState;
    FIcon: TPicture;
    FOnClick: TNotifyEvent;
  public
    MouseOnButton: Boolean;
    property Caption: String read FCaption write FCaption;
    property Font: TFont read FFont write FFont;
    property Enabled: Boolean read FEnabled write FEnabled;
    property Icon: TPicture read FIcon write FIcon;
    property NormalPicture: TPicture read FNormalPicture;
    property HoverPicture: TPicture read FHoverPicture;
    property ClickPicture: TPicture read FClickPicture;
    property DisablePicture: TPicture read FDisablePicture;
    property DrawState: TButtonState read FDrawState;
    // events
    property OnClick: TNotifyEvent read FOnClick write FOnClick;

    constructor Create(Owner: TObject);override;
    destructor Destroy();override;

    procedure Draw(Canvas: TCanvas; drawParam: Integer);override;
  end;

  { TDrawUIImage }
  TDrawUIImage = class(TDrawUIBaseControl)
  private
    FImage: TImage;
  public
    property Image: TImage read FImage write FImage;

    constructor Create(Owner: TObject);override;
    destructor Destroy();override;

    procedure Draw(Canvas: TCanvas; drawParam: Integer);override;
  end;

implementation

{ TDrawBaseControl }

constructor TDrawUIBaseControl.Create(Owner: TObject);
begin
  FOwner := Owner;
  FEnabled := True;
  FVisbile := True;
end;

constructor TDrawUIBaseControl.Create;
begin
  FEnabled := True;
  FVisbile := True;
end;

destructor TDrawUIBaseControl.Destroy;
begin

  inherited;
end;

function TDrawUIBaseControl.GetBrushRect: TRect;
begin
  Result.Left := Left + 1;
  Result.Top := Top + 1;
  Result.Right := Left + Width - 1;
  Result.Bottom := Top + Height - 1;
end;

function TDrawUIBaseControl.GetEdgeRect: TRect;
begin
  Result.Left := Left;
  Result.Top := Top;
  Result.Right := Left + Width;
  Result.Bottom := Top + Height;
end;

procedure TDrawUIBaseControl.SetBounds(ALeft, ATop, AWidth, AHeight: Integer);
begin
  FLeft := ALeft;
  FTop := ATop;
  FWidth := AWidth;
  FHeight := AHeight;
end;

procedure TDrawUIBaseControl.SetEdgeRect(value: TRect);
begin
  with value do
    SetBounds(Left, Top, Right - Left, Bottom - Top);
end;

{ TDrawButton }

constructor TDrawUIButton.Create(Owner: TObject);
begin
  inherited Create(Owner);
  FFont := TFont.Create;
  FEnabled := True;
  FIcon := TPicture.Create;
  FNormalPicture := TPicture.Create;
  FHoverPicture := TPicture.Create;
  FClickPicture := TPicture.Create;
  FDisablePicture := TPicture.Create;
end;

destructor TDrawUIButton.Destroy;
begin
  FFont.Free;
  FIcon.Free;
  FNormalPicture.Free;
  FHoverPicture.Free;
  FClickPicture.Free;
  FDisablePicture.Free;
  inherited;
end;

procedure TDrawUIButton.Draw(Canvas: TCanvas; drawParam: Integer);
var
  rBtn, rEdge, iconRect: TRect;
  fontOri: TFont;
begin
  rBtn := BrushRect;
  rEdge := Self.EdgeRect;
  iconRect := Classes.Rect(0, 0, 0, 0);
  if Assigned(FIcon.Graphic) and (not FIcon.Graphic.Empty) then begin
    iconRect := Classes.Rect(rEdge.Left + 2, rEdge.Top + 1,
      rEdge.Left + Self.Height - 2, rEdge.Top + Self.Height - 1);
    Canvas.StretchDraw(iconRect, FIcon.Graphic);
  end;

  rBtn.Left := rBtn.Left + RectWidth(iconRect);
  if not Enabled then begin
    if Assigned(DisablePicture.Graphic) and (not DisablePicture.Graphic.Empty) then begin
//      Canvas.Draw(rEdge.Left, rEdge.Top, DisablePicture.Graphic);
      Canvas.StretchDraw(rEdge, DisablePicture.Graphic);
    end else begin
      Canvas.Brush.Color := $F4F4F4;
      DrawEdge(Canvas.Handle, rEdge, EDGE_RAISED, BF_RECT);
      Canvas.FillRect(rBtn);
    end;
  end else begin
     Canvas.Brush.Color := Color;
     if drawParam = BUTTON_DRAW_CLICK then begin
      if Assigned(ClickPicture.Graphic) and (not ClickPicture.Graphic.Empty) then begin
//        Canvas.Draw(rEdge.Left, rEdge.Top, ClickPicture.Graphic);
        Canvas.StretchDraw(rEdge, ClickPicture.Graphic);
      end else begin
        DrawEdge(Canvas.Handle, rEdge, EDGE_ETCHED, BF_RECT);
        Canvas.FillRect(rBtn);
      end;
    end else if drawParam = BUTTON_DRAW_HOVER then begin
      if Assigned(HoverPicture.Graphic) and (not HoverPicture.Graphic.Empty) then begin
//        Canvas.Draw(rEdge.Left, rEdge.Top, HoverPicture.Graphic);
        Canvas.StretchDraw(rEdge, HoverPicture.Graphic);
      end else begin
        InflateRect(rEdge, 1, 1);
        DrawEdge(Canvas.Handle, rEdge, EDGE_RAISED, BF_RECT);
        Canvas.FillRect(rBtn);
      end;
    end else begin
      if Assigned(NormalPicture.Graphic) and (not NormalPicture.Graphic.Empty) then begin
//        Canvas.Draw(rEdge.Left, rEdge.Top, NormalPicture.Graphic);
        Canvas.StretchDraw(rEdge, NormalPicture.Graphic);
      end else begin
        DrawEdge(Canvas.Handle, rEdge, EDGE_RAISED, BF_RECT);
        Canvas.FillRect(rBtn);
      end;
    end;
  end;

  fontOri := TFont.Create;
  fontOri.Assign(Canvas.Font);
  try
    if Enabled then
      Canvas.Font.Color := Self.Font.Color
    else
      Canvas.Font.Color := clGrayText;
    Canvas.Font.Name := Self.Font.Name;
    Canvas.Font.Size := Self.Font.Size;
    SetBkMode(Canvas.Handle, TRANSPARENT);
    DrawText(Canvas.Handle, Caption,
      Length(Caption), rBtn, DT_CENTER + DT_SINGLELINE + DT_VCENTER);
    Canvas.Font.Assign(fontOri);
  finally
    fontOri.Free;
  end;
end;

{ TDrawImage }

constructor TDrawUIImage.Create(Owner: TObject);
begin
  inherited Create(Owner);
  FImage := TImage.Create(nil);
end;

destructor TDrawUIImage.Destroy;
begin
  FImage.Free;
  inherited;
end;

procedure TDrawUIImage.Draw(Canvas: TCanvas; drawParam: Integer);
begin
  Canvas.Draw(Left, Top, Image.Picture.Graphic);
end;

{ TDrawUIUtil }

class function TDrawUIUtil.GetLineCount(wrapChar: Char; s: WideString): Integer;
var
  npos: Integer;
begin
  Result := 0;
  while True do
  begin
    npos := Pos( wrapChar, s );
    if npos = 0 then
    begin
      if s <> '' then
        Inc(result);

      Break;
    end
    else
    begin
      Delete(s, 1, npos);
      Inc(result);
    end;
  end;
end;

class procedure TDrawUIUtil.HighlightTextOutLine(ACanvas: TCanvas; X, Y: Integer;
  const hlData: THighlightData);
var
  sLine, sln1,sln2,sln3:string;
  topBase, leftBase: Integer;
begin
  topBase := Y;
  leftBase := X;
  sLine := hlData.Str;
  if (hlData.Pos > Length(sLine)) or (hlData.EndPos <= 0) then
  begin
    ACanvas.TextOut(leftBase,  topBase, sLine);
    Exit;
  end;
  if hlData.UseInverse then
    ACanvas.Font.Color := hlData.MainInverseColor
  else
    ACanvas.Font.Color := hlData.MainColor;
  sln1 := Copy(sLine, 1, hlData.Pos - 1);
  ACanvas.TextOut(leftBase, topBase, sln1);
  if hlData.UseInverse then
    ACanvas.Font.Color := hlData.KeyInverseColor
  else
    ACanvas.Font.Color := hlData.KeyColor;
  if hlData.Pos > 0 then
    sln2 := Copy(sLine, hlData.Pos, hlData.Len)
  else
    sln2 := Copy(sLine, 1, hlData.Len);
  ACanvas.TextOut(leftBase + ACanvas.TextWidth(sln1), topBase, sln2);
  if hlData.UseInverse then
    ACanvas.Font.Color := hlData.MainInverseColor
  else
    ACanvas.Font.Color := hlData.MainColor;
  sln3 := Copy(sLine, hlData.EndPos + 1, Length(sLine) - hlData.EndPos);
  ACanvas.TextOut(leftBase + ACanvas.TextWidth(sln1 + sln2), topBase, sln3);
end;

class function TDrawUIUtil.HighlightTextOut(ACanvas: TCanvas; X, Y: Integer;
  const hlData: THighlightData): Integer;
var
  i, iPos, iEndPos, nRows, lineHeight: Integer;
  sTmp, sln: String;
  hiLine: THighlightData;
begin
  sTmp := hlData.Str;
  iPos := hlData.Pos;
  iEndPos := hlData.EndPos;
  hiLine := hlData.Clone;
  nRows := 0;
  if Trim(sTmp) <> '' then begin
    lineHeight := ACanvas.TextHeight(Copy(sTmp, 1, 1));
    // 逐行输出，并记录输出了多少行，返回给调用者
    while true do begin
      I := Pos(#10,sTmp);
      if I <> 0 then begin
        sln := Copy(sTmp,1,I-1);
        sTmp := Copy(sTmp,I+1,Length(sTmp));

        hiLine.Str := sln;
        hiLine.Pos := iPos;
        hiLine.Len := iEndPos - iPos + 1;
        HighlightTextOutLine(ACanvas, X, Y + lineHeight * nRows, hiLine);
        Inc(nRows);
        iPos := iPos - Length(sln);
        iEndPos := iEndPos - Length(sln);
      end else begin
        if Length(sTmp) <> 0 then begin
          hiLine.Str := sTmp;
          hiLine.Pos := iPos;
          hiLine.Len := iEndPos - iPos + 1;
          HighlightTextOutLine(ACanvas, X, Y + lineHeight * nRows, hiLine);
          Inc(nRows);
        end;
        System.Break;
      end;
    end;
  end;
  Result := nRows;
end;

class function TDrawUIUtil.TrimStrRows(ACanvas: TCanvas; text: string; nWidth,
  maxLine: Integer): String;
var
  i, rowsCount: Integer;
  nOverLen, nOverRows: Integer;
  subStr, sLastLine: String;
begin
  nOverRows := GetLineCount(#10, text) - maxLine;
  if nOverRows <= 0 then begin
    Result := text;
    Exit;
  end;

  rowsCount := 0;
  i := Length(text);
  while i > 0 do begin
    if text[i] = #10 then
      Inc(rowsCount);
    if rowsCount = nOverRows then
      System.Break;
    Dec(i);
  end;
  subStr := Copy(text, 1, i - 1); // 从开头截取到换行符之前

  if Trim(subStr) <> '' then begin
    i := Length(subStr) - 1;
    while i> 0 do begin
      if subStr[i] = #10 then
        System.Break;
      Dec(i);
    end;
    if i = 0 then
      sLastLine := subStr
    else
      sLastLine := Copy(subStr, i + 1, MaxInt);
    nOverLen := ACanvas.TextWidth(sLastLine + '...') - nWidth;

    while nOverLen > 0 do begin
      sLastLine := Copy(sLastLine, 1, Length(sLastLine) - 1);
      nOverLen := ACanvas.TextWidth(sLastLine + '...') - nWidth;
    end;
    subStr := Copy(subStr, 1, i) + sLastLine + '...';
  end;
  Result := subStr;
end;

class function TDrawUIUtil.WrapText(ACanvas: TCanvas; text: string;
  nWidth: Integer): string;
var
  iWid,i,iPos:integer;
  sTmp:string;
begin
  iPos:=1;
  iWid:=nWidth;
  Result := '';
  for i := 1 to Length(text) do
  begin
    sTmp:=Copy(text,iPos, i - iPos + 1);
    if ACanvas.TextWidth(sTmp)>iWid then
    begin
      Result:=Result+Copy(text,iPos,i - iPos)+#10;
      iPos:= i;
    end;
  end;

  Result:=Result+Copy(text,iPos,Length(text) - iPos + 1)
end;

class function TDrawUIUtil.WrapText(ACanvas: TCanvas; text: string; nWidth,
  maxLine: Integer): string;
var
  iWid,i,iPos:integer;
  sTmp:string;
begin
  iPos:=1;
  iWid:=nWidth;
  Result := '';
  for i := 1 to Length(text) do
  begin
    if text[i] = #10 then begin
      Result := Result + Copy(text, iPos, i - iPos + 1);
      iPos := i + 1;
      Continue;
    end;

    sTmp := Copy(text,iPos,i - iPos + 1);
    if ACanvas.TextWidth(sTmp)>iWid then
    begin
      Result:=Result + Copy(text,iPos,i - iPos)+#10;
      iPos:= i;
    end;
  end;

  Result := Result + Copy(text,iPos,Length(text) - iPos + 1);
  Result := TrimStrRows(ACanvas, Result, nWidth, maxLine);
end;

{ THighlightData }

procedure THighlightData.SetStr(value: String);
begin
  FStr := value;
end;

procedure THighlightData.SetPos(value: Integer);
begin
  FPos := value;
end;

procedure THighlightData.SetMainColor(value: TColor);
begin
  FMainColor := value;
end;

procedure THighlightData.SetMainInverseColor(value: TColor);
begin
  FMainInverseColor := value;
end;

procedure THighlightData.SetKeyColor(value: TColor);
begin
  FKeyColor := value;
end;

procedure THighlightData.SetKeyInverseColor(value: TColor);
begin
  FKeyInverseColor := value;
end;

procedure THighlightData.SetLen(value: Integer);
begin
  FLen := value;
end;

function THighlightData.GetEndPos: Integer;
begin
  Result := Pos + Len - 1;
end;

procedure THighlightData.InitColor(mainColor, mainInverse, keyColor,
  keyInverse: TColor);
begin
  Self.MainColor := mainColor;
  Self.MainInverseColor := mainInverse;
  Self.KeyColor := keyColor;
  Self.KeyInverseColor := keyInverse;
end;

function THighlightData.Clone: THighlightData;
var
  data: THighlightData;
begin
  data.Pos := Self.Pos;
  data.Len := Self.Len;
  data.Str := Self.Str;
  data.InitColor(Self.MainColor, Self.MainInverseColor, Self.KeyColor, Self.KeyInverseColor);
  data.UseInverse := Self.UseInverse;
  Result := data;
end;

end.
