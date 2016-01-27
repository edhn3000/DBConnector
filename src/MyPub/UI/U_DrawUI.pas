unit U_DrawUI;

{ 用于在界面绘制控件UI时的数据对象
  author: edhn
}

interface

uses
  Generics.Collections, Windows, Forms, ComCtrls, Controls, Classes,
  Types, Messages, Graphics, ExtCtrls, SysUtils, StdCtrls, Buttons;

const
  BUTTON_DRAW_NORMAL = 1;
  BUTTON_DRAW_HOVER = 2;
  BUTTON_DRAW_CLICK = 3;

type
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

    constructor Create();overload; virtual;
    constructor Create(Owner: TObject);overload; virtual;
    destructor Destroy();override;

    procedure Draw(Canvas: TCanvas; param: Integer);virtual; abstract;
  end;

  { TDrawUIButton }
  TDrawUIButton = class(TDrawUIBaseControl)
  private
    FCaption: String;
    FFont: TFont;
    FEnabled: Boolean;
    FOnClick: TNotifyEvent;
    FNormalPicture: TPicture;
    FHoverPicture: TPicture;
    FClickPicture: TPicture;
    FDisablePicture: TPicture;
    FDrawState: TButtonState;
    FIcon: TPicture;
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
    property OnClick: TNotifyEvent read FOnClick write FOnClick;

    constructor Create(Owner: TObject);override;
    destructor Destroy();override;

    procedure Draw(Canvas: TCanvas; param: Integer);override;
  end;

  { TDrawUIImage }
  TDrawUIImage = class(TDrawUIBaseControl)
  private
    FImage: TImage;
  public
    property Image: TImage read FImage write FImage;

    constructor Create(Owner: TObject);override;
    destructor Destroy();override;

    procedure Draw(Canvas: TCanvas; param: Integer);override;
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

procedure TDrawUIButton.Draw(Canvas: TCanvas; param: Integer);
var
  rBtn, rEdge, iconRect: TRect;
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
    Canvas.Brush.Color := $F4F4F4;
    if Assigned(DisablePicture.Graphic) and (not DisablePicture.Graphic.Empty) then begin
      Canvas.Draw(rEdge.Left, rEdge.Top, DisablePicture.Graphic);
    end else begin
      DrawEdge(Canvas.Handle, rEdge, EDGE_RAISED, BF_RECT);
      Canvas.FillRect(rBtn);
    end;
  end else begin
     Canvas.Brush.Color := Color;
     if param = BUTTON_DRAW_CLICK then begin
      if Assigned(ClickPicture.Graphic) and (not ClickPicture.Graphic.Empty) then begin
        Canvas.Draw(rEdge.Left, rEdge.Top, ClickPicture.Graphic);
      end else begin
        DrawEdge(Canvas.Handle, rEdge, EDGE_ETCHED, BF_RECT);
        Canvas.FillRect(rBtn);
      end;
    end else if param = BUTTON_DRAW_HOVER then begin
      if Assigned(HoverPicture.Graphic) and (not HoverPicture.Graphic.Empty) then begin
        Canvas.Draw(rEdge.Left, rEdge.Top, HoverPicture.Graphic);
      end else begin
        InflateRect(rEdge, 1, 1);
        DrawEdge(Canvas.Handle, rEdge, EDGE_RAISED, BF_RECT);
        Canvas.FillRect(rBtn);
      end;
    end else begin
      if Assigned(NormalPicture.Graphic) and (not NormalPicture.Graphic.Empty) then begin
        Canvas.Draw(rEdge.Left, rEdge.Top, NormalPicture.Graphic);
      end else begin
        DrawEdge(Canvas.Handle, rEdge, EDGE_RAISED, BF_RECT);
        Canvas.FillRect(rBtn);
      end;
    end;
  end;

  if Enabled then
    Canvas.Font.Color := Self.Font.Color
  else
    Canvas.Font.Color := clGrayText;
  Canvas.Font.Name := '微软雅黑';
  Canvas.Font.Size := 9;
  SetBkMode(Canvas.Handle, TRANSPARENT);
  DrawText(Canvas.Handle, Caption,
    Length(Caption), rBtn, DT_CENTER + DT_SINGLELINE + DT_VCENTER);
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

procedure TDrawUIImage.Draw(Canvas: TCanvas; param: Integer);
begin
  Canvas.Draw(Left, Top, Image.Picture.Bitmap);
end;

end.
