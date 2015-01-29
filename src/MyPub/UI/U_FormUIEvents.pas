unit U_FormUIEvents;

interface

uses
  Generics.Collections, Windows, Forms, ComCtrls, Controls, Classes,
  Types, Messages, Graphics, ExtCtrls, SysUtils, StdCtrls;

type
  { 控件在鼠标拖动时可改变大小的8种方向 }
  TResizePosition = (rpUp, rpDown, rpLeft, rpRight, rpLeftUp, rpLeftDown,
    rpRightUp, rpRightDown);
  TResizePositions = Set of TResizePosition;
  { 窗体停靠方式，相当于对TResizePosition的归类 }
  TResizeAlign = (raTop, raDown, raLeft, raRight, raClient, raInner);
  { 窗体右上角按钮 }
  TFormSysButton = (fsbClose, fsbMin, fsMenu, fsSkin);
  TFormSysButtons = Set of TFormSysButton;

  { 事件定义 }
  TMouseDownEvent = procedure (Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer) of object;
  TMouseMoveEvent = procedure (Sender: TObject; Shift: TShiftState; X,
      Y: Integer) of object;

type
  TBaseUIEvents = class
  protected

  public
    constructor Create();
    destructor Destroy;override;

  end;

  { 实现窗体在MouseDown和MouseMove时的移动、改变大小功能，无边框窗体需要用到 }
  TFormPositionEvents = class(TBaseUIEvents)
  private
    FForm: TCustomForm;
  public
    property Form: TCustomForm read FForm write FForm;

    constructor Create(form: TCustomForm);
    destructor Destroy;override;


    // common event
    procedure TopMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure TopMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure BottomMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure BottomMouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure LeftMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure LeftMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure RightMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure RightMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure ClientMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ClientMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure InnerMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure InnerMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);

    class function ResizePositionTest(rect: TRect; X, Y: Integer;
      var outResizeWParam: Integer; var ourCursor: TCursor;
      rps: TResizePositions): Boolean;

  end;

  { 使用Image控件代替按钮时需要使用到的事件对象 }
  TImageButtonEvents = class(TBaseUIEvents)
  private
    FImgBtn: TImage;
    FForm: TCustomForm;
    FMouseDownPicture: TPicture;
    FMouseEnterPicture: TPicture;
    FMouseLeavePicture: TPicture;
  protected
    procedure ImageMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ImageMouseEnter(Sender: TObject);
    procedure ImageMouseLeave(Sender: TObject);
  public
    CursorHoverHand: Boolean;
    property MouseDownPicture: TPicture read FMouseDownPicture;
    property MouseEnterPicture: TPicture read FMouseEnterPicture;
    property MouseLeavePicture: TPicture read FMouseLeavePicture;
    property Form: TCustomForm read FForm write FForm;
    constructor Create(imgBtn: TImage; form: TCustomForm);
    destructor Destroy;override;

    procedure ImgCloseClick(Sender: TObject);
    procedure ImgMinClick(Sender: TObject);
  end;

  { TLabelLinkEvents }
  TLabelLinkEvents = class(TBaseUIEvents)
  private
    FLabel: TLabel;

  protected
    procedure LabelMouseEnter(Sender: TObject);
    procedure LabelMouseLeave(Sender: TObject);
  public
    CursorHoverHand: Boolean;
    constructor Create(lbl: TLabel);
    destructor Destroy;override;

  end;

const
  // SC is WM_SYSCOMMAND
  SC_wParam_FormResize_Left = $F001;
  SC_wParam_FormResize_Right = $F002;
  SC_wParam_FormResize_Top = $F003;
  SC_wParam_FormResize_TopLeft = $F004;
  SC_wParam_FormResize_TopRight = $F005;
  SC_wParam_FormResize_Bottom = $F006;
  SC_wParam_FormResize_BottomLeft = $F007;
  SC_wParam_FormResize_BottomRight = $F008;

  SC_wParam_FormMove = $F012;

implementation

{ TBaseUIEvents }

constructor TBaseUIEvents.Create;
begin

end;

destructor TBaseUIEvents.Destroy;
begin

  inherited;
end;

{ TFormPositionEvents }

constructor TFormPositionEvents.Create(form: TCustomForm);
begin
  FForm := form;
end;

destructor TFormPositionEvents.Destroy;
begin

  inherited;
end;

class function TFormPositionEvents.ResizePositionTest(rect: TRect; X, Y: Integer;
  var outResizeWParam: Integer; var ourCursor: TCursor;
  rps: TResizePositions): Boolean;
var
  cr: TCursor;
  nWidth, nHeight: Integer;
begin
  Result := True;
  cr := Screen.Cursor;
  nWidth := rect.Right - rect.Left;
  nHeight := rect.Bottom - rect.Top;
  if (rpLeftUp in rps) and ((X < 5) and (Y < 5)) then begin   // 左上
    outResizeWParam := SC_wParam_FormResize_TopLeft;
    cr := crSizeNWSE;
  end
  else if (rpRightUp in rps) and ((X > nWidth - 5) and (Y < 5)) then begin // 右上
    outResizeWParam := SC_wParam_FormResize_TopRight;
    cr := crSizeNESW;
  end
  else if (rpRightDown in rps) and ((X > nWidth - 5) and (Y > nHeight - 5)) then begin // 右下
    outResizeWParam := SC_wParam_FormResize_BottomRight;
    cr := crSizeNWSE;
  end
  else if (rpLeftDown in rps) and ((X < 5) and (Y > nHeight - 5)) then begin // 左下
    outResizeWParam := SC_wParam_FormResize_BottomLeft;
    cr := crSizeNESW;
  end
  else if (rpLeft in rps) and (X < 5) then begin // 左
    outResizeWParam := SC_wParam_FormResize_Left;
    cr := crSizeWE;
  end
  else if (rpUp in rps) and (Y < 5) then begin // 上
    outResizeWParam := SC_wParam_FormResize_Top;
    cr := crSizeNS;
  end
  else if (rpRight in rps) and (X > nWidth - 5) then begin // 右
    outResizeWParam := SC_wParam_FormResize_Right;
    cr := crSizeWE;
  end
  else if (rpDown in rps) and (Y > nHeight - 5) then begin // 下
    outResizeWParam := SC_wParam_FormResize_Bottom;
    cr := crSizeNS;
  end
  else
    Result := False;

  ourCursor := cr;
end;

procedure TFormPositionEvents.TopMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  resizeParam: Integer;
  cr: TCursor;
begin
  if ResizePositionTest((Sender as TControl).BoundsRect, X, Y, resizeParam, cr,
    [rpLeft, rpRight, rpUp, rpLeftUp, rpRightUp]) then begin
    ReleaseCapture;
    FForm.Perform(WM_SYSCOMMAND, resizeParam, 0);
  end else begin
    ReleaseCapture;
    FForm.PerForm(WM_SYSCOMMAND,SC_wParam_FormMove,0);
  end;
end;

procedure TFormPositionEvents.TopMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var
  n: Integer;
  cr: TCursor;
begin
  ResizePositionTest((Sender as TControl).BoundsRect, X, Y, n, cr,
    [rpLeft, rpRight, rpUp, rpLeftUp, rpRightUp]);
  (Sender as TControl).Cursor := cr;
end;

procedure TFormPositionEvents.BottomMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  resizeParam: Integer;
  cr: TCursor;
begin
  if ResizePositionTest((Sender as TControl).BoundsRect, X, Y, resizeParam, cr,
    [rpLeft, rpRight, rpDown, rpLeftDown, rpRightDown]) then begin
    ReleaseCapture;
    FForm.Perform(WM_SYSCOMMAND, resizeParam, 0);
  end else begin
    ReleaseCapture;
    FForm.PerForm(WM_SYSCOMMAND,SC_wParam_FormMove,0);
  end;
end;

procedure TFormPositionEvents.BottomMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var
  n: Integer;
  cr: TCursor;
begin
  ResizePositionTest((Sender as TControl).BoundsRect, X, Y, n, cr,
    [rpLeft, rpRight, rpDown, rpLeftDown, rpRightDown]);
  (Sender as TControl).Cursor := cr;
end;

procedure TFormPositionEvents.ClientMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  resizeParam: Integer;
  cr: TCursor;
begin
  if ResizePositionTest((Sender as TControl).BoundsRect, X, Y, resizeParam, cr,
    [rpLeft, rpRight]) then begin
    ReleaseCapture;
    FForm.Perform(WM_SYSCOMMAND, resizeParam, 0);
  end else begin
    ReleaseCapture;
    FForm.PerForm(WM_SYSCOMMAND,SC_wParam_FormMove,0);
  end;
end;

procedure TFormPositionEvents.ClientMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var
  n: Integer;
  cr: TCursor;
begin
  ResizePositionTest((Sender as TControl).BoundsRect, X, Y, n, cr,
    [rpLeft, rpRight]);
  (Sender as TControl).Cursor := cr;
end;

procedure TFormPositionEvents.InnerMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  resizeParam: Integer;
  cr: TCursor;
begin
  if ResizePositionTest((Sender as TControl).BoundsRect, X, Y, resizeParam, cr,
    []) then begin
    ReleaseCapture;
    FForm.Perform(WM_SYSCOMMAND, resizeParam, 0);
  end else begin
    ReleaseCapture;
    FForm.PerForm(WM_SYSCOMMAND,SC_wParam_FormMove,0);
  end;
end;

procedure TFormPositionEvents.InnerMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var
  n: Integer;
  cr: TCursor;
begin
  ResizePositionTest((Sender as TControl).BoundsRect, X, Y, n, cr,
    []);
//  (Sender as TControl).Cursor := cr;
end;

procedure TFormPositionEvents.LeftMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  resizeParam: Integer;
  cr: TCursor;
begin
  if ResizePositionTest((Sender as TControl).BoundsRect, X, Y, resizeParam, cr,
    [rpLeft]) then begin
    ReleaseCapture;
    FForm.Perform(WM_SYSCOMMAND, resizeParam, 0);
  end else begin
    ReleaseCapture;
    FForm.PerForm(WM_SYSCOMMAND,SC_wParam_FormMove,0);
  end;
end;

procedure TFormPositionEvents.LeftMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var
  n: Integer;
  cr: TCursor;
begin
  ResizePositionTest((Sender as TControl).BoundsRect, X, Y, n, cr,
    [rpLeft]);
  (Sender as TControl).Cursor := cr;
end;

procedure TFormPositionEvents.RightMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  resizeParam: Integer;
  cr: TCursor;
begin
  if ResizePositionTest((Sender as TControl).BoundsRect, X, Y, resizeParam, cr,
    [rpRight]) then begin
    ReleaseCapture;
    FForm.Perform(WM_SYSCOMMAND, resizeParam, 0);
  end else begin
    ReleaseCapture;
    FForm.PerForm(WM_SYSCOMMAND,SC_wParam_FormMove,0);
  end;
end;

procedure TFormPositionEvents.RightMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var
  n: Integer;
  cr: TCursor;
begin
  ResizePositionTest((Sender as TControl).BoundsRect, X, Y, n, cr,
    [ rpRight]);
  (Sender as TControl).Cursor := cr;
end;

{ TImageButtonEvents }

constructor TImageButtonEvents.Create(imgBtn: TImage; form: TCustomForm);
begin
  FForm := form;
  FImgBtn := imgBtn;
  FImgBtn.OnMouseDown := Self.ImageMouseDown;
  FImgBtn.OnMouseEnter := Self.ImageMouseEnter;
  FImgBtn.OnMouseLeave := Self.ImageMouseLeave;
  FMouseDownPicture := TPicture.Create;
  FMouseEnterPicture := TPicture.Create;
  FMouseLeavePicture := TPicture.Create;
end;

destructor TImageButtonEvents.Destroy;
begin
  FImgBtn.OnMouseDown := nil;
  FImgBtn.OnMouseEnter := nil;
  FImgBtn.OnMouseLeave := nil;
  FMouseDownPicture.Free;
  FMouseEnterPicture.Free;
  FMouseLeavePicture.Free;
  inherited;
end;

procedure TImageButtonEvents.ImgCloseClick(Sender: TObject);
begin
  if Assigned(FForm) then
    FForm.Close;
end;

procedure TImageButtonEvents.ImgMinClick(Sender: TObject);
begin
  if Assigned(FForm) then
    FForm.WindowState := wsMinimized;
end;

procedure TImageButtonEvents.ImageMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Assigned(MouseDownPicture.Graphic) then
    FImgBtn.Picture.Assign(MouseDownPicture);
end;

procedure TImageButtonEvents.ImageMouseEnter(Sender: TObject);
begin
  if CursorHoverHand
     and (FImgBtn.Cursor = crDefault) then
    FImgBtn.Cursor := crHandPoint;
  if Assigned(MouseEnterPicture.Graphic) then
    FImgBtn.Picture.Assign(MouseEnterPicture);
end;

procedure TImageButtonEvents.ImageMouseLeave(Sender: TObject);
begin
  if CursorHoverHand
     and (FImgBtn.Cursor = crHandPoint) then
    FImgBtn.Cursor := crDefault;
  if Assigned(MouseLeavePicture.Graphic) then
    FImgBtn.Picture.Assign(MouseLeavePicture);
end;

{ TLabelLinkEvents }

constructor TLabelLinkEvents.Create(lbl: TLabel);
begin
  FLabel := lbl;
  FLabel.Font.Color := clHighlight;
  FLabel.OnMouseEnter := LabelMouseEnter;
  FLabel.OnMouseLeave := LabelMouseLeave;
end;

destructor TLabelLinkEvents.Destroy;
begin

  inherited;
end;

procedure TLabelLinkEvents.LabelMouseEnter(Sender: TObject);
begin
  if CursorHoverHand
     and (FLabel.Cursor = crDefault) then
    FLabel.Cursor := crHandPoint;
  FLabel.Font.Style := FLabel.Font.Style + [fsUnderline];
end;

procedure TLabelLinkEvents.LabelMouseLeave(Sender: TObject);
begin
  if CursorHoverHand
     and (FLabel.Cursor = crHandPoint) then
    FLabel.Cursor := crDefault;
  FLabel.Font.Style := FLabel.Font.Style - [fsUnderline];
end;

end.
