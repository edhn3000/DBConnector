unit U_fHintTree;
interface

uses
  Windows, Messages, SysUtils, Classes, Controls, ComCtrls, Graphics, ExtCtrls;

type
  TShowHintEvent = procedure( Sender : TObject; Node : TTreeNode ) of Object;

  TfHintTree = class(TTreeView)
  private
    { Private declarations }
    FPreText : String;
    FPoint : TPoint;
    FHintBgColor : TColor;
    FHintRemainTime : Cardinal;
    FMaxWidth : Integer;
    FOnShowHint : TShowHintEvent;
    Hinter : THintWindow;
    HintTimer : TTimer;
    procedure FOnTimer( Sender : TObject );
  protected
    procedure CMHintShow;
    procedure CMMouseLeave;
    procedure SetHintBgColor( Value : TColor );
    procedure SetHintRemainTime( Value : Cardinal );
    procedure DoExit;override;
    procedure Resize;override;
    procedure WndProc(var Message : TMessage); override;
    { Protected declarations }
  public
    { Public declarations }
    constructor Create(AOwner: TComponent); override;
    procedure ShowHint( sHint : String );
  published
    { Published declarations }
    property HintMaxWidth : Integer read FMaxWidth write FMaxWidth default 500;
    property HintBgColor : TColor read FHintBgColor write SetHintBgColor;
    property HintRemainTime : Cardinal read FHintRemainTime write SetHintRemainTime;
    property OnShowHint: TShowHintEvent read FOnShowHint write FOnShowHint;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('DBTool', [TfHintTree]);
end;

constructor TfHintTree.Create(AOwner: TComponent);
begin
  inherited Create( AOwner );
  Hinter := THintWindow.Create( self );
  SetHintBgColor( clInfoBk );
  HintTimer := TTimer.Create( self );
  HintTimer.Enabled := false;
  HintTimer.OnTimer := FOnTimer;
  SetHintRemainTime( 5000 );
  FMaxWidth := 500;

  FPreText := '';
end;

procedure TfHintTree.SetHintBgColor( Value : TColor );
begin
  FHintBgColor := Value;
  Hinter.Canvas.Brush.Color := Value;
  Hinter.Color := Value;
end;

procedure TfHintTree.SetHintRemainTime( Value : Cardinal );
begin
  FHintRemainTime := Value;
  HintTimer.Interval := Value;
end;

procedure TfHintTree.WndProc(var Message : TMessage);
begin
  case Message.Msg of
    CM_HINTSHOW:
    begin
      CMHintShow;
    end;
    CM_MOUSELEAVE:
    begin
      CMMouseLeave;
    end
    else
  end;
  inherited WndProc(Message);
end;

procedure TfHintTree.CMHintShow;
var
  Point : TPoint;
  Node : TTreeNode;
begin

  if GetCursorPos( FPoint ) then
  begin
    Point := ScreenToClient( FPoint );
    Node := GetNodeAt( Point.X, Point.Y );
    if Assigned( Node ) then
    begin
      if Node.Text = FPreText then
        Exit
      else
        FPreText := Node.Text;
    end
    else
      FPreText := '';
  end
  else
    Node := nil;

  if Assigned( FOnShowHint ) then FOnShowHint( self, Node );
end;

procedure TfHintTree.CMMouseLeave;
begin
  FPreText := '';
  Hinter.ReleaseHandle;
  HintTimer.Enabled := false;
end;

procedure TfHintTree.ShowHint( sHint : String );
var
  Rect : TRect;
  nWidth, nHeight : Integer;
begin
  if sHint <> '' then
  begin
    Rect := Hinter.CalcHintRect( FMaxWidth, sHint, nil );
    nWidth := Rect.Right - Rect.Left;
    nHeight := Rect.Bottom - Rect.Top;
    Rect.Top := FPoint.Y + 20;
    Rect.Left := FPoint.X;
    Rect.Right := Rect.Left + nWidth;
    Rect.Bottom := Rect.Top + nHeight;

    Hinter.ActivateHint( Rect, sHint );
    if FHintRemainTime <> 0 then
      HintTimer.Enabled := true;
  end
  else
  begin
    FPreText := '';
    Hinter.ReleaseHandle;
    HintTimer.Enabled := false;
  end;
end;

procedure TfHintTree.FOnTimer( Sender : TObject );
begin
  Hinter.ReleaseHandle;
  HintTimer.Enabled := false;
end;

procedure TfHintTree.DoExit;
begin
  inherited;
  FPreText := '';
  Hinter.ReleaseHandle;
  HintTimer.Enabled := false;
end;

procedure TfHintTree.Resize;
begin
  inherited;
  FPreText := '';
  Hinter.ReleaseHandle;
  HintTimer.Enabled := false;
end;

end.
