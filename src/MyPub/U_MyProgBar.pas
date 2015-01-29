unit U_MyProgBar;

interface

uses
  ComCtrls, Classes, Controls, Types;

type
  { TMyProgbar }
  TMyProgbar = class(TProgressBar)
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure Init( AParent: TWinControl;
        ShowOnCreate: Boolean = False);
    procedure SetPosition(APosition:Integer);
    procedure Reset;
    procedure Finish(Hide: Boolean = True);
    procedure AdjustRect(rect: TRect);
    procedure AddPercent(fsPercent: Single; MaxTo: Single = 1);
  end;

implementation

uses
  Forms;

{ TProgbar begin }

constructor TMyProgbar.Create(AOwner: TComponent);
begin
  inherited;

end;

destructor TMyProgbar.Destroy;
begin
  inherited;
end;

procedure TMyProgbar.Init( AParent: TWinControl; ShowOnCreate: Boolean);
begin
  Parent := AParent;
  Visible := ShowOnCreate; // 默认刚创建时不把进度调显示出来
  Min := 0;
  Max := 100;
  Step := 1; 
  SetPosition(0);
end;

procedure TMyProgbar.Reset;
begin
  SetPosition(0);
  Visible := True;
end;

procedure TMyProgbar.Finish(Hide: Boolean);
begin
  SetPosition(Max);
  if Hide then
    Visible := False
end;

procedure TMyProgbar.AdjustRect(rect: TRect);
begin
  Top := rect.Top;
  Left := rect.Left;
  Height := rect.Bottom - rect.Top;
  Width := rect.Right - rect.Left;
end;

procedure TMyProgbar.AddPercent(fsPercent: Single; MaxTo: Single);
var
  nProgPos, TrueMax, nFinalPos: Integer;
begin
  if Max > Round(Max*MaxTo) then
    TrueMax := Round(Max*MaxTo)
  else
    TrueMax := Max;

  nProgPos := Round(fsPercent * TrueMax);
  if Position + nProgPos > TrueMax then
    nFinalPos := TrueMax
  else
    nFinalPos := Position + nProgPos;
  SetPosition(nFinalPos);
end;     

procedure TMyProgbar.SetPosition(APosition: Integer);
begin
  Position := APosition;
  Application.ProcessMessages;
end;

end.
