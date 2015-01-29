unit U_fListView;

interface

uses
  SysUtils, Classes, Controls, ComCtrls, Windows, CommCtrl, Variants;

type
  TfListView = class(TListView)
  private
    img: TImageList;
    FCheckFalseIndex: Integer;
    FCheckTrueIndex : Integer;
    FIsCheckList: Boolean;
    { Private declarations }
  protected
    { Protected declarations } 
    procedure DoStartDrag(var DragObject: TDragObject);override;
    function CustomDrawItem(Item: TListItem; State: TCustomDrawState;
      Stage: TCustomDrawStage): Boolean;override;
  public
    { Public declarations } 
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure AddOneItem(s: string);
  published
    { Published declarations }
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('DBTool', [TfListView]);
end;

{ TfListView }

procedure TfListView.AddOneItem(s: string);
begin
  with Items.Add do
  begin
  Caption := s;
  ImageIndex := FCheckTrueIndex;
  end;
end;

constructor TfListView.Create(AOwner: TComponent);
var
  icon: HIcon;
begin
  inherited;
  FIsCheckList := True;
  img := TImageList.Create(Self);     
  icon := LoadIcon(HInstance, 'check_false_icon');
  FCheckFalseIndex := ImageList_AddIcon(img.Handle, icon);
  icon := LoadIcon(HInstance, 'check_true_icon');
  FCheckTrueIndex  := ImageList_AddIcon(img.Handle, icon);
  SmallImages := img;
end;

function TfListView.CustomDrawItem(Item: TListItem;
  State: TCustomDrawState; Stage: TCustomDrawStage): Boolean;
begin
  Result := inherited CustomDrawItem(Item, State, Stage);
  if FIsCheckList then
    item.ImageIndex := FCheckTrueIndex;
end;

destructor TfListView.Destroy;
begin
  img.Free;
  inherited;
end;

procedure TfListView.DoStartDrag(var DragObject: TDragObject);
begin
  inherited;
end;

end.
