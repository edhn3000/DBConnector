unit U_CodeComboBox;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Variants, Forms, StdCtrls,
  ComCtrls, ExtCtrls;

type
  TCodeGetValueMode = (cvmCode=1, cvmName=2, cvmCodeAndName=3);

  TCodeComboBox = class(TCustomEdit)
  private
    FList: TListBox;
    FShowMode: TCodeGetValueMode;
    FValueMode: TCodeGetValueMode;
//    FKeys: TStrings;
//    FValues: TStrings;
    { Private declarations }
    procedure AdjustListPosition;
  protected
    { Protected declarations }
  public
    property ShowMode: TCodeGetValueMode read FShowMode write FShowMode;  
    property ValueMode: TCodeGetValueMode read FValueMode write FValueMode;
  public
    { Public declarations }
    constructor Create(AOwner: TComponent); override;
    destructor Destroy;override;

//    procedure DoExit;override;
//    procedure KeyDown(var Key: Word; Shift: TShiftState);
    procedure DropDown;
    procedure CloseDrop;    
  published
    { Published declarations } 
    property Anchors;
    property AutoSelect;
    property AutoSize;
    property BevelEdges;
    property BevelInner;
    property BevelKind default bkNone;
    property BevelOuter;
    property BiDiMode;
    property BorderStyle;
    property CharCase;
    property Color;
    property Constraints;
    property Ctl3D;
    property DragCursor;
    property DragKind;
    property DragMode;
    property Enabled;
    property Font;
    property HideSelection;
    property ImeMode;
    property ImeName;
    property MaxLength;
    property OEMConvert;
    property ParentBiDiMode;
    property ParentColor;
    property ParentCtl3D;
    property ParentFont;
    property ParentShowHint;
    property PasswordChar;
    property PopupMenu;
    property ReadOnly;
    property ShowHint;
    property TabOrder;
    property TabStop;
    property Text;
    property Visible;
    property OnChange;
    property OnClick;
    property OnContextPopup;
    property OnDblClick;
    property OnDragDrop;
    property OnDragOver;
    property OnEndDock;
    property OnEndDrag;
    property OnEnter;
    property OnExit;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property OnMouseDown;
    property OnMouseMove;
    property OnMouseUp;
    property OnStartDock;
    property OnStartDrag;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('Samples', [TCodeComboBox]);
end;

{ TfComboBox }  

constructor TCodeComboBox.Create(AOwner: TComponent);
begin
  inherited;
  FList := TListBox.Create(AOwner);
  FList.Visible := False;
  AdjustListPosition;
  FList.Width := Self.Width;
end;  

destructor TCodeComboBox.Destroy;
begin
  FList.Free;
  inherited;
end;

procedure TCodeComboBox.AdjustListPosition;
begin
  FList.Top := Self.Top + Self.Height;
  FList.Left := Self.Left;
end;

procedure TCodeComboBox.CloseDrop;
begin
  FList.Visible := False;
end;

procedure TCodeComboBox.DropDown;
begin
  AdjustListPosition;
  FList.Visible := True;
end;

end.
