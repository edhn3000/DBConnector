unit U_TnEdit;

interface

uses
  Windows, Messages, Graphics, Forms, Dialogs,
  SysUtils, Classes, Controls, StdCtrls;

type
  TIntMode = (imTrunc, imRound);

  TnkEdit = class(TCustomEdit)
  private
    { Private declarations }
    FfsText: Double;                          // edit���еĸ�������
    FnText : Longint;                         // edit���е���������
    FnDecimalDigits:Integer;                  // ����С��λ��
    FIntMode: TIntMode;

    function GetTextAsFloat: Double;
    function GetDecimalDIgits:Integer;
    function GetTextAsInt: Longint;

    procedure SetTextByFloat(value: Double);
    procedure SetDecimalDIgits(value: Integer);
    procedure SetTextByInt(value: Longint);

  protected
    { Protected declarations }

    // ��edit�а������¼�
    procedure KeyPress(var Key: Char); override;
    procedure DoExit;override;

  public
    { Public declarations }
    constructor Create(AOwner:TComponent); override;

  published
    { Published declarations }
    // TEdit
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

    // ��ȡ������������ı�ת���ɵĸ�������ʽ
    property TextAsInt: Longint read GetTextAsInt write SetTextByInt;

    // ��ȡ������������ı�ת���ɵĸ�������ʽ
    property TextAsFloat : Double read GetTextAsFloat write SetTextByFloat;

    // ��ȡ������С��λ��,Ĭ����0
    property DecimalDigits:Longint read GetDecimalDIgits write SetDecimalDIgits;

    property IntMode: TIntMode read FIntMode write FIntMode;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('TfEdnt', [TnkEdit]);
end;

constructor TnkEdit.Create(AOwner:TComponent);
begin
  inherited Create(AOwner);
  FnDecimalDIgits := 0;
  FIntMode := imRound;
end;

Function Tnkedit.GetDecimalDIgits:Integer;
begin
  result := FnDecimalDIgits;
end;  

procedure Tnkedit.SetDecimalDIgits(Value:Integer);
begin
  if Value >= 0 then
    FnDecimalDIgits := Value
  else
    FnDecimalDIgits := 0;
end;

function TnkEdit.GetTextAsFloat:Double;
begin
  Result := StrToFloatDef(Text, 0.0)
end;

procedure TnkEdit.SetTextByFloat(value:Double);
begin
  FfsText := value;
  Text := FloatToStr( FfsText );
end;

function TnkEdit.GetTextAsInt: Longint;
begin
  if DecimalDigits = 0 then
    Result := StrToInt64Def( Text, 0 )
  else
    Result := Round(GetTextasFloat);
end;

procedure TnkEdit.SetTextByInt(value: Longint);
begin
  FnText := value;
  Text := IntToStr( FnText );
end;

// KeyPress�¼�
procedure TnkEdit.KeyPress(var key:char);
begin
  case key of
  // �����������,�ı�����С���㲢��С��λ�����ú�ʵ�ʵ�С��λ��һ��,���������
  '0'.. '9': if ( Pos( '.', Text) <> 0 ) and
                ( FnDecimalDIgits = Length( Text ) - Pos( DecimalSeparator, Text ) ) then
               key := #0;
  // ����-�� �ı����Ȳ���0ʱ����?
  '-': if Length( Text ) <> 0 then
         key := #0;
  #8, #13:;
  // ���С��λ����0 ����������С����
  '.': if FnDecimalDIgits = 0 then
         Key := #0;
  // �����������ַ�,������Ƿָ��������ı��д��ڷָ���,�����
  else
  if not ( ( key = DecimalSeparator ) and
    ( Pos( DecimalSeparator, Text ) = 0) ) then
    key :=  #0;
  end;
  inherited KeyPress(key);
end;

procedure Tnkedit.DoExit;
var i:Longint;
begin
  if Length( Text ) <> 0 then
    begin
    // ����ʱû����С������,�ҹ涨��С��λ����Ϊ0
    if ( Pos( '.', text ) = 0 ) and ( FnDecimalDIgits <> 0 ) then
    begin
      Text := Text + '.';
    end;
    // ����ʱĩβ��.
    if Pos( '.',Text ) = Length( Text ) then
    begin
      for i := 1 to FnDecimalDIgits do
      begin
        Text := Text + '0';
      end;
    end;
    // ����С��λ������
    if FnDecimalDIgits > ( Length( Text ) - Pos( '.', Text ) ) then
    begin
      for i := Length( Text ) - Pos('.', text) to FnDecimalDIgits do
      begin
        Text := Text + '0';
      end;
    end
    // ���������ֵ����������nDecimalDIgits��ֵ����С��λ��ʱȥ������Ĳ���
    else if ( FnDecimalDIgits <> 0 ) and
      ( FnDecimalDIgits  < ( Length( Text ) - Pos( '.', Text ) ) ) then
    begin
      Text := Copy( Text, 0, Pos( '.',Text ) + FnDecimalDIgits );
    end
    // ������С���������,����nDecimalDIgits = 0 �������ı����������ʽ�Ĵ���
    else if ( FnDecimalDIgits = 0 ) and ( Pos( '.', Text ) <> 0 ) then
    begin
      Text := Copy( Text, 0, Pos( '.',Text ) - 1 );
    end;
    inherited DoExit;
  end;
end;

end.
