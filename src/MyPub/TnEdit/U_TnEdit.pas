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
    FfsText: Double;                          // edit框中的浮点数字
    FnText : Longint;                         // edit框中的整数数字
    FnDecimalDigits:Integer;                  // 代表小数位数
    FIntMode: TIntMode;

    function GetTextAsFloat: Double;
    function GetDecimalDIgits:Integer;
    function GetTextAsInt: Longint;

    procedure SetTextByFloat(value: Double);
    procedure SetDecimalDIgits(value: Integer);
    procedure SetTextByInt(value: Longint);

  protected
    { Protected declarations }

    // 在edit中按键的事件
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

    // 获取或设置输入的文本转换成的浮点数形式
    property TextAsInt: Longint read GetTextAsInt write SetTextByInt;

    // 获取或设置输入的文本转换成的浮点数形式
    property TextAsFloat : Double read GetTextAsFloat write SetTextByFloat;

    // 获取或设置小数位数,默认是0
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

// KeyPress事件
procedure TnkEdit.KeyPress(var key:char);
begin
  case key of
  // 输入的是数字,文本中有小数点并且小数位数设置和实际的小数位数一致,则忽略输入
  '0'.. '9': if ( Pos( '.', Text) <> 0 ) and
                ( FnDecimalDIgits = Length( Text ) - Pos( DecimalSeparator, Text ) ) then
               key := #0;
  // 输入-号 文本长度不是0时忽略?
  '-': if Length( Text ) <> 0 then
         key := #0;
  #8, #13:;
  // 如果小数位数是0 不允许输入小数点
  '.': if FnDecimalDIgits = 0 then
         Key := #0;
  // 其他的输入字符,如果不是分隔符或者文本中存在分隔符,则忽略
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
    // 输入时没有输小数部分,且规定的小数位数不为0
    if ( Pos( '.', text ) = 0 ) and ( FnDecimalDIgits <> 0 ) then
    begin
      Text := Text + '.';
    end;
    // 输入时末尾是.
    if Pos( '.',Text ) = Length( Text ) then
    begin
      for i := 1 to FnDecimalDIgits do
      begin
        Text := Text + '0';
      end;
    end;
    // 输入小数位数不够
    if FnDecimalDIgits > ( Length( Text ) - Pos( '.', Text ) ) then
    begin
      for i := Length( Text ) - Pos('.', text) to FnDecimalDIgits do
      begin
        Text := Text + '0';
      end;
    end
    // 在已有数字的情况下设置nDecimalDIgits的值减少小数位数时去掉多余的部分
    else if ( FnDecimalDIgits <> 0 ) and
      ( FnDecimalDIgits  < ( Length( Text ) - Pos( '.', Text ) ) ) then
    begin
      Text := Copy( Text, 0, Pos( '.',Text ) + FnDecimalDIgits );
    end
    // 在已有小数的情况下,设置nDecimalDIgits = 0 则做把文本变成整数形式的处理
    else if ( FnDecimalDIgits = 0 ) and ( Pos( '.', Text ) <> 0 ) then
    begin
      Text := Copy( Text, 0, Pos( '.',Text ) - 1 );
    end;
    inherited DoExit;
  end;
end;

end.
