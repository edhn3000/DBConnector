unit U_WaitingForm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, U_CommonFunc, ExtCtrls, U_MyProgBar;

type
  TWaitingForm = class(TForm)
    lblHint: TLabel;
    pnl1: TPanel;
    lblTitle: TLabel;
    btnCancel: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
  private
    { Private declarations }
    FProgBar: TMyProgbar;
    FCanceled: Boolean;  //ÍË³ö±êÖ¾

    function GetHintMsg: string;
    procedure SetHintMsg(value: string);
    function GetTitleMsg: string;
    procedure SetTitleMsg(value: string);
    procedure AdjustLayout;  
  public
    property ProgBar: TMyProgbar read FProgBar;
    property HintMsg: string read GetHintMsg write SetHintMsg;
    property TitleMsg: string read GetTitleMsg write SetTitleMsg;
    property Canceled: Boolean read FCanceled write FCanceled;
  public
    procedure AddProgBarPercent(fspercent: Single; MaxTo: Single);
    { Public declarations }
  end;
  
var
  WaitingForm: TWaitingForm;

implementation

{$R *.dfm}

procedure TWaitingForm.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TWaitingForm.FormDestroy(Sender: TObject);
begin
  WaitingForm := nil;
end;

procedure TWaitingForm.FormCreate(Sender: TObject);
begin
  lblTitle.Caption := '';
  lblHint.Caption := '';
  FProgBar := TMyProgbar.Create(Self);
  with FProgBar do
  begin     
    Init(pnl1, False);
    AdjustRect(pnl1.ClientRect);
    Visible := True;
  end;
  FCanceled := False;
end;

procedure TWaitingForm.AddProgBarPercent(fspercent: Single; MaxTo: Single);
begin
  ProgBar.AddPercent(fspercent, MaxTo);
end;

function TWaitingForm.GetHintMsg: string;
begin
  Result := lblHint.Caption;
end;

procedure TWaitingForm.SetHintMsg(value: string);
begin
  lblHint.Caption := value;
  Application.ProcessMessages;
end;

function TWaitingForm.GetTitleMsg: string;
begin
  Result := lblTitle.Caption;
end;

procedure TWaitingForm.SetTitleMsg(value: string);
begin
  lblTitle.AutoSize := True;
  lblTitle.Caption := value;
  AdjustLayout;
  Application.ProcessMessages;
end;

procedure TWaitingForm.AdjustLayout;
const
  TITLE_MARGIN = 24;
  PNL_MARGIN = 16;
var
  titlewidth: Integer;
begin
  titlewidth := Canvas.TextWidth(lblTitle.Caption);

  if titlewidth > Width then
  begin
    lblTitle.Left := TITLE_MARGIN;
    Width := lblTitle.Left + lblTitle.Width + TITLE_MARGIN
  end
  else
  begin
    lblTitle.Left := Trunc((Width - titlewidth)/2);
  end;
  pnl1.Left := PNL_MARGIN;
  pnl1.Width := Width - PNL_MARGIN*2;
  ProgBar.AdjustRect(pnl1.ClientRect);  

  lblTitle.Top := 8;
  lblHint.Top := lblTitle.Top*2 + lblTitle.Height;
  pnl1.Top := lblHint.Top + lblHint.Height + 5;
  btnCancel.Top := pnl1.Top + pnl1.Height + 5;
  btnCancel.Left := Trunc((Width - btnCancel.Width)/2);
  Height := pnl1.Top + pnl1.Height + 40;
end;

procedure TWaitingForm.btnCancelClick(Sender: TObject);
begin
  FCanceled := True;
  btnCancel.Enabled := False;
end;

end.
