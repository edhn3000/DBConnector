unit U_About;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls;

type
  TF_About = class(TForm)
    btnOk: TButton;
    lblWinVersion: TLabel;
    lblSoftVersion: TLabel;
    bvl1: TBevel;
    img1: TImage;
    tmr1: TTimer;
    lblDescription: TLabel;
    tmr2: TTimer;
    lblURL: TLabel;
    procedure FormShow(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure tmr1Timer(Sender: TObject);
    procedure tmr2Timer(Sender: TObject);
  private  
    FsSoftName: string;
    FsDescription: string;

    procedure AdjustLayOut;
  public
    procedure InitDescription(SoftName, Description: string);
  end;

implementation

uses
  U_CommonFunc, U_UIUtil;

var
  m,n:word;

{$R *.dfm} 

procedure TF_About.FormShow(Sender: TObject);
var
  a, b, c: string;
  tempico: string;
begin
  commonFunc.GetWinVersionInfo(a, b, c);
  lblWinVersion.Caption := '²Ù×÷ÏµÍ³£º' + #13 + a + #13 + b + #13 + c;

  lblSoftVersion.Caption := '';
  lblDescription.Caption := '';

  if '' <> FsSoftName then
    lblSoftVersion.Caption := FsSoftName;
  if '' <> FsDescription then
     lblDescription.Caption := FsDescription;

  AdjustLayOut;

  tempico := SysConfig.WinTemp + '\about.ico';
  Application.Icon.SaveToFile(tempico);
  UIUtil.LoadPhoto(img1, tempico);
//  img1.Stretch := True;
  DeleteFile(tempico);   
end;

procedure TF_About.btnOkClick(Sender: TObject);
begin
//  tmr2.Enabled := True;
  Close;
end;

procedure TF_About.tmr1Timer(Sender: TObject);
var
  Row, Ht: Word ;
begin
  Ht := (img1.Height + 255) div 276 ;
  for Row := 0 to 255 do
    with img1.Canvas do
    begin
       Brush.Color := RGB(Row, m, n) ;
       FillRect(Rect(0, Row * 1, Width, (Row + 1) * Ht)) ;
    end;
  m := m + 2;
  if m > 170 then m := 0;
  n := n + 5;
  if n > 150 then n := 0;
end;

procedure TF_About.tmr2Timer(Sender: TObject);
begin
  ClientHeight := ClientHeight - 10;
  ClientWidth  := ClientWidth  - 10;
  if ClientHeight <= 0 then
    Close;
end;

procedure TF_About.InitDescription(SoftName, Description: string);
begin
  FsSoftName := SoftName;
  FsDescription := Description;
end;

procedure TF_About.AdjustLayOut;
begin 
  bvl1.Top := lblDescription.Top + lblDescription.Height + 30;
  btnOk.Top := bvl1.Top + bvl1.Height + 10;
  ClientHeight := btnOk.Top + btnOk.Height + 15;
end;

end.
