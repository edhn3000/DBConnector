unit UF_About;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, CnWaterImage, jpeg;

type
  TF_About = class(TForm)
    btnOk: TButton;
    lblWinVersion: TLabel;
    lblSoftVersion: TLabel;
    bvl1: TBevel;
    lblDescription: TLabel;
    lblURL: TLabel;
    lbl1: TLabel;
    lblLianxi: TLabel;
    CnWaterImage1: TCnWaterImage;
    procedure FormShow(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure lblURLClick(Sender: TObject);
    procedure lblURLMouseEnter(Sender: TObject);
    procedure lblURLMouseLeave(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private  
  public
  end;

implementation

uses
  U_CommonFunc, U_FileUtil, U_Const;

{$R *.dfm} 

procedure TF_About.FormShow(Sender: TObject);
var
  a, b, c: string;
begin
  commonFunc.GetWinVersionInfo(a, b, c);
  lblWinVersion.Caption := '²Ù×÷ÏµÍ³£º' + #13 + a + #13 + b + #13 + c;
end;

procedure TF_About.btnOkClick(Sender: TObject);
begin
//  tmr2.Enabled := True;   
  Close;
end;

procedure TF_About.lblURLClick(Sender: TObject);
begin
  WinExec(PAnsiChar('mailto:edhn@163.com'), SW_SHOW);
//  ShellExecute(Handle,PChar('mailto:edhn@163.com'), '', '', '', SW_SHOWNOACTIVATE);
end;

procedure TF_About.lblURLMouseEnter(Sender: TObject);
begin
  lblURL.Font.Color := clRed;
  Cursor := crHandPoint;
end;

procedure TF_About.lblURLMouseLeave(Sender: TObject);
begin
  lblURL.Font.Color := clBlue;
  Cursor := crDefault;
end;

procedure TF_About.FormCreate(Sender: TObject);
begin
  lblSoftVersion.Caption := gC_AppName + ' ' +
    GetFileVersion(Application.ExeName);
end;

end.
