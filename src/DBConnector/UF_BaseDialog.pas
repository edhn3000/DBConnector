unit UF_BaseDialog;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, jpeg;

type
  TF_BaseDialog = class(TForm)
    panelTitleArea: TPanel;
    panelCtlArea: TPanel;
    Bevel1: TBevel;
    Bevel2: TBevel;
    panelMainArea: TPanel;
    lblMainDescription: TLabel;
    lblSubDescription: TLabel;
    Image1: TImage;
    labAppTitle: TLabel;
    labVer: TLabel;
    btnCloseMe: TButton;
    btnOK: TButton;
    labTips: TLabel;
    procedure FormShow(Sender: TObject);
    procedure btnCloseMeClick(Sender: TObject);
  private
  protected
    { Protected declarations }
    procedure ChangeTip(sMsg:string);
    procedure ChangeMainDescription(s: string);
    procedure ChangeSubDescription(s: string);
  public
    { Public declarations }
  end;
  

implementation

{$R *.dfm}

uses
  U_FileUtil;

procedure TF_BaseDialog.FormShow(Sender: TObject);
begin
  self.Caption:=Application.Title;
  self.labAppTitle.Caption := Application.Title;
  ChangeMainDescription('');
  ChangeSubDescription('');
  ChangeTip('');  
  self.labVer.Caption:=GetFileMainVersion(Application.ExeName);
end;

procedure TF_BaseDialog.btnCloseMeClick(Sender: TObject);
begin
  Close;
end; 

procedure TF_BaseDialog.ChangeMainDescription(s: string);
begin
  Self.lblMainDescription.Caption:=s;
end;

procedure TF_BaseDialog.ChangeSubDescription(s: string);
begin      
  Self.lblSubDescription.Caption:=s;
end;

procedure TF_BaseDialog.ChangeTip(sMsg: string);
begin 
  self.labTips.Caption:=sMsg;
end;

end.
