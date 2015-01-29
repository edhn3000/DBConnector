unit UF_Tip;

interface    

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls;

type
  TF_Tip = class(TForm)
    mmoTip: TMemo;
    pnl1: TPanel;
    bvl1: TBevel;
    btn1: TButton;
    procedure FormShow(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    procedure AdjustTheSize;
  public
    { Public declarations }
  end;

var
  tipForm: TF_Tip;

implementation

{$R *.dfm}

uses
  U_Pub;

procedure TF_Tip.FormShow(Sender: TObject);
var
  sFile: string;
  i: Integer;
begin
  sFile := GetAppRootPath +  ChangeFileExt(ExtractFileName(
    Application.ExeName), '')+ 'Tip.txt';
  if FileExists(sFile) then
    mmoTip.Lines.LoadFromFile(sFile);
  for i := 0 to mmoTip.Lines.Count - 1 do
    mmoTip.Lines[i] := StringReplace(mmoTip.Lines[i], '\n', #13#10,[rfIgnoreCase]);
  AdjustTheSize;
end;

procedure TF_Tip.AdjustTheSize;
begin
  //
end;

procedure TF_Tip.btn1Click(Sender: TObject);
begin
  Close;
end;

procedure TF_Tip.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TF_Tip.FormDestroy(Sender: TObject);
begin
  tipForm := nil;
end;

end.
