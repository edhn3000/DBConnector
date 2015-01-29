unit UF_Rename;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls;

type
  TF_Rename = class(TForm)
    btnOK: TButton;
    LabeledEdit1: TLabeledEdit;
    procedure FormShow(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    OldName: string;
    NewName: string;
  end;

implementation

{$R *.dfm}

procedure TF_Rename.FormShow(Sender: TObject);
begin
  LabeledEdit1.Text := OldName;
end;

procedure TF_Rename.btnOKClick(Sender: TObject);
begin
  NewName := LabeledEdit1.Text;
  ModalResult := mrOk;
end;

end.
