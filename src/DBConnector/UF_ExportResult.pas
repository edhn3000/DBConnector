{**
 * <p>Title: U_ExportResult  </p>
 * <p>Copyright: Copyright (c) 2008/4/28</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 导出结果 </p>
 *}
 unit UF_ExportResult;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ShellApi;

type
  TF_ExportResult = class(TForm)
    btnOpenWidhEditor: TButton;
    lbl1: TLabel;
    edtExportFile: TEdit;
    bvl1: TBevel;
    btnOpen: TButton;
    btnOpenDir: TButton;
    procedure btnOpenWidhEditorClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnOpenClick(Sender: TObject);
    procedure btnOpenDirClick(Sender: TObject);
  private
    { Private declarations }
    FsExportedFile: string;
  public
    { Public declarations }
    procedure ShowExportFile(sFileName: string);
  end;

implementation

uses UF_MAIN;

{$R *.dfm}

{ TF_ExportResult }

procedure TF_ExportResult.btnCancelClick(Sender: TObject);
begin
  Close;
end;

procedure TF_ExportResult.btnOpenWidhEditorClick(Sender: TObject);
begin
  F_MAIN.OpenWithEditor(FsExportedFile);
  Close;
end;

procedure TF_ExportResult.ShowExportFile(sFileName: string);
begin
  FsExportedFile := sFileName;
  edtExportFile.Text := sFileName;
end;

procedure TF_ExportResult.FormShow(Sender: TObject);
begin
  if ExtractFileExt(FsExportedFile) = '.xls' then
    btnOpenWidhEditor.Enabled := False;
end;

procedure TF_ExportResult.btnOpenClick(Sender: TObject);
begin
  ShellExecute(Application.Handle, nil, PChar(FsExportedFile),
               nil, nil, SW_SHOWNOACTIVATE);
end;

procedure TF_ExportResult.btnOpenDirClick(Sender: TObject);
begin   
  ShellExecute(Application.Handle, nil, PChar(ExtractFileDir(FsExportedFile)),
               nil, nil, SW_SHOWNOACTIVATE);
end;

end.
