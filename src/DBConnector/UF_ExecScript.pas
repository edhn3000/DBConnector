unit UF_ExecScript;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, U_Pub, UD_TIPS, U_DBConnectInterface;

type
  TF_ExecScript = class(TForm)
    lbledtScript: TLabeledEdit;
    btnOk: TButton;
    btnChoose: TButton;
    Memo1: TMemo;
    lblCurrLine: TLabel;
    procedure btnChooseClick(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    procedure DBLog(sLog: string);
    procedure OnExecLine(nCurrLine: Integer);
  public
    { Public declarations }
  end;


var
  F_ExecScript: TF_ExecScript;

implementation

{$R *.dfm}

uses
  U_DBConnnectManager, U_ini;

const
  C_nHeight_Init = 150;
  C_nHeight_ShowLog = 350;

procedure TF_ExecScript.btnCancelClick(Sender: TObject);
begin
  Close;
end;

procedure TF_ExecScript.btnChooseClick(Sender: TObject);
begin
  with TOpenDialog.Create(Self) do
  begin
    if lbledtScript.Text <> '' then
      InitialDir := ExtractFileDir(lbledtScript.Text)
    else if GlobalParams.LastScriptFile <> '' then
      InitialDir := ExtractFilePath(GlobalParams.LastScriptFile)
    else
      InitialDir := ExtractFilePath(Application.ExeName);
    if Execute() then
      lbledtScript.Text := FileName;
  end;
end;

procedure TF_ExecScript.btnOkClick(Sender: TObject);
var
  dbConnect: IDBConnect;
begin
  Height := C_nHeight_ShowLog;
  Memo1.Visible := True;
  Memo1.Clear;
  if (lbledtScript.Text = '') then
  begin
    Memo1.Lines.Add('脚本文件为空');
    Exit;
  end;
  if not FileExists(lbledtScript.Text) then
  begin                                    
    Memo1.Lines.Add('脚本文件不存在');
    Exit;
  end;
  
  GlobalParams.LastScriptFile := lbledtScript.Text;
  dbConnect := TDBConnectManager.CreateDBConnect(g_Global.DBConnect.DBEngine.GetDBType);
  btnOk.Enabled := False;
  lblCurrLine.Visible := True;
  lblCurrLine.Caption := '';
  try
    dbConnect.SystemObject := g_Global.DBConnect.SystemObject;
    dbConnect.ShareEngine(g_Global.DBConnect.DBEngine);
    dbConnect.OnLog := DBLog;
    dbConnect.OnExecLineChange := OnExecLine;
    dbConnect.ExecScriptFile(lbledtScript.Text);
    dbConnect.ClearLog;
  finally
    dbConnect := nil;
    btnOk.Enabled := True;    
    lblCurrLine.Visible := False;
  end;
end;

procedure TF_ExecScript.DBLog(sLog: string);
begin
  if Assigned(Memo1) then
  begin
    Memo1.Lines.Add(sLog);
    Application.ProcessMessages;
  end;
end;    

procedure TF_ExecScript.OnExecLine(nCurrLine: Integer);
begin
  lblCurrLine.Caption := Format('正在执行第%d行', [nCurrLine]);
  Application.ProcessMessages;
//  if Assigned(Memo1) then
//  begin
//    Memo1.Lines.Add(sLog);
//    Application.ProcessMessages;
//  end;
end;

procedure TF_ExecScript.FormShow(Sender: TObject);
begin
  Height := C_nHeight_Init;
end;

procedure TF_ExecScript.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  if not btnOk.Enabled then
    Action := caNone
  else
    Action := caFree;
end;

procedure TF_ExecScript.FormDestroy(Sender: TObject);
begin
  F_ExecScript := nil;
end;

end.
