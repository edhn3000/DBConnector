unit UF_ConsoleView;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, U_Const, ExtCtrls;

type
  TF_ConsoleView = class(TForm)
    mmoConsole: TMemo;
    tmrClose: TTimer;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure tmrCloseTimer(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure mmoConsoleKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    FStopClose: Boolean;
    { Private declarations }
  public
    { Public declarations }
    procedure WriteLn(s: string);
    procedure Clear;
    procedure OnDBLog(sLog: string);
    procedure WMUserConsoleRun(var msg: TMsg); message WMUSER_CONSOLE_RUN;
    procedure DelayClosed(nSeconds: Integer);
  end;

var
  g_consoleForm: TF_ConsoleView;

implementation

{$R *.dfm}

uses
  U_DBConnectorInConsole;

var
  nDelayClose: Integer;

{ TF_ConsoleView }

procedure TF_ConsoleView.Clear;
begin
  mmoConsole.Clear;
end;

procedure TF_ConsoleView.WriteLn(s: string);
begin
  mmoConsole.Lines.Add(s);
end;

procedure TF_ConsoleView.DelayClosed(nSeconds: Integer);
begin
  nDelayClose := nSeconds;
  FStopClose := False;      
  WriteLn(Format('%d秒内本窗口将自动关闭，如果不想关闭请按ESC键', [nSeconds]));
  tmrClose.Enabled := True;
end;

procedure TF_ConsoleView.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
//  Action := caFree;
end;

procedure TF_ConsoleView.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_ESCAPE then
  begin
    FStopClose := True;
  end;  
end;

procedure TF_ConsoleView.OnDBLog(sLog: string);
begin
  WriteLn(sLog);
end;

procedure TF_ConsoleView.tmrCloseTimer(Sender: TObject);
begin
  if FStopClose then
  begin
    tmrClose.Enabled := False;     
    WriteLn(Format('停止自动关闭', []));
    Exit;
  end;  
  WriteLn(Format('自动关闭，剩余时间%d', [nDelayClose]));
  if nDelayClose <= 0 then
    Close;
  Dec(nDelayClose);
end;

procedure TF_ConsoleView.WMUserConsoleRun(var msg: TMsg);
begin
  RunInConsole(wcmView);
end;

procedure TF_ConsoleView.FormShow(Sender: TObject);
begin
  PostMessage(Self.Handle, WMUSER_CONSOLE_RUN, 0, 0);
end;

procedure TF_ConsoleView.mmoConsoleKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_ESCAPE then
  begin
    FStopClose := True;
  end;
end;

end.
