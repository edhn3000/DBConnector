unit U_WaitingForm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ActiveX, GIFImg, ExtCtrls, StdCtrls;

type
  TThreadProc = procedure of object;
  TTaskThread = class(TThread)
  private
    Fhandle: HWND;
  public
    sWritContent: string;
    sCheckSetting: string;
    result: string;
    Proc: TThreadProc;
    constructor Create(CreateSuspended: Boolean;formhandle: HWND);overload;
    procedure Execute; override;
  end;

  { TWaitingForm }
  TWaitingForm = class(TForm)
    Image1: TImage;
    lblHint: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    FThread: TTaskThread;
    procedure SetMesasge(s: String);
    procedure StartThread(AProc: TThreadProc);
  end;

implementation

{$R *.dfm}

{ TTaskThread }

constructor TTaskThread.Create(CreateSuspended: Boolean; formhandle: HWND);
begin
  Fhandle := formhandle;
  inherited Create(CreateSuspended);
end;

procedure TTaskThread.Execute;
begin
  inherited;
  if Assigned(Proc) then begin
    Proc;
  end;
  PostMessage(Fhandle,WM_CLOSE,0,0);
end;

{ TTThreadForm }

procedure TWaitingForm.FormCreate(Sender: TObject);
begin
  FThread := TTaskThread.Create(True,Handle);
end;

procedure TWaitingForm.FormDestroy(Sender: TObject);
begin
  if Assigned(FThread) and (not FThread.Terminated) then
    FThread.Terminate;
  FreeAndNil(FThread);
end;

procedure TWaitingForm.FormShow(Sender: TObject);
begin
  TGIFImage(image1.Picture.Graphic).AnimationSpeed := 120;
  TGIFImage(Image1.Picture.Graphic).Animate := True;
end;

procedure TWaitingForm.SetMesasge(s: String);
begin
  lblHint.Caption := s;
  lblHint.Left := Round((Self.Width - lblHint.Width)/2);
end;

procedure TWaitingForm.startThread(AProc: TThreadProc);
begin
  FThread.Proc := AProc;
  FThread.Start;
end;

end.
