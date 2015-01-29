unit UD_TIPS;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, U_ThreadUtil, Buttons;

const
  C_nTipLevel_Info  = 1;
  C_nTipLevel_Warn  = 2;
  C_nTipLevel_Error = 3;

type
  TdlgTips = class(TForm)
    shpTips: TShape;
    lblTips: TLabel;
    img1: TImage;
    tmrclose: TTimer;
    img2: TImage;
    Img3: TImage;
    btnClose: TSpeedButton;
    tmrflash: TTimer;
    lblCaption: TLabel;
    procedure lblTipsMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure tmrcloseTimer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnCloseClick(Sender: TObject);
    procedure tmrflashTimer(Sender: TObject);
  private
    { Private declarations }
    FsMsg: string;
    procedure InitByLevel(nLevel: Integer);
  protected
    procedure SetMsg(sMsg: string);
    procedure CreateParams(var param: TCreateParams);override;
  public
    { Public declarations }
  end;
  TShowTipOption = class
  public 
    tipForm: TdlgTips;
    AutoClose: Boolean;
    Msg: string;
    InitialTop: Integer;
    FinalTop: Integer;
    Flash: Boolean;
  end;
  { TShowTipThread }
  TShowTipThread = class(TThread)
  public
    tipOption: TShowTipOption;
  public
    constructor Create;
    procedure Execute;override;
  end;

  procedure ShowTipDlg(sMsg:String;iLevel:Integer=C_nTipLevel_Info;
    bAutoClose:boolean=true);
  procedure CloseAllTips;
  procedure ShowFlashTip(sMsg:String;iLevel:Integer=C_nTipLevel_Info;
    bAutoClose:boolean=true);
  procedure CloseFlashTip;

var
//  dlgTips: TdlgTips;
  g_TipFormList: TLockStringList;
  MultiTipForm: Boolean = True;

implementation

{$R *.dfm}

uses
  ShellAPI, Math;
var
  m_DisableTips: Boolean = False;  // 关闭自动分析tip位置的操作
//  m_mutexDialog: TMutex;

function getTaskBarHeight: Integer;
var
  tbdata: Tappbardata;
begin
  tbdata.cbsize := sizeof(tappbardata);
  SHappbarmessage(abm_gettaskbarpos,tbdata);
  result := tbdata.rc.Bottom-tbdata.rc.Top;
end;

function GetNewTipOption(sMsg: string; bAutoClose: Boolean; nLevel: Integer): TShowTipOption;
var
  tipForm: TdlgTips;
  nTaskBarHeight: Integer;
  tipOption: TShowTipOption;
  nUsedScreenHeight: Integer;
  optionLast: TShowTipOption;
begin
  nTaskBarHeight := getTaskBarHeight;
  g_TipFormList.Lock;
  try
    if (g_TipFormList.Count = 0) or MultiTipForm then
    begin
      tipForm := TdlgTips.Create(nil);
      tipForm.Top := Screen.Height;//超出屏幕下 使其不被看到
      tipOption := TShowTipOption.Create;
      tipOption.tipForm := tipForm; 
    end
    else
    begin
      tipOption := TShowTipOption(g_tipFormList.Objects[0]);
      tipForm := tipOption.tipForm;
    end;

    // 刚创建时要先修改大小，再修改位置
    tipForm.Width:=180;
    tipForm.Height:=80;

    nUsedScreenHeight := Screen.Height-nTaskBarHeight;

    if (g_TipFormList.Count = 0)
       or (not MultiTipForm)
       then     // 第一个固定位置
    begin
      tipForm.Left:=Screen.Width-tipForm.Width-2;
      tipOption.InitialTop := nUsedScreenHeight;
    end
    else                               // 看前一个的位置而定
    begin
      optionLast := TShowTipOption(g_TipFormList.Objects[g_TipFormList.Count-1]);
      if optionLast.tipForm.Top < tipForm.Height + 2 then // 顶部空间显示不了下一个
      begin 
        tipForm.Left:=optionLast.tipForm.Left-tipForm.Width-2;
        tipOption.InitialTop := nUsedScreenHeight;
      end
      else
      begin
        tipOption.InitialTop := optionLast.tipForm.Top-2;
        tipForm.Left := optionLast.tipForm.Left;
      end;
    end;
    tipOption.FinalTop := tipOption.InitialTop-tipForm.Height-2;

    tipOption.Msg := sMsg;
    tipOption.AutoClose := bAutoClose;
    tipForm.InitByLevel(nLevel);
    if (g_TipFormList.Count = 0) or MultiTipForm then
      g_TipFormList.AddObject(IntToStr(tipForm.Handle), tipOption);
  finally
    g_TipFormList.UnLock;
  end;
  Result := tipOption;
end;

procedure ShowTipDialog(Aoption: TShowTipOption);
var
  thread: TShowTipThread;
begin
  thread := TShowTipThread.Create;
  thread.tipOption := Aoption;
  thread.Resume;
end;

//iLevel:1tips;2warning;3error.
procedure ShowTipDlg(sMsg:String;iLevel:Integer;bAutoClose:boolean);
var
  option: TShowTipOption;
begin
  option := GetNewTipOption(sMsg, bAutoClose, iLevel);
  ShowTipDialog(option);
end;

procedure CloseAllTips;
var
  i: Integer;
  op: TShowTipOption;
begin
  m_DisableTips := True;
  g_TipFormList.Lock;
  try
    for i := g_TipFormList.Count - 1 downto 0 do
    begin
      op := TShowTipOption(g_TipFormList.Objects[i]);
      op.tipForm.tmrflash.Enabled := False;
      op.tipForm.Close;
    end;
  finally
    m_DisableTips := False;
    g_TipFormList.UnLock;
  end;
end;

procedure ShowFlashTip(sMsg:String;iLevel:Integer;bAutoClose:boolean);
var
  option: TShowTipOption;
begin
  option := GetNewTipOption(sMsg, bAutoClose, iLevel);
  option.Flash := True;
  ShowTipDialog(option);
end;

procedure CloseFlashTip;
var
  i: Integer;
  op: TShowTipOption;
begin
  g_TipFormList.Lock;
  try
    for i := g_TipFormList.Count - 1 downto 0 do
    begin
      op := TShowTipOption(g_TipFormList.Objects[i]);
      if op.Flash then
      begin
        op.tipForm.tmrflash.Enabled := False;
        op.tipForm.Close;
      end;
    end;
  finally
    g_TipFormList.UnLock;
  end;
end;  

procedure TdlgTips.lblTipsMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  ShowWindow(Application.Handle, SW_SHOWNORMAL);  // 主窗口最小化时可还原
  Application.MainForm.SetFocus;                  // 激活
end;

procedure TdlgTips.tmrcloseTimer(Sender: TObject);
begin
  Self.Close;
  tmrclose.Enabled := False;
end;

procedure TdlgTips.CreateParams(var param: TCreateParams);
begin
  inherited;
  with param do
  begin
    WndParent := GetDesktopWindow;
  end;  
end;

procedure TdlgTips.FormCreate(Sender: TObject);
begin
  // 提示窗口不在任务栏显示
  SetWindowLong(Self.Handle, GWL_EXSTYLE, WS_EX_TOOLWINDOW);
end;     

procedure TdlgTips.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TdlgTips.FormDestroy(Sender: TObject);   
var
  nIndex: Integer;
  i: Integer;
  option, option2: TShowTipOption;
begin
  g_TipFormList.Lock;
  try
    nIndex := g_TipFormList.IndexOf(IntToStr(Self.Handle));
    if nindex <> - 1 then
    begin
      g_TipFormList.Objects[nIndex].Free;  // Free tipOption, tipForm已经有Action := caFree;
      g_TipFormList.Delete(nIndex);
      if m_DisableTips then Exit;          // 不分析剩余tip位置,因为即将关闭所有tip
      for i := nIndex to g_TipFormList.Count - 1 do
      begin
        option := TShowTipOption(g_TipFormList.Objects[i]); 
        if i = 0 then   // 第一个固定位置
          option.tipForm.Top := Screen.Height-getTaskBarHeight-option.tipForm.Height-2
        else            // 以后的参照前面的位置
        begin
          option2 := TShowTipOption(g_TipFormList.Objects[i-1]);
          if option2.tipForm.Top > option2.tipForm.Height + 2 then  // 顶部剩余空间可以显示一个窗口
          begin
            option.tipForm.Left := option2.tipForm.Left;
            option.tipForm.Top := option2.tipForm.Top-2-option.tipForm.Height;
          end
          else
            option.tipForm.Top := option.tipForm.Top+2+option.tipForm.Height;
        end;
      end;
    end;
  finally
    g_TipFormList.UnLock;
  end;
end;  

procedure TdlgTips.InitByLevel(nLevel: Integer);
begin
  case nLevel of
    C_nTipLevel_Info:
    begin
      Self.lblTips.Font.Color:=  clMaroon;
      Self.img1.Visible:=true;
      Self.img2.Visible:=False;
      Self.Img3.Visible:=false;
      Self.Caption:='提醒';
      lblCaption.Caption := '提醒';
    end;
    C_nTipLevel_Warn:
    begin
      Self.lblTips.Font.Color:= clMaroon;
      Self.img1.Visible:=false;
      Self.img2.Visible:=true;
      Self.Img3.Visible:=false;
      Self.Caption:='警告';
      lblCaption.Caption := '警告';
    end;
    C_nTipLevel_Error:
    begin
      Self.lblTips.Font.Color:=clRed;
      Self.img1.Visible:=false;
      Self.img2.Visible:=False;
      Self.Img3.Visible:=true;
      Self.Caption:='错误';
      lblCaption.Caption := '错误';
    end;
  end;
end;

{ TShowTipThread }

constructor TShowTipThread.Create;
begin
  inherited Create(True);
  FreeOnTerminate := True;
end;

procedure TShowTipThread.Execute;
var
  i: Integer;
begin
  inherited;
  tipOption.tipForm.SetMsg(tipOption.Msg);

  tipOption.tipForm.tmrclose.Enabled := False;
  if not tipOption.tipForm.Showing then
  begin
    tipOption.tipForm.Top := tipOption.InitialTop;
    tipOption.tipForm.Show;
    for i := tipOption.InitialTop-1 downto tipOption.FinalTop do
    begin
      tipOption.tipForm.Top := i;
    end;
  end;
  if (not tipOption.Flash) and tipOption.AutoClose then
    tipOption.tipForm.tmrclose.Enabled := True;
  if tipOption.Flash then
    tipOption.tipForm.tmrflash.Enabled := True;
  
  SetWindowPos(tipOption.tipForm.Handle, HWND_TOPMOST, 0, 0, 0, 0,
    SWP_NOMOVE or SWP_NOSIZE); //在所有窗口最上
end;

procedure TdlgTips.btnCloseClick(Sender: TObject);
begin 
  tmrclose.Enabled := False;
  Self.Close;
end;

procedure TdlgTips.tmrflashTimer(Sender: TObject);
begin
  if lblTips.Caption = FsMsg then
    lblTips.Caption := ''
  else
    lblTips.Caption := FsMsg;
end;

procedure TdlgTips.SetMsg(sMsg: string);
begin
  FsMsg := sMsg;
  lblTips.Caption := sMsg;
  lblTips.Hint := sMsg;
end;

initialization
//  dlgTips:=nil;
  g_TipFormList := TLockStringList.Create;

finalization
//  if Assigned(dlgTips) then
//    FreeAndNil(dlgTips);
  if Assigned(g_TipFormList) then
    FreeAndNil(g_TipFormList);

end.
