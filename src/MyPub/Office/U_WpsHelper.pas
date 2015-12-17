unit U_WPSHelper;

interface

uses
  SysUtils, Variants, Classes, Controls, Forms, Windows,
  ComCtrls, ComObj, ActiveX, WPS_TLB, Office_TLB, U_OfficeHelper;

type
  TWpsHelper = class(TOfficeHelper)
  private

  protected

  public
    constructor Create; overload;
    constructor Create(useActiveApp: Boolean); overload;
    destructor Destroy;override;

    // 打开Wps文档
    procedure NewFile(sFileName: String); override;
    procedure OpenFile(sFileName: String; readOnly: Boolean);overload; override;
    // 关闭Wps文档
    procedure CloseFile(bSave: Boolean = false); override;

    function ExecuteControl(controlType: TOleEnum; tag: String): Boolean; override;
  end;

  TWpsV9Helper = class(TWpsHelper)
  protected

  public
    constructor Create(useActiveApp: Boolean); overload;

  end;


implementation

{ TWpsHelper }

constructor TWpsHelper.Create;
begin
  FOpend := False;
  Create(False);
end;

constructor TWpsHelper.Create(useActiveApp: Boolean);
begin
  FIsInstalled := CreateApplication(useActiveApp, 'Wps.Application');
  FCloseOnFree := not FUseActiveApp;
end;

destructor TWpsHelper.Destroy;
begin
  if FCloseOnFree then
  begin
    try
      if FOpend then
        CloseFile(False);
    except
      on e: Exception do
        OutputDebugString(PChar('TWpsHelper.Destroy close file error！' + e.Message));
    end;
  end;
  try
    if not FUseActiveApp then
      FApplicatioin.Quit(wdDoNotSaveChanges);
  except
    on e: Exception do
      OutputDebugString(PChar('TWpsHelper.Destroy quit application error！' + e.Message));
  end;

  FApplicatioin := Unassigned;
  FDocument := Unassigned;
  inherited;
end;

procedure TWpsHelper.CloseFile(bSave: Boolean);
begin
  if bSave then begin
    FDocument.SaveAs(FFileName);
    FDocument.Close(wdDoNotSaveChanges);
  end else begin
    FDocument.Close(wdDoNotSaveChanges);
  end;
  FOpend := False;
end;

procedure TWpsHelper.NewFile(sFileName: String);
begin
  FFileName := sFileName;
  FApplicatioin.Documents.Add;
//  FWordApp := WordApplication(IDispatch(FApplicatioin));
  FDocument := FApplicatioin.ActiveDocument;
end;

procedure TWpsHelper.OpenFile(sFileName: String; readOnly: Boolean);
begin
  FFileName := sFileName;
  if readOnly then
    FApplicatioin.Documents.Open(sFileName, False, True)
  else
    FApplicatioin.Documents.Open(sFileName);
  FApplicatioin.Options.SaveNormalPrompt := False;
//  FWordApp := WordApplication(IDispatch(FApplicatioin));
  FDocument := FApplicatioin.ActiveDocument;
  FWindow := FApplicatioin.ActiveWindow;
  FOpend := True;
end;

function TWpsHelper.ExecuteControl(controlType: TOleEnum; tag: String): Boolean;
var
  cmdBar: OleVariant;
  cmdCtrl: OleVariant;
begin
  Result := False;
  cmdBar := FApplicatioin.CommandBars.Item['Standard'];
  cmdCtrl := cmdBar.FindControl(msoControlButton, 1, tag, 1, true);
  // TODO 这个VarisNull检测常常不好使
  if not VarisNull(cmdCtrl) then
  begin
    try
      cmdCtrl.Execute;
      Result := True;
    except
      Result := False;
    end;
  end;
end;

{ TWpsV9Helper }

constructor TWpsV9Helper.Create(useActiveApp: Boolean);
begin
  // V9版的API，className是Kwps.Application，如配置工具设置了兼容Word，也可以使用Word.Application
  FIsInstalled := CreateApplication(useActiveApp, 'Kwps.Application');
  if not FIsInstalled then begin
    FIsInstalled := CreateApplication(useActiveApp, 'Word.Application');
  end;

  FCloseOnFree := not FUseActiveApp;
end;

end.
