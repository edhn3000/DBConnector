{
 @author  fengyq
 @comment Wps助手单元
 @version 1.0
 @version 2015/11/09
}
unit U_WPSHelper;

interface

uses
  SysUtils, Variants, Classes, Controls, Forms, Windows,
  ComCtrls, ComObj, ActiveX, WPS_TLB, Office_TLB, U_OfficeHelper;

type
  { TWpsAPIVer }
  TWpsAPIVer = (wavAuto, wavV8, wavV9);
  { TWpsHelper }
  TWpsHelper = class(TOfficeHelper)
  private
    FApiVer: TWpsAPIVer;

  protected

    function InitApplication: Boolean; override;

  public
    constructor Create;

    // 打开Wps文档
    procedure NewFile(sFileName: String); override;
    procedure OpenFile(sFileName: String; readOnly: Boolean);overload; override;
    // 关闭Wps文档
    procedure CloseFile(bSave: Boolean = false); override;

    function GetDocumentWindow(doc: IDispatch): IDispatch; override;

    function ExecuteControl(controlType: TOleEnum; tag: String): Boolean; override;

  end;

implementation

{ TWpsHelper }

constructor TWpsHelper.Create;
begin
  inherited Create;
end;

function TWpsHelper.InitApplication: Boolean;
begin
  if FIsAppInited then begin
    Result := True;
    Exit;
  end;

  FIsInstalled := False;
  FIsAppInited := False;
  if FApiVer = wavV9 then begin
    FIsAppInited := CreateApplication(FUseActiveApp, 'Kwps.Application');
    if not FIsAppInited then
      FIsAppInited := CreateApplication(FUseActiveApp, 'Word.Application');
  end else if FApiVer = wavV8 then begin
    FIsAppInited := CreateApplication(FUseActiveApp, 'Wps.Application');
  end else begin
    FIsAppInited := CreateApplication(FUseActiveApp, 'Kwps.Application');
    if not FIsAppInited then
      FIsAppInited := CreateApplication(FUseActiveApp, 'Word.Application');
    if not FIsAppInited then
      FIsAppInited := CreateApplication(FUseActiveApp, 'Wps.Application');
  end;

  if FIsAppInited then begin
    FIsInstalled := True;
    FIsAppInited := True;
    FCloseOnFree := not FUseActiveApp;
  end;
  Result := FIsAppInited;
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
  InitApplication;

  FFileName := sFileName;
  FApplication.Documents.Add;
  FDocument := FApplication.ActiveDocument;
end;

procedure TWpsHelper.OpenFile(sFileName: String; readOnly: Boolean);
begin
  InitApplication;

  FFileName := sFileName;
  if readOnly then
    FApplication.Documents.Open(sFileName, False, True)
  else
    FApplication.Documents.Open(sFileName);
  FApplication.Options.SaveNormalPrompt := False;
  FDocument := FApplication.ActiveDocument;
  FWindow := FApplication.ActiveWindow;
  FOpend := True;
end;

function TWpsHelper.GetDocumentWindow(doc: IDispatch): IDispatch;
var
  i: Integer;
  wndIndex: OleVariant;
  d: WPS_TLB._Document;
  wnd: WPS_TLB.Window;
begin
  d := (doc as WPS_TLB._Document);
  Result := nil;
  if d.Windows.Count = 1 then begin
    wndIndex := 1;
    Result := d.Windows.Item(wndIndex);
  end else begin
    for i := 1 to d.Windows.Count do begin
      wndIndex := i;
      wnd := d.Windows.Item(wndIndex);
      if wnd.Document = doc then begin
        Result := wnd;
        System.Break;
      end;
    end;
  end;
end;

function TWpsHelper.ExecuteControl(controlType: TOleEnum; tag: String): Boolean;
var
  cmdBar: OleVariant;
  cmdCtrl: OleVariant;
begin
  Result := False;
  cmdBar := FApplication.CommandBars.Item['Standard'];
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

end.
