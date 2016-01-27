unit U_WordHelper;

interface

uses
  SysUtils, Variants, Classes, Controls, Forms,
  ComCtrls, ComObj, ActiveX, Word_TLB, Office_TLB, U_OfficeHelper;

type
  { TWordHelper }
  TWordHelper = class(TOfficeHelper)
  private

  protected

  protected

  public
    constructor Create; overload;
    constructor Create(useActiveApp: Boolean);overload;
    destructor Destroy;override;

    // 打开Word文档
    procedure NewFile(sFileName: String); override;
    procedure OpenFile(sFileName: String; readOnly: Boolean);overload; override;
    // 关闭Word文档
    procedure CloseFile(bSave: Boolean = false); override;

    function ExecuteControl(controlType: TOleEnum; tag: String): Boolean; override;

  end;

implementation

uses
  Windows;

{ TWordHelper }

constructor TWordHelper.Create;
begin
  FOpend := False;
  Create(False);
end;

constructor TWordHelper.Create(useActiveApp: Boolean);
begin
  if CreateApplication(useActiveApp, 'Word.Application') then begin
    // 不保存到Normal.dotm，避免文档被重复打开后在关闭时报错
    FApplication.Options.SaveNormalPrompt := False;
  //  FWordApp := WordApplication(IDispatch(FApplication));
  //  if not CheckWordApp then
  //  begin
  //    //
  //  end;
    FIsInstalled := True;
    FCloseOnFree := not FUseActiveApp;
  end;
end;

destructor TWordHelper.Destroy;
begin
  if FCloseOnFree then
  begin
    try
      if FOpend then
        CloseFile(False);
    except
      on e: Exception do
        OutputDebugString(PChar('TWordHelper.Destroy close file error！' + e.Message));
    end;
  end;
  try
    if not FUseActiveApp then begin
      FApplication.NormalTemplate.Saved := True; // 为了不保存NormalTemplate
      FApplication.Quit(wdDoNotSaveChanges);
    end;
  except
    on e: Exception do
      OutputDebugString(PChar('TWordHelper.Destroy quit application error！' + e.Message));
  end;

  FApplication := Unassigned;
  FDocument := Unassigned;
  inherited;
end;

procedure TWordHelper.NewFile(sFileName: String);
begin
  FFileName := sFileName;
  FApplication.Documents.Add;
//  FWordApp := WordApplication(IDispatch(FApplication));
  FDocument := FApplication.ActiveDocument;
end;

procedure TWordHelper.OpenFile(sFileName: String; readOnly: Boolean);
begin
  FFileName := sFileName;
  if readOnly then
    FApplication.Documents.Open(sFileName, False, True)
  else
    FApplication.Documents.Open(sFileName);
  FApplication.Options.SaveNormalPrompt := False;
//  FWordApp := WordApplication(IDispatch(FApplication));
  FDocument := FApplication.ActiveDocument;
  FWindow := FApplication.ActiveWindow;
  FOpend := True;
end;

procedure TWordHelper.CloseFile(bSave: Boolean);
begin
  if bSave then begin
    FDocument.Save;
//    FDocument.SaveAs(FFileName);
    FDocument.Close(wdDoNotSaveChanges);
  end else begin
    FDocument.Saved := True;
    FDocument.Close(wdDoNotSaveChanges);
  end;
  FOpend := False;
end;

function TWordHelper.ExecuteControl(controlType: TOleEnum; tag: String): Boolean;
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
