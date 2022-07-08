{
 @author  fengyq
 @comment Word助手单元
 @version 1.0
 @version 2015/11/09
}
unit U_WordHelper;

interface

uses
  Windows, SysUtils, Variants, Classes, Controls, Forms,
  ComCtrls, ComObj, ActiveX, Word_TLB, Office_TLB, U_OfficeHelper;

type
  { TWordHelper }
  TWordHelper = class(TOfficeHelper)
  private

  protected

  protected
    function InitApplication: Boolean; override;

  public
    constructor Create;

    destructor Destroy;override;

    // 打开Word文档
    procedure NewFile(sFileName: String); override;
    procedure OpenFile(sFileName: String; readOnly: Boolean);overload; override;
    // 关闭Word文档
    procedure CloseFile(bSave: Boolean = false); override;

    function GetDocumentWindow(doc: IDispatch): IDispatch; override;

    function ExecuteControl(controlType: TOleEnum; tag: String): Boolean; override;
  end;

implementation


{ TWordHelper }

constructor TWordHelper.Create;
begin
  inherited Create;
end;

destructor TWordHelper.Destroy;
begin

  inherited;
end;

function TWordHelper.InitApplication: Boolean;
begin
  if FIsAppInited then begin
    Result := True;
    Exit;
  end;

  FIsInstalled := False;
  FIsAppInited := CreateApplication(FUseActiveApp, 'Word.Application');
  if FIsAppInited then begin
    // 不保存到Normal.dotm，避免文档被重复打开后在关闭时报错
    FApplication.Options.SaveNormalPrompt := False;
  //  FWordApp := WordApplication(IDispatch(FApplication));
  //  if not CheckWordApp then
  //  begin
  //    //
  //  end;
    FIsInstalled := True;
    FIsAppInited := True;
    FCloseOnFree := not FUseActiveApp;
  end;
  Result := FIsAppInited;
end;

procedure TWordHelper.NewFile(sFileName: String);
begin
  InitApplication;

  FFileName := sFileName;
  FApplication.Documents.Add;
  FDocument := FApplication.ActiveDocument;
end;

procedure TWordHelper.OpenFile(sFileName: String; readOnly: Boolean);
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

function TWordHelper.GetDocumentWindow(doc: IDispatch): IDispatch;
var
  i: Integer;
  wndIndex: OleVariant;
  d: Word_TLB._Document;
  wnd: Word_TLB.Window;
begin
  d := (doc as Word_TLB._Document);
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
