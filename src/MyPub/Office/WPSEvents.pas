{ *****************************************************************************
  WARNING: This component file was generated using the EventSinkImp utility.
           The contents of this file will be overwritten everytime EventSinkImp
           is asked to regenerate this sink component.

  NOTE:    When using this component at the same time with the XXX_TLB.pas in
           your Delphi projects, make sure you always put the XXX_TLB unit name
           AFTER this component unit name in the USES clause of the interface
           section of your unit; otherwise you may get interface conflict
           errors from the Delphi compiler.

           EventSinkImp is written by Binh Ly (bly@techvanguards.com)
  *****************************************************************************
  //Sink Classes//
  TWPSOCXEvents
  TWPSApplicationEvents
  TWPSDocumentEvents
}

//SinkUnitName//
unit WPSEvents;

interface

uses
  Windows, ActiveX, Classes, ComObj, OleCtrls, StdVCL, KSO_TLB, WPS_TLB;

type

  TWPSEventsBaseSink = class(TObject, IUnknown, IDispatch)
  protected
    { IUnknown }
    function QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;

    { IDispatch }
    function GetIDsOfNames(const IID: TGUID; Names: Pointer;
      NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; virtual; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; virtual; stdcall;
    function GetTypeInfoCount(out Count: Integer): HResult; virtual; stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; stdcall;
  protected
    FCookie: integer;
    FCP: IConnectionPoint;
    FSinkIID: TGUID;
    FSource: IUnknown;
    function DoInvoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var dps: TDispParams; pDispIds: PDispIdList;
      VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; abstract;
  public
    destructor Destroy; override;
    procedure Connect(const ASource: IUnknown);
    procedure Disconnect;
    property SinkIID: TGUID read FSinkIID write FSinkIID;
    property Source: IUnknown read FSource;
  end;

  //SinkImportsForwards//

  //SinkImports//

  //SinkIntfStart//

  //SinkEventsForwards//
  TOCXEventsGotFocusEvent = procedure(Sender: TObject) of object;
  TOCXEventsLostFocusEvent = procedure(Sender: TObject) of object;

  //SinkComponent//
  TWPSOCXEvents = class(TWPSEventsBaseSink
    //ISinkInterface//
      )
  protected
    function DoInvoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var dps: TDispParams; pDispIds: PDispIdList;
      VarResult, ExcepInfo, ArgErr: Pointer): HResult; override;

    //ISinkInterfaceMethods//
  public
    { system methods }
    constructor Create; virtual;
  protected
    //SinkInterface//
    procedure DoGotFocus; safecall;
    procedure DoLostFocus; safecall;
  protected
    //SinkEventsProtected//
    FGotFocus: TOCXEventsGotFocusEvent;
    FLostFocus: TOCXEventsLostFocusEvent;
  published
    //SinkEventsPublished//
    property GotFocus: TOCXEventsGotFocusEvent read FGotFocus write FGotFocus;
    property LostFocus: TOCXEventsLostFocusEvent read FLostFocus write FLostFocus;
  end;

  //SinkEventsForwards//
  //TNotifyEvent =TNotifyEvent;
  //TApplicationQuit = procedure(Sender: TObject) of object;
 // TApplicationDocumentChange = procedure(Sender: TObject) of object;
  TApplicationDocumentOpen = procedure(Sender: TObject; const Doc: _Document) of object;
  TApplicationDocumentBeforeClose = procedure(Sender: TObject; const Doc: _Document; var Cancel: WordBool) of object;
  TApplicationDocumentBeforePrint = procedure(Sender: TObject; const Doc: _Document; var Cancel: WordBool) of object;
  TApplicationDocumentBeforeSave = procedure(Sender: TObject; const Doc: _Document; var SaveAsUI: WordBool; var Cancel: WordBool) of object;
  TApplicationNewDocument = procedure(Sender: TObject; const Doc: _Document) of object;
  TApplicationWindowActivate = procedure(Sender: TObject; const Doc: _Document; const Wn: Window) of object;
  TApplicationWindowDeactivate = procedure(Sender: TObject; const Doc: _Document; const Wn: Window) of object;
  TApplicationWindowSelectionChange = procedure(Sender: TObject; const Sel: Selection) of object;
  TApplicationWindowBeforeRightClick = procedure(Sender: TObject; const Sel: Selection; var Cancel: WordBool) of object;
  TApplicationWindowBeforeDoubleClick = procedure(Sender: TObject; const Sel: Selection; var Cancel: WordBool) of object;
  TApplicationWindowSize = procedure(Sender: TObject; const Doc: _Document; const Wn: Window) of object;

  //SinkComponent//
  TWPSApplicationEvents = class(TWPSEventsBaseSink
    //ISinkInterface//
      )
  protected
    function DoInvoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var dps: TDispParams; pDispIds: PDispIdList;
      VarResult, ExcepInfo, ArgErr: Pointer): HResult; override;

    //ISinkInterfaceMethods//
  public
    { system methods }
    constructor Create; virtual;
  protected
    //SinkInterface//
    procedure DoStartup; safecall;
    procedure DoQuit; safecall;
    procedure DoDocumentChange; safecall;
    procedure DoDocumentOpen(const Doc: _Document); safecall;
    procedure DoDocumentBeforeClose(const Doc: _Document; var Cancel: WordBool); safecall;
    procedure DoDocumentBeforePrint(const Doc: _Document; var Cancel: WordBool); safecall;
    procedure DoDocumentBeforeSave(const Doc: _Document; var SaveAsUI: WordBool; var Cancel: WordBool); safecall;
    procedure DoNewDocument(const Doc: _Document); safecall;
    procedure DoWindowActivate(const Doc: _Document; const Wn: Window); safecall;
    procedure DoWindowDeactivate(const Doc: _Document; const Wn: Window); safecall;
    procedure DoWindowSelectionChange(const Sel: Selection); safecall;
    procedure DoWindowBeforeRightClick(const Sel: Selection; var Cancel: WordBool); safecall;
    procedure DoWindowBeforeDoubleClick(const Sel: Selection; var Cancel: WordBool); safecall;
    procedure DoWindowSize(const Doc: _Document; const Wn: Window); safecall;
  protected
    //SinkEventsProtected//
    FStartup: TNotifyEvent;
    FQuit: TNotifyEvent;
    FDocumentChange: TNotifyEvent;
    FDocumentOpen: TApplicationDocumentOpen;
    FDocumentBeforeClose: TApplicationDocumentBeforeClose;
    FDocumentBeforePrint: TApplicationDocumentBeforePrint;
    FDocumentBeforeSave: TApplicationDocumentBeforeSave;
    FNewDocument: TApplicationNewDocument;
    FWindowActivate: TApplicationWindowActivate;
    FWindowDeactivate: TApplicationWindowDeactivate;
    FWindowSelectionChange: TApplicationWindowSelectionChange;
    FWindowBeforeRightClick: TApplicationWindowBeforeRightClick;
    FWindowBeforeDoubleClick: TApplicationWindowBeforeDoubleClick;
    FWindowSize: TApplicationWindowSize;
  published
    //SinkEventsPublished//
    property Startup: TNotifyEvent read FStartup write FStartup;
    property Quit: TNotifyEvent read FQuit write FQuit;
    property DocumentChange: TNotifyEvent read FDocumentChange write FDocumentChange;
    property DocumentOpen: TApplicationDocumentOpen read FDocumentOpen write FDocumentOpen;
    property DocumentBeforeClose: TApplicationDocumentBeforeClose read FDocumentBeforeClose write FDocumentBeforeClose;
    property DocumentBeforePrint: TApplicationDocumentBeforePrint read FDocumentBeforePrint write FDocumentBeforePrint;
    property DocumentBeforeSave: TApplicationDocumentBeforeSave read FDocumentBeforeSave write FDocumentBeforeSave;
    property NewDocument: TApplicationNewDocument read FNewDocument write FNewDocument;
    property WindowActivate: TApplicationWindowActivate read FWindowActivate write FWindowActivate;
    property WindowDeactivate: TApplicationWindowDeactivate read FWindowDeactivate write FWindowDeactivate;
    property WindowSelectionChange: TApplicationWindowSelectionChange read FWindowSelectionChange write FWindowSelectionChange;
    property WindowBeforeRightClick: TApplicationWindowBeforeRightClick read FWindowBeforeRightClick write FWindowBeforeRightClick;
    property WindowBeforeDoubleClick: TApplicationWindowBeforeDoubleClick read FWindowBeforeDoubleClick write FWindowBeforeDoubleClick;
    property WindowSize: TApplicationWindowSize read FWindowSize write FWindowSize;
  end;

  //SinkEventsForwards//
  //SinkComponent//
  TWPSDocumentEvents = class(TWPSEventsBaseSink
    //ISinkInterface//
      )
  protected
    function DoInvoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var dps: TDispParams; pDispIds: PDispIdList;
      VarResult, ExcepInfo, ArgErr: Pointer): HResult; override;

    //ISinkInterfaceMethods//
  public
    { system methods }
    constructor Create; virtual;
  protected
    //SinkInterface//
    procedure DoNew; safecall;
    procedure DoOpen; safecall;
    procedure DoClose; safecall;
  protected
    //SinkEventsProtected//
    FNew: TNotifyEvent;
    FOpen: TNotifyEvent;
    FClose: TNotifyEvent;
  published
    //SinkEventsPublished//
    property New: TNotifyEvent read FNew write FNew;
    property Open: TNotifyEvent read FOpen write FOpen;
    property Close: TNotifyEvent read FClose write FClose;
  end;

  //SinkIntfEnd//

implementation

uses
  SysUtils;

{ globals }

procedure BuildPositionalDispIds(pDispIds: PDispIdList; const dps: TDispParams);
var
  i: integer;
begin
  Assert(pDispIds <> nil);

  { by default, directly arrange in reverse order }
  for i := 0 to dps.cArgs - 1 do
    pDispIds^[i] := dps.cArgs - 1 - i;

  { check for named args }
  if (dps.cNamedArgs <= 0) then Exit;

  { parse named args }
  for i := 0 to dps.cNamedArgs - 1 do
    pDispIds^[dps.rgdispidNamedArgs^[i]] := i;
end;

{ TWPSEventsBaseSink }

function TWPSEventsBaseSink.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;
begin
  Result := E_NOTIMPL;
end;

function TWPSEventsBaseSink.GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult;
begin
  Result := E_NOTIMPL;
  pointer(TypeInfo) := nil;
end;

function TWPSEventsBaseSink.GetTypeInfoCount(out Count: Integer): HResult;
begin
  Result := E_NOTIMPL;
  Count := 0;
end;

function TWPSEventsBaseSink.Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
  Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult;
var
  dps: TDispParams absolute Params;
  bHasParams: boolean;
  pDispIds: PDispIdList;
  iDispIdsSize: integer;
begin
  { validity checks }
  if (Flags and DISPATCH_METHOD = 0) then
    raise Exception.Create(
      Format('%s only supports sinking of method calls!', [ClassName]
      ));

  { build pDispIds array. this maybe a bit of overhead but it allows us to
    sink named-argument calls such as Excel's AppEvents, etc!
  }
  pDispIds := nil;
  iDispIdsSize := 0;
  bHasParams := (dps.cArgs > 0);
  if (bHasParams) then
  begin
    iDispIdsSize := dps.cArgs * SizeOf(TDispId);
    GetMem(pDispIds, iDispIdsSize);
  end; { if }

  try
    { rearrange dispids properly }
    if (bHasParams) then BuildPositionalDispIds(pDispIds, dps);
    Result := DoInvoke(DispId, IID, LocaleID, Flags, dps, pDispIds, VarResult, ExcepInfo, ArgErr);
  finally
    { free pDispIds array }
    if (bHasParams) then FreeMem(pDispIds, iDispIdsSize);
  end; { finally }
end;

function TWPSEventsBaseSink.QueryInterface(const IID: TGUID; out Obj): HResult;
begin
  if (GetInterface(IID, Obj)) then
  begin
    Result := S_OK;
    Exit;
  end
  else if (IsEqualIID(IID, FSinkIID)) then
    if (GetInterface(IDispatch, Obj)) then
    begin
      Result := S_OK;
      Exit;
    end;
  Result := E_NOINTERFACE;
  pointer(Obj) := nil;
end;

function TWPSEventsBaseSink._AddRef: Integer;
begin
  Result := 2;
end;

function TWPSEventsBaseSink._Release: Integer;
begin
  Result := 1;
end;

destructor TWPSEventsBaseSink.Destroy;
begin
  Disconnect;
  inherited;
end;

procedure TWPSEventsBaseSink.Connect(const ASource: IUnknown);
var
  pcpc: IConnectionPointContainer;
begin
  Assert(ASource <> nil);
  Disconnect;
  try
    OleCheck(ASource.QueryInterface(IConnectionPointContainer, pcpc));
    OleCheck(pcpc.FindConnectionPoint(FSinkIID, FCP));
    OleCheck(FCP.Advise(Self, FCookie));
    FSource := ASource;
  except
    raise Exception.Create(Format('Unable to connect %s.'#13'%s',
      [ClassName, Exception(ExceptObject).Message]
      ));
  end; { finally }
end;

procedure TWPSEventsBaseSink.Disconnect;
begin
  if (FSource = nil) then Exit;
  try
    OleCheck(FCP.Unadvise(FCookie));
    FCP := nil;
    FSource := nil;
  except
    pointer(FCP) := nil;
    pointer(FSource) := nil;
  end; { except }
end;

//SinkImplStart//

function TWPSOCXEvents.DoInvoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
  Flags: Word; var dps: TDispParams; pDispIds: PDispIdList;
  VarResult, ExcepInfo, ArgErr: Pointer): HResult;
type
  POleVariant = ^OleVariant;
begin
  Result := DISP_E_MEMBERNOTFOUND;

  //SinkInvoke//
  case DispId of
    -2147417888:
      begin
        DoGotFocus();
        Result := S_OK;
      end;
    -2147417887:
      begin
        DoLostFocus();
        Result := S_OK;
      end;
  end; { case }
  //SinkInvokeEnd//
end;

constructor TWPSOCXEvents.Create;
begin
  inherited Create;
  //SinkInit//
 // FSinkIID := OCXEvents;
  FSinkIID := DIID_OCXEvents;
end;

//SinkImplementation//

procedure TWPSOCXEvents.DoGotFocus;
begin
  if not Assigned(GotFocus) then System.Exit;
  GotFocus(Self);
end;

procedure TWPSOCXEvents.DoLostFocus;
begin
  if not Assigned(LostFocus) then System.Exit;
  LostFocus(Self);
end;

function TWPSApplicationEvents.DoInvoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
  Flags: Word; var dps: TDispParams; pDispIds: PDispIdList;
  VarResult, ExcepInfo, ArgErr: Pointer): HResult;
type
  POleVariant = ^OleVariant;
begin
  Result := DISP_E_MEMBERNOTFOUND;

  //SinkInvoke//
  case DispId of
    1:
      begin
        DoStartup();
        Result := S_OK;
      end;
    2:
      begin
        DoQuit();
        Result := S_OK;
      end;
    3:
      begin
        DoDocumentChange();
        Result := S_OK;
      end;
    4:
      begin
        DoDocumentOpen(IUnknown(dps.rgvarg^[pDispIds^[0]].unkval) as _Document);
        Result := S_OK;
      end;
    6:
      begin
        DoDocumentBeforeClose(IUnknown(dps.rgvarg^[pDispIds^[0]].unkval) as _Document, dps.rgvarg^[pDispIds^[1]].pbool^);
        Result := S_OK;
      end;
    7:
      begin
        DoDocumentBeforePrint(IUnknown(dps.rgvarg^[pDispIds^[0]].unkval) as _Document, dps.rgvarg^[pDispIds^[1]].pbool^);
        Result := S_OK;
      end;
    8:
      begin
        DoDocumentBeforeSave(IUnknown(dps.rgvarg^[pDispIds^[0]].unkval) as _Document, dps.rgvarg^[pDispIds^[1]].pbool^, dps.rgvarg^[pDispIds^[2]].pbool^);
        Result := S_OK;
      end;
    9:
      begin
        DoNewDocument(IUnknown(dps.rgvarg^[pDispIds^[0]].unkval) as _Document);
        Result := S_OK;
      end;
    10:
      begin
        try
          DoWindowActivate(IUnknown(dps.rgvarg^[pDispIds^[0]].unkval) as _Document, IUnknown(dps.rgvarg^[pDispIds^[1]].unkval) as Window);
        except
        end;
        Result := S_OK;
      end;
    11:
      begin
        DoWindowDeactivate(IUnknown(dps.rgvarg^[pDispIds^[0]].unkval) as _Document, IUnknown(dps.rgvarg^[pDispIds^[1]].unkval) as Window);
        Result := S_OK;
      end;
    12:
      begin
        DoWindowSelectionChange(IUnknown(dps.rgvarg^[pDispIds^[0]].unkval) as Selection);
        Result := S_OK;
      end;
    13:
      begin
        DoWindowBeforeRightClick(IUnknown(dps.rgvarg^[pDispIds^[0]].unkval) as Selection, dps.rgvarg^[pDispIds^[1]].pbool^);
        Result := S_OK;
      end;
    14:
      begin
        DoWindowBeforeDoubleClick(IUnknown(dps.rgvarg^[pDispIds^[0]].unkval) as Selection, dps.rgvarg^[pDispIds^[1]].pbool^);
        Result := S_OK;
      end;
    25:
      begin
        DoWindowSize(IUnknown(dps.rgvarg^[pDispIds^[0]].unkval) as _Document, IUnknown(dps.rgvarg^[pDispIds^[1]].unkval) as Window);
        Result := S_OK;
      end;
  end; { case }
  //SinkInvokeEnd//
end;

constructor TWPSApplicationEvents.Create;
begin
  inherited Create;
  //SinkInit//
 // FSinkIID := ApplicationEvents;
  FSinkIID := DIID_ApplicationEvents;
end;

//SinkImplementation//

procedure TWPSApplicationEvents.DoStartup;
begin
  if not Assigned(Startup) then System.Exit;
  Startup(Self);
end;

procedure TWPSApplicationEvents.DoQuit;
begin
  if not Assigned(Quit) then System.Exit;
  Quit(Self);
end;

procedure TWPSApplicationEvents.DoDocumentChange;
begin
  if not Assigned(DocumentChange) then System.Exit;
  DocumentChange(Self);
end;

procedure TWPSApplicationEvents.DoDocumentOpen(const Doc: _Document);
begin
  if not Assigned(DocumentOpen) then System.Exit;
  DocumentOpen(Self, Doc);
end;

procedure TWPSApplicationEvents.DoDocumentBeforeClose(const Doc: _Document; var Cancel: WordBool);
begin
  if not Assigned(DocumentBeforeClose) then System.Exit;
  DocumentBeforeClose(Self, Doc, Cancel);
end;

procedure TWPSApplicationEvents.DoDocumentBeforePrint(const Doc: _Document; var Cancel: WordBool);
begin
  if not Assigned(DocumentBeforePrint) then System.Exit;
  DocumentBeforePrint(Self, Doc, Cancel);
end;

procedure TWPSApplicationEvents.DoDocumentBeforeSave(const Doc: _Document; var SaveAsUI: WordBool; var Cancel: WordBool);
begin
  if not Assigned(DocumentBeforeSave) then System.Exit;
  DocumentBeforeSave(Self, Doc, SaveAsUI, Cancel);
end;

procedure TWPSApplicationEvents.DoNewDocument(const Doc: _Document);
begin
  if not Assigned(NewDocument) then System.Exit;
  NewDocument(Self, Doc);
end;

procedure TWPSApplicationEvents.DoWindowActivate(const Doc: _Document; const Wn: Window);
begin
  if not Assigned(WindowActivate) then System.Exit;
  WindowActivate(Self, Doc, Wn);
end;

procedure TWPSApplicationEvents.DoWindowDeactivate(const Doc: _Document; const Wn: Window);
begin
  if not Assigned(WindowDeactivate) then System.Exit;
  WindowDeactivate(Self, Doc, Wn);
end;

procedure TWPSApplicationEvents.DoWindowSelectionChange(const Sel: Selection);
begin
  if not Assigned(WindowSelectionChange) then System.Exit;
  WindowSelectionChange(Self, Sel);
end;

procedure TWPSApplicationEvents.DoWindowBeforeRightClick(const Sel: Selection; var Cancel: WordBool);
begin
  if not Assigned(WindowBeforeRightClick) then System.Exit;
  WindowBeforeRightClick(Self, Sel, Cancel);
end;

procedure TWPSApplicationEvents.DoWindowBeforeDoubleClick(const Sel: Selection; var Cancel: WordBool);
begin
  if not Assigned(WindowBeforeDoubleClick) then System.Exit;
  WindowBeforeDoubleClick(Self, Sel, Cancel);
end;

procedure TWPSApplicationEvents.DoWindowSize(const Doc: _Document; const Wn: Window);
begin
  if not Assigned(WindowSize) then System.Exit;
  WindowSize(Self, Doc, Wn);
end;

function TWPSDocumentEvents.DoInvoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
  Flags: Word; var dps: TDispParams; pDispIds: PDispIdList;
  VarResult, ExcepInfo, ArgErr: Pointer): HResult;
type
  POleVariant = ^OleVariant;
begin
  Result := DISP_E_MEMBERNOTFOUND;

  //SinkInvoke//
  case DispId of
    4:
      begin
        DoNew();
        Result := S_OK;
      end;
    5:
      begin
        DoOpen();
        Result := S_OK;
      end;
    6:
      begin
        DoClose();
        Result := S_OK;
      end;
  end; { case }
  //SinkInvokeEnd//
end;

constructor TWPSDocumentEvents.Create;
begin
  inherited Create;
  //SinkInit//
  //FSinkIID := DocumentEvents;
  FSinkIID := DIID_DocumentEvents;
end;

//SinkImplementation//

procedure TWPSDocumentEvents.DoNew;
begin
  if not Assigned(New) then System.Exit;
  New(Self);
end;

procedure TWPSDocumentEvents.DoOpen;
begin
  if not Assigned(Open) then System.Exit;
  Open(Self);
end;

procedure TWPSDocumentEvents.DoClose;
begin
  if not Assigned(Close) then System.Exit;
  Close(Self);
end;

//SinkImplEnd//

end.

