unit WordAppEvents;

interface

uses
  OleServer, Word_TLB, Classes, ActiveX, Variants;

type
  TWordApplicationDocumentOpen = procedure(ASender: TObject; const Doc: WordDocument) of object;
  TWordApplicationDocumentBeforeClose = procedure(ASender: TObject; const Doc: WordDocument;
                                                                    var Cancel: WordBool) of object;
  TWordApplicationDocumentBeforePrint = procedure(ASender: TObject; const Doc: WordDocument;
                                                                    var Cancel: WordBool) of object;
  TWordApplicationDocumentBeforeSave = procedure(ASender: TObject; const Doc: WordDocument;
                                                                   var SaveAsUI: WordBool;
                                                                   var Cancel: WordBool) of object;
  TWordApplicationNewDocument = procedure(ASender: TObject; const Doc: WordDocument) of object;
  TWordApplicationWindowActivate = procedure(ASender: TObject; const Doc: WordDocument;
                                                               const Wn: Window) of object;
  TWordApplicationWindowDeactivate = procedure(ASender: TObject; const Doc: WordDocument;
                                                                 const Wn: Window) of object;
  TWordApplicationWindowSelectionChange = procedure(ASender: TObject; const Sel: WordSelection) of object;

  TWordApplication = class(TOleServer)
  private
    Fintf: WordApplication;
    FOnStartup: TNotifyEvent;
    FOnQuit: TNotifyEvent;
    FOnDocumentChange: TNotifyEvent;
    FOnDocumentOpen: TWordApplicationDocumentOpen;
    FOnDocumentBeforeClose: TWordApplicationDocumentBeforeClose;
    FOnDocumentBeforePrint: TWordApplicationDocumentBeforePrint;
    FOnDocumentBeforeSave: TWordApplicationDocumentBeforeSave;
    FOnNewDocument: TWordApplicationNewDocument;
    FOnWindowActivate: TWordApplicationWindowActivate;
    FOnWindowDeactivate: TWordApplicationWindowDeactivate;
    FOnWindowSelectionChange: TWordApplicationWindowSelectionChange;
    FAutoQuit:    Boolean;
    function GetDefaultInterface: _Application;
  protected
    procedure InitServerData; override;
    procedure InvokeEvent(DispID: TDispID; var Params: TVariantArray); override;
  public

    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: _Application);
    procedure Disconnect; override;
    procedure Quit; overload;
    property DefaultInterface: _Application read GetDefaultInterface;
  published
    property OnStartup: TNotifyEvent read FOnStartup write FOnStartup;
    property OnQuit: TNotifyEvent read FOnQuit write FOnQuit;
    property OnDocumentChange: TNotifyEvent read FOnDocumentChange write FOnDocumentChange;
    property OnDocumentOpen: TWordApplicationDocumentOpen read FOnDocumentOpen write FOnDocumentOpen;
    property OnDocumentBeforeClose: TWordApplicationDocumentBeforeClose read FOnDocumentBeforeClose write FOnDocumentBeforeClose;
    property OnDocumentBeforePrint: TWordApplicationDocumentBeforePrint read FOnDocumentBeforePrint write FOnDocumentBeforePrint;
    property OnDocumentBeforeSave: TWordApplicationDocumentBeforeSave read FOnDocumentBeforeSave write FOnDocumentBeforeSave;
    property OnNewDocument: TWordApplicationNewDocument read FOnNewDocument write FOnNewDocument;
    property OnWindowActivate: TWordApplicationWindowActivate read FOnWindowActivate write FOnWindowActivate;
    property OnWindowDeactivate: TWordApplicationWindowDeactivate read FOnWindowDeactivate write FOnWindowDeactivate;
    property OnWindowSelectionChange: TWordApplicationWindowSelectionChange read FOnWindowSelectionChange write FOnWindowSelectionChange;
  end;


implementation

{ TWordApplication }

procedure TWordApplication.Connect;
var
  punk: IUnknown;
begin
  if Fintf = nil then
  begin
    punk := GetServer;
    ConnectEvents(punk);
    Fintf:= punk as _Application;
  end;
end;

procedure TWordApplication.ConnectTo(svrIntf: _Application);
begin
  Disconnect;
  Fintf := svrIntf;
  ConnectEvents(Fintf);
end;

constructor TWordApplication.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;


destructor TWordApplication.Destroy;
begin
  inherited Destroy;
end;


procedure TWordApplication.DisConnect;
begin
  if Fintf <> nil then
  begin
    DisconnectEvents(FIntf);
    if FAutoQuit then
      Quit();
    FIntf := nil;
  end;
end;


function TWordApplication.GetDefaultInterface: _Application;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

procedure TWordApplication.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{000209FF-0000-0000-C000-000000000046}';
    IntfIID:   '{00020970-0000-0000-C000-000000000046}';
    EventIID:  '{00020A00-0000-0000-C000-000000000046}';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;


procedure TWordApplication.InvokeEvent(DispID: TDispID;
  var Params: TVariantArray);
begin
 case DispID of
    -1: Exit;  // DISPID_UNKNOWN
    1: if Assigned(FOnStartup) then
         FOnStartup(Self);
    2: if Assigned(FOnQuit) then
         FOnQuit(Self);
    3: if Assigned(FOnDocumentChange) then
         FOnDocumentChange(Self);
    4: if Assigned(FOnDocumentOpen) then
         FOnDocumentOpen(Self, IUnknown(TVarData(Params[0]).VPointer) as WordDocument {const WordDocument});
    6: if Assigned(FOnDocumentBeforeClose) then
         FOnDocumentBeforeClose(Self,
                                IUnknown(TVarData(Params[0]).VPointer) as WordDocument {const WordDocument},
                                WordBool((TVarData(Params[1]).VPointer)^) {var WordBool});
    7: if Assigned(FOnDocumentBeforePrint) then
         FOnDocumentBeforePrint(Self,
                                IUnknown(TVarData(Params[0]).VPointer) as WordDocument {const WordDocument},
                                WordBool((TVarData(Params[1]).VPointer)^) {var WordBool});
    8: if Assigned(FOnDocumentBeforeSave) then
         FOnDocumentBeforeSave(Self,
                               IUnknown(TVarData(Params[0]).VPointer) as WordDocument {const WordDocument},
                               WordBool((TVarData(Params[1]).VPointer)^) {var WordBool},
                               WordBool((TVarData(Params[2]).VPointer)^) {var WordBool});
    9: if Assigned(FOnNewDocument) then
         FOnNewDocument(Self, IUnknown(TVarData(Params[0]).VPointer) as WordDocument {const WordDocument});
    10: if Assigned(FOnWindowActivate) then
         FOnWindowActivate(Self,
                           IUnknown(TVarData(Params[0]).VPointer) as WordDocument {const WordDocument},
                           IUnknown(TVarData(Params[1]).VPointer) as Window {const Window});
    11: if Assigned(FOnWindowDeactivate) then
         FOnWindowDeactivate(Self,
                             IUnknown(TVarData(Params[0]).VPointer) as WordDocument {const WordDocument},
                             IUnknown(TVarData(Params[1]).VPointer) as Window {const Window});
    12: if Assigned(FOnWindowSelectionChange) then
         FOnWindowSelectionChange(Self, IUnknown(TVarData(Params[0]).VPointer) as WordSelection {const WordSelection});
  end; {case DispID}
end;

procedure TWordApplication.Quit();
begin
  DefaultInterface.Quit(EmptyParam, EmptyParam, EmptyParam);
end;

end.
