unit AddInDesignerObjects_TLB_WPS;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// $Rev: 17244 $
// File generated on 2011-4-28 16:48:00 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Program Files\Kingsoft\WPS Office Personal\office6\ksaddndr.dll (1)
// LIBID: {A537E638-AB2A-4308-A502-2EFF280C6E98}
// LCID: 0
// Helpfile: 
// HelpString: Kingsoft Add-In Designer
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\system32\stdole2.tlb)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}
interface

uses Windows, ActiveX, Classes, Graphics, OleServer, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  AddInDesignerObjectsMajorVersion = 1;
  AddInDesignerObjectsMinorVersion = 0;

  LIBID_AddInDesignerObjects: TGUID = '{A537E638-AB2A-4308-A502-2EFF280C6E98}';

  IID__IDTExtensibility2: TGUID = '{B65AD801-ABAF-11D0-BB8B-00A0C90F2744}';
  CLASS_IDTExtensibility2: TGUID = '{EC847E18-AD46-4757-AF56-7EFDC5BE13F5}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum ext_ConnectMode
type
  ext_ConnectMode = TOleEnum;
const
  ext_cm_AfterStartup = $00000000;
  ext_cm_Startup = $00000001;
  ext_cm_External = $00000002;
  ext_cm_CommandLine = $00000003;

// Constants for enum ext_DisconnectMode
type
  ext_DisconnectMode = TOleEnum;
const
  ext_dm_HostShutdown = $00000000;
  ext_dm_UserClosed = $00000001;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  _IDTExtensibility2 = interface;
  _IDTExtensibility2Disp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  IDTExtensibility2 = _IDTExtensibility2;


// *********************************************************************//
// Declaration of structures, unions and aliases.                         
// *********************************************************************//
  PPSafeArray1 = ^PSafeArray; {*}


// *********************************************************************//
// Interface: _IDTExtensibility2
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {B65AD801-ABAF-11D0-BB8B-00A0C90F2744}
// *********************************************************************//
  _IDTExtensibility2 = interface(IDispatch)
    ['{B65AD801-ABAF-11D0-BB8B-00A0C90F2744}']
    procedure OnConnection(const Application: IDispatch; ConnectMode: ext_ConnectMode; 
                           const AddInInst: IDispatch; var custom: PSafeArray); safecall;
    procedure OnDisconnection(RemoveMode: ext_DisconnectMode; var custom: PSafeArray); safecall;
    procedure OnAddInsUpdate(var custom: PSafeArray); safecall;
    procedure OnStartupComplete(var custom: PSafeArray); safecall;
    procedure OnBeginShutdown(var custom: PSafeArray); safecall;
  end;

// *********************************************************************//
// DispIntf:  _IDTExtensibility2Disp
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {B65AD801-ABAF-11D0-BB8B-00A0C90F2744}
// *********************************************************************//
  _IDTExtensibility2Disp = dispinterface
    ['{B65AD801-ABAF-11D0-BB8B-00A0C90F2744}']
    procedure OnConnection(const Application: IDispatch; ConnectMode: ext_ConnectMode; 
                           const AddInInst: IDispatch; var custom: {??PSafeArray}OleVariant); dispid 1;
    procedure OnDisconnection(RemoveMode: ext_DisconnectMode; var custom: {??PSafeArray}OleVariant); dispid 2;
    procedure OnAddInsUpdate(var custom: {??PSafeArray}OleVariant); dispid 3;
    procedure OnStartupComplete(var custom: {??PSafeArray}OleVariant); dispid 4;
    procedure OnBeginShutdown(var custom: {??PSafeArray}OleVariant); dispid 5;
  end;

// *********************************************************************//
// The Class CoIDTExtensibility2 provides a Create and CreateRemote method to          
// create instances of the default interface _IDTExtensibility2 exposed by              
// the CoClass IDTExtensibility2. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoIDTExtensibility2 = class
    class function Create: _IDTExtensibility2;
    class function CreateRemote(const MachineName: string): _IDTExtensibility2;
  end;

implementation

uses ComObj;

class function CoIDTExtensibility2.Create: _IDTExtensibility2;
begin
  Result := CreateComObject(CLASS_IDTExtensibility2) as _IDTExtensibility2;
end;

class function CoIDTExtensibility2.CreateRemote(const MachineName: string): _IDTExtensibility2;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_IDTExtensibility2) as _IDTExtensibility2;
end;

end.
