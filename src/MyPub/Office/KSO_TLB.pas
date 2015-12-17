unit KSO_TLB;

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
// File generated on 2011-4-28 16:48:33 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Program Files\Kingsoft\WPS Office Personal\office6\kso10.dll (1)
// LIBID: {4A1D9D13-2EC6-495B-A5B5-848228E0A1CE}
// LCID: 0
// Helpfile: C:\Program Files\Kingsoft\WPS Office Personal\office6\kso10.dll\ksoapi.chm
// HelpString: Kingsoft Office 1.0 Object Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\system32\stdole2.tlb)
// Parent TypeLibrary:
//   (0) v2.0 WPS, (C:\Program Files\Kingsoft\WPS Office Personal\office6\wpscore.dll)
// Errors:
//   Hint: Parameter 'Type' of _CommandBars.FindControl changed to 'Type_'
//   Hint: Parameter 'Type' of _CommandBars.FindControls changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of CommandBar.FindControl changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of CommandBarControls.Add changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of _KsoDiagram.Type changed to 'Type_'
//   Hint: Parameter 'Type' of _KsoDiagram.Convert changed to 'Type_'
//   Hint: Member 'Parent' of '_KsoDiagramNode' changed to 'Parent_'
//   Hint: Parameter 'Type' of _KsoDiagramNode.Layout changed to 'Type_'
//   Hint: Parameter 'Type' of _KsoDiagramNode.Layout changed to 'Type_'
//   Hint: Member 'Parent' of '_KsoDiagramNodeChildren' changed to 'Parent_'
//   Hint: Parameter 'Type' of KsoShape._Type changed to 'Type_'
//   Hint: Parameter 'Type' of KsoShapeRange._Type changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of KsoCalloutFormat.Type changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of KsoCalloutFormat.Type changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of KsoConnectorFormat.Type changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of KsoConnectorFormat.Type changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of KsoFillFormat.Type changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of KsoColorFormat.Type changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of KsoShadowFormat.Type changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of KsoShadowFormat.Type changed to 'Type_'
//   Hint: Parameter 'Object' of KsoScript.Shape changed to 'Object_'
//   Hint: Parameter 'Type' of KsoCanvasShapes._AddCallout changed to 'Type_'
//   Hint: Parameter 'Type' of KsoCanvasShapes._AddConnector changed to 'Type_'
//   Hint: Parameter 'Label' of KsoCanvasShapes._AddLabel changed to 'Label_'
//   Hint: Parameter 'Type' of KsoCanvasShapes._AddShape changed to 'Type_'
//   Hint: Parameter 'Type' of KsoShapes._AddCallout changed to 'Type_'
//   Hint: Parameter 'Type' of KsoShapes._AddConnector changed to 'Type_'
//   Hint: Parameter 'Label' of KsoShapes._AddLabel changed to 'Label_'
//   Hint: Parameter 'Type' of KsoShapes._AddShape changed to 'Type_'
//   Hint: Parameter 'Type' of KsoShapes._AddDiagram changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of DocumentProperties.Add changed to 'Type_'
//   Hint: Member 'Object' of 'COMAddIn' changed to 'Object_'
//   Hint: Member 'Protected' of 'KeyBinding' changed to 'Protected_'
//   Hint: Parameter 'Type' of AdvApiRoot.AdvApi changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'File' of AdvApplication.CreateFilterMediaOnFile changed to 'File_'
//   Hint: Member 'Object' of '_TaskPane' changed to 'Object_'
//   Hint: Parameter 'Object' of _TaskPane.Object changed to 'Object_'
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
  KSOMajorVersion = 1;
  KSOMinorVersion = 0;

  LIBID_KSO: TGUID = '{4A1D9D13-2EC6-495B-A5B5-848228E0A1CE}';

  IID_IKsoDispObj: TGUID = '{000C0300-FFFF-0000-C000-000000000046}';
  IID_Collection: TGUID = '{000C0301-FFFF-0000-C000-000000000046}';
  IID_SecurityOptions: TGUID = '{0002B003-FFFE-0000-C000-000000111146}';
  IID__CommandBars: TGUID = '{0002EB0D-FFFE-0000-C000-000000111146}';
  IID_CommandBarControl: TGUID = '{0002C882-FFFE-0000-C000-000000111146}';
  IID_CommandBar: TGUID = '{00023F42-FFFE-0000-C000-000000111146}';
  IID_CommandBarControls: TGUID = '{0002448B-FFFE-0000-C000-000000111146}';
  IID__CommandBarButton: TGUID = '{00023C90-FFFE-0000-C000-000000111146}';
  IID_CommandBarPopup: TGUID = '{0002BF47-FFFE-0000-C000-000000111146}';
  IID__CommandBarComboBox: TGUID = '{0002ABB5-FFFE-0000-C000-000000111146}';
  IID_ICommandBarsEvents: TGUID = '{000205F4-FFFE-0000-C000-000000111146}';
  DIID__CommandBarsEvents: TGUID = '{10227926-9CB9-45B9-ADF2-803D8358EC6A}';
  IID_ICommandBarComboBoxEvents: TGUID = '{0002861A-FFFE-0000-C000-000000111146}';
  DIID__CommandBarComboBoxEvents: TGUID = '{2AC87A8B-FABC-4F54-8234-B637DFEC5528}';
  CLASS_CommandBarComboBox: TGUID = '{BD7A5BD8-B94E-4EEB-9217-C9A8375CE30B}';
  DIID__CommandBarButtonEvents: TGUID = '{B480CA3F-0571-458A-82E4-FEDE4CB3FF96}';
  CLASS_CommandBarButton: TGUID = '{264C279A-9BB9-4565-BA1E-353C0A4903C7}';
  IID_ICommandBarButtonEvents: TGUID = '{000258F0-FFFE-0000-C000-000000111146}';
  CLASS_CommandBars: TGUID = '{D4D7D5A3-E8B0-414A-B266-69CBBB27699F}';
  IID__KsoDiagram: TGUID = '{000C036D-FFFF-0000-C000-000000000046}';
  IID__KsoDiagramNodes: TGUID = '{000C036E-FFFF-0000-C000-000000000046}';
  IID__KsoDiagramNode: TGUID = '{000C0370-FFFF-0000-C000-000000000046}';
  IID__KsoDiagramNodeChildren: TGUID = '{000C036F-FFFF-0000-C000-000000000046}';
  IID_KsoShape: TGUID = '{000C031C-FFFF-0000-C000-000000000046}';
  IID_KsoShapeRange: TGUID = '{000C031D-FFFF-0000-C000-000000000046}';
  IID_KsoAdjustments: TGUID = '{000C0310-FFFF-0000-C000-000000000046}';
  IID_KsoCalloutFormat: TGUID = '{000C0311-FFFF-0000-C000-000000000046}';
  IID_KsoConnectorFormat: TGUID = '{000C0313-FFFF-0000-C000-000000000046}';
  IID_KsoFillFormat: TGUID = '{000C0314-FFFF-0000-C000-000000000046}';
  IID_KsoColorFormat: TGUID = '{000C0312-FFFF-0000-C000-000000000046}';
  IID_KsoGroupShapes: TGUID = '{000C0316-FFFF-0000-C000-000000000046}';
  IID_KsoLineFormat: TGUID = '{000C0317-FFFF-0000-C000-000000000046}';
  IID_KsoShapeNodes: TGUID = '{000C0319-FFFF-0000-C000-000000000046}';
  IID_KsoShapeNode: TGUID = '{000C0318-FFFF-0000-C000-000000000046}';
  IID_KsoPictureFormat: TGUID = '{000C031A-FFFF-0000-C000-000000000046}';
  IID_KsoShadowFormat: TGUID = '{000C031B-FFFF-0000-C000-000000000046}';
  IID_KsoTextEffectFormat: TGUID = '{000C031F-FFFF-0000-C000-000000000046}';
  IID_KsoTextFrame: TGUID = '{000C3484-FFFF-0000-C000-000000000046}';
  IID_KsoThreeDFormat: TGUID = '{000C0321-FFFF-0000-C000-000000000046}';
  IID_KsoScript: TGUID = '{000C0341-FFFF-0000-C000-000000000046}';
  IID_KsoCanvasShapes: TGUID = '{000C0371-FFFF-0000-C000-000000000046}';
  IID_KsoFreeformBuilder: TGUID = '{000C0315-FFFF-0000-C000-000000000046}';
  IID_KsoOLEFormat: TGUID = '{000C0322-FFFF-0000-C000-000000000046}';
  IID_KsoShapes: TGUID = '{000C031E-FFFF-0000-C000-000000000046}';
  IID_DocumentProperty: TGUID = '{CA2A9BB6-76E4-4509-835A-EAC029FBC7EF}';
  IID_DocumentProperties: TGUID = '{8CC27FD1-177E-4527-B03B-30A07962D256}';
  IID_COMAddIn: TGUID = '{0002033A-FFFE-0000-C000-000000111146}';
  IID_COMAddIns: TGUID = '{00020339-FFFE-0000-C000-000000111146}';
  IID_KeyBinding: TGUID = '{00020998-FFFE-0000-C000-000000111146}';
  IID_KeyBindings: TGUID = '{00020996-FFFE-0000-C000-000000111146}';
  IID_KeysBoundTo: TGUID = '{00020997-FFFE-0000-C000-000000111146}';
  IID_AdvApiRoot: TGUID = '{00023EF4-FFFE-0000-C000-000000111146}';
  IID_AdvAddins: TGUID = '{0002F52F-FFFE-0000-C000-000000111146}';
  IID_AdvAddin: TGUID = '{0002B7E6-FFFE-0000-C000-000000111146}';
  IID_AdvRight: TGUID = '{0002EE54-FFFE-0000-C000-000000111146}';
  IID_AdvSeal: TGUID = '{000299FC-FFFE-0000-C000-000000111146}';
  IID_Seals: TGUID = '{0002BE8B-FFFE-0000-C000-000000111146}';
  IID_Seal: TGUID = '{0002F78B-FFFE-0000-C000-000000111146}';
  IID_IFilterMedium: TGUID = '{0002D7C3-FFFE-0000-C000-000000111146}';
  IID_AdvApplication: TGUID = '{0002AFDA-FFFE-0000-C000-000000111146}';
  DIID_AdvApplicationEvents: TGUID = '{B94C31F7-4947-4CEB-ABDD-9B37CE8422B5}';
  DIID_AdvRightEvents: TGUID = '{A0984A28-9CF6-4826-86A9-E8D592950965}';
  DIID_AdvSealEvents: TGUID = '{F84DA428-2F66-4B38-B7D9-84C4E6EB4B28}';
  IID__TaskPane: TGUID = '{0002B802-FFFE-0000-C000-000000111146}';
  IID_TaskPanes: TGUID = '{0002B801-FFFE-0000-C000-000000111146}';
  DIID__ITaskPaneEvent: TGUID = '{1D4270E9-E494-47B5-B7A0-2280BCD2A352}';
  CLASS_TaskPane: TGUID = '{FB1ACD80-7BCC-4B67-ABF7-08B2257AEC17}';
  IID_PluginPlatform: TGUID = '{00020F01-FFFE-0000-C000-000000111146}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum KsoControlOLEUsage
type
  KsoControlOLEUsage = TOleEnum;
const
  ksoControlOLEUsageNeither = $00000000;
  ksoControlOLEUsageServer = $00000001;
  ksoControlOLEUsageClient = $00000002;
  ksoControlOLEUsageBoth = $00000003;

// Constants for enum KsoControlType
type
  KsoControlType = TOleEnum;
const
  ksoControlCustom = $00000000;
  ksoControlButton = $00000001;
  ksoControlEdit = $00000002;
  ksoControlDropdown = $00000003;
  ksoControlComboBox = $00000004;
  ksoControlButtonDropdown = $00000005;
  ksoControlSplitDropdown = $00000006;
  ksoControlOCXDropdown = $00000007;
  ksoControlGenericDropdown = $00000008;
  ksoControlGraphicDropdown = $00000009;
  ksoControlPopup = $0000000A;
  ksoControlGraphicPopup = $0000000B;
  ksoControlButtonPopup = $0000000C;
  ksoControlSplitButtonPopup = $0000000D;
  ksoControlSplitButtonMRUPopup = $0000000E;
  ksoControlLabel = $0000000F;
  ksoControlExpandingGrid = $00000010;
  ksoControlSplitExpandingGrid = $00000011;
  ksoControlGrid = $00000012;
  ksoControlGauge = $00000013;
  ksoControlGraphicCombo = $00000014;
  ksoControlPane = $00000015;
  ksoControlActiveX = $00000016;
  ksoControlSpinner = $00000017;
  ksoControlLabelEx = $00000018;
  ksoControlWorkPane = $00000019;
  ksoControlAutoCompleteCombo = $0000001A;

// Constants for enum KsoBarPosition
type
  KsoBarPosition = TOleEnum;
const
  ksoBarLeft = $00000000;
  ksoBarTop = $00000001;
  ksoBarRight = $00000002;
  ksoBarBottom = $00000003;
  ksoBarFloating = $00000004;
  ksoBarPopup = $00000005;
  ksoBarMenuBar = $00000006;

// Constants for enum KsoBarProtection
type
  KsoBarProtection = TOleEnum;
const
  ksoBarNoProtection = $00000000;
  ksoBarNoCustomize = $00000001;
  ksoBarNoResize = $00000002;
  ksoBarNoMove = $00000004;
  ksoBarNoChangeVisible = $00000008;
  ksoBarNoChangeDock = $00000010;
  ksoBarNoVerticalDock = $00000020;
  ksoBarNoHorizontalDock = $00000040;

// Constants for enum KsoBarType
type
  KsoBarType = TOleEnum;
const
  ksoBarTypeNormal = $00000000;
  ksoBarTypeMenuBar = $00000001;
  ksoBarTypePopup = $00000002;

// Constants for enum KsoMenuAnimation
type
  KsoMenuAnimation = TOleEnum;
const
  ksoMenuAnimationNone = $00000000;
  ksoMenuAnimationRandom = $00000001;
  ksoMenuAnimationUnfold = $00000002;
  ksoMenuAnimationSlide = $00000003;

// Constants for enum KsoButtonState
type
  KsoButtonState = TOleEnum;
const
  ksoButtonUp = $00000000;
  ksoButtonDown = $FFFFFFFF;
  ksoButtonMixed = $00000002;

// Constants for enum KsoButtonStyle
type
  KsoButtonStyle = TOleEnum;
const
  ksoButtonAutomatic = $00000000;
  ksoButtonIcon = $00000001;
  ksoButtonCaption = $00000002;
  ksoButtonIconAndCaption = $00000003;
  ksoButtonIconAndWrapCaption = $00000007;
  ksoButtonIconAndCaptionBelow = $0000000B;
  ksoButtonWrapCaption = $0000000E;
  ksoButtonIconAndWrapCaptionBelow = $0000000F;

// Constants for enum KsoCommandBarButtonHyperlinkType
type
  KsoCommandBarButtonHyperlinkType = TOleEnum;
const
  ksoCommandBarButtonHyperlinkNone = $00000000;
  ksoCommandBarButtonHyperlinkOpen = $00000001;
  ksoCommandBarButtonHyperlinkInsertPicture = $00000002;

// Constants for enum KsoOLEMenuGroup
type
  KsoOLEMenuGroup = TOleEnum;
const
  ksoOLEMenuGroupNone = $FFFFFFFF;
  ksoOLEMenuGroupFile = $00000000;
  ksoOLEMenuGroupEdit = $00000001;
  ksoOLEMenuGroupContainer = $00000002;
  ksoOLEMenuGroupObject = $00000003;
  ksoOLEMenuGroupWindow = $00000004;
  ksoOLEMenuGroupHelp = $00000005;

// Constants for enum KsoComboStyle
type
  KsoComboStyle = TOleEnum;
const
  ksoComboNormal = $00000000;
  ksoComboLabel = $00000001;

// Constants for enum KsoKeyCategory
type
  KsoKeyCategory = TOleEnum;
const
  ksoKeyCategoryNil = $FFFFFFFF;
  ksoKeyCategoryDisable = $00000000;
  ksoKeyCategoryCommand = $00000001;
  ksoKeyCategoryMacro = $00000002;
  ksoKeyCategoryFont = $00000003;
  ksoKeyCategoryAutoText = $00000004;
  ksoKeyCategoryStyle = $00000005;
  ksoKeyCategorySymbol = $00000006;
  ksoKeyCategoryPrefix = $00000007;

// Constants for enum KsoKey
type
  KsoKey = TOleEnum;
const
  ksoNoKey = $000000FF;
  ksoKeyShift = $00000100;
  ksoKeyControl = $00000200;
  ksoKeyCommand = $00000200;
  ksoKeyAlt = $00000400;
  ksoKeyOption = $00000400;
  ksoKeyA = $00000041;
  ksoKeyB = $00000042;
  ksoKeyC = $00000043;
  ksoKeyD = $00000044;
  ksoKeyE = $00000045;
  ksoKeyF = $00000046;
  ksoKeyG = $00000047;
  ksoKeyH = $00000048;
  ksoKeyI = $00000049;
  ksoKeyJ = $0000004A;
  ksoKeyK = $0000004B;
  ksoKeyL = $0000004C;
  ksoKeyM = $0000004D;
  ksoKeyN = $0000004E;
  ksoKeyO = $0000004F;
  ksoKeyP = $00000050;
  ksoKeyQ = $00000051;
  ksoKeyR = $00000052;
  ksoKeyS = $00000053;
  ksoKeyT = $00000054;
  ksoKeyU = $00000055;
  ksoKeyV = $00000056;
  ksoKeyW = $00000057;
  ksoKeyX = $00000058;
  ksoKeyY = $00000059;
  ksoKeyZ = $0000005A;
  ksoKey0 = $00000030;
  ksoKey1 = $00000031;
  ksoKey2 = $00000032;
  ksoKey3 = $00000033;
  ksoKey4 = $00000034;
  ksoKey5 = $00000035;
  ksoKey6 = $00000036;
  ksoKey7 = $00000037;
  ksoKey8 = $00000038;
  ksoKey9 = $00000039;
  ksoKeyBackspace = $00000008;
  ksoKeyTab = $00000009;
  ksoKeyNumeric5Special = $0000000C;
  ksoKeyReturn = $0000000D;
  ksoKeyPause = $00000013;
  ksoKeyEsc = $0000001B;
  ksoKeySpacebar = $00000020;
  ksoKeyPageUp = $00000021;
  ksoKeyPageDown = $00000022;
  ksoKeyEnd = $00000023;
  ksoKeyHome = $00000024;
  ksoKeyInsert = $0000002D;
  ksoKeyDelete = $0000002E;
  ksoKeyNumeric0 = $00000060;
  ksoKeyNumeric1 = $00000061;
  ksoKeyNumeric2 = $00000062;
  ksoKeyNumeric3 = $00000063;
  ksoKeyNumeric4 = $00000064;
  ksoKeyNumeric5 = $00000065;
  ksoKeyNumeric6 = $00000066;
  ksoKeyNumeric7 = $00000067;
  ksoKeyNumeric8 = $00000068;
  ksoKeyNumeric9 = $00000069;
  ksoKeyNumericMultiply = $0000006A;
  ksoKeyNumericAdd = $0000006B;
  ksoKeyNumericSubtract = $0000006D;
  ksoKeyNumericDecimal = $0000006E;
  ksoKeyNumericDivide = $0000006F;
  ksoKeyF1 = $00000070;
  ksoKeyF2 = $00000071;
  ksoKeyF3 = $00000072;
  ksoKeyF4 = $00000073;
  ksoKeyF5 = $00000074;
  ksoKeyF6 = $00000075;
  ksoKeyF7 = $00000076;
  ksoKeyF8 = $00000077;
  ksoKeyF9 = $00000078;
  ksoKeyF10 = $00000079;
  ksoKeyF11 = $0000007A;
  ksoKeyF12 = $0000007B;
  ksoKeyF13 = $0000007C;
  ksoKeyF14 = $0000007D;
  ksoKeyF15 = $0000007E;
  ksoKeyF16 = $0000007F;
  ksoKeyScrollLock = $00000091;
  ksoKeySemiColon = $000000BA;
  ksoKeyEquals = $000000BB;
  ksoKeyComma = $000000BC;
  ksoKeyHyphen = $000000BD;
  ksoKeyPeriod = $000000BE;
  ksoKeySlash = $000000BF;
  ksoKeyBackSingleQuote = $000000C0;
  ksoKeyOpenSquareBrace = $000000DB;
  ksoKeyBackSlash = $000000DC;
  ksoKeyCloseSquareBrace = $000000DD;
  ksoKeySingleQuote = $000000DE;

// Constants for enum MEDIUMTYPE
type
  MEDIUMTYPE = TOleEnum;
const
  MEDIUM_TYPE_FILE = $00000002;
  MEDIUM_TYPE_ISTREAM = $00000004;
  MEDIUM_TYPE_ISTORAGE = $00000008;
  MEDIUM_TYPE_NULL = $00000000;

// Constants for enum AdvApiType
type
  AdvApiType = TOleEnum;
const
  AdvApiType_AdvApplication = $00000001;
  AdvApiType_AdvAddins = $00000002;
  AdvApiType_AdvRight = $00000003;
  AdvApiType_AdvSeal = $00000004;

// Constants for enum AdvCertType
type
  AdvCertType = TOleEnum;
const
  AdvCertType_None = $00000000;
  AdvCertType_AdvRight = $00000001;
  AdvCertType_AdvSeal = $00000002;

// Constants for enum KsoSecurityLevel
type
  KsoSecurityLevel = TOleEnum;
const
  ksoSecurity_None = $00000000;
  ksoSecurity_Low = $00000001;
  ksoSecurity_Medium = $00000002;
  ksoSecurity_High = $00000003;

// Constants for enum KsoLineDashStyle
type
  KsoLineDashStyle = TOleEnum;
const
  KsoLineDashStyleMixed = $FFFFFFFE;
  KsoLineSolid = $00000001;
  KsoLineSquareDot = $00000002;
  KsoLineRoundDot = $00000003;
  KsoLineDash = $00000004;
  KsoLineDashDot = $00000005;
  KsoLineDashDotDot = $00000006;
  KsoLineLongDash = $00000007;
  KsoLineLongDashDot = $00000008;

// Constants for enum KsoLineStyle
type
  KsoLineStyle = TOleEnum;
const
  KsoLineStyleMixed = $FFFFFFFE;
  KsoLineSingle = $00000001;
  KsoLineThinThin = $00000002;
  KsoLineThinThick = $00000003;
  KsoLineThickThin = $00000004;
  KsoLineThickBetweenThin = $00000005;

// Constants for enum KsoArrowheadStyle
type
  KsoArrowheadStyle = TOleEnum;
const
  KsoArrowheadStyleMixed = $FFFFFFFE;
  KsoArrowheadNone = $00000001;
  KsoArrowheadTriangle = $00000002;
  KsoArrowheadOpen = $00000003;
  KsoArrowheadStealth = $00000004;
  KsoArrowheadDiamond = $00000005;
  KsoArrowheadOval = $00000006;

// Constants for enum KsoArrowheadWidth
type
  KsoArrowheadWidth = TOleEnum;
const
  KsoArrowheadWidthMixed = $FFFFFFFE;
  KsoArrowheadNarrow = $00000001;
  KsoArrowheadWidthMedium = $00000002;
  KsoArrowheadWide = $00000003;

// Constants for enum KsoArrowheadLength
type
  KsoArrowheadLength = TOleEnum;
const
  KsoArrowheadLengthMixed = $FFFFFFFE;
  KsoArrowheadShort = $00000001;
  KsoArrowheadLengthMedium = $00000002;
  KsoArrowheadLong = $00000003;

// Constants for enum KsoFillType
type
  KsoFillType = TOleEnum;
const
  KsoFillMixed = $FFFFFFFE;
  KsoFillSolid = $00000001;
  KsoFillPatterned = $00000002;
  KsoFillGradient = $00000003;
  KsoFillTextured = $00000004;
  KsoFillBackground = $00000005;
  KsoFillPicture = $00000006;

// Constants for enum KsoGradientStyle
type
  KsoGradientStyle = TOleEnum;
const
  KsoGradientMixed = $FFFFFFFE;
  KsoGradientHorizontal = $00000001;
  KsoGradientVertical = $00000002;
  KsoGradientDiagonalUp = $00000003;
  KsoGradientDiagonalDown = $00000004;
  KsoGradientFromCorner = $00000005;
  KsoGradientFromTitle = $00000006;
  KsoGradientFromCenter = $00000007;

// Constants for enum KsoGradientColorType
type
  KsoGradientColorType = TOleEnum;
const
  KsoGradientColorMixed = $FFFFFFFE;
  KsoGradientOneColor = $00000001;
  KsoGradientTwoColors = $00000002;
  KsoGradientPresetColors = $00000003;

// Constants for enum KsoTextureType
type
  KsoTextureType = TOleEnum;
const
  KsoTextureTypeMixed = $FFFFFFFE;
  KsoTexturePreset = $00000001;
  KsoTextureUserDefined = $00000002;

// Constants for enum KsoPresetTexture
type
  KsoPresetTexture = TOleEnum;
const
  KsoPresetTextureMixed = $FFFFFFFE;
  KsoTexturePane1 = $00000001;
  KsoTexturePane2 = $00000002;
  KsoTextureTraditional1 = $00000003;
  KsoTextureTraditional2 = $00000004;
  KsoTextureCrossband = $00000005;
  KsoTextureAnimal_skin = $00000006;
  KsoTextureAoarse_cloth = $00000007;
  KsoTextureKingsoft = $00000008;
  KsoTexturePaper1 = $00000009;
  KsoTexturePaper2 = $0000000A;
  KsoTexturePane_woven = $0000000B;
  KsoTextureOld_cottonfabric = $0000000C;
  KsoTextureStar_sky = $0000000D;
  KsoTextureColored_paper1 = $0000000E;
  KsoTextureColored_paper2 = $0000000F;
  KsoTextureColored_paper3 = $00000010;
  KsoTextureWeave = $00000011;
  KsoTextureNap_list = $00000012;
  KsoTextureFell = $00000013;
  KsoTextureWater = $00000014;
  KsoTextureEarth1 = $00000015;
  KsoTextureEarth2 = $00000016;
  KsoTextureCircle = $00000017;
  KsoTextureTwine = $00000018;

// Constants for enum KsoPatternType
type
  KsoPatternType = TOleEnum;
const
  KsoPatternMixed = $FFFFFFFE;
  KsoPattern5Percent = $00000001;
  KsoPattern10Percent = $00000002;
  KsoPattern20Percent = $00000003;
  KsoPattern25Percent = $00000004;
  KsoPattern30Percent = $00000005;
  KsoPattern40Percent = $00000006;
  KsoPattern50Percent = $00000007;
  KsoPattern60Percent = $00000008;
  KsoPattern70Percent = $00000009;
  KsoPattern75Percent = $0000000A;
  KsoPattern80Percent = $0000000B;
  KsoPattern90Percent = $0000000C;
  KsoPatternDarkHorizontal = $0000000D;
  KsoPatternDarkVertical = $0000000E;
  KsoPatternDarkDownwardDiagonal = $0000000F;
  KsoPatternDarkUpwardDiagonal = $00000010;
  KsoPatternSmallCheckerBoard = $00000011;
  KsoPatternTrellis = $00000012;
  KsoPatternLightHorizontal = $00000013;
  KsoPatternLightVertical = $00000014;
  KsoPatternLightDownwardDiagonal = $00000015;
  KsoPatternLightUpwardDiagonal = $00000016;
  KsoPatternSmallGrid = $00000017;
  KsoPatternDottedDiamond = $00000018;
  KsoPatternWideDownwardDiagonal = $00000019;
  KsoPatternWideUpwardDiagonal = $0000001A;
  KsoPatternDashedUpwardDiagonal = $0000001B;
  KsoPatternDashedDownwardDiagonal = $0000001C;
  KsoPatternNarrowVertical = $0000001D;
  KsoPatternNarrowHorizontal = $0000001E;
  KsoPatternDashedVertical = $0000001F;
  KsoPatternDashedHorizontal = $00000020;
  KsoPatternLargeConfetti = $00000021;
  KsoPatternLargeGrid = $00000022;
  KsoPatternHorizontalBrick = $00000023;
  KsoPatternLargeCheckerBoard = $00000024;
  KsoPatternSmallConfetti = $00000025;
  KsoPatternZigZag = $00000026;
  KsoPatternSolidDiamond = $00000027;
  KsoPatternDiagonalBrick = $00000028;
  KsoPatternOutlinedDiamond = $00000029;
  KsoPatternPlaid = $0000002A;
  KsoPatternSphere = $0000002B;
  KsoPatternWeave = $0000002C;
  KsoPatternDottedGrid = $0000002D;
  KsoPatternDivot = $0000002E;
  KsoPatternShingle = $0000002F;
  KsoPatternWave = $00000030;

// Constants for enum KsoPresetGradientType
type
  KsoPresetGradientType = TOleEnum;
const
  KsoPresetGradientMixed = $FFFFFFFE;
  KsoGradientEarlySunset = $00000001;
  KsoGradientLateSunset = $00000002;
  KsoGradientNightfall = $00000003;
  KsoGradientDaybreak = $00000004;
  KsoGradientHorizon = $00000005;
  KsoGradientDesert = $00000006;
  KsoGradientOcean = $00000007;
  KsoGradientCalmWater = $00000008;
  KsoGradientFire = $00000009;
  KsoGradientFog = $0000000A;
  KsoGradientMoss = $0000000B;
  KsoGradientPeacock = $0000000C;
  KsoGradientWheat = $0000000D;
  KsoGradientParchment = $0000000E;
  KsoGradientMahogany = $0000000F;
  KsoGradientRainbow = $00000010;
  KsoGradientRainbowII = $00000011;
  KsoGradientGold = $00000012;
  KsoGradientGoldII = $00000013;
  KsoGradientBrass = $00000014;
  KsoGradientChrome = $00000015;
  KsoGradientChromeII = $00000016;
  KsoGradientSilver = $00000017;
  KsoGradientSapphire = $00000018;
  KsoGradientSpring = $00000019;
  KsoGradientGreen = $0000001A;
  KsoGradientCoffee = $0000001B;
  KsoGradientMirage = $0000001C;
  KsoGradientCurtainOfNight = $0000001D;
  KsoGradientRed = $0000001E;

// Constants for enum KsoShadowType
type
  KsoShadowType = TOleEnum;
const
  KsoShadowMixed = $FFFFFFFE;
  KsoShadow1 = $00000001;
  KsoShadow2 = $00000002;
  KsoShadow3 = $00000003;
  KsoShadow4 = $00000004;
  KsoShadow5 = $00000005;
  KsoShadow6 = $00000006;
  KsoShadow7 = $00000007;
  KsoShadow8 = $00000008;
  KsoShadow9 = $00000009;
  KsoShadow10 = $0000000A;
  KsoShadow11 = $0000000B;
  KsoShadow12 = $0000000C;
  KsoShadow13 = $0000000D;
  KsoShadow14 = $0000000E;
  KsoShadow15 = $0000000F;
  KsoShadow16 = $00000010;
  KsoShadow17 = $00000011;
  KsoShadow18 = $00000012;
  KsoShadow19 = $00000013;
  KsoShadow20 = $00000014;

// Constants for enum KsoPresetTextEffect
type
  KsoPresetTextEffect = TOleEnum;
const
  ksoTextEffectMixed = $FFFFFFFE;
  ksoTextEffect1 = $00000000;
  ksoTextEffect2 = $00000001;
  ksoTextEffect3 = $00000002;
  ksoTextEffect4 = $00000003;
  ksoTextEffect5 = $00000004;
  ksoTextEffect6 = $00000005;
  ksoTextEffect7 = $00000006;
  ksoTextEffect8 = $00000007;
  ksoTextEffect9 = $00000008;
  ksoTextEffect10 = $00000009;
  ksoTextEffect11 = $0000000A;
  ksoTextEffect12 = $0000000B;
  ksoTextEffect13 = $0000000C;
  ksoTextEffect14 = $0000000D;
  ksoTextEffect15 = $0000000E;
  ksoTextEffect16 = $0000000F;
  ksoTextEffect17 = $00000010;
  ksoTextEffect18 = $00000011;
  ksoTextEffect19 = $00000012;
  ksoTextEffect20 = $00000013;
  ksoTextEffect21 = $00000014;
  ksoTextEffect22 = $00000015;
  ksoTextEffect23 = $00000016;
  ksoTextEffect24 = $00000017;
  ksoTextEffect25 = $00000018;
  ksoTextEffect26 = $00000019;
  ksoTextEffect27 = $0000001A;
  ksoTextEffect28 = $0000001B;
  ksoTextEffect29 = $0000001C;
  ksoTextEffect30 = $0000001D;

// Constants for enum KsoPresetTextEffectShape
type
  KsoPresetTextEffectShape = TOleEnum;
const
  ksoTextEffectShapeMixed = $FFFFFFFE;
  ksoTextEffectShapePlainText = $00000001;
  ksoTextEffectShapeStop = $00000002;
  ksoTextEffectShapeTriangleUp = $00000003;
  ksoTextEffectShapeTriangleDown = $00000004;
  ksoTextEffectShapeChevronUp = $00000005;
  ksoTextEffectShapeChevronDown = $00000006;
  ksoTextEffectShapeRingInside = $00000007;
  ksoTextEffectShapeRingOutside = $00000008;
  ksoTextEffectShapeArchUpCurve = $00000009;
  ksoTextEffectShapeArchDownCurve = $0000000A;
  ksoTextEffectShapeCircleCurve = $0000000B;
  ksoTextEffectShapeButtonCurve = $0000000C;
  ksoTextEffectShapeArchUpPour = $0000000D;
  ksoTextEffectShapeArchDownPour = $0000000E;
  ksoTextEffectShapeCirclePour = $0000000F;
  ksoTextEffectShapeButtonPour = $00000010;
  ksoTextEffectShapeCurveUp = $00000011;
  ksoTextEffectShapeCurveDown = $00000012;
  ksoTextEffectShapeCanUp = $00000013;
  ksoTextEffectShapeCanDown = $00000014;
  ksoTextEffectShapeWave1 = $00000015;
  ksoTextEffectShapeWave2 = $00000016;
  ksoTextEffectShapeDoubleWave1 = $00000017;
  ksoTextEffectShapeDoubleWave2 = $00000018;
  ksoTextEffectShapeInflate = $00000019;
  ksoTextEffectShapeDeflate = $0000001A;
  ksoTextEffectShapeInflateBottom = $0000001B;
  ksoTextEffectShapeDeflateBottom = $0000001C;
  ksoTextEffectShapeInflateTop = $0000001D;
  ksoTextEffectShapeDeflateTop = $0000001E;
  ksoTextEffectShapeDeflateInflate = $0000001F;
  ksoTextEffectShapeDeflateInflateDeflate = $00000020;
  ksoTextEffectShapeFadeRight = $00000021;
  ksoTextEffectShapeFadeLeft = $00000022;
  ksoTextEffectShapeFadeUp = $00000023;
  ksoTextEffectShapeFadeDown = $00000024;
  ksoTextEffectShapeSlantUp = $00000025;
  ksoTextEffectShapeSlantDown = $00000026;
  ksoTextEffectShapeCascadeUp = $00000027;
  ksoTextEffectShapeCascadeDown = $00000028;

// Constants for enum KsoTextEffectAlignment
type
  KsoTextEffectAlignment = TOleEnum;
const
  ksoTextEffectAlignmentMixed = $FFFFFFFE;
  ksoTextEffectAlignmentLeft = $00000001;
  ksoTextEffectAlignmentCentered = $00000002;
  ksoTextEffectAlignmentRight = $00000003;
  ksoTextEffectAlignmentLetterJustify = $00000004;
  ksoTextEffectAlignmentWordJustify = $00000005;
  ksoTextEffectAlignmentStretchJustify = $00000006;

// Constants for enum KsoPresetLightingDirection
type
  KsoPresetLightingDirection = TOleEnum;
const
  ksoPresetLightingDirectionMixed = $FFFFFFFE;
  ksoLightingTopLeft = $00000001;
  ksoLightingTop = $00000002;
  ksoLightingTopRight = $00000003;
  ksoLightingLeft = $00000004;
  ksoLightingNone = $00000005;
  ksoLightingRight = $00000006;
  ksoLightingBottomLeft = $00000007;
  ksoLightingBottom = $00000008;
  ksoLightingBottomRight = $00000009;

// Constants for enum KsoPresetLightingSoftness
type
  KsoPresetLightingSoftness = TOleEnum;
const
  ksoPresetLightingSoftnessMixed = $FFFFFFFE;
  ksoLightingDim = $00000001;
  ksoLightingNormal = $00000002;
  ksoLightingBright = $00000003;

// Constants for enum KsoPresetMaterial
type
  KsoPresetMaterial = TOleEnum;
const
  ksoPresetMaterialMixed = $FFFFFFFE;
  ksoMaterialMatte = $00000001;
  ksoMaterialPlastic = $00000002;
  ksoMaterialMetal = $00000003;
  ksoMaterialWireFrame = $00000004;

// Constants for enum KsoPresetExtrusionDirection
type
  KsoPresetExtrusionDirection = TOleEnum;
const
  ksoPresetExtrusionDirectionMixed = $FFFFFFFE;
  ksoExtrusionBottomRight = $00000001;
  ksoExtrusionBottom = $00000002;
  ksoExtrusionBottomLeft = $00000003;
  ksoExtrusionRight = $00000004;
  ksoExtrusionNone = $00000005;
  ksoExtrusionLeft = $00000006;
  ksoExtrusionTopRight = $00000007;
  ksoExtrusionTop = $00000008;
  ksoExtrusionTopLeft = $00000009;

// Constants for enum KsoPresetThreeDFormat
type
  KsoPresetThreeDFormat = TOleEnum;
const
  ksoPresetThreeDFormatMixed = $FFFFFFFE;
  ksoThreeD1 = $00000001;
  ksoThreeD2 = $00000002;
  ksoThreeD3 = $00000003;
  ksoThreeD4 = $00000004;
  ksoThreeD5 = $00000005;
  ksoThreeD6 = $00000006;
  ksoThreeD7 = $00000007;
  ksoThreeD8 = $00000008;
  ksoThreeD9 = $00000009;
  ksoThreeD10 = $0000000A;
  ksoThreeD11 = $0000000B;
  ksoThreeD12 = $0000000C;
  ksoThreeD13 = $0000000D;
  ksoThreeD14 = $0000000E;
  ksoThreeD15 = $0000000F;
  ksoThreeD16 = $00000010;
  ksoThreeD17 = $00000011;
  ksoThreeD18 = $00000012;
  ksoThreeD19 = $00000013;
  ksoThreeD20 = $00000014;

// Constants for enum KsoExtrusionColorType
type
  KsoExtrusionColorType = TOleEnum;
const
  ksoExtrusionColorTypeMixed = $FFFFFFFE;
  ksoExtrusionColorAutomatic = $00000001;
  ksoExtrusionColorCustom = $00000002;

// Constants for enum KsoAlignCmd
type
  KsoAlignCmd = TOleEnum;
const
  ksoAlignLefts = $00000000;
  ksoAlignCenters = $00000001;
  ksoAlignRights = $00000002;
  ksoAlignTops = $00000003;
  ksoAlignMiddles = $00000004;
  ksoAlignBottoms = $00000005;

// Constants for enum KsoDistributeCmd
type
  KsoDistributeCmd = TOleEnum;
const
  ksoDistributeHorizontally = $00000000;
  ksoDistributeVertically = $00000001;

// Constants for enum KsoConnectorType
type
  KsoConnectorType = TOleEnum;
const
  ksoConnectorTypeMixed = $FFFFFFFE;
  ksoConnectorStraight = $00000001;
  ksoConnectorElbow = $00000002;
  ksoConnectorCurve = $00000003;

// Constants for enum KsoHorizontalAnchor
type
  KsoHorizontalAnchor = TOleEnum;
const
  ksoHorizontalAnchorMixed = $FFFFFFFE;
  ksoAnchorNone = $00000001;
  ksoAnchorCenter = $00000002;

// Constants for enum KsoVerticalAnchor
type
  KsoVerticalAnchor = TOleEnum;
const
  ksoVerticalAnchorMixed = $FFFFFFFE;
  ksoAnchorTop = $00000001;
  ksoAnchorTopBaseline = $00000002;
  ksoAnchorMiddle = $00000003;
  ksoAnchorBottom = $00000004;
  ksoAnchorBottomBaseLine = $00000005;

// Constants for enum KsoOrientation
type
  KsoOrientation = TOleEnum;
const
  ksoOrientationMixed = $FFFFFFFE;
  ksoOrientationHorizontal = $00000001;
  ksoOrientationVertical = $00000002;

// Constants for enum KsoZOrderCmd
type
  KsoZOrderCmd = TOleEnum;
const
  ksoBringToFront = $00000000;
  ksoSendToBack = $00000001;
  ksoBringForward = $00000002;
  ksoSendBackward = $00000003;
  ksoBringInFrontOfText = $00000004;
  ksoSendBehindText = $00000005;

// Constants for enum KsoSegmentType
type
  KsoSegmentType = TOleEnum;
const
  ksoSegmentLine = $00000000;
  ksoSegmentCurve = $00000001;

// Constants for enum KsoEditingType
type
  KsoEditingType = TOleEnum;
const
  ksoEditingAuto = $00000000;
  ksoEditingCorner = $00000001;
  ksoEditingSmooth = $00000002;
  ksoEditingSymmetric = $00000003;

// Constants for enum KsoAutoShapeType
type
  KsoAutoShapeType = TOleEnum;
const
  ksoShapeMixed = $FFFFFFFE;
  ksoShapeRectangle = $00000001;
  ksoShapeParallelogram = $00000002;
  ksoShapeTrapezoid = $00000003;
  ksoShapeDiamond = $00000004;
  ksoShapeRoundedRectangle = $00000005;
  ksoShapeOctagon = $00000006;
  ksoShapeIsoscelesTriangle = $00000007;
  ksoShapeRightTriangle = $00000008;
  ksoShapeOval = $00000009;
  ksoShapeHexagon = $0000000A;
  ksoShapeCross = $0000000B;
  ksoShapeRegularPentagon = $0000000C;
  ksoShapeCan = $0000000D;
  ksoShapeCube = $0000000E;
  ksoShapeBevel = $0000000F;
  ksoShapeFoldedCorner = $00000010;
  ksoShapeSmileyFace = $00000011;
  ksoShapeDonut = $00000012;
  ksoShapeNoSymbol = $00000013;
  ksoShapeBlockArc = $00000014;
  ksoShapeHeart = $00000015;
  ksoShapeLightningBolt = $00000016;
  ksoShapeSun = $00000017;
  ksoShapeMoon = $00000018;
  ksoShapeArc = $00000019;
  ksoShapeDoubleBracket = $0000001A;
  ksoShapeDoubleBrace = $0000001B;
  ksoShapePlaque = $0000001C;
  ksoShapeLeftBracket = $0000001D;
  ksoShapeRightBracket = $0000001E;
  ksoShapeLeftBrace = $0000001F;
  ksoShapeRightBrace = $00000020;
  ksoShapeRightArrow = $00000021;
  ksoShapeLeftArrow = $00000022;
  ksoShapeUpArrow = $00000023;
  ksoShapeDownArrow = $00000024;
  ksoShapeLeftRightArrow = $00000025;
  ksoShapeUpDownArrow = $00000026;
  ksoShapeQuadArrow = $00000027;
  ksoShapeLeftRightUpArrow = $00000028;
  ksoShapeBentArrow = $00000029;
  ksoShapeUTurnArrow = $0000002A;
  ksoShapeLeftUpArrow = $0000002B;
  ksoShapeBentUpArrow = $0000002C;
  ksoShapeCurvedRightArrow = $0000002D;
  ksoShapeCurvedLeftArrow = $0000002E;
  ksoShapeCurvedUpArrow = $0000002F;
  ksoShapeCurvedDownArrow = $00000030;
  ksoShapeStripedRightArrow = $00000031;
  ksoShapeNotchedRightArrow = $00000032;
  ksoShapePentagon = $00000033;
  ksoShapeChevron = $00000034;
  ksoShapeRightArrowCallout = $00000035;
  ksoShapeLeftArrowCallout = $00000036;
  ksoShapeUpArrowCallout = $00000037;
  ksoShapeDownArrowCallout = $00000038;
  ksoShapeLeftRightArrowCallout = $00000039;
  ksoShapeUpDownArrowCallout = $0000003A;
  ksoShapeQuadArrowCallout = $0000003B;
  ksoShapeCircularArrow = $0000003C;
  ksoShapeFlowchartProcess = $0000003D;
  ksoShapeFlowchartAlternateProcess = $0000003E;
  ksoShapeFlowchartDecision = $0000003F;
  ksoShapeFlowchartData = $00000040;
  ksoShapeFlowchartPredefinedProcess = $00000041;
  ksoShapeFlowchartInternalStorage = $00000042;
  ksoShapeFlowchartDocument = $00000043;
  ksoShapeFlowchartMultidocument = $00000044;
  ksoShapeFlowchartTerminator = $00000045;
  ksoShapeFlowchartPreparation = $00000046;
  ksoShapeFlowchartManualInput = $00000047;
  ksoShapeFlowchartManualOperation = $00000048;
  ksoShapeFlowchartConnector = $00000049;
  ksoShapeFlowchartOffpageConnector = $0000004A;
  ksoShapeFlowchartCard = $0000004B;
  ksoShapeFlowchartPunchedTape = $0000004C;
  ksoShapeFlowchartSummingJunction = $0000004D;
  ksoShapeFlowchartOr = $0000004E;
  ksoShapeFlowchartCollate = $0000004F;
  ksoShapeFlowchartSort = $00000050;
  ksoShapeFlowchartExtract = $00000051;
  ksoShapeFlowchartMerge = $00000052;
  ksoShapeFlowchartStoredData = $00000053;
  ksoShapeFlowchartDelay = $00000054;
  ksoShapeFlowchartSequentialAccessStorage = $00000055;
  ksoShapeFlowchartMagneticDisk = $00000056;
  ksoShapeFlowchartDirectAccessStorage = $00000057;
  ksoShapeFlowchartDisplay = $00000058;
  ksoShapeExplosion1 = $00000059;
  ksoShapeExplosion2 = $0000005A;
  ksoShape4pointStar = $0000005B;
  ksoShape5pointStar = $0000005C;
  ksoShape8pointStar = $0000005D;
  ksoShape16pointStar = $0000005E;
  ksoShape24pointStar = $0000005F;
  ksoShape32pointStar = $00000060;
  ksoShapeUpRibbon = $00000061;
  ksoShapeDownRibbon = $00000062;
  ksoShapeCurvedUpRibbon = $00000063;
  ksoShapeCurvedDownRibbon = $00000064;
  ksoShapeVerticalScroll = $00000065;
  ksoShapeHorizontalScroll = $00000066;
  ksoShapeWave = $00000067;
  ksoShapeDoubleWave = $00000068;
  ksoShapeRectangularCallout = $00000069;
  ksoShapeRoundedRectangularCallout = $0000006A;
  ksoShapeOvalCallout = $0000006B;
  ksoShapeCloudCallout = $0000006C;
  ksoShapeLineCallout1 = $0000006D;
  ksoShapeLineCallout2 = $0000006E;
  ksoShapeLineCallout3 = $0000006F;
  ksoShapeLineCallout4 = $00000070;
  ksoShapeLineCallout1AccentBar = $00000071;
  ksoShapeLineCallout2AccentBar = $00000072;
  ksoShapeLineCallout3AccentBar = $00000073;
  ksoShapeLineCallout4AccentBar = $00000074;
  ksoShapeLineCallout1NoBorder = $00000075;
  ksoShapeLineCallout2NoBorder = $00000076;
  ksoShapeLineCallout3NoBorder = $00000077;
  ksoShapeLineCallout4NoBorder = $00000078;
  ksoShapeLineCallout1BorderandAccentBar = $00000079;
  ksoShapeLineCallout2BorderandAccentBar = $0000007A;
  ksoShapeLineCallout3BorderandAccentBar = $0000007B;
  ksoShapeLineCallout4BorderandAccentBar = $0000007C;
  ksoShapeActionButtonCustom = $0000007D;
  ksoShapeActionButtonHome = $0000007E;
  ksoShapeActionButtonHelp = $0000007F;
  ksoShapeActionButtonInformation = $00000080;
  ksoShapeActionButtonBackorPrevious = $00000081;
  ksoShapeActionButtonForwardorNext = $00000082;
  ksoShapeActionButtonBeginning = $00000083;
  ksoShapeActionButtonEnd = $00000084;
  ksoShapeActionButtonReturn = $00000085;
  ksoShapeActionButtonDocument = $00000086;
  ksoShapeActionButtonSound = $00000087;
  ksoShapeActionButtonMovie = $00000088;
  ksoShapeBalloon = $00000089;
  ksoShapeNotPrimitive = $0000008A;

// Constants for enum KsoShapeType
type
  KsoShapeType = TOleEnum;
const
  ksoShapeTypeMixed = $FFFFFFFE;
  ksoAutoShape = $00000001;
  ksoCallout = $00000002;
  ksoChart = $00000003;
  ksoComment = $00000004;
  ksoFreeform = $00000005;
  ksoGroup = $00000006;
  ksoEmbeddedOLEObject = $00000007;
  ksoFormControl = $00000008;
  ksoLine = $00000009;
  ksoLinkedOLEObject = $0000000A;
  ksoLinkedPicture = $0000000B;
  ksoOLEControlObject = $0000000C;
  ksoPicture = $0000000D;
  ksoPlaceholder = $0000000E;
  ksoTextEffect = $0000000F;
  ksoMedia = $00000010;
  ksoTextBox = $00000011;
  ksoScriptAnchor = $00000012;
  ksoTable = $00000013;
  ksoCanvas = $00000014;
  ksoDiagram = $00000015;
  ksoInk = $00000016;
  ksoInkComment = $00000017;
  ksoIndependentControl = $00000018;

// Constants for enum KsoFlipCmd
type
  KsoFlipCmd = TOleEnum;
const
  ksoFlipHorizontal = $00000000;
  ksoFlipVertical = $00000001;

// Constants for enum KsoTriState
type
  KsoTriState = TOleEnum;
const
  ksoTrue = $FFFFFFFF;
  ksoFalse = $00000000;
  ksoCTrue = $00000001;
  ksoTriStateToggle = $FFFFFFFD;
  ksoTriStateMixed = $FFFFFFFE;

// Constants for enum KsoColorType
type
  KsoColorType = TOleEnum;
const
  ksoColorTypeMixed = $FFFFFFFE;
  ksoColorTypeRGB = $00000001;
  ksoColorTypeScheme = $00000002;
  ksoColorTypeCMYK = $00000003;
  ksoColorTypeCMS = $00000004;
  ksoColorTypeInk = $00000005;

// Constants for enum KsoPictureColorType
type
  KsoPictureColorType = TOleEnum;
const
  ksoPictureMixed = $FFFFFFFE;
  ksoPictureAutomatic = $00000001;
  ksoPictureGrayscale = $00000002;
  ksoPictureBlackAndWhite = $00000003;
  ksoPictureWatermark = $00000004;

// Constants for enum KsoCalloutAngleType
type
  KsoCalloutAngleType = TOleEnum;
const
  ksoCalloutAngleMixed = $FFFFFFFE;
  ksoCalloutAngleAutomatic = $00000001;
  ksoCalloutAngle30 = $00000002;
  ksoCalloutAngle45 = $00000003;
  ksoCalloutAngle60 = $00000004;
  ksoCalloutAngle90 = $00000005;

// Constants for enum KsoCalloutDropType
type
  KsoCalloutDropType = TOleEnum;
const
  ksoCalloutDropMixed = $FFFFFFFE;
  ksoCalloutDropCustom = $00000001;
  ksoCalloutDropTop = $00000002;
  ksoCalloutDropCenter = $00000003;
  ksoCalloutDropBottom = $00000004;

// Constants for enum KsoCalloutType
type
  KsoCalloutType = TOleEnum;
const
  ksoCalloutMixed = $FFFFFFFE;
  ksoCalloutOne = $00000001;
  ksoCalloutTwo = $00000002;
  ksoCalloutThree = $00000003;
  ksoCalloutFour = $00000004;

// Constants for enum KsoBlackWhiteMode
type
  KsoBlackWhiteMode = TOleEnum;
const
  ksoBlackWhiteMixed = $FFFFFFFE;
  ksoBlackWhiteAutomatic = $00000001;
  ksoBlackWhiteGrayScale = $00000002;
  ksoBlackWhiteLightGrayScale = $00000003;
  ksoBlackWhiteInverseGrayScale = $00000004;
  ksoBlackWhiteGrayOutline = $00000005;
  ksoBlackWhiteBlackTextAndLine = $00000006;
  ksoBlackWhiteHighContrast = $00000007;
  ksoBlackWhiteBlack = $00000008;
  ksoBlackWhiteWhite = $00000009;
  ksoBlackWhiteDontShow = $0000000A;

// Constants for enum KsoMixedType
type
  KsoMixedType = TOleEnum;
const
  ksoIntegerMixed = $00008000;
  ksoSingleMixed = $44653600;

// Constants for enum KsoTextOrientation
type
  KsoTextOrientation = TOleEnum;
const
  ksoTextOrientationMixed = $FFFFFFFE;
  ksoTextOrientationHorizontal = $00000001;
  ksoTextOrientationUpward = $00000002;
  ksoTextOrientationDownward = $00000003;
  ksoTextOrientationVerticalFarEast = $00000004;
  ksoTextOrientationVertical = $00000005;
  ksoTextOrientationHorizontalRotatedFarEast = $00000006;
  ksoTextOrientationVerticalTrans = $00000007;

// Constants for enum KsoScaleFrom
type
  KsoScaleFrom = TOleEnum;
const
  ksoScaleFromTopLeft = $00000000;
  ksoScaleFromMiddle = $00000001;
  ksoScaleFromBottomRight = $00000002;

// Constants for enum KsoScriptLanguage
type
  KsoScriptLanguage = TOleEnum;
const
  ksoScriptLanguageJava = $00000001;
  ksoScriptLanguageVisualBasic = $00000002;
  ksoScriptLanguageASP = $00000003;
  ksoScriptLanguageOther = $00000004;

// Constants for enum KsoScriptLocation
type
  KsoScriptLocation = TOleEnum;
const
  ksoScriptLocationInHead = $00000001;
  ksoScriptLocationInBody = $00000002;

// Constants for enum KsoAnimEffect
type
  KsoAnimEffect = TOleEnum;
const
  ksoAnimEffectCustom = $00000000;
  ksoAnimEffectAppear = $00000001;
  ksoAnimEffectFly = $00000002;
  ksoAnimEffectBlinds = $00000003;
  ksoAnimEffectBox = $00000004;
  ksoAnimEffectCheckerboard = $00000005;
  ksoAnimEffectCircle = $00000006;
  ksoAnimEffectCrawl = $00000007;
  ksoAnimEffectDiamond = $00000008;
  ksoAnimEffectDissolve = $00000009;
  ksoAnimEffectFade = $0000000A;
  ksoAnimEffectFlashOnce = $0000000B;
  ksoAnimEffectPeek = $0000000C;
  ksoAnimEffectPlus = $0000000D;
  ksoAnimEffectRandomBars = $0000000E;
  ksoAnimEffectSpiral = $0000000F;
  ksoAnimEffectSplit = $00000010;
  ksoAnimEffectStretch = $00000011;
  ksoAnimEffectStrips = $00000012;
  ksoAnimEffectSwivel = $00000013;
  ksoAnimEffectWedge = $00000014;
  ksoAnimEffectWheel = $00000015;
  ksoAnimEffectWipe = $00000016;
  ksoAnimEffectZoom = $00000017;
  ksoAnimEffectRandomEffects = $00000018;
  ksoAnimEffectBoomerang = $00000019;
  ksoAnimEffectBounce = $0000001A;
  ksoAnimEffectColorReveal = $0000001B;
  ksoAnimEffectCredits = $0000001C;
  ksoAnimEffectEaseIn = $0000001D;
  ksoAnimEffectFloat = $0000001E;
  ksoAnimEffectGrowAndTurn = $0000001F;
  ksoAnimEffectLightSpeed = $00000020;
  ksoAnimEffectPinwheel = $00000021;
  ksoAnimEffectRiseUp = $00000022;
  ksoAnimEffectSwish = $00000023;
  ksoAnimEffectThinLine = $00000024;
  ksoAnimEffectUnfold = $00000025;
  ksoAnimEffectWhip = $00000026;
  ksoAnimEffectAscend = $00000027;
  ksoAnimEffectCenterRevolve = $00000028;
  ksoAnimEffectFadedSwivel = $00000029;
  ksoAnimEffectDescend = $0000002A;
  ksoAnimEffectSling = $0000002B;
  ksoAnimEffectSpinner = $0000002C;
  ksoAnimEffectStretchy = $0000002D;
  ksoAnimEffectZip = $0000002E;
  ksoAnimEffectArcUp = $0000002F;
  ksoAnimEffectFadedZoom = $00000030;
  ksoAnimEffectGlide = $00000031;
  ksoAnimEffectExpand = $00000032;
  ksoAnimEffectFlip = $00000033;
  ksoAnimEffectShimmer = $00000034;
  ksoAnimEffectFold = $00000035;
  ksoAnimEffectChangeFillColor = $00000036;
  ksoAnimEffectChangeFont = $00000037;
  ksoAnimEffectChangeFontColor = $00000038;
  ksoAnimEffectChangeFontSize = $00000039;
  ksoAnimEffectChangeFontStyle = $0000003A;
  ksoAnimEffectGrowShrink = $0000003B;
  ksoAnimEffectChangeLineColor = $0000003C;
  ksoAnimEffectSpin = $0000003D;
  ksoAnimEffectTransparency = $0000003E;
  ksoAnimEffectBoldFlash = $0000003F;
  ksoAnimEffectBlast = $00000040;
  ksoAnimEffectBoldReveal = $00000041;
  ksoAnimEffectBrushOnColor = $00000042;
  ksoAnimEffectBrushOnUnderline = $00000043;
  ksoAnimEffectColorBlend = $00000044;
  ksoAnimEffectColorWave = $00000045;
  ksoAnimEffectComplementaryColor = $00000046;
  ksoAnimEffectComplementaryColor2 = $00000047;
  ksoAnimEffectContrastingColor = $00000048;
  ksoAnimEffectDarken = $00000049;
  ksoAnimEffectDesaturate = $0000004A;
  ksoAnimEffectFlashBulb = $0000004B;
  ksoAnimEffectFlicker = $0000004C;
  ksoAnimEffectGrowWithColor = $0000004D;
  ksoAnimEffectLighten = $0000004E;
  ksoAnimEffectStyleEmphasis = $0000004F;
  ksoAnimEffectTeeter = $00000050;
  ksoAnimEffectVerticalGrow = $00000051;
  ksoAnimEffectWave = $00000052;
  ksoAnimEffectMediaPlay = $00000053;
  ksoAnimEffectMediaPause = $00000054;
  ksoAnimEffectMediaStop = $00000055;
  ksoAnimEffectPathCircle = $00000056;
  ksoAnimEffectPathRightTriangle = $00000057;
  ksoAnimEffectPathDiamond = $00000058;
  ksoAnimEffectPathHexagon = $00000059;
  ksoAnimEffectPath5PointStar = $0000005A;
  ksoAnimEffectPathCrescentMoon = $0000005B;
  ksoAnimEffectPathSquare = $0000005C;
  ksoAnimEffectPathTrapezoid = $0000005D;
  ksoAnimEffectPathHeart = $0000005E;
  ksoAnimEffectPathOctagon = $0000005F;
  ksoAnimEffectPath6PointStar = $00000060;
  ksoAnimEffectPathFootball = $00000061;
  ksoAnimEffectPathEqualTriangle = $00000062;
  ksoAnimEffectPathParallelogram = $00000063;
  ksoAnimEffectPathPentagon = $00000064;
  ksoAnimEffectPath4PointStar = $00000065;
  ksoAnimEffectPath8PointStar = $00000066;
  ksoAnimEffectPathTeardrop = $00000067;
  ksoAnimEffectPathPointyStar = $00000068;
  ksoAnimEffectPathCurvedSquare = $00000069;
  ksoAnimEffectPathCurvedX = $0000006A;
  ksoAnimEffectPathVerticalFigure8 = $0000006B;
  ksoAnimEffectPathCurvyStar = $0000006C;
  ksoAnimEffectPathLoopdeLoop = $0000006D;
  ksoAnimEffectPathBuzzsaw = $0000006E;
  ksoAnimEffectPathHorizontalFigure8 = $0000006F;
  ksoAnimEffectPathPeanut = $00000070;
  ksoAnimEffectPathFigure8Four = $00000071;
  ksoAnimEffectPathNeutron = $00000072;
  ksoAnimEffectPathSwoosh = $00000073;
  ksoAnimEffectPathBean = $00000074;
  ksoAnimEffectPathPlus = $00000075;
  ksoAnimEffectPathInvertedTriangle = $00000076;
  ksoAnimEffectPathInvertedSquare = $00000077;
  ksoAnimEffectPathLeft = $00000078;
  ksoAnimEffectPathTurnRight = $00000079;
  ksoAnimEffectPathArcDown = $0000007A;
  ksoAnimEffectPathZigzag = $0000007B;
  ksoAnimEffectPathSCurve2 = $0000007C;
  ksoAnimEffectPathSineWave = $0000007D;
  ksoAnimEffectPathBounceLeft = $0000007E;
  ksoAnimEffectPathDown = $0000007F;
  ksoAnimEffectPathTurnUp = $00000080;
  ksoAnimEffectPathArcUp = $00000081;
  ksoAnimEffectPathHeartbeat = $00000082;
  ksoAnimEffectPathSpiralRight = $00000083;
  ksoAnimEffectPathWave = $00000084;
  ksoAnimEffectPathCurvyLeft = $00000085;
  ksoAnimEffectPathDiagonalDownRight = $00000086;
  ksoAnimEffectPathTurnDown = $00000087;
  ksoAnimEffectPathArcLeft = $00000088;
  ksoAnimEffectPathFunnel = $00000089;
  ksoAnimEffectPathSpring = $0000008A;
  ksoAnimEffectPathBounceRight = $0000008B;
  ksoAnimEffectPathSpiralLeft = $0000008C;
  ksoAnimEffectPathDiagonalUpRight = $0000008D;
  ksoAnimEffectPathTurnUpRight = $0000008E;
  ksoAnimEffectPathArcRight = $0000008F;
  ksoAnimEffectPathSCurve1 = $00000090;
  ksoAnimEffectPathDecayingWave = $00000091;
  ksoAnimEffectPathCurvyRight = $00000092;
  ksoAnimEffectPathStairsDown = $00000093;
  ksoAnimEffectPathUp = $00000094;
  ksoAnimEffectPathRight = $00000095;
  ksoAnimEffectWink = $00000096;
  ksoAnimEffectFlash = $00000097;
  ksoAnimEffectPathCustom = $00000098;

// Constants for enum KsoAnimateByLevel
type
  KsoAnimateByLevel = TOleEnum;
const
  ksoAnimateLevelMixed = $FFFFFFFF;
  ksoAnimateLevelNone = $00000000;
  ksoAnimateTextByAllLevels = $00000001;
  ksoAnimateTextByFirstLevel = $00000002;
  ksoAnimateTextBySecondLevel = $00000003;
  ksoAnimateTextByThirdLevel = $00000004;
  ksoAnimateTextByFourthLevel = $00000005;
  ksoAnimateTextByFifthLevel = $00000006;
  ksoAnimateChartAllAtOnce = $00000007;
  ksoAnimateChartByCategory = $00000008;
  ksoAnimateChartByCategoryElements = $00000009;
  ksoAnimateChartBySeries = $0000000A;
  ksoAnimateChartBySeriesElements = $0000000B;
  ksoAnimateDiagramAllAtOnce = $0000000C;
  ksoAnimateDiagramDepthByNode = $0000000D;
  ksoAnimateDiagramDepthByBranch = $0000000E;
  ksoAnimateDiagramBreadthByNode = $0000000F;
  ksoAnimateDiagramBreadthByLevel = $00000010;
  ksoAnimateDiagramClockwise = $00000011;
  ksoAnimateDiagramClockwiseIn = $00000012;
  ksoAnimateDiagramClockwiseOut = $00000013;
  ksoAnimateDiagramCounterClockwise = $00000014;
  ksoAnimateDiagramCounterClockwiseIn = $00000015;
  ksoAnimateDiagramCounterClockwiseOut = $00000016;
  ksoAnimateDiagramInByRing = $00000017;
  ksoAnimateDiagramOutByRing = $00000018;
  ksoAnimateDiagramUp = $00000019;
  ksoAnimateDiagramDown = $0000001A;

// Constants for enum KsoAnimateRepeateType
type
  KsoAnimateRepeateType = TOleEnum;
const
  ksoAnimateUntilNextClick = $FFFFFFFE;
  ksoAnimateUntilEndOfSlide = $FFFFFFFF;

// Constants for enum KsoAnimTriggerType
type
  KsoAnimTriggerType = TOleEnum;
const
  ksoAnimTriggerMixed = $FFFFFFFF;
  ksoAnimTriggerNone = $00000000;
  ksoAnimTriggerOnPageClick = $00000001;
  ksoAnimTriggerWithPrevious = $00000002;
  ksoAnimTriggerAfterPrevious = $00000003;
  ksoAnimTriggerOnShapeClick = $00000004;

// Constants for enum KsoAnimAfterEffect
type
  KsoAnimAfterEffect = TOleEnum;
const
  ksoAnimAfterEffectMixed = $FFFFFFFF;
  ksoAnimAfterEffectNone = $00000000;
  ksoAnimAfterEffectDim = $00000001;
  ksoAnimAfterEffectHide = $00000002;
  ksoAnimAfterEffectHideOnNextClick = $00000003;

// Constants for enum KsoAnimTextUnitEffect
type
  KsoAnimTextUnitEffect = TOleEnum;
const
  ksoAnimTextUnitEffectMixed = $FFFFFFFF;
  ksoAnimTextUnitEffectByParagraph = $00000000;
  ksoAnimTextUnitEffectByCharacter = $00000001;
  ksoAnimTextUnitEffectByWord = $00000002;

// Constants for enum KsoAnimEffectRestart
type
  KsoAnimEffectRestart = TOleEnum;
const
  ksoAnimEffectRestartAlways = $00000001;
  ksoAnimEffectRestartWhenOff = $00000002;
  ksoAnimEffectRestartNever = $00000003;

// Constants for enum KsoAnimEffectAfter
type
  KsoAnimEffectAfter = TOleEnum;
const
  ksoAnimEffectAfterFreeze = $00000001;
  ksoAnimEffectAfterRemove = $00000002;
  ksoAnimEffectAfterHold = $00000003;
  ksoAnimEffectAfterTransition = $00000004;

// Constants for enum KsoAnimDirection
type
  KsoAnimDirection = TOleEnum;
const
  ksoAnimDirectionNone = $00000000;
  ksoAnimDirectionUp = $00000001;
  ksoAnimDirectionRight = $00000002;
  ksoAnimDirectionDown = $00000003;
  ksoAnimDirectionLeft = $00000004;
  ksoAnimDirectionOrdinalMask = $00000005;
  ksoAnimDirectionUpLeft = $00000006;
  ksoAnimDirectionUpRight = $00000007;
  ksoAnimDirectionDownRight = $00000008;
  ksoAnimDirectionDownLeft = $00000009;
  ksoAnimDirectionTop = $0000000A;
  ksoAnimDirectionBottom = $0000000B;
  ksoAnimDirectionTopLeft = $0000000C;
  ksoAnimDirectionTopRight = $0000000D;
  ksoAnimDirectionBottomRight = $0000000E;
  ksoAnimDirectionBottomLeft = $0000000F;
  ksoAnimDirectionHorizontal = $00000010;
  ksoAnimDirectionVertical = $00000011;
  ksoAnimDirectionAcross = $00000012;
  ksoAnimDirectionIn = $00000013;
  ksoAnimDirectionOut = $00000014;
  ksoAnimDirectionClockwise = $00000015;
  ksoAnimDirectionCounterclockwise = $00000016;
  ksoAnimDirectionHorizontalIn = $00000017;
  ksoAnimDirectionHorizontalOut = $00000018;
  ksoAnimDirectionVerticalIn = $00000019;
  ksoAnimDirectionVerticalOut = $0000001A;
  ksoAnimDirectionSlightly = $0000001B;
  ksoAnimDirectionCenter = $0000001C;
  ksoAnimDirectionInSlightly = $0000001D;
  ksoAnimDirectionInCenter = $0000001E;
  ksoAnimDirectionInBottom = $0000001F;
  ksoAnimDirectionOutSlightly = $00000020;
  ksoAnimDirectionOutCenter = $00000021;
  ksoAnimDirectionOutBottom = $00000022;
  ksoAnimDirectionFontBold = $00000023;
  ksoAnimDirectionFontItalic = $00000024;
  ksoAnimDirectionFontUnderline = $00000025;
  ksoAnimDirectionFontStrikethrough = $00000026;
  ksoAnimDirectionFontShadow = $00000027;
  ksoAnimDirectionFontAllCaps = $00000028;
  ksoAnimDirectionInstant = $00000029;
  ksoAnimDirectionGradual = $0000002A;
  ksoAnimDirectionCycleClockwise = $0000002B;
  ksoAnimDirectionCycleCounterclockwise = $0000002C;
  ksoAnimDirectionWheel1 = $0000002D;
  ksoAnimDirectionWheel2 = $0000002E;
  ksoAnimDirectionWheel3 = $0000002F;
  ksoAnimDirectionWheel4 = $00000030;
  ksoAnimDirectionWheel8 = $00000031;

// Constants for enum KsoAnimType
type
  KsoAnimType = TOleEnum;
const
  ksoAnimTypeMixed = $FFFFFFFE;
  ksoAnimTypeNone = $00000000;
  ksoAnimTypeMotion = $00000001;
  ksoAnimTypeColor = $00000002;
  ksoAnimTypeScale = $00000003;
  ksoAnimTypeRotation = $00000004;
  ksoAnimTypeProperty = $00000005;
  ksoAnimTypeCommand = $00000006;
  ksoAnimTypeFilter = $00000007;
  ksoAnimTypeSet = $00000008;

// Constants for enum KsoAnimAdditive
type
  KsoAnimAdditive = TOleEnum;
const
  ksoAnimAdditiveAddBase = $00000001;
  ksoAnimAdditiveAddSum = $00000002;

// Constants for enum KsoAnimAccumulate
type
  KsoAnimAccumulate = TOleEnum;
const
  ksoAnimAccumulateNone = $00000001;
  ksoAnimAccumulateAlways = $00000002;

// Constants for enum KsoAnimProperty
type
  KsoAnimProperty = TOleEnum;
const
  ksoAnimNone = $00000000;
  ksoAnimX = $00000001;
  ksoAnimY = $00000002;
  ksoAnimWidth = $00000003;
  ksoAnimHeight = $00000004;
  ksoAnimOpacity = $00000005;
  ksoAnimRotation = $00000006;
  ksoAnimColor = $00000007;
  ksoAnimVisibility = $00000008;
  ksoAnimXshear = $00000009;
  ksoAnimTextFontBold = $00000064;
  ksoAnimTextFontColor = $00000065;
  ksoAnimTextFontEmboss = $00000066;
  ksoAnimTextFontItalic = $00000067;
  ksoAnimTextFontName = $00000068;
  ksoAnimTextFontShadow = $00000069;
  ksoAnimTextFontSize = $0000006A;
  ksoAnimTextFontSubscript = $0000006B;
  ksoAnimTextFontSuperscript = $0000006C;
  ksoAnimTextFontUnderline = $0000006D;
  ksoAnimTextFontStrikeThrough = $0000006E;
  ksoAnimTextBulletCharacter = $0000006F;
  ksoAnimTextBulletFontName = $00000070;
  ksoAnimTextBulletNumber = $00000071;
  ksoAnimTextBulletColor = $00000072;
  ksoAnimTextBulletRelativeSize = $00000073;
  ksoAnimTextBulletStyle = $00000074;
  ksoAnimTextBulletType = $00000075;
  ksoAnimShapePictureContrast = $000003E8;
  ksoAnimShapePictureBrightness = $000003E9;
  ksoAnimShapePictureGamma = $000003EA;
  ksoAnimShapePictureGrayscale = $000003EB;
  ksoAnimShapeFillOn = $000003EC;
  ksoAnimShapeFillColor = $000003ED;
  ksoAnimShapeFillOpacity = $000003EE;
  ksoAnimShapeFillBackColor = $000003EF;
  ksoAnimShapeLineOn = $000003F0;
  ksoAnimShapeLineColor = $000003F1;
  ksoAnimShapeShadowOn = $000003F2;
  ksoAnimShapeShadowType = $000003F3;
  ksoAnimShapeShadowColor = $000003F4;
  ksoAnimShapeShadowOpacity = $000003F5;
  ksoAnimShapeShadowOffsetX = $000003F6;
  ksoAnimShapeShadowOffsetY = $000003F7;
  ksoAnimShapeFillType = $000003F8;
  ksoAnimShapeStyleRotation = $000003F9;
  ksoAnimXY = $000003FA;

// Constants for enum KsoAnimCommandType
type
  KsoAnimCommandType = TOleEnum;
const
  ksoAnimCommandTypeEvent = $00000000;
  ksoAnimCommandTypeCall = $00000001;
  ksoAnimCommandTypeVerb = $00000002;

// Constants for enum KsoAnimFilterEffectType
type
  KsoAnimFilterEffectType = TOleEnum;
const
  ksoAnimFilterEffectTypeNone = $00000000;
  ksoAnimFilterEffectTypeBarn = $00000001;
  ksoAnimFilterEffectTypeBlinds = $00000002;
  ksoAnimFilterEffectTypeBox = $00000003;
  ksoAnimFilterEffectTypeCheckerboard = $00000004;
  ksoAnimFilterEffectTypeCircle = $00000005;
  ksoAnimFilterEffectTypeDiamond = $00000006;
  ksoAnimFilterEffectTypeDissolve = $00000007;
  ksoAnimFilterEffectTypeFade = $00000008;
  ksoAnimFilterEffectTypeImage = $00000009;
  ksoAnimFilterEffectTypePixelate = $0000000A;
  ksoAnimFilterEffectTypePlus = $0000000B;
  ksoAnimFilterEffectTypeRandomBar = $0000000C;
  ksoAnimFilterEffectTypeSlide = $0000000D;
  ksoAnimFilterEffectTypeStretch = $0000000E;
  ksoAnimFilterEffectTypeStrips = $0000000F;
  ksoAnimFilterEffectTypeWedge = $00000010;
  ksoAnimFilterEffectTypeWheel = $00000011;
  ksoAnimFilterEffectTypeWipe = $00000012;

// Constants for enum KsoAnimFilterEffectSubtype
type
  KsoAnimFilterEffectSubtype = TOleEnum;
const
  ksoAnimFilterEffectSubtypeNone = $00000000;
  ksoAnimFilterEffectSubtypeInVertical = $00000001;
  ksoAnimFilterEffectSubtypeOutVertical = $00000002;
  ksoAnimFilterEffectSubtypeInHorizontal = $00000003;
  ksoAnimFilterEffectSubtypeOutHorizontal = $00000004;
  ksoAnimFilterEffectSubtypeHorizontal = $00000005;
  ksoAnimFilterEffectSubtypeVertical = $00000006;
  ksoAnimFilterEffectSubtypeIn = $00000007;
  ksoAnimFilterEffectSubtypeOut = $00000008;
  ksoAnimFilterEffectSubtypeAcross = $00000009;
  ksoAnimFilterEffectSubtypeFromLeft = $0000000A;
  ksoAnimFilterEffectSubtypeFromRight = $0000000B;
  ksoAnimFilterEffectSubtypeFromTop = $0000000C;
  ksoAnimFilterEffectSubtypeFromBottom = $0000000D;
  ksoAnimFilterEffectSubtypeDownLeft = $0000000E;
  ksoAnimFilterEffectSubtypeUpLeft = $0000000F;
  ksoAnimFilterEffectSubtypeDownRight = $00000010;
  ksoAnimFilterEffectSubtypeUpRight = $00000011;
  ksoAnimFilterEffectSubtypeSpokes1 = $00000012;
  ksoAnimFilterEffectSubtypeSpokes2 = $00000013;
  ksoAnimFilterEffectSubtypeSpokes3 = $00000014;
  ksoAnimFilterEffectSubtypeSpokes4 = $00000015;
  ksoAnimFilterEffectSubtypeSpokes8 = $00000016;
  ksoAnimFilterEffectSubtypeLeft = $00000017;
  ksoAnimFilterEffectSubtypeRight = $00000018;
  ksoAnimFilterEffectSubtypeDown = $00000019;
  ksoAnimFilterEffectSubtypeUp = $0000001A;

// Constants for enum KsoExtraInfoMethod
type
  KsoExtraInfoMethod = TOleEnum;
const
  ksoMethodGet = $00000000;
  ksoMethodPost = $00000001;

// Constants for enum KsoFarEastLineBreakLanguageID
type
  KsoFarEastLineBreakLanguageID = TOleEnum;
const
  ksoFarEastLineBreakLanguageJapanese = $00000411;
  ksoFarEastLineBreakLanguageKorean = $00000412;
  ksoFarEastLineBreakLanguageSimplifiedChinese = $00000804;
  ksoFarEastLineBreakLanguageTraditionalChinese = $00000404;

// Constants for enum KsoLanguageID
type
  KsoLanguageID = TOleEnum;
const
  ksoLanguageIDMixed = $FFFFFFFE;
  ksoLanguageIDNone = $00000000;
  ksoLanguageIDNoProofing = $00000400;
  ksoLanguageIDAfrikaans = $00000436;
  ksoLanguageIDAlbanian = $0000041C;
  ksoLanguageIDAmharic = $0000045E;
  ksoLanguageIDArabicAlgeria = $00001401;
  ksoLanguageIDArabicBahrain = $00003C01;
  ksoLanguageIDArabicEgypt = $00000C01;
  ksoLanguageIDArabicIraq = $00000801;
  ksoLanguageIDArabicJordan = $00002C01;
  ksoLanguageIDArabicKuwait = $00003401;
  ksoLanguageIDArabicLebanon = $00003001;
  ksoLanguageIDArabicLibya = $00001001;
  ksoLanguageIDArabicMorocco = $00001801;
  ksoLanguageIDArabicOman = $00002001;
  ksoLanguageIDArabicQatar = $00004001;
  ksoLanguageIDArabic = $00000401;
  ksoLanguageIDArabicSyria = $00002801;
  ksoLanguageIDArabicTunisia = $00001C01;
  ksoLanguageIDArabicUAE = $00003801;
  ksoLanguageIDArabicYemen = $00002401;
  ksoLanguageIDArmenian = $0000042B;
  ksoLanguageIDAssamese = $0000044D;
  ksoLanguageIDAzeriCyrillic = $0000082C;
  ksoLanguageIDAzeriLatin = $0000042C;
  ksoLanguageIDBasque = $0000042D;
  ksoLanguageIDByelorussian = $00000423;
  ksoLanguageIDBengali = $00000445;
  ksoLanguageIDBulgarian = $00000402;
  ksoLanguageIDBurmese = $00000455;
  ksoLanguageIDCatalan = $00000403;
  ksoLanguageIDChineseHongKongSAR = $00000C04;
  ksoLanguageIDChineseMacaoSAR = $00001404;
  ksoLanguageIDSimplifiedChinese = $00000804;
  ksoLanguageIDChineseSingapore = $00001004;
  ksoLanguageIDTraditionalChinese = $00000404;
  ksoLanguageIDCherokee = $0000045C;
  ksoLanguageIDCroatian = $0000041A;
  ksoLanguageIDCzech = $00000405;
  ksoLanguageIDDanish = $00000406;
  ksoLanguageIDDivehi = $00000465;
  ksoLanguageIDBelgianDutch = $00000813;
  ksoLanguageIDDutch = $00000413;
  ksoLanguageIDDzongkhaBhutan = $00000851;
  ksoLanguageIDEdo = $00000466;
  ksoLanguageIDEnglishAUS = $00000C09;
  ksoLanguageIDEnglishBelize = $00002809;
  ksoLanguageIDEnglishCanadian = $00001009;
  ksoLanguageIDEnglishCaribbean = $00002409;
  ksoLanguageIDEnglishIndonesia = $00003809;
  ksoLanguageIDEnglishIreland = $00001809;
  ksoLanguageIDEnglishJamaica = $00002009;
  ksoLanguageIDEnglishNewZealand = $00001409;
  ksoLanguageIDEnglishPhilippines = $00003409;
  ksoLanguageIDEnglishSouthAfrica = $00001C09;
  ksoLanguageIDEnglishTrinidadTobago = $00002C09;
  ksoLanguageIDEnglishUK = $00000809;
  ksoLanguageIDEnglishUS = $00000409;
  ksoLanguageIDEnglishZimbabwe = $00003009;
  ksoLanguageIDEstonian = $00000425;
  ksoLanguageIDFaeroese = $00000438;
  ksoLanguageIDFarsi = $00000429;
  ksoLanguageIDFilipino = $00000464;
  ksoLanguageIDFinnish = $0000040B;
  ksoLanguageIDBelgianFrench = $0000080C;
  ksoLanguageIDFrenchCameroon = $00002C0C;
  ksoLanguageIDFrenchCanadian = $00000C0C;
  ksoLanguageIDFrenchCotedIvoire = $0000300C;
  ksoLanguageIDFrench = $0000040C;
  ksoLanguageIDFrenchHaiti = $00003C0C;
  ksoLanguageIDFrenchLuxembourg = $0000140C;
  ksoLanguageIDFrenchMali = $0000340C;
  ksoLanguageIDFrenchMonaco = $0000180C;
  ksoLanguageIDFrenchMorocco = $0000380C;
  ksoLanguageIDFrenchReunion = $0000200C;
  ksoLanguageIDFrenchSenegal = $0000280C;
  ksoLanguageIDSwissFrench = $0000100C;
  ksoLanguageIDFrenchWestIndies = $00001C0C;
  ksoLanguageIDFrenchZaire = $0000240C;
  ksoLanguageIDFrisianNetherlands = $00000462;
  ksoLanguageIDFulfulde = $00000467;
  ksoLanguageIDGaelicIreland = $0000083C;
  ksoLanguageIDGaelicScotland = $0000043C;
  ksoLanguageIDGalician = $00000456;
  ksoLanguageIDGeorgian = $00000437;
  ksoLanguageIDGermanAustria = $00000C07;
  ksoLanguageIDGerman = $00000407;
  ksoLanguageIDGermanLiechtenstein = $00001407;
  ksoLanguageIDGermanLuxembourg = $00001007;
  ksoLanguageIDSwissGerman = $00000807;
  ksoLanguageIDGreek = $00000408;
  ksoLanguageIDGuarani = $00000474;
  ksoLanguageIDGujarati = $00000447;
  ksoLanguageIDHausa = $00000468;
  ksoLanguageIDHawaiian = $00000475;
  ksoLanguageIDHebrew = $0000040D;
  ksoLanguageIDHindi = $00000439;
  ksoLanguageIDHungarian = $0000040E;
  ksoLanguageIDIbibio = $00000469;
  ksoLanguageIDIcelandic = $0000040F;
  ksoLanguageIDIgbo = $00000470;
  ksoLanguageIDIndonesian = $00000421;
  ksoLanguageIDInuktitut = $0000045D;
  ksoLanguageIDItalian = $00000410;
  ksoLanguageIDSwissItalian = $00000810;
  ksoLanguageIDJapanese = $00000411;
  ksoLanguageIDKannada = $0000044B;
  ksoLanguageIDKanuri = $00000471;
  ksoLanguageIDKashmiri = $00000460;
  ksoLanguageIDKashmiriDevanagari = $00000860;
  ksoLanguageIDKazakh = $0000043F;
  ksoLanguageIDKhmer = $00000453;
  ksoLanguageIDKirghiz = $00000440;
  ksoLanguageIDKonkani = $00000457;
  ksoLanguageIDKorean = $00000412;
  ksoLanguageIDKyrgyz = $00000440;
  ksoLanguageIDLatin = $00000476;
  ksoLanguageIDLao = $00000454;
  ksoLanguageIDLatvian = $00000426;
  ksoLanguageIDLithuanian = $00000427;
  ksoLanguageIDMacedonian = $0000042F;
  ksoLanguageIDMalaysian = $0000043E;
  ksoLanguageIDMalayBruneiDarussalam = $0000083E;
  ksoLanguageIDMalayalam = $0000044C;
  ksoLanguageIDMaltese = $0000043A;
  ksoLanguageIDManipuri = $00000458;
  ksoLanguageIDMarathi = $0000044E;
  ksoLanguageIDMongolian = $00000450;
  ksoLanguageIDNepali = $00000461;
  ksoLanguageIDNorwegianBokmol = $00000414;
  ksoLanguageIDNorwegianNynorsk = $00000814;
  ksoLanguageIDOriya = $00000448;
  ksoLanguageIDOromo = $00000472;
  ksoLanguageIDPashto = $00000463;
  ksoLanguageIDPolish = $00000415;
  ksoLanguageIDBrazilianPortuguese = $00000416;
  ksoLanguageIDPortuguese = $00000816;
  ksoLanguageIDPunjabi = $00000446;
  ksoLanguageIDRhaetoRomanic = $00000417;
  ksoLanguageIDRomanianMoldova = $00000818;
  ksoLanguageIDRomanian = $00000418;
  ksoLanguageIDRussianMoldova = $00000819;
  ksoLanguageIDRussian = $00000419;
  ksoLanguageIDSamiLappish = $0000043B;
  ksoLanguageIDSanskrit = $0000044F;
  ksoLanguageIDSerbianCyrillic = $00000C1A;
  ksoLanguageIDSerbianLatin = $0000081A;
  ksoLanguageIDSesotho = $00000430;
  ksoLanguageIDSindhi = $00000459;
  ksoLanguageIDSindhiPakistan = $00000859;
  ksoLanguageIDSinhalese = $0000045B;
  ksoLanguageIDSlovak = $0000041B;
  ksoLanguageIDSlovenian = $00000424;
  ksoLanguageIDSomali = $00000477;
  ksoLanguageIDSorbian = $0000042E;
  ksoLanguageIDSpanishArgentina = $00002C0A;
  ksoLanguageIDSpanishBolivia = $0000400A;
  ksoLanguageIDSpanishChile = $0000340A;
  ksoLanguageIDSpanishColombia = $0000240A;
  ksoLanguageIDSpanishCostaRica = $0000140A;
  ksoLanguageIDSpanishDominicanRepublic = $00001C0A;
  ksoLanguageIDSpanishEcuador = $0000300A;
  ksoLanguageIDSpanishElSalvador = $0000440A;
  ksoLanguageIDSpanishGuatemala = $0000100A;
  ksoLanguageIDSpanishHonduras = $0000480A;
  ksoLanguageIDMexicanSpanish = $0000080A;
  ksoLanguageIDSpanishNicaragua = $00004C0A;
  ksoLanguageIDSpanishPanama = $0000180A;
  ksoLanguageIDSpanishParaguay = $00003C0A;
  ksoLanguageIDSpanishPeru = $0000280A;
  ksoLanguageIDSpanishPuertoRico = $0000500A;
  ksoLanguageIDSpanishModernSort = $00000C0A;
  ksoLanguageIDSpanish = $0000040A;
  ksoLanguageIDSpanishUruguay = $0000380A;
  ksoLanguageIDSpanishVenezuela = $0000200A;
  ksoLanguageIDSutu = $00000430;
  ksoLanguageIDSwahili = $00000441;
  ksoLanguageIDSwedishFinland = $0000081D;
  ksoLanguageIDSwedish = $0000041D;
  ksoLanguageIDSyriac = $0000045A;
  ksoLanguageIDTajik = $00000428;
  ksoLanguageIDTamil = $00000449;
  ksoLanguageIDTamazight = $0000045F;
  ksoLanguageIDTamazightLatin = $0000085F;
  ksoLanguageIDTatar = $00000444;
  ksoLanguageIDTelugu = $0000044A;
  ksoLanguageIDThai = $0000041E;
  ksoLanguageIDTibetan = $00000451;
  ksoLanguageIDTigrignaEthiopic = $00000473;
  ksoLanguageIDTigrignaEritrea = $00000873;
  ksoLanguageIDTsonga = $00000431;
  ksoLanguageIDTswana = $00000432;
  ksoLanguageIDTurkish = $0000041F;
  ksoLanguageIDTurkmen = $00000442;
  ksoLanguageIDUkrainian = $00000422;
  ksoLanguageIDUrdu = $00000420;
  ksoLanguageIDUzbekCyrillic = $00000843;
  ksoLanguageIDUzbekLatin = $00000443;
  ksoLanguageIDVenda = $00000433;
  ksoLanguageIDVietnamese = $0000042A;
  ksoLanguageIDWelsh = $00000452;
  ksoLanguageIDXhosa = $00000434;
  ksoLanguageIDYi = $00000478;
  ksoLanguageIDYiddish = $0000043D;
  ksoLanguageIDYoruba = $0000046A;
  ksoLanguageIDZulu = $00000435;

// Constants for enum KsoEncoding
type
  KsoEncoding = TOleEnum;
const
  ksoEncodingThai = $0000036A;
  ksoEncodingJapaneseShiftJIS = $000003A4;
  ksoEncodingSimplifiedChineseGBK = $000003A8;
  ksoEncodingKorean = $000003B5;
  ksoEncodingTraditionalChineseBig5 = $000003B6;
  ksoEncodingUnicodeLittleEndian = $000004B0;
  ksoEncodingUnicodeBigEndian = $000004B1;
  ksoEncodingCentralEuropean = $000004E2;
  ksoEncodingCyrillic = $000004E3;
  ksoEncodingWestern = $000004E4;
  ksoEncodingGreek = $000004E5;
  ksoEncodingTurkish = $000004E6;
  ksoEncodingHebrew = $000004E7;
  ksoEncodingArabic = $000004E8;
  ksoEncodingBaltic = $000004E9;
  ksoEncodingVietnamese = $000004EA;
  ksoEncodingAutoDetect = $0000C351;
  ksoEncodingJapaneseAutoDetect = $0000C6F4;
  ksoEncodingSimplifiedChineseAutoDetect = $0000C6F8;
  ksoEncodingKoreanAutoDetect = $0000C705;
  ksoEncodingTraditionalChineseAutoDetect = $0000C706;
  ksoEncodingCyrillicAutoDetect = $0000C833;
  ksoEncodingGreekAutoDetect = $0000C835;
  ksoEncodingArabicAutoDetect = $0000C838;
  ksoEncodingISO88591Latin1 = $00006FAF;
  ksoEncodingISO88592CentralEurope = $00006FB0;
  ksoEncodingISO88593Latin3 = $00006FB1;
  ksoEncodingISO88594Baltic = $00006FB2;
  ksoEncodingISO88595Cyrillic = $00006FB3;
  ksoEncodingISO88596Arabic = $00006FB4;
  ksoEncodingISO88597Greek = $00006FB5;
  ksoEncodingISO88598Hebrew = $00006FB6;
  ksoEncodingISO88599Turkish = $00006FB7;
  ksoEncodingISO885915Latin9 = $00006FBD;
  ksoEncodingISO2022JPNoHalfwidthKatakana = $0000C42C;
  ksoEncodingISO2022JPJISX02021984 = $0000C42D;
  ksoEncodingISO2022JPJISX02011989 = $0000C42E;
  ksoEncodingISO2022KR = $0000C431;
  ksoEncodingISO2022CNTraditionalChinese = $0000C433;
  ksoEncodingISO2022CNSimplifiedChinese = $0000C435;
  ksoEncodingMacRoman = $00002710;
  ksoEncodingMacJapanese = $00002711;
  ksoEncodingMacTraditionalChineseBig5 = $00002712;
  ksoEncodingMacKorean = $00002713;
  ksoEncodingMacArabic = $00002714;
  ksoEncodingMacHebrew = $00002715;
  ksoEncodingMacGreek1 = $00002716;
  ksoEncodingMacCyrillic = $00002717;
  ksoEncodingMacSimplifiedChineseGB2312 = $00002718;
  ksoEncodingMacRomania = $0000271A;
  ksoEncodingMacUkraine = $00002721;
  ksoEncodingMacLatin2 = $0000272D;
  ksoEncodingMacIcelandic = $0000275F;
  ksoEncodingMacTurkish = $00002761;
  ksoEncodingMacCroatia = $00002762;
  ksoEncodingEBCDICUSCanada = $00000025;
  ksoEncodingEBCDICInternational = $000001F4;
  ksoEncodingEBCDICMultilingualROECELatin2 = $00000366;
  ksoEncodingEBCDICGreekModern = $0000036B;
  ksoEncodingEBCDICTurkishLatin5 = $00000402;
  ksoEncodingEBCDICGermany = $00004F31;
  ksoEncodingEBCDICDenmarkNorway = $00004F35;
  ksoEncodingEBCDICFinlandSweden = $00004F36;
  ksoEncodingEBCDICItaly = $00004F38;
  ksoEncodingEBCDICLatinAmericaSpain = $00004F3C;
  ksoEncodingEBCDICUnitedKingdom = $00004F3D;
  ksoEncodingEBCDICJapaneseKatakanaExtended = $00004F42;
  ksoEncodingEBCDICFrance = $00004F49;
  ksoEncodingEBCDICArabic = $00004FC4;
  ksoEncodingEBCDICGreek = $00004FC7;
  ksoEncodingEBCDICHebrew = $00004FC8;
  ksoEncodingEBCDICKoreanExtended = $00005161;
  ksoEncodingEBCDICThai = $00005166;
  ksoEncodingEBCDICIcelandic = $00005187;
  ksoEncodingEBCDICTurkish = $000051A9;
  ksoEncodingEBCDICRussian = $00005190;
  ksoEncodingEBCDICSerbianBulgarian = $00005221;
  ksoEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = $0000C6F2;
  ksoEncodingEBCDICUSCanadaAndJapanese = $0000C6F3;
  ksoEncodingEBCDICKoreanExtendedAndKorean = $0000C6F5;
  ksoEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = $0000C6F7;
  ksoEncodingEBCDICUSCanadaAndTraditionalChinese = $0000C6F9;
  ksoEncodingEBCDICJapaneseLatinExtendedAndJapanese = $0000C6FB;
  ksoEncodingOEMUnitedStates = $000001B5;
  ksoEncodingOEMGreek437G = $000002E1;
  ksoEncodingOEMBaltic = $00000307;
  ksoEncodingOEMMultilingualLatinI = $00000352;
  ksoEncodingOEMMultilingualLatinII = $00000354;
  ksoEncodingOEMCyrillic = $00000357;
  ksoEncodingOEMTurkish = $00000359;
  ksoEncodingOEMPortuguese = $0000035C;
  ksoEncodingOEMIcelandic = $0000035D;
  ksoEncodingOEMHebrew = $0000035E;
  ksoEncodingOEMCanadianFrench = $0000035F;
  ksoEncodingOEMArabic = $00000360;
  ksoEncodingOEMNordic = $00000361;
  ksoEncodingOEMCyrillicII = $00000362;
  ksoEncodingOEMModernGreek = $00000365;
  ksoEncodingEUCJapanese = $0000CADC;
  ksoEncodingEUCChineseSimplifiedChinese = $0000CAE0;
  ksoEncodingEUCKorean = $0000CAED;
  ksoEncodingEUCTaiwaneseTraditionalChinese = $0000CAEE;
  ksoEncodingISCIIDevanagari = $0000DEAA;
  ksoEncodingISCIIBengali = $0000DEAB;
  ksoEncodingISCIITamil = $0000DEAC;
  ksoEncodingISCIITelugu = $0000DEAD;
  ksoEncodingISCIIAssamese = $0000DEAE;
  ksoEncodingISCIIOriya = $0000DEAF;
  ksoEncodingISCIIKannada = $0000DEB0;
  ksoEncodingISCIIMalayalam = $0000DEB1;
  ksoEncodingISCIIGujarati = $0000DEB2;
  ksoEncodingISCIIPunjabi = $0000DEB3;
  ksoEncodingArabicASMO = $000002C4;
  ksoEncodingArabicTransparentASMO = $000002D0;
  ksoEncodingKoreanJohab = $00000551;
  ksoEncodingTaiwanCNS = $00004E20;
  ksoEncodingTaiwanTCA = $00004E21;
  ksoEncodingTaiwanEten = $00004E22;
  ksoEncodingTaiwanIBM5550 = $00004E23;
  ksoEncodingTaiwanTeleText = $00004E24;
  ksoEncodingTaiwanWang = $00004E25;
  ksoEncodingIA5IRV = $00004E89;
  ksoEncodingIA5German = $00004E8A;
  ksoEncodingIA5Swedish = $00004E8B;
  ksoEncodingIA5Norwegian = $00004E8C;
  ksoEncodingUSASCII = $00004E9F;
  ksoEncodingT61 = $00004F25;
  ksoEncodingISO6937NonSpacingAccent = $00004F2D;
  ksoEncodingKOI8R = $00005182;
  ksoEncodingExtAlphaLowercase = $00005223;
  ksoEncodingKOI8U = $0000556A;
  ksoEncodingEuropa3 = $00007149;
  ksoEncodingHZGBSimplifiedChinese = $0000CEC8;
  ksoEncodingUTF7 = $0000FDE8;
  ksoEncodingUTF8 = $0000FDE9;
  ksoEncodingSimplifiedChinese = $000051C8;

// Constants for enum KsoSyncErrorType
type
  KsoSyncErrorType = TOleEnum;
const
  ksoSyncErrorNone = $00000000;
  ksoSyncErrorUnauthorizedUser = $00000001;
  ksoSyncErrorCouldNotConnect = $00000002;
  ksoSyncErrorOutOfSpace = $00000003;
  ksoSyncErrorFileNotFound = $00000004;
  ksoSyncErrorFileTooLarge = $00000005;
  ksoSyncErrorFileInUse = $00000006;
  ksoSyncErrorVirusUpload = $00000007;
  ksoSyncErrorVirusDownload = $00000008;
  ksoSyncErrorUnknownUpload = $00000009;
  ksoSyncErrorUnknownDownload = $0000000A;
  ksoSyncErrorCouldNotOpen = $0000000B;
  ksoSyncErrorCouldNotUpdate = $0000000C;
  ksoSyncErrorCouldNotCompare = $0000000D;
  ksoSyncErrorCouldNotResolve = $0000000E;
  ksoSyncErrorNoNetwork = $0000000F;
  ksoSyncErrorUnknown = $00000010;

// Constants for enum KsoSyncStatusType
type
  KsoSyncStatusType = TOleEnum;
const
  ksoSyncStatusNoSharedWorkspace = $00000000;
  ksoSyncStatusLatest = $00000001;
  ksoSyncStatusNewerAvailable = $00000002;
  ksoSyncStatusLocalChanges = $00000003;
  ksoSyncStatusConflict = $00000004;
  ksoSyncStatusSuspended = $00000005;
  ksoSyncStatusError = $00000006;

// Constants for enum KsoSyncVersionType
type
  KsoSyncVersionType = TOleEnum;
const
  ksoSyncVersionLastViewed = $00000000;
  ksoSyncVersionServer = $00000001;

// Constants for enum KsoSyncConflictResolutionType
type
  KsoSyncConflictResolutionType = TOleEnum;
const
  ksoSyncConflictClientWins = $00000000;
  ksoSyncConflictServerWins = $00000001;
  ksoSyncConflictMerge = $00000002;

// Constants for enum KsoHTMLProjectState
type
  KsoHTMLProjectState = TOleEnum;
const
  ksoHTMLProjectStateDocumentLocked = $00000001;
  ksoHTMLProjectStateProjectLocked = $00000002;
  ksoHTMLProjectStateDocumentProjectUnlocked = $00000003;

// Constants for enum KsoHTMLProjectOpen
type
  KsoHTMLProjectOpen = TOleEnum;
const
  ksoHTMLProjectOpenSourceView = $00000001;
  ksoHTMLProjectOpenTextView = $00000002;

// Constants for enum KsoSharedWorkspaceTaskStatus
type
  KsoSharedWorkspaceTaskStatus = TOleEnum;
const
  ksoSharedWorkspaceTaskStatusNotStarted = $00000001;
  ksoSharedWorkspaceTaskStatusInProgress = $00000002;
  ksoSharedWorkspaceTaskStatusCompleted = $00000003;
  ksoSharedWorkspaceTaskStatusDeferred = $00000004;
  ksoSharedWorkspaceTaskStatusWaiting = $00000005;

// Constants for enum KsoSharedWorkspaceTaskPriority
type
  KsoSharedWorkspaceTaskPriority = TOleEnum;
const
  ksoSharedWorkspaceTaskPriorityHigh = $00000001;
  ksoSharedWorkspaceTaskPriorityNormal = $00000002;
  ksoSharedWorkspaceTaskPriorityLow = $00000003;

// Constants for enum KsoSyncEventType
type
  KsoSyncEventType = TOleEnum;
const
  ksoSyncEventDownloadInitiated = $00000000;
  ksoSyncEventDownloadSucceeded = $00000001;
  ksoSyncEventDownloadFailed = $00000002;
  ksoSyncEventUploadInitiated = $00000003;
  ksoSyncEventUploadSucceeded = $00000004;
  ksoSyncEventUploadFailed = $00000005;
  ksoSyncEventDownloadNoChange = $00000006;
  ksoSyncEventOffline = $00000007;

// Constants for enum KsoAutoSize
type
  KsoAutoSize = TOleEnum;
const
  ksoAutoSizeMixed = $FFFFFFFE;
  ksoAutoSizeNone = $00000000;
  ksoAutoSizeShapeToFitText = $00000001;

// Constants for enum KsoDiagramType
type
  KsoDiagramType = TOleEnum;
const
  ksoDiagramMixed = $FFFFFFFE;
  ksoDiagramOrgChart = $00000001;
  ksoDiagramCycle = $00000002;
  ksoDiagramRadial = $00000003;
  ksoDiagramPyramid = $00000004;
  ksoDiagramVenn = $00000005;
  ksoDiagramTarget = $00000006;

// Constants for enum KsoOrgChartLayoutType
type
  KsoOrgChartLayoutType = TOleEnum;
const
  ksoOrgChartLayoutMixed = $FFFFFFFE;
  ksoOrgChartLayoutStandard = $00000001;
  ksoOrgChartLayoutBothHanging = $00000002;
  ksoOrgChartLayoutLeftHanging = $00000003;
  ksoOrgChartLayoutRightHanging = $00000004;

// Constants for enum KsoDiagramNodeType
type
  KsoDiagramNodeType = TOleEnum;
const
  ksoDiagramNode = $00000001;
  ksoDiagramAssistant = $00000002;

// Constants for enum KsoRelativeNodePosition
type
  KsoRelativeNodePosition = TOleEnum;
const
  ksoBeforeNode = $00000001;
  ksoAfterNode = $00000002;
  ksoBeforeFirstSibling = $00000003;
  ksoAfterLastSibling = $00000004;

// Constants for enum KsoHyperlinkType
type
  KsoHyperlinkType = TOleEnum;
const
  ksoHyperlinkRange = $00000000;
  ksoHyperlinkShape = $00000001;
  ksoHyperlinkInlineShape = $00000002;

// Constants for enum ksoSuspensive
type
  ksoSuspensive = TOleEnum;
const
  ksoSuspensiveLong = $0098967F;
  ksoSuspensiveString = $00000000;
  ksoSuspensiveInterface = $00000000;

// Constants for enum KsoDocProperties
type
  KsoDocProperties = TOleEnum;
const
  ksoPropertyTypeNumber = $00000001;
  ksoPropertyTypeBoolean = $00000002;
  ksoPropertyTypeDate = $00000003;
  ksoPropertyTypeString = $00000004;
  ksoPropertyTypeFloat = $00000005;
  ksoPropertyTypeMetaFile = $00000006;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IKsoDispObj = interface;
  IKsoDispObjDisp = dispinterface;
  Collection = interface;
  CollectionDisp = dispinterface;
  SecurityOptions = interface;
  SecurityOptionsDisp = dispinterface;
  _CommandBars = interface;
  _CommandBarsDisp = dispinterface;
  CommandBarControl = interface;
  CommandBarControlDisp = dispinterface;
  CommandBar = interface;
  CommandBarDisp = dispinterface;
  CommandBarControls = interface;
  CommandBarControlsDisp = dispinterface;
  _CommandBarButton = interface;
  _CommandBarButtonDisp = dispinterface;
  CommandBarPopup = interface;
  CommandBarPopupDisp = dispinterface;
  _CommandBarComboBox = interface;
  _CommandBarComboBoxDisp = dispinterface;
  ICommandBarsEvents = interface;
  ICommandBarsEventsDisp = dispinterface;
  _CommandBarsEvents = dispinterface;
  ICommandBarComboBoxEvents = interface;
  ICommandBarComboBoxEventsDisp = dispinterface;
  _CommandBarComboBoxEvents = dispinterface;
  _CommandBarButtonEvents = dispinterface;
  ICommandBarButtonEvents = interface;
  ICommandBarButtonEventsDisp = dispinterface;
  _KsoDiagram = interface;
  _KsoDiagramDisp = dispinterface;
  _KsoDiagramNodes = interface;
  _KsoDiagramNodesDisp = dispinterface;
  _KsoDiagramNode = interface;
  _KsoDiagramNodeDisp = dispinterface;
  _KsoDiagramNodeChildren = interface;
  _KsoDiagramNodeChildrenDisp = dispinterface;
  KsoShape = interface;
  KsoShapeDisp = dispinterface;
  KsoShapeRange = interface;
  KsoShapeRangeDisp = dispinterface;
  KsoAdjustments = interface;
  KsoAdjustmentsDisp = dispinterface;
  KsoCalloutFormat = interface;
  KsoCalloutFormatDisp = dispinterface;
  KsoConnectorFormat = interface;
  KsoConnectorFormatDisp = dispinterface;
  KsoFillFormat = interface;
  KsoFillFormatDisp = dispinterface;
  KsoColorFormat = interface;
  KsoColorFormatDisp = dispinterface;
  KsoGroupShapes = interface;
  KsoGroupShapesDisp = dispinterface;
  KsoLineFormat = interface;
  KsoLineFormatDisp = dispinterface;
  KsoShapeNodes = interface;
  KsoShapeNodesDisp = dispinterface;
  KsoShapeNode = interface;
  KsoShapeNodeDisp = dispinterface;
  KsoPictureFormat = interface;
  KsoPictureFormatDisp = dispinterface;
  KsoShadowFormat = interface;
  KsoShadowFormatDisp = dispinterface;
  KsoTextEffectFormat = interface;
  KsoTextEffectFormatDisp = dispinterface;
  KsoTextFrame = interface;
  KsoTextFrameDisp = dispinterface;
  KsoThreeDFormat = interface;
  KsoThreeDFormatDisp = dispinterface;
  KsoScript = interface;
  KsoScriptDisp = dispinterface;
  KsoCanvasShapes = interface;
  KsoCanvasShapesDisp = dispinterface;
  KsoFreeformBuilder = interface;
  KsoFreeformBuilderDisp = dispinterface;
  KsoOLEFormat = interface;
  KsoOLEFormatDisp = dispinterface;
  KsoShapes = interface;
  KsoShapesDisp = dispinterface;
  DocumentProperty = interface;
  DocumentPropertyDisp = dispinterface;
  DocumentProperties = interface;
  DocumentPropertiesDisp = dispinterface;
  COMAddIn = interface;
  COMAddInDisp = dispinterface;
  COMAddIns = interface;
  COMAddInsDisp = dispinterface;
  KeyBinding = interface;
  KeyBindingDisp = dispinterface;
  KeyBindings = interface;
  KeyBindingsDisp = dispinterface;
  KeysBoundTo = interface;
  KeysBoundToDisp = dispinterface;
  AdvApiRoot = interface;
  AdvApiRootDisp = dispinterface;
  AdvAddins = interface;
  AdvAddinsDisp = dispinterface;
  AdvAddin = interface;
  AdvAddinDisp = dispinterface;
  AdvRight = interface;
  AdvRightDisp = dispinterface;
  AdvSeal = interface;
  AdvSealDisp = dispinterface;
  Seals = interface;
  SealsDisp = dispinterface;
  Seal = interface;
  SealDisp = dispinterface;
  IFilterMedium = interface;
  IFilterMediumDisp = dispinterface;
  AdvApplication = interface;
  AdvApplicationDisp = dispinterface;
  AdvApplicationEvents = dispinterface;
  AdvRightEvents = dispinterface;
  AdvSealEvents = dispinterface;
  _TaskPane = interface;
  _TaskPaneDisp = dispinterface;
  TaskPanes = interface;
  TaskPanesDisp = dispinterface;
  _ITaskPaneEvent = dispinterface;
  PluginPlatform = interface;
  PluginPlatformDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  CommandBarComboBox = _CommandBarComboBox;
  CommandBarButton = _CommandBarButton;
  CommandBars = _CommandBars;
  TaskPane = _TaskPane;


// *********************************************************************//
// Declaration of structures, unions and aliases.                         
// *********************************************************************//
  POleVariant1 = ^OleVariant; {*}
  PByte1 = ^Byte; {*}


// *********************************************************************//
// Interface: IKsoDispObj
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0300-FFFF-0000-C000-000000000046}
// *********************************************************************//
  IKsoDispObj = interface(IDispatch)
    ['{000C0300-FFFF-0000-C000-000000000046}']
    function Get_Application: IDispatch; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    procedure zimp_DispObj_Reserved1; safecall;
    procedure zimp_DispObj_Reserved2; safecall;
    procedure zimp_DispObj_Reserved3; safecall;
    procedure zimp_DispObj_Reserved4; safecall;
    procedure zimp_DispObj_Reserved5; safecall;
    procedure zimp_DispObj_Reserved6; safecall;
    property Application: IDispatch read Get_Application;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
  end;

// *********************************************************************//
// DispIntf:  IKsoDispObjDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0300-FFFF-0000-C000-000000000046}
// *********************************************************************//
  IKsoDispObjDisp = dispinterface
    ['{000C0300-FFFF-0000-C000-000000000046}']
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: Collection
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0301-FFFF-0000-C000-000000000046}
// *********************************************************************//
  Collection = interface(IKsoDispObj)
    ['{000C0301-FFFF-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function _Index(Index: SYSINT): OleVariant; safecall;
    function Get_Count: Integer; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  CollectionDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0301-FFFF-0000-C000-000000000046}
// *********************************************************************//
  CollectionDisp = dispinterface
    ['{000C0301-FFFF-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    function _Index(Index: SYSINT): OleVariant; dispid 10;
    property Count: Integer readonly dispid 11;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: SecurityOptions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002B003-FFFE-0000-C000-000000111146}
// *********************************************************************//
  SecurityOptions = interface(IKsoDispObj)
    ['{0002B003-FFFE-0000-C000-000000111146}']
    function Get_AutoClearTempFile: KsoTriState; safecall;
    procedure Set_AutoClearTempFile(prop: KsoTriState); safecall;
    function Get_AutoClearBackupFile: KsoTriState; safecall;
    procedure Set_AutoClearBackupFile(prop: KsoTriState); safecall;
    function Get_AutoClearRecentFileLists: KsoTriState; safecall;
    procedure Set_AutoClearRecentFileLists(prop: KsoTriState); safecall;
    function Get_AutoClearClipboard: KsoTriState; safecall;
    procedure Set_AutoClearClipboard(prop: KsoTriState); safecall;
    function Get_AutoCheckBeforeSM: KsoTriState; safecall;
    procedure Set_AutoCheckBeforeSM(prop: KsoTriState); safecall;
    function Get_AutoCheckBeforeCloseFile: KsoTriState; safecall;
    procedure Set_AutoCheckBeforeCloseFile(prop: KsoTriState); safecall;
    property AutoClearTempFile: KsoTriState read Get_AutoClearTempFile write Set_AutoClearTempFile;
    property AutoClearBackupFile: KsoTriState read Get_AutoClearBackupFile write Set_AutoClearBackupFile;
    property AutoClearRecentFileLists: KsoTriState read Get_AutoClearRecentFileLists write Set_AutoClearRecentFileLists;
    property AutoClearClipboard: KsoTriState read Get_AutoClearClipboard write Set_AutoClearClipboard;
    property AutoCheckBeforeSM: KsoTriState read Get_AutoCheckBeforeSM write Set_AutoCheckBeforeSM;
    property AutoCheckBeforeCloseFile: KsoTriState read Get_AutoCheckBeforeCloseFile write Set_AutoCheckBeforeCloseFile;
  end;

// *********************************************************************//
// DispIntf:  SecurityOptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002B003-FFFE-0000-C000-000000111146}
// *********************************************************************//
  SecurityOptionsDisp = dispinterface
    ['{0002B003-FFFE-0000-C000-000000111146}']
    property AutoClearTempFile: KsoTriState dispid 1;
    property AutoClearBackupFile: KsoTriState dispid 2;
    property AutoClearRecentFileLists: KsoTriState dispid 3;
    property AutoClearClipboard: KsoTriState dispid 4;
    property AutoCheckBeforeSM: KsoTriState dispid 5;
    property AutoCheckBeforeCloseFile: KsoTriState dispid 6;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: _CommandBars
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002EB0D-FFFE-0000-C000-000000111146}
// *********************************************************************//
  _CommandBars = interface(IKsoDispObj)
    ['{0002EB0D-FFFE-0000-C000-000000111146}']
    function Get_ActionControl: CommandBarControl; safecall;
    function Get_ActiveMenuBar: CommandBar; safecall;
    function Add(Name: OleVariant; Position: OleVariant; MenuBar: OleVariant; Temporary: OleVariant): CommandBar; safecall;
    function Get_Count: SYSINT; safecall;
    function Get_DisplayTooltips: WordBool; safecall;
    procedure Set_DisplayTooltips(pvarfDisplayTooltips: WordBool); safecall;
    function Get_DisplayKeysInTooltips: WordBool; safecall;
    procedure Set_DisplayKeysInTooltips(pvarfDisplayKeys: WordBool); safecall;
    function FindControl(Type_: OleVariant; Id: OleVariant; Tag: OleVariant; Visible: OleVariant): CommandBarControl; safecall;
    function Get_Item(Index: OleVariant): CommandBar; safecall;
    function Get_LargeButtons: WordBool; safecall;
    procedure Set_LargeButtons(pvarfLargeButtons: WordBool); safecall;
    function Get_MenuAnimationStyle: KsoMenuAnimation; safecall;
    procedure Set_MenuAnimationStyle(pma: KsoMenuAnimation); safecall;
    function Get__NewEnum: IUnknown; safecall;
    procedure ReleaseFocus; safecall;
    function Get_IdsString(ids: SYSINT; out pbstrName: WideString): SYSINT; safecall;
    function Get_TmcGetName(tmc: SYSINT; out pbstrName: WideString): SYSINT; safecall;
    function Get_AdaptiveMenus: WordBool; safecall;
    procedure Set_AdaptiveMenus(pvarfAdaptiveMenus: WordBool); safecall;
    function FindControls(Type_: OleVariant; Id: OleVariant; Tag: OleVariant; Visible: OleVariant): CommandBarControls; safecall;
    function AddEx(TbidOrName: OleVariant; Position: OleVariant; MenuBar: OleVariant; 
                   Temporary: OleVariant; TbtrProtection: OleVariant): CommandBar; safecall;
    function Get_DisplayFonts: WordBool; safecall;
    procedure Set_DisplayFonts(pvarfDisplayFonts: WordBool); safecall;
    function Get_DisableCustomize: WordBool; safecall;
    procedure Set_DisableCustomize(pvarfDisableCustomize: WordBool); safecall;
    function Get_DisableAskAQuestionDropdown: WordBool; safecall;
    procedure Set_DisableAskAQuestionDropdown(pvarfDisableAskAQuestionDropdown: WordBool); safecall;
    property ActionControl: CommandBarControl read Get_ActionControl;
    property ActiveMenuBar: CommandBar read Get_ActiveMenuBar;
    property Count: SYSINT read Get_Count;
    property DisplayTooltips: WordBool read Get_DisplayTooltips write Set_DisplayTooltips;
    property DisplayKeysInTooltips: WordBool read Get_DisplayKeysInTooltips write Set_DisplayKeysInTooltips;
    property Item[Index: OleVariant]: CommandBar read Get_Item; default;
    property LargeButtons: WordBool read Get_LargeButtons write Set_LargeButtons;
    property MenuAnimationStyle: KsoMenuAnimation read Get_MenuAnimationStyle write Set_MenuAnimationStyle;
    property _NewEnum: IUnknown read Get__NewEnum;
    property IdsString[ids: SYSINT; out pbstrName: WideString]: SYSINT read Get_IdsString;
    property TmcGetName[tmc: SYSINT; out pbstrName: WideString]: SYSINT read Get_TmcGetName;
    property AdaptiveMenus: WordBool read Get_AdaptiveMenus write Set_AdaptiveMenus;
    property DisplayFonts: WordBool read Get_DisplayFonts write Set_DisplayFonts;
    property DisableCustomize: WordBool read Get_DisableCustomize write Set_DisableCustomize;
    property DisableAskAQuestionDropdown: WordBool read Get_DisableAskAQuestionDropdown write Set_DisableAskAQuestionDropdown;
  end;

// *********************************************************************//
// DispIntf:  _CommandBarsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002EB0D-FFFE-0000-C000-000000111146}
// *********************************************************************//
  _CommandBarsDisp = dispinterface
    ['{0002EB0D-FFFE-0000-C000-000000111146}']
    property ActionControl: CommandBarControl readonly dispid 1610809344;
    property ActiveMenuBar: CommandBar readonly dispid 1610809345;
    function Add(Name: OleVariant; Position: OleVariant; MenuBar: OleVariant; Temporary: OleVariant): CommandBar; dispid 1610809346;
    property Count: SYSINT readonly dispid 1610809347;
    property DisplayTooltips: WordBool dispid 1610809348;
    property DisplayKeysInTooltips: WordBool dispid 1610809350;
    function FindControl(Type_: OleVariant; Id: OleVariant; Tag: OleVariant; Visible: OleVariant): CommandBarControl; dispid 1610809352;
    property Item[Index: OleVariant]: CommandBar readonly dispid 0; default;
    property LargeButtons: WordBool dispid 1610809354;
    property MenuAnimationStyle: KsoMenuAnimation dispid 1610809356;
    property _NewEnum: IUnknown readonly dispid -4;
    procedure ReleaseFocus; dispid 1610809360;
    property IdsString[ids: SYSINT; out pbstrName: WideString]: SYSINT readonly dispid 1610809361;
    property TmcGetName[tmc: SYSINT; out pbstrName: WideString]: SYSINT readonly dispid 1610809362;
    property AdaptiveMenus: WordBool dispid 1610809363;
    function FindControls(Type_: OleVariant; Id: OleVariant; Tag: OleVariant; Visible: OleVariant): CommandBarControls; dispid 1610809365;
    function AddEx(TbidOrName: OleVariant; Position: OleVariant; MenuBar: OleVariant; 
                   Temporary: OleVariant; TbtrProtection: OleVariant): CommandBar; dispid 1610809366;
    property DisplayFonts: WordBool dispid 1610809367;
    property DisableCustomize: WordBool dispid 1610809369;
    property DisableAskAQuestionDropdown: WordBool dispid 1610809371;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: CommandBarControl
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002C882-FFFE-0000-C000-000000111146}
// *********************************************************************//
  CommandBarControl = interface(IKsoDispObj)
    ['{0002C882-FFFE-0000-C000-000000111146}']
    function Get_BeginGroup: WordBool; safecall;
    procedure Set_BeginGroup(pvarfBeginGroup: WordBool); safecall;
    function Get_BuiltIn: WordBool; safecall;
    function Get_Caption: WideString; safecall;
    procedure Set_Caption(const pbstrCaption: WideString); safecall;
    function Get_Control: IDispatch; safecall;
    function Copy(Bar: OleVariant; Before: OleVariant): CommandBarControl; safecall;
    procedure Delete(Temporary: OleVariant); safecall;
    function Get_DescriptionText: WideString; safecall;
    procedure Set_DescriptionText(const pbstrText: WideString); safecall;
    function Get_Enabled: WordBool; safecall;
    procedure Set_Enabled(pvarfEnabled: WordBool); safecall;
    procedure Execute; safecall;
    function Get_Height: SYSINT; safecall;
    procedure Set_Height(pdy: SYSINT); safecall;
    function Get_HelpContextId: SYSINT; safecall;
    procedure Set_HelpContextId(pid: SYSINT); safecall;
    function Get_HelpFile: WideString; safecall;
    procedure Set_HelpFile(const pbstrFilename: WideString); safecall;
    function Get_Id: SYSINT; safecall;
    function Get_Index: SYSINT; safecall;
    function Get_InstanceId: Integer; safecall;
    function Move(Bar: OleVariant; Before: OleVariant): CommandBarControl; safecall;
    function Get_Left: SYSINT; safecall;
    function Get_OLEUsage: KsoControlOLEUsage; safecall;
    procedure Set_OLEUsage(pcou: KsoControlOLEUsage); safecall;
    function Get_OnAction: WideString; safecall;
    procedure Set_OnAction(const pbstrOnAction: WideString); safecall;
    function Get_Parameter: WideString; safecall;
    procedure Set_Parameter(const pbstrParam: WideString); safecall;
    function Get_Priority: SYSINT; safecall;
    procedure Set_Priority(pnPri: SYSINT); safecall;
    procedure Reset; safecall;
    procedure SetFocus; safecall;
    function Get_Tag: WideString; safecall;
    procedure Set_Tag(const pbstrTag: WideString); safecall;
    function Get_TooltipText: WideString; safecall;
    procedure Set_TooltipText(const pbstrTooltip: WideString); safecall;
    function Get_Top: SYSINT; safecall;
    function Get_type_: KsoControlType; safecall;
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(pvarfVisible: WordBool); safecall;
    function Get_Width: SYSINT; safecall;
    procedure Set_Width(pdx: SYSINT); safecall;
    function Get_IsPriorityDropped: WordBool; safecall;
    procedure Reserved1; safecall;
    procedure Reserved2; safecall;
    procedure Reserved3; safecall;
    procedure Reserved4; safecall;
    procedure Reserved5; safecall;
    procedure Reserved6; safecall;
    procedure Reserved7; safecall;
    property BeginGroup: WordBool read Get_BeginGroup write Set_BeginGroup;
    property BuiltIn: WordBool read Get_BuiltIn;
    property Caption: WideString read Get_Caption write Set_Caption;
    property Control: IDispatch read Get_Control;
    property DescriptionText: WideString read Get_DescriptionText write Set_DescriptionText;
    property Enabled: WordBool read Get_Enabled write Set_Enabled;
    property Height: SYSINT read Get_Height write Set_Height;
    property HelpContextId: SYSINT read Get_HelpContextId write Set_HelpContextId;
    property HelpFile: WideString read Get_HelpFile write Set_HelpFile;
    property Id: SYSINT read Get_Id;
    property Index: SYSINT read Get_Index;
    property InstanceId: Integer read Get_InstanceId;
    property Left: SYSINT read Get_Left;
    property OLEUsage: KsoControlOLEUsage read Get_OLEUsage write Set_OLEUsage;
    property OnAction: WideString read Get_OnAction write Set_OnAction;
    property Parameter: WideString read Get_Parameter write Set_Parameter;
    property Priority: SYSINT read Get_Priority write Set_Priority;
    property Tag: WideString read Get_Tag write Set_Tag;
    property TooltipText: WideString read Get_TooltipText write Set_TooltipText;
    property Top: SYSINT read Get_Top;
    property type_: KsoControlType read Get_type_;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property Width: SYSINT read Get_Width write Set_Width;
    property IsPriorityDropped: WordBool read Get_IsPriorityDropped;
  end;

// *********************************************************************//
// DispIntf:  CommandBarControlDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002C882-FFFE-0000-C000-000000111146}
// *********************************************************************//
  CommandBarControlDisp = dispinterface
    ['{0002C882-FFFE-0000-C000-000000111146}']
    property BeginGroup: WordBool dispid 1610874880;
    property BuiltIn: WordBool readonly dispid 1610874882;
    property Caption: WideString dispid 1610874883;
    property Control: IDispatch readonly dispid 1610874885;
    function Copy(Bar: OleVariant; Before: OleVariant): CommandBarControl; dispid 1610874886;
    procedure Delete(Temporary: OleVariant); dispid 1610874887;
    property DescriptionText: WideString dispid 1610874888;
    property Enabled: WordBool dispid 1610874890;
    procedure Execute; dispid 1610874892;
    property Height: SYSINT dispid 1610874893;
    property HelpContextId: SYSINT dispid 1610874895;
    property HelpFile: WideString dispid 1610874897;
    property Id: SYSINT readonly dispid 1610874899;
    property Index: SYSINT readonly dispid 1610874900;
    property InstanceId: Integer readonly dispid 1610874901;
    function Move(Bar: OleVariant; Before: OleVariant): CommandBarControl; dispid 1610874902;
    property Left: SYSINT readonly dispid 1610874903;
    property OLEUsage: KsoControlOLEUsage dispid 1610874904;
    property OnAction: WideString dispid 1610874906;
    property Parameter: WideString dispid 1610874909;
    property Priority: SYSINT dispid 1610874911;
    procedure Reset; dispid 1610874913;
    procedure SetFocus; dispid 1610874914;
    property Tag: WideString dispid 1610874915;
    property TooltipText: WideString dispid 1610874917;
    property Top: SYSINT readonly dispid 1610874919;
    property type_: KsoControlType readonly dispid 1610874920;
    property Visible: WordBool dispid 1610874921;
    property Width: SYSINT dispid 1610874923;
    property IsPriorityDropped: WordBool readonly dispid 1610874925;
    procedure Reserved1; dispid 1610874926;
    procedure Reserved2; dispid 1610874927;
    procedure Reserved3; dispid 1610874928;
    procedure Reserved4; dispid 1610874929;
    procedure Reserved5; dispid 1610874930;
    procedure Reserved6; dispid 1610874931;
    procedure Reserved7; dispid 1610874932;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: CommandBar
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00023F42-FFFE-0000-C000-000000111146}
// *********************************************************************//
  CommandBar = interface(IKsoDispObj)
    ['{00023F42-FFFE-0000-C000-000000111146}']
    function Get_BuiltIn: WordBool; safecall;
    function Get_Context: WideString; safecall;
    procedure Set_Context(const pbstrContext: WideString); safecall;
    function Get_Controls: CommandBarControls; safecall;
    procedure Delete; safecall;
    function Get_Enabled: WordBool; safecall;
    procedure Set_Enabled(pvarfEnabled: WordBool); safecall;
    function FindControl(Type_: OleVariant; Id: OleVariant; Tag: OleVariant; Visible: OleVariant; 
                         Recursive: OleVariant): CommandBarControl; safecall;
    function Get_Height: SYSINT; safecall;
    procedure Set_Height(pdy: SYSINT); safecall;
    function Get_Index: SYSINT; safecall;
    function Get_InstanceId: Integer; safecall;
    function Get_Left: SYSINT; safecall;
    procedure Set_Left(pxpLeft: SYSINT); safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const pbstrName: WideString); safecall;
    function Get_NameLocal: WideString; safecall;
    procedure Set_NameLocal(const pbstrNameLocal: WideString); safecall;
    function Get_Position: KsoBarPosition; safecall;
    procedure Set_Position(ppos: KsoBarPosition); safecall;
    function Get_RowIndex: SYSINT; safecall;
    procedure Set_RowIndex(piRow: SYSINT); safecall;
    function Get_Protection: KsoBarProtection; safecall;
    procedure Set_Protection(pprot: KsoBarProtection); safecall;
    procedure Reset; safecall;
    procedure ShowPopup(x: OleVariant; y: OleVariant); safecall;
    function Get_Top: SYSINT; safecall;
    procedure Set_Top(pypTop: SYSINT); safecall;
    function Get_type_: KsoBarType; safecall;
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(pvarfVisible: WordBool); safecall;
    function Get_Width: SYSINT; safecall;
    procedure Set_Width(pdx: SYSINT); safecall;
    function Get_AdaptiveMenu: WordBool; safecall;
    procedure Set_AdaptiveMenu(pvarfAdaptiveMenu: WordBool); safecall;
    function Get_Id: SYSINT; safecall;
    property BuiltIn: WordBool read Get_BuiltIn;
    property Context: WideString read Get_Context write Set_Context;
    property Controls: CommandBarControls read Get_Controls;
    property Enabled: WordBool read Get_Enabled write Set_Enabled;
    property Height: SYSINT read Get_Height write Set_Height;
    property Index: SYSINT read Get_Index;
    property InstanceId: Integer read Get_InstanceId;
    property Left: SYSINT read Get_Left write Set_Left;
    property Name: WideString read Get_Name write Set_Name;
    property NameLocal: WideString read Get_NameLocal write Set_NameLocal;
    property Position: KsoBarPosition read Get_Position write Set_Position;
    property RowIndex: SYSINT read Get_RowIndex write Set_RowIndex;
    property Protection: KsoBarProtection read Get_Protection write Set_Protection;
    property Top: SYSINT read Get_Top write Set_Top;
    property type_: KsoBarType read Get_type_;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property Width: SYSINT read Get_Width write Set_Width;
    property AdaptiveMenu: WordBool read Get_AdaptiveMenu write Set_AdaptiveMenu;
    property Id: SYSINT read Get_Id;
  end;

// *********************************************************************//
// DispIntf:  CommandBarDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00023F42-FFFE-0000-C000-000000111146}
// *********************************************************************//
  CommandBarDisp = dispinterface
    ['{00023F42-FFFE-0000-C000-000000111146}']
    property BuiltIn: WordBool readonly dispid 1610874880;
    property Context: WideString dispid 1610874881;
    property Controls: CommandBarControls readonly dispid 1610874883;
    procedure Delete; dispid 1610874884;
    property Enabled: WordBool dispid 1610874885;
    function FindControl(Type_: OleVariant; Id: OleVariant; Tag: OleVariant; Visible: OleVariant; 
                         Recursive: OleVariant): CommandBarControl; dispid 1610874887;
    property Height: SYSINT dispid 1610874888;
    property Index: SYSINT readonly dispid 1610874890;
    property InstanceId: Integer readonly dispid 1610874891;
    property Left: SYSINT dispid 1610874892;
    property Name: WideString dispid 1610874894;
    property NameLocal: WideString dispid 1610874896;
    property Position: KsoBarPosition dispid 1610874899;
    property RowIndex: SYSINT dispid 1610874901;
    property Protection: KsoBarProtection dispid 1610874903;
    procedure Reset; dispid 1610874905;
    procedure ShowPopup(x: OleVariant; y: OleVariant); dispid 1610874906;
    property Top: SYSINT dispid 1610874907;
    property type_: KsoBarType readonly dispid 1610874909;
    property Visible: WordBool dispid 1610874910;
    property Width: SYSINT dispid 1610874912;
    property AdaptiveMenu: WordBool dispid 1610874914;
    property Id: SYSINT readonly dispid 1610874916;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: CommandBarControls
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002448B-FFFE-0000-C000-000000111146}
// *********************************************************************//
  CommandBarControls = interface(IKsoDispObj)
    ['{0002448B-FFFE-0000-C000-000000111146}']
    function Add(Type_: OleVariant; Id: OleVariant; Parameter: OleVariant; Before: OleVariant; 
                 Temporary: OleVariant): CommandBarControl; safecall;
    function Get_Count: SYSINT; safecall;
    function Get_Item(Index: OleVariant): CommandBarControl; safecall;
    function Get__NewEnum: IUnknown; safecall;
    property Count: SYSINT read Get_Count;
    property Item[Index: OleVariant]: CommandBarControl read Get_Item; default;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  CommandBarControlsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002448B-FFFE-0000-C000-000000111146}
// *********************************************************************//
  CommandBarControlsDisp = dispinterface
    ['{0002448B-FFFE-0000-C000-000000111146}']
    function Add(Type_: OleVariant; Id: OleVariant; Parameter: OleVariant; Before: OleVariant; 
                 Temporary: OleVariant): CommandBarControl; dispid 1610809344;
    property Count: SYSINT readonly dispid 1610809345;
    property Item[Index: OleVariant]: CommandBarControl readonly dispid 0; default;
    property _NewEnum: IUnknown readonly dispid -4;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: _CommandBarButton
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00023C90-FFFE-0000-C000-000000111146}
// *********************************************************************//
  _CommandBarButton = interface(CommandBarControl)
    ['{00023C90-FFFE-0000-C000-000000111146}']
    function Get_BuiltInFace: WordBool; safecall;
    procedure Set_BuiltInFace(pvarfBuiltIn: WordBool); safecall;
    procedure CopyFace; safecall;
    function Get_FaceId: SYSINT; safecall;
    procedure Set_FaceId(pid: SYSINT); safecall;
    procedure PasteFace; safecall;
    function Get_ShortcutText: WideString; safecall;
    procedure Set_ShortcutText(const pbstrText: WideString); safecall;
    function Get_State: KsoButtonState; safecall;
    procedure Set_State(pstate: KsoButtonState); safecall;
    function Get_Style: KsoButtonStyle; safecall;
    procedure Set_Style(pstyle: KsoButtonStyle); safecall;
    function Get_HyperlinkType: KsoCommandBarButtonHyperlinkType; safecall;
    procedure Set_HyperlinkType(phlType: KsoCommandBarButtonHyperlinkType); safecall;
    function Get_Picture: IDispatch; safecall;
    procedure Set_Picture(const ppdispPicture: IDispatch); safecall;
    function Get_Mask: IDispatch; safecall;
    procedure Set_Mask(const ppipictdispMask: IDispatch); safecall;
    property BuiltInFace: WordBool read Get_BuiltInFace write Set_BuiltInFace;
    property FaceId: SYSINT read Get_FaceId write Set_FaceId;
    property ShortcutText: WideString read Get_ShortcutText write Set_ShortcutText;
    property State: KsoButtonState read Get_State write Set_State;
    property Style: KsoButtonStyle read Get_Style write Set_Style;
    property HyperlinkType: KsoCommandBarButtonHyperlinkType read Get_HyperlinkType write Set_HyperlinkType;
    property Picture: IDispatch read Get_Picture write Set_Picture;
    property Mask: IDispatch read Get_Mask write Set_Mask;
  end;

// *********************************************************************//
// DispIntf:  _CommandBarButtonDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00023C90-FFFE-0000-C000-000000111146}
// *********************************************************************//
  _CommandBarButtonDisp = dispinterface
    ['{00023C90-FFFE-0000-C000-000000111146}']
    property BuiltInFace: WordBool dispid 1610940416;
    procedure CopyFace; dispid 1610940418;
    property FaceId: SYSINT dispid 1610940419;
    procedure PasteFace; dispid 1610940421;
    property ShortcutText: WideString dispid 1610940422;
    property State: KsoButtonState dispid 1610940424;
    property Style: KsoButtonStyle dispid 1610940426;
    property HyperlinkType: KsoCommandBarButtonHyperlinkType dispid 1610940428;
    property Picture: IDispatch dispid 1610940430;
    property Mask: IDispatch dispid 1610940432;
    property BeginGroup: WordBool dispid 1610874880;
    property BuiltIn: WordBool readonly dispid 1610874882;
    property Caption: WideString dispid 1610874883;
    property Control: IDispatch readonly dispid 1610874885;
    function Copy(Bar: OleVariant; Before: OleVariant): CommandBarControl; dispid 1610874886;
    procedure Delete(Temporary: OleVariant); dispid 1610874887;
    property DescriptionText: WideString dispid 1610874888;
    property Enabled: WordBool dispid 1610874890;
    procedure Execute; dispid 1610874892;
    property Height: SYSINT dispid 1610874893;
    property HelpContextId: SYSINT dispid 1610874895;
    property HelpFile: WideString dispid 1610874897;
    property Id: SYSINT readonly dispid 1610874899;
    property Index: SYSINT readonly dispid 1610874900;
    property InstanceId: Integer readonly dispid 1610874901;
    function Move(Bar: OleVariant; Before: OleVariant): CommandBarControl; dispid 1610874902;
    property Left: SYSINT readonly dispid 1610874903;
    property OLEUsage: KsoControlOLEUsage dispid 1610874904;
    property OnAction: WideString dispid 1610874906;
    property Parameter: WideString dispid 1610874909;
    property Priority: SYSINT dispid 1610874911;
    procedure Reset; dispid 1610874913;
    procedure SetFocus; dispid 1610874914;
    property Tag: WideString dispid 1610874915;
    property TooltipText: WideString dispid 1610874917;
    property Top: SYSINT readonly dispid 1610874919;
    property type_: KsoControlType readonly dispid 1610874920;
    property Visible: WordBool dispid 1610874921;
    property Width: SYSINT dispid 1610874923;
    property IsPriorityDropped: WordBool readonly dispid 1610874925;
    procedure Reserved1; dispid 1610874926;
    procedure Reserved2; dispid 1610874927;
    procedure Reserved3; dispid 1610874928;
    procedure Reserved4; dispid 1610874929;
    procedure Reserved5; dispid 1610874930;
    procedure Reserved6; dispid 1610874931;
    procedure Reserved7; dispid 1610874932;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: CommandBarPopup
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002BF47-FFFE-0000-C000-000000111146}
// *********************************************************************//
  CommandBarPopup = interface(CommandBarControl)
    ['{0002BF47-FFFE-0000-C000-000000111146}']
    function Get_CommandBar: CommandBar; safecall;
    function Get_Controls: CommandBarControls; safecall;
    function Get_OLEMenuGroup: KsoOLEMenuGroup; safecall;
    procedure Set_OLEMenuGroup(pomg: KsoOLEMenuGroup); safecall;
    property CommandBar: CommandBar read Get_CommandBar;
    property Controls: CommandBarControls read Get_Controls;
    property OLEMenuGroup: KsoOLEMenuGroup read Get_OLEMenuGroup write Set_OLEMenuGroup;
  end;

// *********************************************************************//
// DispIntf:  CommandBarPopupDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002BF47-FFFE-0000-C000-000000111146}
// *********************************************************************//
  CommandBarPopupDisp = dispinterface
    ['{0002BF47-FFFE-0000-C000-000000111146}']
    property CommandBar: CommandBar readonly dispid 1610940416;
    property Controls: CommandBarControls readonly dispid 1610940417;
    property OLEMenuGroup: KsoOLEMenuGroup dispid 1610940418;
    property BeginGroup: WordBool dispid 1610874880;
    property BuiltIn: WordBool readonly dispid 1610874882;
    property Caption: WideString dispid 1610874883;
    property Control: IDispatch readonly dispid 1610874885;
    function Copy(Bar: OleVariant; Before: OleVariant): CommandBarControl; dispid 1610874886;
    procedure Delete(Temporary: OleVariant); dispid 1610874887;
    property DescriptionText: WideString dispid 1610874888;
    property Enabled: WordBool dispid 1610874890;
    procedure Execute; dispid 1610874892;
    property Height: SYSINT dispid 1610874893;
    property HelpContextId: SYSINT dispid 1610874895;
    property HelpFile: WideString dispid 1610874897;
    property Id: SYSINT readonly dispid 1610874899;
    property Index: SYSINT readonly dispid 1610874900;
    property InstanceId: Integer readonly dispid 1610874901;
    function Move(Bar: OleVariant; Before: OleVariant): CommandBarControl; dispid 1610874902;
    property Left: SYSINT readonly dispid 1610874903;
    property OLEUsage: KsoControlOLEUsage dispid 1610874904;
    property OnAction: WideString dispid 1610874906;
    property Parameter: WideString dispid 1610874909;
    property Priority: SYSINT dispid 1610874911;
    procedure Reset; dispid 1610874913;
    procedure SetFocus; dispid 1610874914;
    property Tag: WideString dispid 1610874915;
    property TooltipText: WideString dispid 1610874917;
    property Top: SYSINT readonly dispid 1610874919;
    property type_: KsoControlType readonly dispid 1610874920;
    property Visible: WordBool dispid 1610874921;
    property Width: SYSINT dispid 1610874923;
    property IsPriorityDropped: WordBool readonly dispid 1610874925;
    procedure Reserved1; dispid 1610874926;
    procedure Reserved2; dispid 1610874927;
    procedure Reserved3; dispid 1610874928;
    procedure Reserved4; dispid 1610874929;
    procedure Reserved5; dispid 1610874930;
    procedure Reserved6; dispid 1610874931;
    procedure Reserved7; dispid 1610874932;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: _CommandBarComboBox
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002ABB5-FFFE-0000-C000-000000111146}
// *********************************************************************//
  _CommandBarComboBox = interface(CommandBarControl)
    ['{0002ABB5-FFFE-0000-C000-000000111146}']
    procedure AddItem(const Text: WideString; Index: OleVariant); safecall;
    procedure Clear; safecall;
    function Get_DropDownLines: SYSINT; safecall;
    procedure Set_DropDownLines(pcLines: SYSINT); safecall;
    function Get_DropDownWidth: SYSINT; safecall;
    procedure Set_DropDownWidth(pdx: SYSINT); safecall;
    function Get_List(Index: SYSINT): WideString; safecall;
    procedure Set_List(Index: SYSINT; const pbstrItem: WideString); safecall;
    function Get_ListCount: SYSINT; safecall;
    function Get_ListHeaderCount: SYSINT; safecall;
    procedure Set_ListHeaderCount(pcItems: SYSINT); safecall;
    function Get_ListIndex: SYSINT; safecall;
    procedure Set_ListIndex(pi: SYSINT); safecall;
    procedure RemoveItem(Index: SYSINT); safecall;
    function Get_Style: KsoComboStyle; safecall;
    procedure Set_Style(pstyle: KsoComboStyle); safecall;
    function Get_Text: WideString; safecall;
    procedure Set_Text(const pbstrText: WideString); safecall;
    property DropDownLines: SYSINT read Get_DropDownLines write Set_DropDownLines;
    property DropDownWidth: SYSINT read Get_DropDownWidth write Set_DropDownWidth;
    property List[Index: SYSINT]: WideString read Get_List write Set_List;
    property ListCount: SYSINT read Get_ListCount;
    property ListHeaderCount: SYSINT read Get_ListHeaderCount write Set_ListHeaderCount;
    property ListIndex: SYSINT read Get_ListIndex write Set_ListIndex;
    property Style: KsoComboStyle read Get_Style write Set_Style;
    property Text: WideString read Get_Text write Set_Text;
  end;

// *********************************************************************//
// DispIntf:  _CommandBarComboBoxDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002ABB5-FFFE-0000-C000-000000111146}
// *********************************************************************//
  _CommandBarComboBoxDisp = dispinterface
    ['{0002ABB5-FFFE-0000-C000-000000111146}']
    procedure AddItem(const Text: WideString; Index: OleVariant); dispid 1610940416;
    procedure Clear; dispid 1610940417;
    property DropDownLines: SYSINT dispid 1610940418;
    property DropDownWidth: SYSINT dispid 1610940420;
    property List[Index: SYSINT]: WideString dispid 1610940422;
    property ListCount: SYSINT readonly dispid 1610940424;
    property ListHeaderCount: SYSINT dispid 1610940425;
    property ListIndex: SYSINT dispid 1610940427;
    procedure RemoveItem(Index: SYSINT); dispid 1610940429;
    property Style: KsoComboStyle dispid 1610940430;
    property Text: WideString dispid 1610940432;
    property BeginGroup: WordBool dispid 1610874880;
    property BuiltIn: WordBool readonly dispid 1610874882;
    property Caption: WideString dispid 1610874883;
    property Control: IDispatch readonly dispid 1610874885;
    function Copy(Bar: OleVariant; Before: OleVariant): CommandBarControl; dispid 1610874886;
    procedure Delete(Temporary: OleVariant); dispid 1610874887;
    property DescriptionText: WideString dispid 1610874888;
    property Enabled: WordBool dispid 1610874890;
    procedure Execute; dispid 1610874892;
    property Height: SYSINT dispid 1610874893;
    property HelpContextId: SYSINT dispid 1610874895;
    property HelpFile: WideString dispid 1610874897;
    property Id: SYSINT readonly dispid 1610874899;
    property Index: SYSINT readonly dispid 1610874900;
    property InstanceId: Integer readonly dispid 1610874901;
    function Move(Bar: OleVariant; Before: OleVariant): CommandBarControl; dispid 1610874902;
    property Left: SYSINT readonly dispid 1610874903;
    property OLEUsage: KsoControlOLEUsage dispid 1610874904;
    property OnAction: WideString dispid 1610874906;
    property Parameter: WideString dispid 1610874909;
    property Priority: SYSINT dispid 1610874911;
    procedure Reset; dispid 1610874913;
    procedure SetFocus; dispid 1610874914;
    property Tag: WideString dispid 1610874915;
    property TooltipText: WideString dispid 1610874917;
    property Top: SYSINT readonly dispid 1610874919;
    property type_: KsoControlType readonly dispid 1610874920;
    property Visible: WordBool dispid 1610874921;
    property Width: SYSINT dispid 1610874923;
    property IsPriorityDropped: WordBool readonly dispid 1610874925;
    procedure Reserved1; dispid 1610874926;
    procedure Reserved2; dispid 1610874927;
    procedure Reserved3; dispid 1610874928;
    procedure Reserved4; dispid 1610874929;
    procedure Reserved5; dispid 1610874930;
    procedure Reserved6; dispid 1610874931;
    procedure Reserved7; dispid 1610874932;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: ICommandBarsEvents
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000205F4-FFFE-0000-C000-000000111146}
// *********************************************************************//
  ICommandBarsEvents = interface(IDispatch)
    ['{000205F4-FFFE-0000-C000-000000111146}']
    procedure OnUpdate; stdcall;
  end;

// *********************************************************************//
// DispIntf:  ICommandBarsEventsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000205F4-FFFE-0000-C000-000000111146}
// *********************************************************************//
  ICommandBarsEventsDisp = dispinterface
    ['{000205F4-FFFE-0000-C000-000000111146}']
    procedure OnUpdate; dispid 1;
  end;

// *********************************************************************//
// DispIntf:  _CommandBarsEvents
// Flags:     (4112) Hidden Dispatchable
// GUID:      {10227926-9CB9-45B9-ADF2-803D8358EC6A}
// *********************************************************************//
  _CommandBarsEvents = dispinterface
    ['{10227926-9CB9-45B9-ADF2-803D8358EC6A}']
    procedure OnUpdate; dispid 1;
  end;

// *********************************************************************//
// Interface: ICommandBarComboBoxEvents
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002861A-FFFE-0000-C000-000000111146}
// *********************************************************************//
  ICommandBarComboBoxEvents = interface(IDispatch)
    ['{0002861A-FFFE-0000-C000-000000111146}']
    procedure Change(const Ctrl: CommandBarComboBox); stdcall;
  end;

// *********************************************************************//
// DispIntf:  ICommandBarComboBoxEventsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002861A-FFFE-0000-C000-000000111146}
// *********************************************************************//
  ICommandBarComboBoxEventsDisp = dispinterface
    ['{0002861A-FFFE-0000-C000-000000111146}']
    procedure Change(const Ctrl: CommandBarComboBox); dispid 1;
  end;

// *********************************************************************//
// DispIntf:  _CommandBarComboBoxEvents
// Flags:     (4112) Hidden Dispatchable
// GUID:      {2AC87A8B-FABC-4F54-8234-B637DFEC5528}
// *********************************************************************//
  _CommandBarComboBoxEvents = dispinterface
    ['{2AC87A8B-FABC-4F54-8234-B637DFEC5528}']
    procedure Change(const Ctrl: CommandBarComboBox); dispid 1;
  end;

// *********************************************************************//
// DispIntf:  _CommandBarButtonEvents
// Flags:     (4112) Hidden Dispatchable
// GUID:      {B480CA3F-0571-458A-82E4-FEDE4CB3FF96}
// *********************************************************************//
  _CommandBarButtonEvents = dispinterface
    ['{B480CA3F-0571-458A-82E4-FEDE4CB3FF96}']
    procedure Click(const Ctrl: CommandBarButton; var CancelDefault: WordBool); dispid 1;
  end;

// *********************************************************************//
// Interface: ICommandBarButtonEvents
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000258F0-FFFE-0000-C000-000000111146}
// *********************************************************************//
  ICommandBarButtonEvents = interface(IDispatch)
    ['{000258F0-FFFE-0000-C000-000000111146}']
    procedure Click(const Ctrl: CommandBarButton; var CancelDefault: WordBool); stdcall;
  end;

// *********************************************************************//
// DispIntf:  ICommandBarButtonEventsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000258F0-FFFE-0000-C000-000000111146}
// *********************************************************************//
  ICommandBarButtonEventsDisp = dispinterface
    ['{000258F0-FFFE-0000-C000-000000111146}']
    procedure Click(const Ctrl: CommandBarButton; var CancelDefault: WordBool); dispid 1;
  end;

// *********************************************************************//
// Interface: _KsoDiagram
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C036D-FFFF-0000-C000-000000000046}
// *********************************************************************//
  _KsoDiagram = interface(IKsoDispObj)
    ['{000C036D-FFFF-0000-C000-000000000046}']
    function Get__Nodes: _KsoDiagramNodes; safecall;
    function Get_type_: KsoDiagramType; safecall;
    function Get_AutoLayout: KsoTriState; safecall;
    procedure Set_AutoLayout(AutoLayout: KsoTriState); safecall;
    function Get_Reverse: KsoTriState; safecall;
    procedure Set_Reverse(Reverse: KsoTriState); safecall;
    function Get_AutoFormat: KsoTriState; safecall;
    procedure Set_AutoFormat(AutoFormat: KsoTriState); safecall;
    procedure Convert(Type_: KsoDiagramType); safecall;
    property _Nodes: _KsoDiagramNodes read Get__Nodes;
    property type_: KsoDiagramType read Get_type_;
    property AutoLayout: KsoTriState read Get_AutoLayout write Set_AutoLayout;
    property Reverse: KsoTriState read Get_Reverse write Set_Reverse;
    property AutoFormat: KsoTriState read Get_AutoFormat write Set_AutoFormat;
  end;

// *********************************************************************//
// DispIntf:  _KsoDiagramDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C036D-FFFF-0000-C000-000000000046}
// *********************************************************************//
  _KsoDiagramDisp = dispinterface
    ['{000C036D-FFFF-0000-C000-000000000046}']
    property _Nodes: _KsoDiagramNodes readonly dispid 4197;
    property type_: KsoDiagramType readonly dispid 4198;
    property AutoLayout: KsoTriState dispid 4199;
    property Reverse: KsoTriState dispid 4200;
    property AutoFormat: KsoTriState dispid 4201;
    procedure Convert(Type_: KsoDiagramType); dispid 4106;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: _KsoDiagramNodes
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C036E-FFFF-0000-C000-000000000046}
// *********************************************************************//
  _KsoDiagramNodes = interface(IKsoDispObj)
    ['{000C036E-FFFF-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    procedure SelectAll; safecall;
    function Get_Count: SYSINT; safecall;
    function _Item(Index: SYSINT): _KsoDiagramNode; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: SYSINT read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  _KsoDiagramNodesDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C036E-FFFF-0000-C000-000000000046}
// *********************************************************************//
  _KsoDiagramNodesDisp = dispinterface
    ['{000C036E-FFFF-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    procedure SelectAll; dispid 4106;
    property Count: SYSINT readonly dispid 4197;
    function _Item(Index: SYSINT): _KsoDiagramNode; dispid 4198;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: _KsoDiagramNode
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0370-FFFF-0000-C000-000000000046}
// *********************************************************************//
  _KsoDiagramNode = interface(IKsoDispObj)
    ['{000C0370-FFFF-0000-C000-000000000046}']
    function _AddNode(Pos: KsoRelativeNodePosition; NodeType: KsoDiagramNodeType): _KsoDiagramNode; safecall;
    procedure Delete; safecall;
    procedure _MoveNode(const TargetNode: _KsoDiagramNode; Pos: KsoRelativeNodePosition); safecall;
    procedure _ReplaceNode(const TargetNode: _KsoDiagramNode); safecall;
    procedure _SwapNode(const TargetNode: _KsoDiagramNode; SwapChildren: WordBool); safecall;
    function _CloneNode(CopyChildren: WordBool; const TargetNode: _KsoDiagramNode; 
                        Pos: KsoRelativeNodePosition): _KsoDiagramNode; safecall;
    procedure _TransferChildren(const ReceivingNode: _KsoDiagramNode); safecall;
    function _NextNode: _KsoDiagramNode; safecall;
    function _PrevNode: _KsoDiagramNode; safecall;
    function Get_Parent_: IDispatch; safecall;
    function Get__Children: _KsoDiagramNodeChildren; safecall;
    function Get__Shape: KsoShape; safecall;
    function Get__Root: _KsoDiagramNode; safecall;
    function Get__Diagram: _KsoDiagram; safecall;
    function Get_Layout: KsoOrgChartLayoutType; safecall;
    procedure Set_Layout(Type_: KsoOrgChartLayoutType); safecall;
    function Get__TextShape: KsoShape; safecall;
    property Parent_: IDispatch read Get_Parent_;
    property _Children: _KsoDiagramNodeChildren read Get__Children;
    property _Shape: KsoShape read Get__Shape;
    property _Root: _KsoDiagramNode read Get__Root;
    property _Diagram: _KsoDiagram read Get__Diagram;
    property Layout: KsoOrgChartLayoutType read Get_Layout write Set_Layout;
    property _TextShape: KsoShape read Get__TextShape;
  end;

// *********************************************************************//
// DispIntf:  _KsoDiagramNodeDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0370-FFFF-0000-C000-000000000046}
// *********************************************************************//
  _KsoDiagramNodeDisp = dispinterface
    ['{000C0370-FFFF-0000-C000-000000000046}']
    function _AddNode(Pos: KsoRelativeNodePosition; NodeType: KsoDiagramNodeType): _KsoDiagramNode; dispid 4106;
    procedure Delete; dispid 4107;
    procedure _MoveNode(const TargetNode: _KsoDiagramNode; Pos: KsoRelativeNodePosition); dispid 4108;
    procedure _ReplaceNode(const TargetNode: _KsoDiagramNode); dispid 4109;
    procedure _SwapNode(const TargetNode: _KsoDiagramNode; SwapChildren: WordBool); dispid 4110;
    function _CloneNode(CopyChildren: WordBool; const TargetNode: _KsoDiagramNode; 
                        Pos: KsoRelativeNodePosition): _KsoDiagramNode; dispid 4111;
    procedure _TransferChildren(const ReceivingNode: _KsoDiagramNode); dispid 4112;
    function _NextNode: _KsoDiagramNode; dispid 4113;
    function _PrevNode: _KsoDiagramNode; dispid 4114;
    property Parent_: IDispatch readonly dispid 4196;
    property _Children: _KsoDiagramNodeChildren readonly dispid 4197;
    property _Shape: KsoShape readonly dispid 4198;
    property _Root: _KsoDiagramNode readonly dispid 4199;
    property _Diagram: _KsoDiagram readonly dispid 4200;
    property Layout: KsoOrgChartLayoutType dispid 4201;
    property _TextShape: KsoShape readonly dispid 4202;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: _KsoDiagramNodeChildren
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C036F-FFFF-0000-C000-000000000046}
// *********************************************************************//
  _KsoDiagramNodeChildren = interface(IKsoDispObj)
    ['{000C036F-FFFF-0000-C000-000000000046}']
    function Get__NewEnum: IUnknown; safecall;
    function _AddNode(Index: SYSINT; NodeType: KsoDiagramNodeType): _KsoDiagramNode; safecall;
    procedure SelectAll; safecall;
    function Get_Parent_: IDispatch; safecall;
    function Get_Count: SYSINT; safecall;
    function Get__FirstChild: _KsoDiagramNode; safecall;
    function Get__LastChild: _KsoDiagramNode; safecall;
    function _Item(Index: SYSINT): _KsoDiagramNode; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Parent_: IDispatch read Get_Parent_;
    property Count: SYSINT read Get_Count;
    property _FirstChild: _KsoDiagramNode read Get__FirstChild;
    property _LastChild: _KsoDiagramNode read Get__LastChild;
  end;

// *********************************************************************//
// DispIntf:  _KsoDiagramNodeChildrenDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C036F-FFFF-0000-C000-000000000046}
// *********************************************************************//
  _KsoDiagramNodeChildrenDisp = dispinterface
    ['{000C036F-FFFF-0000-C000-000000000046}']
    property _NewEnum: IUnknown readonly dispid -4;
    function _AddNode(Index: SYSINT; NodeType: KsoDiagramNodeType): _KsoDiagramNode; dispid 4106;
    procedure SelectAll; dispid 4107;
    property Parent_: IDispatch readonly dispid 4196;
    property Count: SYSINT readonly dispid 4197;
    property _FirstChild: _KsoDiagramNode readonly dispid 4199;
    property _LastChild: _KsoDiagramNode readonly dispid 4200;
    function _Item(Index: SYSINT): _KsoDiagramNode; dispid 4201;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoShape
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C031C-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoShape = interface(IKsoDispObj)
    ['{000C031C-FFFF-0000-C000-000000000046}']
    procedure Apply; safecall;
    procedure Delete; safecall;
    function _Duplicate: KsoShape; safecall;
    procedure Flip(FlipCmd: KsoFlipCmd); safecall;
    procedure IncrementLeft(Increment: Single); safecall;
    procedure IncrementRotation(Increment: Single); safecall;
    procedure IncrementTop(Increment: Single); safecall;
    procedure PickUp; safecall;
    procedure RerouteConnections; safecall;
    procedure ScaleHeight(Factor: Single; RelativeToOriginalSize: KsoTriState; fScale: KsoScaleFrom); safecall;
    procedure ScaleWidth(Factor: Single; RelativeToOriginalSize: KsoTriState; fScale: KsoScaleFrom); safecall;
    procedure Select(Replace: KsoTriState); safecall;
    procedure SetShapesDefaultProperties; safecall;
    function _Ungroup: KsoShapeRange; safecall;
    procedure ZOrder(ZOrderCmd: KsoZOrderCmd); safecall;
    function Get__Adjustments: KsoAdjustments; safecall;
    function Get_AutoShapeType: KsoAutoShapeType; safecall;
    procedure Set_AutoShapeType(AutoShapeType: KsoAutoShapeType); safecall;
    function Get_BlackWhiteMode: KsoBlackWhiteMode; safecall;
    procedure Set_BlackWhiteMode(BlackWhiteMode: KsoBlackWhiteMode); safecall;
    function Get__Callout: KsoCalloutFormat; safecall;
    function Get_ConnectionSiteCount: SYSINT; safecall;
    function Get_Connector: KsoTriState; safecall;
    function Get__ConnectorFormat: KsoConnectorFormat; safecall;
    function Get__Fill: KsoFillFormat; safecall;
    function Get__GroupItems: KsoGroupShapes; safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(Height: Single); safecall;
    function Get_HorizontalFlip: KsoTriState; safecall;
    function Get_Left: Single; safecall;
    procedure Set_Left(Left: Single); safecall;
    function Get__Line: KsoLineFormat; safecall;
    function Get_LockAspectRatio: KsoTriState; safecall;
    procedure Set_LockAspectRatio(LockAspectRatio: KsoTriState); safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const Name: WideString); safecall;
    function Get__Nodes: KsoShapeNodes; safecall;
    function Get_Rotation: Single; safecall;
    procedure Set_Rotation(Rotation: Single); safecall;
    function Get__PictureFormat: KsoPictureFormat; safecall;
    function Get__Shadow: KsoShadowFormat; safecall;
    function Get__TextEffect: KsoTextEffectFormat; safecall;
    function Get__TextFrame: KsoTextFrame; safecall;
    function Get__ThreeD: KsoThreeDFormat; safecall;
    function Get_Top: Single; safecall;
    procedure Set_Top(Top: Single); safecall;
    function Get__Type: KsoShapeType; safecall;
    function Get_VerticalFlip: KsoTriState; safecall;
    function Get_Vertices: OleVariant; safecall;
    function Get_Visible: KsoTriState; safecall;
    procedure Set_Visible(Visible: KsoTriState); safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(Width: Single); safecall;
    function Get_ZOrderPosition: SYSINT; safecall;
    function Get__Script: KsoScript; safecall;
    function Get_AlternativeText: WideString; safecall;
    procedure Set_AlternativeText(const AlternativeText: WideString); safecall;
    function Get_HasDiagram: KsoTriState; safecall;
    function Get__Diagram: _KsoDiagram; safecall;
    function Get_HasDiagramNode: KsoTriState; safecall;
    function Get__DiagramNode: _KsoDiagramNode; safecall;
    function Get_Child: KsoTriState; safecall;
    function Get__ParentGroup: KsoShape; safecall;
    function Get__CanvasItems: KsoCanvasShapes; safecall;
    function Get_Id: SYSINT; safecall;
    procedure CanvasCropLeft(Increment: Single); safecall;
    procedure CanvasCropTop(Increment: Single); safecall;
    procedure CanvasCropRight(Increment: Single); safecall;
    procedure CanvasCropBottom(Increment: Single); safecall;
    procedure Set_RTF(const Param1: WideString); safecall;
    function Get__OLEFormat: KsoOLEFormat; safecall;
    property _Adjustments: KsoAdjustments read Get__Adjustments;
    property AutoShapeType: KsoAutoShapeType read Get_AutoShapeType write Set_AutoShapeType;
    property BlackWhiteMode: KsoBlackWhiteMode read Get_BlackWhiteMode write Set_BlackWhiteMode;
    property _Callout: KsoCalloutFormat read Get__Callout;
    property ConnectionSiteCount: SYSINT read Get_ConnectionSiteCount;
    property Connector: KsoTriState read Get_Connector;
    property _ConnectorFormat: KsoConnectorFormat read Get__ConnectorFormat;
    property _Fill: KsoFillFormat read Get__Fill;
    property _GroupItems: KsoGroupShapes read Get__GroupItems;
    property Height: Single read Get_Height write Set_Height;
    property HorizontalFlip: KsoTriState read Get_HorizontalFlip;
    property Left: Single read Get_Left write Set_Left;
    property _Line: KsoLineFormat read Get__Line;
    property LockAspectRatio: KsoTriState read Get_LockAspectRatio write Set_LockAspectRatio;
    property Name: WideString read Get_Name write Set_Name;
    property _Nodes: KsoShapeNodes read Get__Nodes;
    property Rotation: Single read Get_Rotation write Set_Rotation;
    property _PictureFormat: KsoPictureFormat read Get__PictureFormat;
    property _Shadow: KsoShadowFormat read Get__Shadow;
    property _TextEffect: KsoTextEffectFormat read Get__TextEffect;
    property _TextFrame: KsoTextFrame read Get__TextFrame;
    property _ThreeD: KsoThreeDFormat read Get__ThreeD;
    property Top: Single read Get_Top write Set_Top;
    property _Type: KsoShapeType read Get__Type;
    property VerticalFlip: KsoTriState read Get_VerticalFlip;
    property Vertices: OleVariant read Get_Vertices;
    property Visible: KsoTriState read Get_Visible write Set_Visible;
    property Width: Single read Get_Width write Set_Width;
    property ZOrderPosition: SYSINT read Get_ZOrderPosition;
    property _Script: KsoScript read Get__Script;
    property AlternativeText: WideString read Get_AlternativeText write Set_AlternativeText;
    property HasDiagram: KsoTriState read Get_HasDiagram;
    property _Diagram: _KsoDiagram read Get__Diagram;
    property HasDiagramNode: KsoTriState read Get_HasDiagramNode;
    property _DiagramNode: _KsoDiagramNode read Get__DiagramNode;
    property Child: KsoTriState read Get_Child;
    property _ParentGroup: KsoShape read Get__ParentGroup;
    property _CanvasItems: KsoCanvasShapes read Get__CanvasItems;
    property Id: SYSINT read Get_Id;
    property RTF: WideString write Set_RTF;
    property _OLEFormat: KsoOLEFormat read Get__OLEFormat;
  end;

// *********************************************************************//
// DispIntf:  KsoShapeDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C031C-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoShapeDisp = dispinterface
    ['{000C031C-FFFF-0000-C000-000000000046}']
    procedure Apply; dispid 4106;
    procedure Delete; dispid 4107;
    function _Duplicate: KsoShape; dispid 4108;
    procedure Flip(FlipCmd: KsoFlipCmd); dispid 4109;
    procedure IncrementLeft(Increment: Single); dispid 4110;
    procedure IncrementRotation(Increment: Single); dispid 4111;
    procedure IncrementTop(Increment: Single); dispid 4112;
    procedure PickUp; dispid 4113;
    procedure RerouteConnections; dispid 4114;
    procedure ScaleHeight(Factor: Single; RelativeToOriginalSize: KsoTriState; fScale: KsoScaleFrom); dispid 4115;
    procedure ScaleWidth(Factor: Single; RelativeToOriginalSize: KsoTriState; fScale: KsoScaleFrom); dispid 4116;
    procedure Select(Replace: KsoTriState); dispid 4117;
    procedure SetShapesDefaultProperties; dispid 4118;
    function _Ungroup: KsoShapeRange; dispid 4119;
    procedure ZOrder(ZOrderCmd: KsoZOrderCmd); dispid 4120;
    property _Adjustments: KsoAdjustments readonly dispid 4196;
    property AutoShapeType: KsoAutoShapeType dispid 4197;
    property BlackWhiteMode: KsoBlackWhiteMode dispid 4198;
    property _Callout: KsoCalloutFormat readonly dispid 4199;
    property ConnectionSiteCount: SYSINT readonly dispid 4200;
    property Connector: KsoTriState readonly dispid 4201;
    property _ConnectorFormat: KsoConnectorFormat readonly dispid 4202;
    property _Fill: KsoFillFormat readonly dispid 4203;
    property _GroupItems: KsoGroupShapes readonly dispid 4204;
    property Height: Single dispid 4205;
    property HorizontalFlip: KsoTriState readonly dispid 4206;
    property Left: Single dispid 4207;
    property _Line: KsoLineFormat readonly dispid 4208;
    property LockAspectRatio: KsoTriState dispid 4209;
    property Name: WideString dispid 4211;
    property _Nodes: KsoShapeNodes readonly dispid 4212;
    property Rotation: Single dispid 4213;
    property _PictureFormat: KsoPictureFormat readonly dispid 4214;
    property _Shadow: KsoShadowFormat readonly dispid 4215;
    property _TextEffect: KsoTextEffectFormat readonly dispid 4216;
    property _TextFrame: KsoTextFrame readonly dispid 69753;
    property _ThreeD: KsoThreeDFormat readonly dispid 4218;
    property Top: Single dispid 4219;
    property _Type: KsoShapeType readonly dispid 69756;
    property VerticalFlip: KsoTriState readonly dispid 4221;
    property Vertices: OleVariant readonly dispid 4222;
    property Visible: KsoTriState dispid 4223;
    property Width: Single dispid 4224;
    property ZOrderPosition: SYSINT readonly dispid 4225;
    property _Script: KsoScript readonly dispid 4226;
    property AlternativeText: WideString dispid 4227;
    property HasDiagram: KsoTriState readonly dispid 4228;
    property _Diagram: _KsoDiagram readonly dispid 4229;
    property HasDiagramNode: KsoTriState readonly dispid 4230;
    property _DiagramNode: _KsoDiagramNode readonly dispid 4231;
    property Child: KsoTriState readonly dispid 4232;
    property _ParentGroup: KsoShape readonly dispid 4233;
    property _CanvasItems: KsoCanvasShapes readonly dispid 4234;
    property Id: SYSINT readonly dispid 4235;
    procedure CanvasCropLeft(Increment: Single); dispid 4236;
    procedure CanvasCropTop(Increment: Single); dispid 4237;
    procedure CanvasCropRight(Increment: Single); dispid 4238;
    procedure CanvasCropBottom(Increment: Single); dispid 4239;
    property RTF: WideString writeonly dispid 4240;
    property _OLEFormat: KsoOLEFormat readonly dispid 4241;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoShapeRange
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C031D-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoShapeRange = interface(IKsoDispObj)
    ['{000C031D-FFFF-0000-C000-000000000046}']
    function Get_Count: SYSINT; safecall;
    function Get__NewEnum: IUnknown; safecall;
    procedure Align(AlignCmd: KsoAlignCmd; RelativeTo: KsoTriState); safecall;
    procedure Apply; safecall;
    procedure Delete; safecall;
    procedure Distribute(DistributeCmd: KsoDistributeCmd; RelativeTo: KsoTriState); safecall;
    function _Duplicate: KsoShapeRange; safecall;
    procedure Flip(FlipCmd: KsoFlipCmd); safecall;
    procedure IncrementLeft(Increment: Single); safecall;
    procedure IncrementRotation(Increment: Single); safecall;
    procedure IncrementTop(Increment: Single); safecall;
    function _Group: KsoShape; safecall;
    procedure PickUp; safecall;
    function _Regroup: KsoShape; safecall;
    procedure RerouteConnections; safecall;
    procedure ScaleHeight(Factor: Single; RelativeToOriginalSize: KsoTriState; fScale: KsoScaleFrom); safecall;
    procedure ScaleWidth(Factor: Single; RelativeToOriginalSize: KsoTriState; fScale: KsoScaleFrom); safecall;
    procedure Select(Replace: KsoTriState); safecall;
    procedure SetShapesDefaultProperties; safecall;
    function _Ungroup: KsoShapeRange; safecall;
    procedure ZOrder(ZOrderCmd: KsoZOrderCmd); safecall;
    function Get__Adjustments: KsoAdjustments; safecall;
    function Get_AutoShapeType: KsoAutoShapeType; safecall;
    procedure Set_AutoShapeType(AutoShapeType: KsoAutoShapeType); safecall;
    function Get_BlackWhiteMode: KsoBlackWhiteMode; safecall;
    procedure Set_BlackWhiteMode(BlackWhiteMode: KsoBlackWhiteMode); safecall;
    function Get__Callout: KsoCalloutFormat; safecall;
    function Get_ConnectionSiteCount: SYSINT; safecall;
    function Get_Connector: KsoTriState; safecall;
    function Get__ConnectorFormat: KsoConnectorFormat; safecall;
    function Get__Fill: KsoFillFormat; safecall;
    function Get__GroupItems: KsoGroupShapes; safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(Height: Single); safecall;
    function Get_HorizontalFlip: KsoTriState; safecall;
    function Get_Left: Single; safecall;
    procedure Set_Left(Left: Single); safecall;
    function Get__Line: KsoLineFormat; safecall;
    function Get_LockAspectRatio: KsoTriState; safecall;
    procedure Set_LockAspectRatio(LockAspectRatio: KsoTriState); safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const Name: WideString); safecall;
    function Get__Nodes: KsoShapeNodes; safecall;
    function Get_Rotation: Single; safecall;
    procedure Set_Rotation(Rotation: Single); safecall;
    function Get__PictureFormat: KsoPictureFormat; safecall;
    function Get__Shadow: KsoShadowFormat; safecall;
    function Get__TextEffect: KsoTextEffectFormat; safecall;
    function Get__TextFrame: KsoTextFrame; safecall;
    function Get__ThreeD: KsoThreeDFormat; safecall;
    function Get_Top: Single; safecall;
    procedure Set_Top(Top: Single); safecall;
    function Get__Type: KsoShapeType; safecall;
    function Get_VerticalFlip: KsoTriState; safecall;
    function Get_Vertices: OleVariant; safecall;
    function Get_Visible: KsoTriState; safecall;
    procedure Set_Visible(Visible: KsoTriState); safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(Width: Single); safecall;
    function Get_ZOrderPosition: SYSINT; safecall;
    function Get__Script: KsoScript; safecall;
    function Get_AlternativeText: WideString; safecall;
    procedure Set_AlternativeText(const AlternativeText: WideString); safecall;
    function Get_HasDiagram: KsoTriState; safecall;
    function Get__Diagram: _KsoDiagram; safecall;
    function Get_HasDiagramNode: KsoTriState; safecall;
    function Get__DiagramNode: _KsoDiagramNode; safecall;
    function Get_Child: KsoTriState; safecall;
    function Get__ParentGroup: KsoShape; safecall;
    function Get__CanvasItems: KsoCanvasShapes; safecall;
    function Get_Id: SYSINT; safecall;
    procedure CanvasCropLeft(Increment: Single); safecall;
    procedure CanvasCropTop(Increment: Single); safecall;
    procedure CanvasCropRight(Increment: Single); safecall;
    procedure CanvasCropBottom(Increment: Single); safecall;
    procedure Set_RTF(const Param1: WideString); safecall;
    function _Item(Index: SYSINT): KsoShape; safecall;
    property Count: SYSINT read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
    property _Adjustments: KsoAdjustments read Get__Adjustments;
    property AutoShapeType: KsoAutoShapeType read Get_AutoShapeType write Set_AutoShapeType;
    property BlackWhiteMode: KsoBlackWhiteMode read Get_BlackWhiteMode write Set_BlackWhiteMode;
    property _Callout: KsoCalloutFormat read Get__Callout;
    property ConnectionSiteCount: SYSINT read Get_ConnectionSiteCount;
    property Connector: KsoTriState read Get_Connector;
    property _ConnectorFormat: KsoConnectorFormat read Get__ConnectorFormat;
    property _Fill: KsoFillFormat read Get__Fill;
    property _GroupItems: KsoGroupShapes read Get__GroupItems;
    property Height: Single read Get_Height write Set_Height;
    property HorizontalFlip: KsoTriState read Get_HorizontalFlip;
    property Left: Single read Get_Left write Set_Left;
    property _Line: KsoLineFormat read Get__Line;
    property LockAspectRatio: KsoTriState read Get_LockAspectRatio write Set_LockAspectRatio;
    property Name: WideString read Get_Name write Set_Name;
    property _Nodes: KsoShapeNodes read Get__Nodes;
    property Rotation: Single read Get_Rotation write Set_Rotation;
    property _PictureFormat: KsoPictureFormat read Get__PictureFormat;
    property _Shadow: KsoShadowFormat read Get__Shadow;
    property _TextEffect: KsoTextEffectFormat read Get__TextEffect;
    property _TextFrame: KsoTextFrame read Get__TextFrame;
    property _ThreeD: KsoThreeDFormat read Get__ThreeD;
    property Top: Single read Get_Top write Set_Top;
    property _Type: KsoShapeType read Get__Type;
    property VerticalFlip: KsoTriState read Get_VerticalFlip;
    property Vertices: OleVariant read Get_Vertices;
    property Visible: KsoTriState read Get_Visible write Set_Visible;
    property Width: Single read Get_Width write Set_Width;
    property ZOrderPosition: SYSINT read Get_ZOrderPosition;
    property _Script: KsoScript read Get__Script;
    property AlternativeText: WideString read Get_AlternativeText write Set_AlternativeText;
    property HasDiagram: KsoTriState read Get_HasDiagram;
    property _Diagram: _KsoDiagram read Get__Diagram;
    property HasDiagramNode: KsoTriState read Get_HasDiagramNode;
    property _DiagramNode: _KsoDiagramNode read Get__DiagramNode;
    property Child: KsoTriState read Get_Child;
    property _ParentGroup: KsoShape read Get__ParentGroup;
    property _CanvasItems: KsoCanvasShapes read Get__CanvasItems;
    property Id: SYSINT read Get_Id;
    property RTF: WideString write Set_RTF;
  end;

// *********************************************************************//
// DispIntf:  KsoShapeRangeDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C031D-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoShapeRangeDisp = dispinterface
    ['{000C031D-FFFF-0000-C000-000000000046}']
    property Count: SYSINT readonly dispid 4098;
    property _NewEnum: IUnknown readonly dispid -4;
    procedure Align(AlignCmd: KsoAlignCmd; RelativeTo: KsoTriState); dispid 4106;
    procedure Apply; dispid 4107;
    procedure Delete; dispid 4108;
    procedure Distribute(DistributeCmd: KsoDistributeCmd; RelativeTo: KsoTriState); dispid 4109;
    function _Duplicate: KsoShapeRange; dispid 4110;
    procedure Flip(FlipCmd: KsoFlipCmd); dispid 4111;
    procedure IncrementLeft(Increment: Single); dispid 4112;
    procedure IncrementRotation(Increment: Single); dispid 4113;
    procedure IncrementTop(Increment: Single); dispid 4114;
    function _Group: KsoShape; dispid 4115;
    procedure PickUp; dispid 4116;
    function _Regroup: KsoShape; dispid 4117;
    procedure RerouteConnections; dispid 4118;
    procedure ScaleHeight(Factor: Single; RelativeToOriginalSize: KsoTriState; fScale: KsoScaleFrom); dispid 4119;
    procedure ScaleWidth(Factor: Single; RelativeToOriginalSize: KsoTriState; fScale: KsoScaleFrom); dispid 4120;
    procedure Select(Replace: KsoTriState); dispid 4121;
    procedure SetShapesDefaultProperties; dispid 4122;
    function _Ungroup: KsoShapeRange; dispid 4123;
    procedure ZOrder(ZOrderCmd: KsoZOrderCmd); dispid 4124;
    property _Adjustments: KsoAdjustments readonly dispid 4196;
    property AutoShapeType: KsoAutoShapeType dispid 4197;
    property BlackWhiteMode: KsoBlackWhiteMode dispid 4198;
    property _Callout: KsoCalloutFormat readonly dispid 4199;
    property ConnectionSiteCount: SYSINT readonly dispid 4200;
    property Connector: KsoTriState readonly dispid 4201;
    property _ConnectorFormat: KsoConnectorFormat readonly dispid 4202;
    property _Fill: KsoFillFormat readonly dispid 4203;
    property _GroupItems: KsoGroupShapes readonly dispid 4204;
    property Height: Single dispid 4205;
    property HorizontalFlip: KsoTriState readonly dispid 4206;
    property Left: Single dispid 4207;
    property _Line: KsoLineFormat readonly dispid 4208;
    property LockAspectRatio: KsoTriState dispid 4209;
    property Name: WideString dispid 4211;
    property _Nodes: KsoShapeNodes readonly dispid 4212;
    property Rotation: Single dispid 4213;
    property _PictureFormat: KsoPictureFormat readonly dispid 4214;
    property _Shadow: KsoShadowFormat readonly dispid 4215;
    property _TextEffect: KsoTextEffectFormat readonly dispid 4216;
    property _TextFrame: KsoTextFrame readonly dispid 69753;
    property _ThreeD: KsoThreeDFormat readonly dispid 4218;
    property Top: Single dispid 4219;
    property _Type: KsoShapeType readonly dispid 4220;
    property VerticalFlip: KsoTriState readonly dispid 4221;
    property Vertices: OleVariant readonly dispid 4222;
    property Visible: KsoTriState dispid 4223;
    property Width: Single dispid 4224;
    property ZOrderPosition: SYSINT readonly dispid 4225;
    property _Script: KsoScript readonly dispid 4226;
    property AlternativeText: WideString dispid 4227;
    property HasDiagram: KsoTriState readonly dispid 4228;
    property _Diagram: _KsoDiagram readonly dispid 4229;
    property HasDiagramNode: KsoTriState readonly dispid 4230;
    property _DiagramNode: _KsoDiagramNode readonly dispid 4231;
    property Child: KsoTriState readonly dispid 4232;
    property _ParentGroup: KsoShape readonly dispid 4233;
    property _CanvasItems: KsoCanvasShapes readonly dispid 4234;
    property Id: SYSINT readonly dispid 4235;
    procedure CanvasCropLeft(Increment: Single); dispid 4236;
    procedure CanvasCropTop(Increment: Single); dispid 4237;
    procedure CanvasCropRight(Increment: Single); dispid 4238;
    procedure CanvasCropBottom(Increment: Single); dispid 4239;
    property RTF: WideString writeonly dispid 4240;
    function _Item(Index: SYSINT): KsoShape; dispid 4241;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoAdjustments
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0310-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoAdjustments = interface(IKsoDispObj)
    ['{000C0310-FFFF-0000-C000-000000000046}']
    function Get_Count: SYSINT; safecall;
    function Get_Item(Index: SYSINT): Single; safecall;
    procedure Set_Item(Index: SYSINT; Val: Single); safecall;
    property Count: SYSINT read Get_Count;
    property Item[Index: SYSINT]: Single read Get_Item write Set_Item;
  end;

// *********************************************************************//
// DispIntf:  KsoAdjustmentsDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0310-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoAdjustmentsDisp = dispinterface
    ['{000C0310-FFFF-0000-C000-000000000046}']
    property Count: SYSINT readonly dispid 4098;
    property Item[Index: SYSINT]: Single dispid 4099;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoCalloutFormat
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0311-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoCalloutFormat = interface(IKsoDispObj)
    ['{000C0311-FFFF-0000-C000-000000000046}']
    procedure AutomaticLength; safecall;
    procedure CustomDrop(Drop: Single); safecall;
    procedure CustomLength(Length: Single); safecall;
    procedure PresetDrop(DropType: KsoCalloutDropType); safecall;
    function Get_Accent: KsoTriState; safecall;
    procedure Set_Accent(Accent: KsoTriState); safecall;
    function Get_Angle: KsoCalloutAngleType; safecall;
    procedure Set_Angle(Angle: KsoCalloutAngleType); safecall;
    function Get_AutoAttach: KsoTriState; safecall;
    procedure Set_AutoAttach(AutoAttach: KsoTriState); safecall;
    function Get_AutoLength: KsoTriState; safecall;
    function Get_Border: KsoTriState; safecall;
    procedure Set_Border(Border: KsoTriState); safecall;
    function Get_Drop: Single; safecall;
    function Get_DropType: KsoCalloutDropType; safecall;
    function Get_Gap: Single; safecall;
    procedure Set_Gap(Gap: Single); safecall;
    function Get_Length: Single; safecall;
    function Get_type_: KsoCalloutType; safecall;
    procedure Set_type_(Type_: KsoCalloutType); safecall;
    property Accent: KsoTriState read Get_Accent write Set_Accent;
    property Angle: KsoCalloutAngleType read Get_Angle write Set_Angle;
    property AutoAttach: KsoTriState read Get_AutoAttach write Set_AutoAttach;
    property AutoLength: KsoTriState read Get_AutoLength;
    property Border: KsoTriState read Get_Border write Set_Border;
    property Drop: Single read Get_Drop;
    property DropType: KsoCalloutDropType read Get_DropType;
    property Gap: Single read Get_Gap write Set_Gap;
    property Length: Single read Get_Length;
    property type_: KsoCalloutType read Get_type_ write Set_type_;
  end;

// *********************************************************************//
// DispIntf:  KsoCalloutFormatDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0311-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoCalloutFormatDisp = dispinterface
    ['{000C0311-FFFF-0000-C000-000000000046}']
    procedure AutomaticLength; dispid 4106;
    procedure CustomDrop(Drop: Single); dispid 4107;
    procedure CustomLength(Length: Single); dispid 4108;
    procedure PresetDrop(DropType: KsoCalloutDropType); dispid 4109;
    property Accent: KsoTriState dispid 4196;
    property Angle: KsoCalloutAngleType dispid 4197;
    property AutoAttach: KsoTriState dispid 4198;
    property AutoLength: KsoTriState readonly dispid 4199;
    property Border: KsoTriState dispid 4200;
    property Drop: Single readonly dispid 4201;
    property DropType: KsoCalloutDropType readonly dispid 4202;
    property Gap: Single dispid 4203;
    property Length: Single readonly dispid 4204;
    property type_: KsoCalloutType dispid 4205;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoConnectorFormat
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0313-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoConnectorFormat = interface(IKsoDispObj)
    ['{000C0313-FFFF-0000-C000-000000000046}']
    procedure _BeginConnect(const ConnectedShape: KsoShape; ConnectionSite: SYSINT); safecall;
    procedure BeginDisconnect; safecall;
    procedure _EndConnect(const ConnectedShape: KsoShape; ConnectionSite: SYSINT); safecall;
    procedure EndDisconnect; safecall;
    function Get_type_: KsoConnectorType; safecall;
    procedure Set_type_(Type_: KsoConnectorType); safecall;
    function Get_BeginConnected: KsoTriState; safecall;
    function Get__BeginConnectedShape: KsoShape; safecall;
    function Get_BeginConnectionSite: SYSINT; safecall;
    function Get_EndConnected: KsoTriState; safecall;
    function Get__EndConnectedShape: KsoShape; safecall;
    function Get_EndConnectionSite: SYSINT; safecall;
    property type_: KsoConnectorType read Get_type_ write Set_type_;
    property BeginConnected: KsoTriState read Get_BeginConnected;
    property _BeginConnectedShape: KsoShape read Get__BeginConnectedShape;
    property BeginConnectionSite: SYSINT read Get_BeginConnectionSite;
    property EndConnected: KsoTriState read Get_EndConnected;
    property _EndConnectedShape: KsoShape read Get__EndConnectedShape;
    property EndConnectionSite: SYSINT read Get_EndConnectionSite;
  end;

// *********************************************************************//
// DispIntf:  KsoConnectorFormatDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0313-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoConnectorFormatDisp = dispinterface
    ['{000C0313-FFFF-0000-C000-000000000046}']
    procedure _BeginConnect(const ConnectedShape: KsoShape; ConnectionSite: SYSINT); dispid 4106;
    procedure BeginDisconnect; dispid 4107;
    procedure _EndConnect(const ConnectedShape: KsoShape; ConnectionSite: SYSINT); dispid 4108;
    procedure EndDisconnect; dispid 4109;
    property type_: KsoConnectorType dispid 4202;
    property BeginConnected: KsoTriState readonly dispid 4196;
    property _BeginConnectedShape: KsoShape readonly dispid 4197;
    property BeginConnectionSite: SYSINT readonly dispid 4198;
    property EndConnected: KsoTriState readonly dispid 4199;
    property _EndConnectedShape: KsoShape readonly dispid 4200;
    property EndConnectionSite: SYSINT readonly dispid 4201;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoFillFormat
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0314-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoFillFormat = interface(IKsoDispObj)
    ['{000C0314-FFFF-0000-C000-000000000046}']
    procedure Background; safecall;
    procedure OneColorGradient(Style: KsoGradientStyle; Variant: SYSINT; Degree: Single); safecall;
    procedure Patterned(Pattern: KsoPatternType); safecall;
    procedure PresetGradient(Style: KsoGradientStyle; Variant: SYSINT; 
                             PresetGradientType: KsoPresetGradientType); safecall;
    procedure PresetTextured(PresetTexture: KsoPresetTexture); safecall;
    procedure Solid; safecall;
    procedure TwoColorGradient(Style: KsoGradientStyle; Variant: SYSINT); safecall;
    procedure UserPicture(const PictureFile: WideString); safecall;
    procedure UserTextured(const TextureFile: WideString); safecall;
    function Get__BackColor: KsoColorFormat; safecall;
    procedure Set__BackColor(const BackColor: KsoColorFormat); safecall;
    function Get__ForeColor: KsoColorFormat; safecall;
    procedure Set__ForeColor(const ForeColor: KsoColorFormat); safecall;
    function Get_GradientColorType: KsoGradientColorType; safecall;
    function Get_GradientDegree: Single; safecall;
    function Get_GradientStyle: KsoGradientStyle; safecall;
    function Get_GradientVariant: SYSINT; safecall;
    function Get_Pattern: KsoPatternType; safecall;
    function Get_PresetGradientType: KsoPresetGradientType; safecall;
    function Get_PresetTexture: KsoPresetTexture; safecall;
    function Get_TextureName: WideString; safecall;
    function Get_TextureType: KsoTextureType; safecall;
    function Get_Transparency: Single; safecall;
    procedure Set_Transparency(Transparency: Single); safecall;
    function Get_type_: KsoFillType; safecall;
    function Get_Visible: KsoTriState; safecall;
    procedure Set_Visible(Visible: KsoTriState); safecall;
    property _BackColor: KsoColorFormat read Get__BackColor write Set__BackColor;
    property _ForeColor: KsoColorFormat read Get__ForeColor write Set__ForeColor;
    property GradientColorType: KsoGradientColorType read Get_GradientColorType;
    property GradientDegree: Single read Get_GradientDegree;
    property GradientStyle: KsoGradientStyle read Get_GradientStyle;
    property GradientVariant: SYSINT read Get_GradientVariant;
    property Pattern: KsoPatternType read Get_Pattern;
    property PresetGradientType: KsoPresetGradientType read Get_PresetGradientType;
    property PresetTexture: KsoPresetTexture read Get_PresetTexture;
    property TextureName: WideString read Get_TextureName;
    property TextureType: KsoTextureType read Get_TextureType;
    property Transparency: Single read Get_Transparency write Set_Transparency;
    property type_: KsoFillType read Get_type_;
    property Visible: KsoTriState read Get_Visible write Set_Visible;
  end;

// *********************************************************************//
// DispIntf:  KsoFillFormatDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0314-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoFillFormatDisp = dispinterface
    ['{000C0314-FFFF-0000-C000-000000000046}']
    procedure Background; dispid 4106;
    procedure OneColorGradient(Style: KsoGradientStyle; Variant: SYSINT; Degree: Single); dispid 4107;
    procedure Patterned(Pattern: KsoPatternType); dispid 4108;
    procedure PresetGradient(Style: KsoGradientStyle; Variant: SYSINT; 
                             PresetGradientType: KsoPresetGradientType); dispid 4109;
    procedure PresetTextured(PresetTexture: KsoPresetTexture); dispid 4110;
    procedure Solid; dispid 4111;
    procedure TwoColorGradient(Style: KsoGradientStyle; Variant: SYSINT); dispid 4112;
    procedure UserPicture(const PictureFile: WideString); dispid 4113;
    procedure UserTextured(const TextureFile: WideString); dispid 4114;
    property _BackColor: KsoColorFormat dispid 4196;
    property _ForeColor: KsoColorFormat dispid 4197;
    property GradientColorType: KsoGradientColorType readonly dispid 4198;
    property GradientDegree: Single readonly dispid 4199;
    property GradientStyle: KsoGradientStyle readonly dispid 4200;
    property GradientVariant: SYSINT readonly dispid 4201;
    property Pattern: KsoPatternType readonly dispid 4202;
    property PresetGradientType: KsoPresetGradientType readonly dispid 4203;
    property PresetTexture: KsoPresetTexture readonly dispid 4204;
    property TextureName: WideString readonly dispid 4205;
    property TextureType: KsoTextureType readonly dispid 4206;
    property Transparency: Single dispid 4207;
    property type_: KsoFillType readonly dispid 4208;
    property Visible: KsoTriState dispid 4209;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoColorFormat
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0312-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoColorFormat = interface(IKsoDispObj)
    ['{000C0312-FFFF-0000-C000-000000000046}']
    function Get_RGB: Integer; safecall;
    procedure Set_RGB(RGB: Integer); safecall;
    function Get_SchemeColor: SYSINT; safecall;
    procedure Set_SchemeColor(SchemeColor: SYSINT); safecall;
    function Get_type_: KsoColorType; safecall;
    function Get_TintAndShade: Single; safecall;
    procedure Set_TintAndShade(pValue: Single); safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    function Get_OverPrint: KsoTriState; safecall;
    procedure Set_OverPrint(prop: KsoTriState); safecall;
    function Get_Ink(Index: Integer): Single; safecall;
    procedure Set_Ink(Index: Integer; prop: Single); safecall;
    function Get_Cyan: Integer; safecall;
    procedure Set_Cyan(prop: Integer); safecall;
    function Get_Magenta: Integer; safecall;
    procedure Set_Magenta(prop: Integer); safecall;
    function Get_Yellow: Integer; safecall;
    procedure Set_Yellow(prop: Integer); safecall;
    function Get_Black: Integer; safecall;
    procedure Set_Black(prop: Integer); safecall;
    procedure SetCMYK(Cyan: Integer; Magenta: Integer; Yellow: Integer; Black: Integer); safecall;
    property RGB: Integer read Get_RGB write Set_RGB;
    property SchemeColor: SYSINT read Get_SchemeColor write Set_SchemeColor;
    property type_: KsoColorType read Get_type_;
    property TintAndShade: Single read Get_TintAndShade write Set_TintAndShade;
    property Name: WideString read Get_Name write Set_Name;
    property OverPrint: KsoTriState read Get_OverPrint write Set_OverPrint;
    property Ink[Index: Integer]: Single read Get_Ink write Set_Ink;
    property Cyan: Integer read Get_Cyan write Set_Cyan;
    property Magenta: Integer read Get_Magenta write Set_Magenta;
    property Yellow: Integer read Get_Yellow write Set_Yellow;
    property Black: Integer read Get_Black write Set_Black;
  end;

// *********************************************************************//
// DispIntf:  KsoColorFormatDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0312-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoColorFormatDisp = dispinterface
    ['{000C0312-FFFF-0000-C000-000000000046}']
    property RGB: Integer dispid 512;
    property SchemeColor: SYSINT dispid 4196;
    property type_: KsoColorType readonly dispid 4197;
    property TintAndShade: Single dispid 4199;
    property Name: WideString dispid 4198;
    property OverPrint: KsoTriState dispid 4200;
    property Ink[Index: Integer]: Single dispid 4201;
    property Cyan: Integer dispid 4202;
    property Magenta: Integer dispid 4203;
    property Yellow: Integer dispid 4204;
    property Black: Integer dispid 4205;
    procedure SetCMYK(Cyan: Integer; Magenta: Integer; Yellow: Integer; Black: Integer); dispid 4206;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoGroupShapes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0316-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoGroupShapes = interface(IKsoDispObj)
    ['{000C0316-FFFF-0000-C000-000000000046}']
    function Get_Count: SYSINT; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function _Range(Index: OleVariant): KsoShapeRange; safecall;
    function _Item(Index: SYSINT): KsoShape; safecall;
    property Count: SYSINT read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  KsoGroupShapesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0316-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoGroupShapesDisp = dispinterface
    ['{000C0316-FFFF-0000-C000-000000000046}']
    property Count: SYSINT readonly dispid 4098;
    property _NewEnum: IUnknown readonly dispid -4;
    function _Range(Index: OleVariant): KsoShapeRange; dispid 4106;
    function _Item(Index: SYSINT): KsoShape; dispid 4107;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoLineFormat
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0317-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoLineFormat = interface(IKsoDispObj)
    ['{000C0317-FFFF-0000-C000-000000000046}']
    function Get__BackColor: KsoColorFormat; safecall;
    procedure Set__BackColor(const BackColor: KsoColorFormat); safecall;
    function Get_BeginArrowheadLength: KsoArrowheadLength; safecall;
    procedure Set_BeginArrowheadLength(BeginArrowheadLength: KsoArrowheadLength); safecall;
    function Get_BeginArrowheadStyle: KsoArrowheadStyle; safecall;
    procedure Set_BeginArrowheadStyle(BeginArrowheadStyle: KsoArrowheadStyle); safecall;
    function Get_BeginArrowheadWidth: KsoArrowheadWidth; safecall;
    procedure Set_BeginArrowheadWidth(BeginArrowheadWidth: KsoArrowheadWidth); safecall;
    function Get_DashStyle: KsoLineDashStyle; safecall;
    procedure Set_DashStyle(DashStyle: KsoLineDashStyle); safecall;
    function Get_EndArrowheadLength: KsoArrowheadLength; safecall;
    procedure Set_EndArrowheadLength(EndArrowheadLength: KsoArrowheadLength); safecall;
    function Get_EndArrowheadStyle: KsoArrowheadStyle; safecall;
    procedure Set_EndArrowheadStyle(EndArrowheadStyle: KsoArrowheadStyle); safecall;
    function Get_EndArrowheadWidth: KsoArrowheadWidth; safecall;
    procedure Set_EndArrowheadWidth(EndArrowheadWidth: KsoArrowheadWidth); safecall;
    function Get__ForeColor: KsoColorFormat; safecall;
    procedure Set__ForeColor(const ForeColor: KsoColorFormat); safecall;
    function Get_Pattern: KsoPatternType; safecall;
    procedure Set_Pattern(Pattern: KsoPatternType); safecall;
    function Get_Style: KsoLineStyle; safecall;
    procedure Set_Style(Style: KsoLineStyle); safecall;
    function Get_Transparency: Single; safecall;
    procedure Set_Transparency(Transparency: Single); safecall;
    function Get_Visible: KsoTriState; safecall;
    procedure Set_Visible(Visible: KsoTriState); safecall;
    function Get_Weight: Single; safecall;
    procedure Set_Weight(Weight: Single); safecall;
    function Get_InsetPen: KsoTriState; safecall;
    procedure Set_InsetPen(InsetPen: KsoTriState); safecall;
    property _BackColor: KsoColorFormat read Get__BackColor write Set__BackColor;
    property BeginArrowheadLength: KsoArrowheadLength read Get_BeginArrowheadLength write Set_BeginArrowheadLength;
    property BeginArrowheadStyle: KsoArrowheadStyle read Get_BeginArrowheadStyle write Set_BeginArrowheadStyle;
    property BeginArrowheadWidth: KsoArrowheadWidth read Get_BeginArrowheadWidth write Set_BeginArrowheadWidth;
    property DashStyle: KsoLineDashStyle read Get_DashStyle write Set_DashStyle;
    property EndArrowheadLength: KsoArrowheadLength read Get_EndArrowheadLength write Set_EndArrowheadLength;
    property EndArrowheadStyle: KsoArrowheadStyle read Get_EndArrowheadStyle write Set_EndArrowheadStyle;
    property EndArrowheadWidth: KsoArrowheadWidth read Get_EndArrowheadWidth write Set_EndArrowheadWidth;
    property _ForeColor: KsoColorFormat read Get__ForeColor write Set__ForeColor;
    property Pattern: KsoPatternType read Get_Pattern write Set_Pattern;
    property Style: KsoLineStyle read Get_Style write Set_Style;
    property Transparency: Single read Get_Transparency write Set_Transparency;
    property Visible: KsoTriState read Get_Visible write Set_Visible;
    property Weight: Single read Get_Weight write Set_Weight;
    property InsetPen: KsoTriState read Get_InsetPen write Set_InsetPen;
  end;

// *********************************************************************//
// DispIntf:  KsoLineFormatDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0317-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoLineFormatDisp = dispinterface
    ['{000C0317-FFFF-0000-C000-000000000046}']
    property _BackColor: KsoColorFormat dispid 4196;
    property BeginArrowheadLength: KsoArrowheadLength dispid 4197;
    property BeginArrowheadStyle: KsoArrowheadStyle dispid 4198;
    property BeginArrowheadWidth: KsoArrowheadWidth dispid 4199;
    property DashStyle: KsoLineDashStyle dispid 4200;
    property EndArrowheadLength: KsoArrowheadLength dispid 4201;
    property EndArrowheadStyle: KsoArrowheadStyle dispid 4202;
    property EndArrowheadWidth: KsoArrowheadWidth dispid 4203;
    property _ForeColor: KsoColorFormat dispid 4204;
    property Pattern: KsoPatternType dispid 4205;
    property Style: KsoLineStyle dispid 4206;
    property Transparency: Single dispid 4207;
    property Visible: KsoTriState dispid 4208;
    property Weight: Single dispid 4209;
    property InsetPen: KsoTriState dispid 4210;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoShapeNodes
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0319-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoShapeNodes = interface(IKsoDispObj)
    ['{000C0319-FFFF-0000-C000-000000000046}']
    function Get_Count: SYSINT; safecall;
    function Get__NewEnum: IUnknown; safecall;
    procedure Delete(Index: SYSINT); safecall;
    procedure Insert(Index: SYSINT; SegmentType: KsoSegmentType; EditingType: KsoEditingType; 
                     X1: Single; Y1: Single; X2: Single; Y2: Single; X3: Single; Y3: Single); safecall;
    procedure SetEditingType(Index: SYSINT; EditingType: KsoEditingType); safecall;
    procedure SetPosition(Index: SYSINT; X1: Single; Y1: Single); safecall;
    procedure SetSegmentType(Index: SYSINT; SegmentType: KsoSegmentType); safecall;
    function _Item(Index: SYSINT): KsoShapeNode; safecall;
    property Count: SYSINT read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  KsoShapeNodesDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0319-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoShapeNodesDisp = dispinterface
    ['{000C0319-FFFF-0000-C000-000000000046}']
    property Count: SYSINT readonly dispid 4098;
    property _NewEnum: IUnknown readonly dispid -4;
    procedure Delete(Index: SYSINT); dispid 4107;
    procedure Insert(Index: SYSINT; SegmentType: KsoSegmentType; EditingType: KsoEditingType; 
                     X1: Single; Y1: Single; X2: Single; Y2: Single; X3: Single; Y3: Single); dispid 4108;
    procedure SetEditingType(Index: SYSINT; EditingType: KsoEditingType); dispid 4109;
    procedure SetPosition(Index: SYSINT; X1: Single; Y1: Single); dispid 4110;
    procedure SetSegmentType(Index: SYSINT; SegmentType: KsoSegmentType); dispid 4111;
    function _Item(Index: SYSINT): KsoShapeNode; dispid 4112;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoShapeNode
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0318-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoShapeNode = interface(IKsoDispObj)
    ['{000C0318-FFFF-0000-C000-000000000046}']
    function Get_EditingType: KsoEditingType; safecall;
    function Get_Points: OleVariant; safecall;
    function Get_SegmentType: KsoSegmentType; safecall;
    property EditingType: KsoEditingType read Get_EditingType;
    property Points: OleVariant read Get_Points;
    property SegmentType: KsoSegmentType read Get_SegmentType;
  end;

// *********************************************************************//
// DispIntf:  KsoShapeNodeDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0318-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoShapeNodeDisp = dispinterface
    ['{000C0318-FFFF-0000-C000-000000000046}']
    property EditingType: KsoEditingType readonly dispid 4196;
    property Points: OleVariant readonly dispid 101;
    property SegmentType: KsoSegmentType readonly dispid 4198;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoPictureFormat
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C031A-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoPictureFormat = interface(IKsoDispObj)
    ['{000C031A-FFFF-0000-C000-000000000046}']
    procedure IncrementBrightness(Increment: Single); safecall;
    procedure IncrementContrast(Increment: Single); safecall;
    function Get_Brightness: Single; safecall;
    procedure Set_Brightness(Brightness: Single); safecall;
    function Get_ColorType: KsoPictureColorType; safecall;
    procedure Set_ColorType(ColorType: KsoPictureColorType); safecall;
    function Get_Contrast: Single; safecall;
    procedure Set_Contrast(Contrast: Single); safecall;
    function Get_CropBottom: Single; safecall;
    procedure Set_CropBottom(CropBottom: Single); safecall;
    function Get_CropLeft: Single; safecall;
    procedure Set_CropLeft(CropLeft: Single); safecall;
    function Get_CropRight: Single; safecall;
    procedure Set_CropRight(CropRight: Single); safecall;
    function Get_CropTop: Single; safecall;
    procedure Set_CropTop(CropTop: Single); safecall;
    function Get_TransparencyColor: Integer; safecall;
    procedure Set_TransparencyColor(TransparencyColor: Integer); safecall;
    function Get_TransparentBackground: KsoTriState; safecall;
    procedure Set_TransparentBackground(TransparentBackground: KsoTriState); safecall;
    property Brightness: Single read Get_Brightness write Set_Brightness;
    property ColorType: KsoPictureColorType read Get_ColorType write Set_ColorType;
    property Contrast: Single read Get_Contrast write Set_Contrast;
    property CropBottom: Single read Get_CropBottom write Set_CropBottom;
    property CropLeft: Single read Get_CropLeft write Set_CropLeft;
    property CropRight: Single read Get_CropRight write Set_CropRight;
    property CropTop: Single read Get_CropTop write Set_CropTop;
    property TransparencyColor: Integer read Get_TransparencyColor write Set_TransparencyColor;
    property TransparentBackground: KsoTriState read Get_TransparentBackground write Set_TransparentBackground;
  end;

// *********************************************************************//
// DispIntf:  KsoPictureFormatDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C031A-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoPictureFormatDisp = dispinterface
    ['{000C031A-FFFF-0000-C000-000000000046}']
    procedure IncrementBrightness(Increment: Single); dispid 4106;
    procedure IncrementContrast(Increment: Single); dispid 4107;
    property Brightness: Single dispid 4196;
    property ColorType: KsoPictureColorType dispid 4197;
    property Contrast: Single dispid 4198;
    property CropBottom: Single dispid 4199;
    property CropLeft: Single dispid 4200;
    property CropRight: Single dispid 4201;
    property CropTop: Single dispid 4202;
    property TransparencyColor: Integer dispid 4203;
    property TransparentBackground: KsoTriState dispid 4204;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoShadowFormat
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C031B-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoShadowFormat = interface(IKsoDispObj)
    ['{000C031B-FFFF-0000-C000-000000000046}']
    procedure IncrementOffsetX(Increment: Single); safecall;
    procedure IncrementOffsetY(Increment: Single); safecall;
    function Get__ForeColor: KsoColorFormat; safecall;
    procedure Set__ForeColor(const ForeColor: KsoColorFormat); safecall;
    function Get_Obscured: KsoTriState; safecall;
    procedure Set_Obscured(Obscured: KsoTriState); safecall;
    function Get_OffsetX: Single; safecall;
    procedure Set_OffsetX(OffsetX: Single); safecall;
    function Get_OffsetY: Single; safecall;
    procedure Set_OffsetY(OffsetY: Single); safecall;
    function Get_Transparency: Single; safecall;
    procedure Set_Transparency(Transparency: Single); safecall;
    function Get_type_: KsoShadowType; safecall;
    procedure Set_type_(Type_: KsoShadowType); safecall;
    function Get_Visible: KsoTriState; safecall;
    procedure Set_Visible(Visible: KsoTriState); safecall;
    property _ForeColor: KsoColorFormat read Get__ForeColor write Set__ForeColor;
    property Obscured: KsoTriState read Get_Obscured write Set_Obscured;
    property OffsetX: Single read Get_OffsetX write Set_OffsetX;
    property OffsetY: Single read Get_OffsetY write Set_OffsetY;
    property Transparency: Single read Get_Transparency write Set_Transparency;
    property type_: KsoShadowType read Get_type_ write Set_type_;
    property Visible: KsoTriState read Get_Visible write Set_Visible;
  end;

// *********************************************************************//
// DispIntf:  KsoShadowFormatDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C031B-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoShadowFormatDisp = dispinterface
    ['{000C031B-FFFF-0000-C000-000000000046}']
    procedure IncrementOffsetX(Increment: Single); dispid 4106;
    procedure IncrementOffsetY(Increment: Single); dispid 4107;
    property _ForeColor: KsoColorFormat dispid 4196;
    property Obscured: KsoTriState dispid 4197;
    property OffsetX: Single dispid 4198;
    property OffsetY: Single dispid 4199;
    property Transparency: Single dispid 4200;
    property type_: KsoShadowType dispid 4201;
    property Visible: KsoTriState dispid 4202;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoTextEffectFormat
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C031F-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoTextEffectFormat = interface(IKsoDispObj)
    ['{000C031F-FFFF-0000-C000-000000000046}']
    procedure ToggleVerticalText; safecall;
    function Get_Alignment: KsoTextEffectAlignment; safecall;
    procedure Set_Alignment(Alignment: KsoTextEffectAlignment); safecall;
    function Get_FontBold: KsoTriState; safecall;
    procedure Set_FontBold(FontBold: KsoTriState); safecall;
    function Get_FontItalic: KsoTriState; safecall;
    procedure Set_FontItalic(FontItalic: KsoTriState); safecall;
    function Get_FontName: WideString; safecall;
    procedure Set_FontName(const FontName: WideString); safecall;
    function Get_FontSize: Single; safecall;
    procedure Set_FontSize(FontSize: Single); safecall;
    function Get_KernedPairs: KsoTriState; safecall;
    procedure Set_KernedPairs(KernedPairs: KsoTriState); safecall;
    function Get_NormalizedHeight: KsoTriState; safecall;
    procedure Set_NormalizedHeight(NormalizedHeight: KsoTriState); safecall;
    function Get_PresetShape: KsoPresetTextEffectShape; safecall;
    procedure Set_PresetShape(PresetShape: KsoPresetTextEffectShape); safecall;
    function Get_PresetTextEffect: KsoPresetTextEffect; safecall;
    procedure Set_PresetTextEffect(Preset: KsoPresetTextEffect); safecall;
    function Get_RotatedChars: KsoTriState; safecall;
    procedure Set_RotatedChars(RotatedChars: KsoTriState); safecall;
    function Get_Text: WideString; safecall;
    procedure Set_Text(const Text: WideString); safecall;
    function Get_Tracking: Single; safecall;
    procedure Set_Tracking(Tracking: Single); safecall;
    property Alignment: KsoTextEffectAlignment read Get_Alignment write Set_Alignment;
    property FontBold: KsoTriState read Get_FontBold write Set_FontBold;
    property FontItalic: KsoTriState read Get_FontItalic write Set_FontItalic;
    property FontName: WideString read Get_FontName write Set_FontName;
    property FontSize: Single read Get_FontSize write Set_FontSize;
    property KernedPairs: KsoTriState read Get_KernedPairs write Set_KernedPairs;
    property NormalizedHeight: KsoTriState read Get_NormalizedHeight write Set_NormalizedHeight;
    property PresetShape: KsoPresetTextEffectShape read Get_PresetShape write Set_PresetShape;
    property PresetTextEffect: KsoPresetTextEffect read Get_PresetTextEffect write Set_PresetTextEffect;
    property RotatedChars: KsoTriState read Get_RotatedChars write Set_RotatedChars;
    property Text: WideString read Get_Text write Set_Text;
    property Tracking: Single read Get_Tracking write Set_Tracking;
  end;

// *********************************************************************//
// DispIntf:  KsoTextEffectFormatDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C031F-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoTextEffectFormatDisp = dispinterface
    ['{000C031F-FFFF-0000-C000-000000000046}']
    procedure ToggleVerticalText; dispid 4106;
    property Alignment: KsoTextEffectAlignment dispid 4196;
    property FontBold: KsoTriState dispid 4197;
    property FontItalic: KsoTriState dispid 4198;
    property FontName: WideString dispid 4199;
    property FontSize: Single dispid 4200;
    property KernedPairs: KsoTriState dispid 4201;
    property NormalizedHeight: KsoTriState dispid 4202;
    property PresetShape: KsoPresetTextEffectShape dispid 4203;
    property PresetTextEffect: KsoPresetTextEffect dispid 4204;
    property RotatedChars: KsoTriState dispid 4205;
    property Text: WideString dispid 4206;
    property Tracking: Single dispid 4207;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoTextFrame
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C3484-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoTextFrame = interface(IKsoDispObj)
    ['{000C3484-FFFF-0000-C000-000000000046}']
    function Get_MarginBottom: Single; safecall;
    procedure Set_MarginBottom(MarginBottom: Single); safecall;
    function Get_MarginLeft: Single; safecall;
    procedure Set_MarginLeft(MarginLeft: Single); safecall;
    function Get_MarginRight: Single; safecall;
    procedure Set_MarginRight(MarginRight: Single); safecall;
    function Get_MarginTop: Single; safecall;
    procedure Set_MarginTop(MarginTop: Single); safecall;
    function Get_Orientation: KsoTextOrientation; safecall;
    procedure Set_Orientation(Orientation: KsoTextOrientation); safecall;
    property MarginBottom: Single read Get_MarginBottom write Set_MarginBottom;
    property MarginLeft: Single read Get_MarginLeft write Set_MarginLeft;
    property MarginRight: Single read Get_MarginRight write Set_MarginRight;
    property MarginTop: Single read Get_MarginTop write Set_MarginTop;
    property Orientation: KsoTextOrientation read Get_Orientation write Set_Orientation;
  end;

// *********************************************************************//
// DispIntf:  KsoTextFrameDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C3484-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoTextFrameDisp = dispinterface
    ['{000C3484-FFFF-0000-C000-000000000046}']
    property MarginBottom: Single dispid 4196;
    property MarginLeft: Single dispid 4197;
    property MarginRight: Single dispid 4198;
    property MarginTop: Single dispid 4199;
    property Orientation: KsoTextOrientation dispid 4200;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoThreeDFormat
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0321-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoThreeDFormat = interface(IKsoDispObj)
    ['{000C0321-FFFF-0000-C000-000000000046}']
    procedure IncrementRotationX(Increment: Single); safecall;
    procedure IncrementRotationY(Increment: Single); safecall;
    procedure ResetRotation; safecall;
    procedure SetThreeDFormat(PresetThreeDFormat: KsoPresetThreeDFormat); safecall;
    procedure SetExtrusionDirection(PresetExtrusionDirection: KsoPresetExtrusionDirection); safecall;
    function Get_Depth: Single; safecall;
    procedure Set_Depth(Depth: Single); safecall;
    function Get__ExtrusionColor: KsoColorFormat; safecall;
    function Get_ExtrusionColorType: KsoExtrusionColorType; safecall;
    procedure Set_ExtrusionColorType(ExtrusionColorType: KsoExtrusionColorType); safecall;
    function Get_Perspective: KsoTriState; safecall;
    procedure Set_Perspective(Perspective: KsoTriState); safecall;
    function Get_PresetExtrusionDirection: KsoPresetExtrusionDirection; safecall;
    function Get_PresetLightingDirection: KsoPresetLightingDirection; safecall;
    procedure Set_PresetLightingDirection(PresetLightingDirection: KsoPresetLightingDirection); safecall;
    function Get_PresetLightingSoftness: KsoPresetLightingSoftness; safecall;
    procedure Set_PresetLightingSoftness(PresetLightingSoftness: KsoPresetLightingSoftness); safecall;
    function Get_PresetMaterial: KsoPresetMaterial; safecall;
    procedure Set_PresetMaterial(PresetMaterial: KsoPresetMaterial); safecall;
    function Get_PresetThreeDFormat: KsoPresetThreeDFormat; safecall;
    function Get_RotationX: Single; safecall;
    procedure Set_RotationX(RotationX: Single); safecall;
    function Get_RotationY: Single; safecall;
    procedure Set_RotationY(RotationY: Single); safecall;
    function Get_Visible: KsoTriState; safecall;
    procedure Set_Visible(Visible: KsoTriState); safecall;
    property Depth: Single read Get_Depth write Set_Depth;
    property _ExtrusionColor: KsoColorFormat read Get__ExtrusionColor;
    property ExtrusionColorType: KsoExtrusionColorType read Get_ExtrusionColorType write Set_ExtrusionColorType;
    property Perspective: KsoTriState read Get_Perspective write Set_Perspective;
    property PresetExtrusionDirection: KsoPresetExtrusionDirection read Get_PresetExtrusionDirection;
    property PresetLightingDirection: KsoPresetLightingDirection read Get_PresetLightingDirection write Set_PresetLightingDirection;
    property PresetLightingSoftness: KsoPresetLightingSoftness read Get_PresetLightingSoftness write Set_PresetLightingSoftness;
    property PresetMaterial: KsoPresetMaterial read Get_PresetMaterial write Set_PresetMaterial;
    property PresetThreeDFormat: KsoPresetThreeDFormat read Get_PresetThreeDFormat;
    property RotationX: Single read Get_RotationX write Set_RotationX;
    property RotationY: Single read Get_RotationY write Set_RotationY;
    property Visible: KsoTriState read Get_Visible write Set_Visible;
  end;

// *********************************************************************//
// DispIntf:  KsoThreeDFormatDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0321-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoThreeDFormatDisp = dispinterface
    ['{000C0321-FFFF-0000-C000-000000000046}']
    procedure IncrementRotationX(Increment: Single); dispid 4106;
    procedure IncrementRotationY(Increment: Single); dispid 4107;
    procedure ResetRotation; dispid 4108;
    procedure SetThreeDFormat(PresetThreeDFormat: KsoPresetThreeDFormat); dispid 4109;
    procedure SetExtrusionDirection(PresetExtrusionDirection: KsoPresetExtrusionDirection); dispid 4110;
    property Depth: Single dispid 4196;
    property _ExtrusionColor: KsoColorFormat readonly dispid 4197;
    property ExtrusionColorType: KsoExtrusionColorType dispid 4198;
    property Perspective: KsoTriState dispid 4199;
    property PresetExtrusionDirection: KsoPresetExtrusionDirection readonly dispid 4200;
    property PresetLightingDirection: KsoPresetLightingDirection dispid 4201;
    property PresetLightingSoftness: KsoPresetLightingSoftness dispid 4202;
    property PresetMaterial: KsoPresetMaterial dispid 4203;
    property PresetThreeDFormat: KsoPresetThreeDFormat readonly dispid 4204;
    property RotationX: Single dispid 4205;
    property RotationY: Single dispid 4206;
    property Visible: KsoTriState dispid 4207;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoScript
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0341-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoScript = interface(IKsoDispObj)
    ['{000C0341-FFFF-0000-C000-000000000046}']
    function Get_Extended: WideString; safecall;
    procedure Set_Extended(const Extended: WideString); safecall;
    function Get_Id: WideString; safecall;
    procedure Set_Id(const Id: WideString); safecall;
    function Get_Language: KsoScriptLanguage; safecall;
    procedure Set_Language(Language: KsoScriptLanguage); safecall;
    function Get_Location: KsoScriptLocation; safecall;
    procedure Delete; safecall;
    function Get_Shape: IDispatch; safecall;
    function Get_ScriptText: WideString; safecall;
    procedure Set_ScriptText(const Script: WideString); safecall;
    property Extended: WideString read Get_Extended write Set_Extended;
    property Id: WideString read Get_Id write Set_Id;
    property Language: KsoScriptLanguage read Get_Language write Set_Language;
    property Location: KsoScriptLocation read Get_Location;
    property Shape: IDispatch read Get_Shape;
    property ScriptText: WideString read Get_ScriptText write Set_ScriptText;
  end;

// *********************************************************************//
// DispIntf:  KsoScriptDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0341-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoScriptDisp = dispinterface
    ['{000C0341-FFFF-0000-C000-000000000046}']
    property Extended: WideString dispid 1610809345;
    property Id: WideString dispid 1610809347;
    property Language: KsoScriptLanguage dispid 1610809349;
    property Location: KsoScriptLocation readonly dispid 1610809351;
    procedure Delete; dispid 1610809352;
    property Shape: IDispatch readonly dispid 1610809353;
    property ScriptText: WideString dispid 0;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoCanvasShapes
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0371-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoCanvasShapes = interface(IKsoDispObj)
    ['{000C0371-FFFF-0000-C000-000000000046}']
    function Get_Count: SYSINT; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function _AddCallout(Type_: KsoCalloutType; Left: Single; Top: Single; Width: Single; 
                         Height: Single): KsoShape; safecall;
    function _AddConnector(Type_: KsoConnectorType; BeginX: Single; BeginY: Single; EndX: Single; 
                           EndY: Single): KsoShape; safecall;
    function _AddCurve(SafeArrayOfPoints: OleVariant): KsoShape; safecall;
    function _AddLabel(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                       Height: Single): KsoShape; safecall;
    function _AddLine(BeginX: Single; BeginY: Single; EndX: Single; EndY: Single): KsoShape; safecall;
    function _AddPicture(const FileName: WideString; LinkToFile: KsoTriState; 
                         SaveWithDocument: KsoTriState; Left: Single; Top: Single; Width: Single; 
                         Height: Single): KsoShape; safecall;
    function _AddPolyline(SafeArrayOfPoints: OleVariant): KsoShape; safecall;
    function _AddShape(Type_: KsoAutoShapeType; Left: Single; Top: Single; Width: Single; 
                       Height: Single): KsoShape; safecall;
    function _AddTextEffect(PresetTextEffect: KsoPresetTextEffect; const Text: WideString; 
                            const FontName: WideString; FontSize: Single; FontBold: KsoTriState; 
                            FontItalic: KsoTriState; Left: Single; Top: Single): KsoShape; safecall;
    function _AddTextbox(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                         Height: Single): KsoShape; safecall;
    function _BuildFreeform(EditingType: KsoEditingType; X1: Single; Y1: Single): KsoFreeformBuilder; safecall;
    function _Range(Index: OleVariant): KsoShapeRange; safecall;
    procedure SelectAll; safecall;
    function Get__Background: KsoShape; safecall;
    function _Item(Index: SYSINT): KsoShape; safecall;
    property Count: SYSINT read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
    property _Background: KsoShape read Get__Background;
  end;

// *********************************************************************//
// DispIntf:  KsoCanvasShapesDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0371-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoCanvasShapesDisp = dispinterface
    ['{000C0371-FFFF-0000-C000-000000000046}']
    property Count: SYSINT readonly dispid 4098;
    property _NewEnum: IUnknown readonly dispid -4;
    function _AddCallout(Type_: KsoCalloutType; Left: Single; Top: Single; Width: Single; 
                         Height: Single): KsoShape; dispid 4106;
    function _AddConnector(Type_: KsoConnectorType; BeginX: Single; BeginY: Single; EndX: Single; 
                           EndY: Single): KsoShape; dispid 4107;
    function _AddCurve(SafeArrayOfPoints: OleVariant): KsoShape; dispid 4108;
    function _AddLabel(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                       Height: Single): KsoShape; dispid 4109;
    function _AddLine(BeginX: Single; BeginY: Single; EndX: Single; EndY: Single): KsoShape; dispid 4110;
    function _AddPicture(const FileName: WideString; LinkToFile: KsoTriState; 
                         SaveWithDocument: KsoTriState; Left: Single; Top: Single; Width: Single; 
                         Height: Single): KsoShape; dispid 4111;
    function _AddPolyline(SafeArrayOfPoints: OleVariant): KsoShape; dispid 4112;
    function _AddShape(Type_: KsoAutoShapeType; Left: Single; Top: Single; Width: Single; 
                       Height: Single): KsoShape; dispid 4113;
    function _AddTextEffect(PresetTextEffect: KsoPresetTextEffect; const Text: WideString; 
                            const FontName: WideString; FontSize: Single; FontBold: KsoTriState; 
                            FontItalic: KsoTriState; Left: Single; Top: Single): KsoShape; dispid 4114;
    function _AddTextbox(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                         Height: Single): KsoShape; dispid 4115;
    function _BuildFreeform(EditingType: KsoEditingType; X1: Single; Y1: Single): KsoFreeformBuilder; dispid 4116;
    function _Range(Index: OleVariant): KsoShapeRange; dispid 4117;
    procedure SelectAll; dispid 4118;
    property _Background: KsoShape readonly dispid 4196;
    function _Item(Index: SYSINT): KsoShape; dispid 4197;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoFreeformBuilder
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0315-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoFreeformBuilder = interface(IKsoDispObj)
    ['{000C0315-FFFF-0000-C000-000000000046}']
    procedure AddNodes(SegmentType: KsoSegmentType; EditingType: KsoEditingType; X1: Single; 
                       Y1: Single; X2: Single; Y2: Single; X3: Single; Y3: Single); safecall;
    function _ConvertToShape: KsoShape; safecall;
  end;

// *********************************************************************//
// DispIntf:  KsoFreeformBuilderDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0315-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoFreeformBuilderDisp = dispinterface
    ['{000C0315-FFFF-0000-C000-000000000046}']
    procedure AddNodes(SegmentType: KsoSegmentType; EditingType: KsoEditingType; X1: Single; 
                       Y1: Single; X2: Single; Y2: Single; X3: Single; Y3: Single); dispid 4106;
    function _ConvertToShape: KsoShape; dispid 4107;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoOLEFormat
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0322-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoOLEFormat = interface(IKsoDispObj)
    ['{000C0322-FFFF-0000-C000-000000000046}']
    function Get__ProgID: WideString; safecall;
    function Get__Object: IDispatch; safecall;
    procedure _Activate; safecall;
    procedure _DoVerb(VerbIndex: OleVariant); safecall;
    function Get__ClassType: WideString; safecall;
    procedure Set__ClassType(const prop: WideString); safecall;
    property _ProgID: WideString read Get__ProgID;
    property _Object: IDispatch read Get__Object;
    property _ClassType: WideString read Get__ClassType write Set__ClassType;
  end;

// *********************************************************************//
// DispIntf:  KsoOLEFormatDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0322-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoOLEFormatDisp = dispinterface
    ['{000C0322-FFFF-0000-C000-000000000046}']
    property _ProgID: WideString readonly dispid 4196;
    property _Object: IDispatch readonly dispid 4197;
    procedure _Activate; dispid 4198;
    procedure _DoVerb(VerbIndex: OleVariant); dispid 4199;
    property _ClassType: WideString dispid 4200;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KsoShapes
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C031E-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoShapes = interface(IKsoDispObj)
    ['{000C031E-FFFF-0000-C000-000000000046}']
    function Get_Count: SYSINT; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function _AddCallout(Type_: KsoCalloutType; Left: Single; Top: Single; Width: Single; 
                         Height: Single): KsoShape; safecall;
    function _AddConnector(Type_: KsoConnectorType; BeginX: Single; BeginY: Single; EndX: Single; 
                           EndY: Single): KsoShape; safecall;
    function _AddCurve(SafeArrayOfPoints: OleVariant): KsoShape; safecall;
    function _AddLabel(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                       Height: Single): KsoShape; safecall;
    function _AddLine(BeginX: Single; BeginY: Single; EndX: Single; EndY: Single): KsoShape; safecall;
    function _AddPicture(const FileName: WideString; LinkToFile: KsoTriState; 
                         SaveWithDocument: KsoTriState; Left: Single; Top: Single; Width: Single; 
                         Height: Single): KsoShape; safecall;
    function _AddPolyline(SafeArrayOfPoints: OleVariant): KsoShape; safecall;
    function _AddShape(Type_: KsoAutoShapeType; Left: Single; Top: Single; Width: Single; 
                       Height: Single): KsoShape; safecall;
    function _AddTextEffect(PresetTextEffect: KsoPresetTextEffect; const Text: WideString; 
                            const FontName: WideString; FontSize: Single; FontBold: KsoTriState; 
                            FontItalic: KsoTriState; Left: Single; Top: Single): KsoShape; safecall;
    function _AddTextbox(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                         Height: Single): KsoShape; safecall;
    function _BuildFreeform(EditingType: KsoEditingType; X1: Single; Y1: Single): KsoFreeformBuilder; safecall;
    function _Range(Index: OleVariant): KsoShapeRange; safecall;
    procedure SelectAll; safecall;
    function _AddDiagram(Type_: KsoDiagramType; Left: Single; Top: Single; Width: Single; 
                         Height: Single): KsoShape; safecall;
    function _AddCanvas(Left: Single; Top: Single; Width: Single; Height: Single): KsoShape; safecall;
    function Get__Background: KsoShape; safecall;
    function Get__Default: KsoShape; safecall;
    function _AddOLEObject(Left: Single; Top: Single; Width: Single; Height: Single; 
                           const ClassName: WideString; const FileName: WideString; 
                           DisplayAsIcon: KsoTriState; const IconFileName: WideString; 
                           IconIndex: SYSINT; const IconLabel: WideString; Link: KsoTriState): KsoShape; safecall;
    function _Item(Index: SYSINT): KsoShape; safecall;
    property Count: SYSINT read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
    property _Background: KsoShape read Get__Background;
    property _Default: KsoShape read Get__Default;
  end;

// *********************************************************************//
// DispIntf:  KsoShapesDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C031E-FFFF-0000-C000-000000000046}
// *********************************************************************//
  KsoShapesDisp = dispinterface
    ['{000C031E-FFFF-0000-C000-000000000046}']
    property Count: SYSINT readonly dispid 4098;
    property _NewEnum: IUnknown readonly dispid -4;
    function _AddCallout(Type_: KsoCalloutType; Left: Single; Top: Single; Width: Single; 
                         Height: Single): KsoShape; dispid 4106;
    function _AddConnector(Type_: KsoConnectorType; BeginX: Single; BeginY: Single; EndX: Single; 
                           EndY: Single): KsoShape; dispid 4107;
    function _AddCurve(SafeArrayOfPoints: OleVariant): KsoShape; dispid 4108;
    function _AddLabel(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                       Height: Single): KsoShape; dispid 4109;
    function _AddLine(BeginX: Single; BeginY: Single; EndX: Single; EndY: Single): KsoShape; dispid 4110;
    function _AddPicture(const FileName: WideString; LinkToFile: KsoTriState; 
                         SaveWithDocument: KsoTriState; Left: Single; Top: Single; Width: Single; 
                         Height: Single): KsoShape; dispid 4111;
    function _AddPolyline(SafeArrayOfPoints: OleVariant): KsoShape; dispid 4112;
    function _AddShape(Type_: KsoAutoShapeType; Left: Single; Top: Single; Width: Single; 
                       Height: Single): KsoShape; dispid 4113;
    function _AddTextEffect(PresetTextEffect: KsoPresetTextEffect; const Text: WideString; 
                            const FontName: WideString; FontSize: Single; FontBold: KsoTriState; 
                            FontItalic: KsoTriState; Left: Single; Top: Single): KsoShape; dispid 4114;
    function _AddTextbox(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                         Height: Single): KsoShape; dispid 4115;
    function _BuildFreeform(EditingType: KsoEditingType; X1: Single; Y1: Single): KsoFreeformBuilder; dispid 4116;
    function _Range(Index: OleVariant): KsoShapeRange; dispid 4117;
    procedure SelectAll; dispid 4118;
    function _AddDiagram(Type_: KsoDiagramType; Left: Single; Top: Single; Width: Single; 
                         Height: Single): KsoShape; dispid 4119;
    function _AddCanvas(Left: Single; Top: Single; Width: Single; Height: Single): KsoShape; dispid 4121;
    property _Background: KsoShape readonly dispid 4196;
    property _Default: KsoShape readonly dispid 4197;
    function _AddOLEObject(Left: Single; Top: Single; Width: Single; Height: Single; 
                           const ClassName: WideString; const FileName: WideString; 
                           DisplayAsIcon: KsoTriState; const IconFileName: WideString; 
                           IconIndex: SYSINT; const IconLabel: WideString; Link: KsoTriState): KsoShape; dispid 4198;
    function _Item(Index: SYSINT): KsoShape; dispid 567;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: DocumentProperty
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {CA2A9BB6-76E4-4509-835A-EAC029FBC7EF}
// *********************************************************************//
  DocumentProperty = interface(IKsoDispObj)
    ['{CA2A9BB6-76E4-4509-835A-EAC029FBC7EF}']
    procedure Delete; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const pbstrRetVal: WideString); safecall;
    function Get_Value: OleVariant; safecall;
    procedure Set_Value(pvargRetVal: OleVariant); safecall;
    function Get_type_: KsoDocProperties; safecall;
    procedure Set_type_(ptypeRetVal: KsoDocProperties); safecall;
    function Get_LinkToContent: WordBool; safecall;
    procedure Set_LinkToContent(pfLinkRetVal: WordBool); safecall;
    function Get_LinkSource: WideString; safecall;
    procedure Set_LinkSource(const pbstrSourceRetVal: WideString); safecall;
    property Name: WideString read Get_Name write Set_Name;
    property Value: OleVariant read Get_Value write Set_Value;
    property type_: KsoDocProperties read Get_type_ write Set_type_;
    property LinkToContent: WordBool read Get_LinkToContent write Set_LinkToContent;
    property LinkSource: WideString read Get_LinkSource write Set_LinkSource;
  end;

// *********************************************************************//
// DispIntf:  DocumentPropertyDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {CA2A9BB6-76E4-4509-835A-EAC029FBC7EF}
// *********************************************************************//
  DocumentPropertyDisp = dispinterface
    ['{CA2A9BB6-76E4-4509-835A-EAC029FBC7EF}']
    procedure Delete; dispid 1;
    property Name: WideString dispid 2;
    property Value: OleVariant dispid 0;
    property type_: KsoDocProperties dispid 4;
    property LinkToContent: WordBool dispid 5;
    property LinkSource: WideString dispid 6;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: DocumentProperties
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8CC27FD1-177E-4527-B03B-30A07962D256}
// *********************************************************************//
  DocumentProperties = interface(IKsoDispObj)
    ['{8CC27FD1-177E-4527-B03B-30A07962D256}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Item(Index: OleVariant): DocumentProperty; safecall;
    function Add(const Name: WideString; LinkToContent: WordBool; Type_: OleVariant; 
                 Value: OleVariant; LinkSource: OleVariant): DocumentProperty; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Item[Index: OleVariant]: DocumentProperty read Get_Item; default;
  end;

// *********************************************************************//
// DispIntf:  DocumentPropertiesDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {8CC27FD1-177E-4527-B03B-30A07962D256}
// *********************************************************************//
  DocumentPropertiesDisp = dispinterface
    ['{8CC27FD1-177E-4527-B03B-30A07962D256}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 11;
    property Item[Index: OleVariant]: DocumentProperty readonly dispid 0; default;
    function Add(const Name: WideString; LinkToContent: WordBool; Type_: OleVariant; 
                 Value: OleVariant; LinkSource: OleVariant): DocumentProperty; dispid 2;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: COMAddIn
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002033A-FFFE-0000-C000-000000111146}
// *********************************************************************//
  COMAddIn = interface(IKsoDispObj)
    ['{0002033A-FFFE-0000-C000-000000111146}']
    function Get_Connect: WordBool; safecall;
    procedure Set_Connect(RetValue: WordBool); safecall;
    function Get_LoadBehavior: Integer; safecall;
    function Get_Location: WideString; safecall;
    function Get_ProgId: WideString; safecall;
    function Get_Guid: WideString; safecall;
    function Get_Object_: IDispatch; safecall;
    procedure Set_Object_(const RetValue: IDispatch); safecall;
    function Get_Description: WideString; safecall;
    procedure Set_Description(const RetValue: WideString); safecall;
    procedure Delete; safecall;
    property Connect: WordBool read Get_Connect write Set_Connect;
    property LoadBehavior: Integer read Get_LoadBehavior;
    property Location: WideString read Get_Location;
    property ProgId: WideString read Get_ProgId;
    property Guid: WideString read Get_Guid;
    property Object_: IDispatch read Get_Object_ write Set_Object_;
    property Description: WideString read Get_Description write Set_Description;
  end;

// *********************************************************************//
// DispIntf:  COMAddInDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002033A-FFFE-0000-C000-000000111146}
// *********************************************************************//
  COMAddInDisp = dispinterface
    ['{0002033A-FFFE-0000-C000-000000111146}']
    property Connect: WordBool dispid 161;
    property LoadBehavior: Integer readonly dispid 23;
    property Location: WideString readonly dispid 24;
    property ProgId: WideString readonly dispid 18;
    property Guid: WideString readonly dispid 19;
    property Object_: IDispatch dispid 22;
    property Description: WideString dispid 32;
    procedure Delete; dispid 34;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: COMAddIns
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020339-FFFE-0000-C000-000000111146}
// *********************************************************************//
  COMAddIns = interface(IKsoDispObj)
    ['{00020339-FFFE-0000-C000-000000111146}']
    function Item(var Index: OleVariant): COMAddIn; safecall;
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IUnknown; safecall;
    procedure Update; safecall;
    procedure SetAppModal(varfModal: WordBool); safecall;
    procedure Add(const FilePath: WideString); safecall;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  COMAddInsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020339-FFFE-0000-C000-000000111146}
// *********************************************************************//
  COMAddInsDisp = dispinterface
    ['{00020339-FFFE-0000-C000-000000111146}']
    function Item(var Index: OleVariant): COMAddIn; dispid 0;
    property Count: Integer readonly dispid 1;
    property _NewEnum: IUnknown readonly dispid -4;
    procedure Update; dispid 2;
    procedure SetAppModal(varfModal: WordBool); dispid 4;
    procedure Add(const FilePath: WideString); dispid 16;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KeyBinding
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020998-FFFE-0000-C000-000000111146}
// *********************************************************************//
  KeyBinding = interface(IDispatch)
    ['{00020998-FFFE-0000-C000-000000111146}']
    function Get_Command: WideString; safecall;
    function Get_KeyString: WideString; safecall;
    function Get_Protected_: WordBool; safecall;
    function Get_KeyCategory: KsoKeyCategory; safecall;
    function Get_KeyCode: Integer; safecall;
    function Get_KeyCode2: Integer; safecall;
    function Get_CommandParameter: WideString; safecall;
    function Get_Context: IDispatch; safecall;
    procedure Clear; safecall;
    procedure Disable; safecall;
    procedure Execute; safecall;
    procedure Rebind(KeyCategory: KsoKeyCategory; const Command: WideString; 
                     var CommandParameter: OleVariant); safecall;
    property Command: WideString read Get_Command;
    property KeyString: WideString read Get_KeyString;
    property Protected_: WordBool read Get_Protected_;
    property KeyCategory: KsoKeyCategory read Get_KeyCategory;
    property KeyCode: Integer read Get_KeyCode;
    property KeyCode2: Integer read Get_KeyCode2;
    property CommandParameter: WideString read Get_CommandParameter;
    property Context: IDispatch read Get_Context;
  end;

// *********************************************************************//
// DispIntf:  KeyBindingDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020998-FFFE-0000-C000-000000111146}
// *********************************************************************//
  KeyBindingDisp = dispinterface
    ['{00020998-FFFE-0000-C000-000000111146}']
    property Command: WideString readonly dispid 1;
    property KeyString: WideString readonly dispid 2;
    property Protected_: WordBool readonly dispid 3;
    property KeyCategory: KsoKeyCategory readonly dispid 4;
    property KeyCode: Integer readonly dispid 6;
    property KeyCode2: Integer readonly dispid 7;
    property CommandParameter: WideString readonly dispid 8;
    property Context: IDispatch readonly dispid 10;
    procedure Clear; dispid 101;
    procedure Disable; dispid 102;
    procedure Execute; dispid 103;
    procedure Rebind(KeyCategory: KsoKeyCategory; const Command: WideString; 
                     var CommandParameter: OleVariant); dispid 104;
  end;

// *********************************************************************//
// Interface: KeyBindings
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020996-FFFE-0000-C000-000000111146}
// *********************************************************************//
  KeyBindings = interface(IKsoDispObj)
    ['{00020996-FFFE-0000-C000-000000111146}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Context: IDispatch; safecall;
    function Item(Index: Integer): KeyBinding; safecall;
    function Add(KeyCategory: KsoKeyCategory; const Command: WideString; KeyCode: Integer; 
                 var KeyCode2: OleVariant; var CommandParameter: OleVariant): KeyBinding; safecall;
    procedure ClearAll; safecall;
    function Key(KeyCode: Integer; var KeyCode2: OleVariant): KeyBinding; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Context: IDispatch read Get_Context;
  end;

// *********************************************************************//
// DispIntf:  KeyBindingsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020996-FFFE-0000-C000-000000111146}
// *********************************************************************//
  KeyBindingsDisp = dispinterface
    ['{00020996-FFFE-0000-C000-000000111146}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Context: IDispatch readonly dispid 10;
    function Item(Index: Integer): KeyBinding; dispid 0;
    function Add(KeyCategory: KsoKeyCategory; const Command: WideString; KeyCode: Integer; 
                 var KeyCode2: OleVariant; var CommandParameter: OleVariant): KeyBinding; dispid 101;
    procedure ClearAll; dispid 102;
    function Key(KeyCode: Integer; var KeyCode2: OleVariant): KeyBinding; dispid 110;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: KeysBoundTo
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020997-FFFE-0000-C000-000000111146}
// *********************************************************************//
  KeysBoundTo = interface(IDispatch)
    ['{00020997-FFFE-0000-C000-000000111146}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_KeyCategory: KsoKeyCategory; safecall;
    function Get_Command: WideString; safecall;
    function Get_CommandParameter: WideString; safecall;
    function Get_Context: IDispatch; safecall;
    function Item(Index: Integer): KeyBinding; safecall;
    function Key(KeyCode: Integer; var KeyCode2: OleVariant): KeyBinding; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property KeyCategory: KsoKeyCategory read Get_KeyCategory;
    property Command: WideString read Get_Command;
    property CommandParameter: WideString read Get_CommandParameter;
    property Context: IDispatch read Get_Context;
  end;

// *********************************************************************//
// DispIntf:  KeysBoundToDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020997-FFFE-0000-C000-000000111146}
// *********************************************************************//
  KeysBoundToDisp = dispinterface
    ['{00020997-FFFE-0000-C000-000000111146}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property KeyCategory: KsoKeyCategory readonly dispid 3;
    property Command: WideString readonly dispid 4;
    property CommandParameter: WideString readonly dispid 5;
    property Context: IDispatch readonly dispid 10;
    function Item(Index: Integer): KeyBinding; dispid 0;
    function Key(KeyCode: Integer; var KeyCode2: OleVariant): KeyBinding; dispid 1;
  end;

// *********************************************************************//
// Interface: AdvApiRoot
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00023EF4-FFFE-0000-C000-000000111146}
// *********************************************************************//
  AdvApiRoot = interface(IKsoDispObj)
    ['{00023EF4-FFFE-0000-C000-000000111146}']
    function Get_AdvApi(Type_: AdvApiType; SessionID: Integer; const Doc: IDispatch; 
                        reserved: OleVariant): IDispatch; safecall;
    property AdvApi[Type_: AdvApiType; SessionID: Integer; const Doc: IDispatch; 
                    reserved: OleVariant]: IDispatch read Get_AdvApi;
  end;

// *********************************************************************//
// DispIntf:  AdvApiRootDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00023EF4-FFFE-0000-C000-000000111146}
// *********************************************************************//
  AdvApiRootDisp = dispinterface
    ['{00023EF4-FFFE-0000-C000-000000111146}']
    property AdvApi[Type_: AdvApiType; SessionID: Integer; const Doc: IDispatch; 
                    reserved: OleVariant]: IDispatch readonly dispid 1;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: AdvAddins
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002F52F-FFFE-0000-C000-000000111146}
// *********************************************************************//
  AdvAddins = interface(IKsoDispObj)
    ['{0002F52F-FFFE-0000-C000-000000111146}']
    function Get_Count: Integer; safecall;
    function Get_Item(Index: Integer): AdvAddin; safecall;
    procedure Add(const prop: AdvAddin); safecall;
    property Count: Integer read Get_Count;
    property Item[Index: Integer]: AdvAddin read Get_Item;
  end;

// *********************************************************************//
// DispIntf:  AdvAddinsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002F52F-FFFE-0000-C000-000000111146}
// *********************************************************************//
  AdvAddinsDisp = dispinterface
    ['{0002F52F-FFFE-0000-C000-000000111146}']
    property Count: Integer readonly dispid 1;
    property Item[Index: Integer]: AdvAddin readonly dispid 2;
    procedure Add(const prop: AdvAddin); dispid 3;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: AdvAddin
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002B7E6-FFFE-0000-C000-000000111146}
// *********************************************************************//
  AdvAddin = interface(IKsoDispObj)
    ['{0002B7E6-FFFE-0000-C000-000000111146}']
    function Get_SysIdentifer: WideString; safecall;
    function Get_ShareMode: HResult; safecall;
    function VerifySessionID(SessionID: Integer): WordBool; safecall;
    function Get_CertType: AdvCertType; safecall;
    property SysIdentifer: WideString read Get_SysIdentifer;
    property ShareMode: HResult read Get_ShareMode;
    property CertType: AdvCertType read Get_CertType;
  end;

// *********************************************************************//
// DispIntf:  AdvAddinDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002B7E6-FFFE-0000-C000-000000111146}
// *********************************************************************//
  AdvAddinDisp = dispinterface
    ['{0002B7E6-FFFE-0000-C000-000000111146}']
    property SysIdentifer: WideString readonly dispid 1;
    property ShareMode: HResult readonly dispid 2;
    function VerifySessionID(SessionID: Integer): WordBool; dispid 3;
    property CertType: AdvCertType readonly dispid 4;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: AdvRight
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002EE54-FFFE-0000-C000-000000111146}
// *********************************************************************//
  AdvRight = interface(IKsoDispObj)
    ['{0002EE54-FFFE-0000-C000-000000111146}']
    procedure OnRightChange; safecall;
  end;

// *********************************************************************//
// DispIntf:  AdvRightDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002EE54-FFFE-0000-C000-000000111146}
// *********************************************************************//
  AdvRightDisp = dispinterface
    ['{0002EE54-FFFE-0000-C000-000000111146}']
    procedure OnRightChange; dispid 1;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: AdvSeal
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000299FC-FFFE-0000-C000-000000111146}
// *********************************************************************//
  AdvSeal = interface(IKsoDispObj)
    ['{000299FC-FFFE-0000-C000-000000111146}']
    function Get_Seals(const Doc: IDispatch): Seals; safecall;
    procedure Set_PagePrintGap(PageGap: Integer); safecall;
    function Get_PagePrintGap: Integer; safecall;
    property Seals[const Doc: IDispatch]: Seals read Get_Seals;
    property PagePrintGap: Integer read Get_PagePrintGap write Set_PagePrintGap;
  end;

// *********************************************************************//
// DispIntf:  AdvSealDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000299FC-FFFE-0000-C000-000000111146}
// *********************************************************************//
  AdvSealDisp = dispinterface
    ['{000299FC-FFFE-0000-C000-000000111146}']
    property Seals[const Doc: IDispatch]: Seals readonly dispid 1;
    property PagePrintGap: Integer dispid 2;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: Seals
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002BE8B-FFFE-0000-C000-000000111146}
// *********************************************************************//
  Seals = interface(IKsoDispObj)
    ['{0002BE8B-FFFE-0000-C000-000000111146}']
    function Get_Count: Integer; safecall;
    function Get_Item(Index: OleVariant): Seal; safecall;
    procedure Delete(Index: OleVariant); safecall;
    procedure InsertSeal(const Picture: IDispatch; InsertDirectly: WordBool; 
                         const sealID: WideString; SealWidth: Integer; SealHeight: Integer; 
                         SealPage: Integer; XPos: Integer; YPos: Integer); safecall;
    procedure LoadSealByPosInfo(const Picture: IDispatch; const sealID: WideString; 
                                SealWidth: Integer; SealHeight: Integer; var SealPosInfo: Byte; 
                                NumberofData: Integer); safecall;
    procedure InsertSealEx(const Picture: IDispatch; InsertDirectly: WordBool; 
                           const sealID: WideString; SealWidth: Integer; SealHeight: Integer; 
                           SealPage: Integer; XPos: Integer; YPos: Integer; Rvalue: Integer; 
                           Gvalue: Integer; Bvalue: Integer); safecall;
    procedure LoadSealByPosInfoEx(const Picture: IDispatch; const sealID: WideString; 
                                  SealWidth: Integer; SealHeight: Integer; var SealPosInfo: Byte; 
                                  NumberofData: Integer; Rvalue: Integer; Gvalue: Integer; 
                                  Bvalue: Integer); safecall;
    property Count: Integer read Get_Count;
    property Item[Index: OleVariant]: Seal read Get_Item;
  end;

// *********************************************************************//
// DispIntf:  SealsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002BE8B-FFFE-0000-C000-000000111146}
// *********************************************************************//
  SealsDisp = dispinterface
    ['{0002BE8B-FFFE-0000-C000-000000111146}']
    property Count: Integer readonly dispid 1;
    property Item[Index: OleVariant]: Seal readonly dispid 2;
    procedure Delete(Index: OleVariant); dispid 3;
    procedure InsertSeal(const Picture: IDispatch; InsertDirectly: WordBool; 
                         const sealID: WideString; SealWidth: Integer; SealHeight: Integer; 
                         SealPage: Integer; XPos: Integer; YPos: Integer); dispid 4;
    procedure LoadSealByPosInfo(const Picture: IDispatch; const sealID: WideString; 
                                SealWidth: Integer; SealHeight: Integer; var SealPosInfo: Byte; 
                                NumberofData: Integer); dispid 5;
    procedure InsertSealEx(const Picture: IDispatch; InsertDirectly: WordBool; 
                           const sealID: WideString; SealWidth: Integer; SealHeight: Integer; 
                           SealPage: Integer; XPos: Integer; YPos: Integer; Rvalue: Integer; 
                           Gvalue: Integer; Bvalue: Integer); dispid 6;
    procedure LoadSealByPosInfoEx(const Picture: IDispatch; const sealID: WideString; 
                                  SealWidth: Integer; SealHeight: Integer; var SealPosInfo: Byte; 
                                  NumberofData: Integer; Rvalue: Integer; Gvalue: Integer; 
                                  Bvalue: Integer); dispid 7;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: Seal
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002F78B-FFFE-0000-C000-000000111146}
// *********************************************************************//
  Seal = interface(IKsoDispObj)
    ['{0002F78B-FFFE-0000-C000-000000111146}']
    function Get_Id: WideString; safecall;
    function Get_Width: Integer; safecall;
    function Get_Height: Integer; safecall;
    function Get_Picture: IDispatch; safecall;
    function Get_PoseInfo(var SealPosInfo: Byte): HResult; safecall;
    function Get_Color(out Rvalue: Integer; out Gvalue: Integer): HResult; safecall;
    procedure SetColor(Rvalue: Integer; Gvalue: Integer; Bvalue: Integer); safecall;
    property Id: WideString read Get_Id;
    property Width: Integer read Get_Width;
    property Height: Integer read Get_Height;
    property Picture: IDispatch read Get_Picture;
    property PoseInfo[var SealPosInfo: Byte]: HResult read Get_PoseInfo;
    property Color[out Rvalue: Integer; out Gvalue: Integer]: HResult read Get_Color;
  end;

// *********************************************************************//
// DispIntf:  SealDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002F78B-FFFE-0000-C000-000000111146}
// *********************************************************************//
  SealDisp = dispinterface
    ['{0002F78B-FFFE-0000-C000-000000111146}']
    property Id: WideString readonly dispid 1;
    property Width: Integer readonly dispid 2;
    property Height: Integer readonly dispid 3;
    property Picture: IDispatch readonly dispid 4;
    property PoseInfo[var SealPosInfo: Byte]: HResult readonly dispid 5;
    property Color[out Rvalue: Integer; out Gvalue: Integer]: HResult readonly dispid 6;
    procedure SetColor(Rvalue: Integer; Gvalue: Integer; Bvalue: Integer); dispid 7;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: IFilterMedium
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002D7C3-FFFE-0000-C000-000000111146}
// *********************************************************************//
  IFilterMedium = interface(IKsoDispObj)
    ['{0002D7C3-FFFE-0000-C000-000000111146}']
    function Get_Name: WideString; safecall;
    function Get_type_: MEDIUMTYPE; safecall;
    procedure Set_Mode(prop: Integer); safecall;
    function Get_Mode: Integer; safecall;
    procedure Read(out prop: OleVariant); safecall;
    procedure StartWrite(out prop: OleVariant); safecall;
    procedure EndWrite; safecall;
    procedure Lock(prop: WordBool); safecall;
    procedure CopyTo(const prop: IFilterMedium); safecall;
    property Name: WideString read Get_Name;
    property type_: MEDIUMTYPE read Get_type_;
    property Mode: Integer read Get_Mode write Set_Mode;
  end;

// *********************************************************************//
// DispIntf:  IFilterMediumDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002D7C3-FFFE-0000-C000-000000111146}
// *********************************************************************//
  IFilterMediumDisp = dispinterface
    ['{0002D7C3-FFFE-0000-C000-000000111146}']
    property Name: WideString readonly dispid 1;
    property type_: MEDIUMTYPE readonly dispid 2;
    property Mode: Integer dispid 3;
    procedure Read(out prop: OleVariant); dispid 4;
    procedure StartWrite(out prop: OleVariant); dispid 5;
    procedure EndWrite; dispid 6;
    procedure Lock(prop: WordBool); dispid 7;
    procedure CopyTo(const prop: IFilterMedium); dispid 8;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: AdvApplication
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002AFDA-FFFE-0000-C000-000000111146}
// *********************************************************************//
  AdvApplication = interface(IKsoDispObj)
    ['{0002AFDA-FFFE-0000-C000-000000111146}']
    function Open(const medium: IFilterMedium; ConfirmConversions: WordBool; ReadOnly: WordBool; 
                  AddToRecentFiles: WordBool; const PasswordDocument: WideString; 
                  const PasswordTemplate: WideString; Revert: WordBool; 
                  const WritePasswordDocument: WideString; const WritePasswordTemplate: WideString; 
                  Format: Integer; Encoding: Integer; Visible: WordBool; OpenAndRepair: WordBool; 
                  DocumentDirection: Integer; NoEncodingDialog: WordBool; UpdateLinks: OleVariant; 
                  WriteResPassword: OleVariant; IgnoreReadOnlyRecommended: OleVariant; 
                  Origin: OleVariant; Delimiter: OleVariant; Editable: OleVariant; 
                  Notify: OleVariant; Converter: OleVariant; AddToMru: OleVariant; 
                  Untitled: WordBool; WithWindow: WordBool): IDispatch; safecall;
    function CreateFilterMediaOnStream(Stream: OleVariant; const FileName: WideString): IFilterMedium; safecall;
    function CreateFilterMediaOnStorage(Storage: OleVariant; const FileName: WideString): IFilterMedium; safecall;
    function CreateFilterMediaOnFile(const File_: WideString): IFilterMedium; safecall;
  end;

// *********************************************************************//
// DispIntf:  AdvApplicationDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002AFDA-FFFE-0000-C000-000000111146}
// *********************************************************************//
  AdvApplicationDisp = dispinterface
    ['{0002AFDA-FFFE-0000-C000-000000111146}']
    function Open(const medium: IFilterMedium; ConfirmConversions: WordBool; ReadOnly: WordBool; 
                  AddToRecentFiles: WordBool; const PasswordDocument: WideString; 
                  const PasswordTemplate: WideString; Revert: WordBool; 
                  const WritePasswordDocument: WideString; const WritePasswordTemplate: WideString; 
                  Format: Integer; Encoding: Integer; Visible: WordBool; OpenAndRepair: WordBool; 
                  DocumentDirection: Integer; NoEncodingDialog: WordBool; UpdateLinks: OleVariant; 
                  WriteResPassword: OleVariant; IgnoreReadOnlyRecommended: OleVariant; 
                  Origin: OleVariant; Delimiter: OleVariant; Editable: OleVariant; 
                  Notify: OleVariant; Converter: OleVariant; AddToMru: OleVariant; 
                  Untitled: WordBool; WithWindow: WordBool): IDispatch; dispid 1;
    function CreateFilterMediaOnStream(Stream: OleVariant; const FileName: WideString): IFilterMedium; dispid 3;
    function CreateFilterMediaOnStorage(Storage: OleVariant; const FileName: WideString): IFilterMedium; dispid 4;
    function CreateFilterMediaOnFile(const File_: WideString): IFilterMedium; dispid 5;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// DispIntf:  AdvApplicationEvents
// Flags:     (4096) Dispatchable
// GUID:      {B94C31F7-4947-4CEB-ABDD-9B37CE8422B5}
// *********************************************************************//
  AdvApplicationEvents = dispinterface
    ['{B94C31F7-4947-4CEB-ABDD-9B37CE8422B5}']
    function AcquireFilterMedium(const Doc: IDispatch; out FileFormat: OleVariant): IDispatch; dispid 1;
    procedure BeforeOpen(const Doc: IDispatch; SrcMediumType: MEDIUMTYPE; SrcMedium: OleVariant; 
                         out DesMediumType: MEDIUMTYPE; out DesMedium: OleVariant; 
                         out Cancle: WordBool); dispid 2;
    procedure AfterOpen(const Doc: IDispatch); dispid 3;
    procedure BeforeSave(const Doc: IDispatch; SrcMediumType: MEDIUMTYPE; SrcMedium: OleVariant; 
                         DesMediumType: MEDIUMTYPE; DesMedium: OleVariant; out Modified: WordBool; 
                         out Cancle: WordBool); dispid 4;
    procedure AfterSave(const Doc: IDispatch; MEDIUMTYPE: MEDIUMTYPE; medium: OleVariant); dispid 5;
    procedure AfterClose(const Doc: IDispatch); dispid 6;
    procedure AcquireIdentifer(out Identifier: WideString); dispid 7;
    function NeedEncrypt(const Doc: IDispatch): WordBool; dispid 8;
  end;

// *********************************************************************//
// DispIntf:  AdvRightEvents
// Flags:     (4096) Dispatchable
// GUID:      {A0984A28-9CF6-4826-86A9-E8D592950965}
// *********************************************************************//
  AdvRightEvents = dispinterface
    ['{A0984A28-9CF6-4826-86A9-E8D592950965}']
    procedure DocCopyEnable(out Enable: WordBool); dispid 1;
    procedure DocEditEnable(out Enable: WordBool); dispid 2;
    procedure DocSaveAsEnable(out Enable: WordBool); dispid 3;
    procedure DocHideRevision(out Hidden: WordBool); dispid 4;
    procedure DocGetContentWithVBAEnable(out Enable: WordBool); dispid 5;
    procedure DocPrintEnable(out Copies: Integer); dispid 6;
    procedure DocBeforePrintPage(CopyIndex: Integer; PageIndex: Integer; out Cancel: WordBool); dispid 7;
    procedure DocAfterPrintPage(CopyIndex: Integer; PageIndex: Integer); dispid 8;
  end;

// *********************************************************************//
// DispIntf:  AdvSealEvents
// Flags:     (4096) Dispatchable
// GUID:      {F84DA428-2F66-4B38-B7D9-84C4E6EB4B28}
// *********************************************************************//
  AdvSealEvents = dispinterface
    ['{F84DA428-2F66-4B38-B7D9-84C4E6EB4B28}']
    procedure BeforeInsertSeal(const Doc: IDispatch; const sealID: WideString; 
                               AddToBlankPage: WordBool; out Cancle: WordBool); dispid 1;
    procedure AfterInsertSeal(const Doc: IDispatch; const SealInfo: Seal); dispid 2;
    procedure CancleInsertSeal(const Doc: IDispatch; const sealID: WideString); dispid 3;
  end;

// *********************************************************************//
// Interface: _TaskPane
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002B802-FFFE-0000-C000-000000111146}
// *********************************************************************//
  _TaskPane = interface(IKsoDispObj)
    ['{0002B802-FFFE-0000-C000-000000111146}']
    function Get_Name: WideString; safecall;
    function Get_Caption: WideString; safecall;
    procedure Set_Caption(const Caption: WideString); safecall;
    function Get_Visible: WordBool; safecall;
    function Get_Id: SYSINT; safecall;
    function Get_Object_: IDispatch; safecall;
    function Get_ObjectContext: WideString; safecall;
    procedure Set_ObjectContext(const ObjectContext: WideString); safecall;
    procedure Active; safecall;
    function Get_Enable: WordBool; safecall;
    procedure Set_Enable(Enable: WordBool); safecall;
    procedure Delete; safecall;
    procedure Set_Visible(Visible: WordBool); safecall;
    property Name: WideString read Get_Name;
    property Caption: WideString read Get_Caption write Set_Caption;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property Id: SYSINT read Get_Id;
    property Object_: IDispatch read Get_Object_;
    property ObjectContext: WideString read Get_ObjectContext write Set_ObjectContext;
    property Enable: WordBool read Get_Enable write Set_Enable;
  end;

// *********************************************************************//
// DispIntf:  _TaskPaneDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002B802-FFFE-0000-C000-000000111146}
// *********************************************************************//
  _TaskPaneDisp = dispinterface
    ['{0002B802-FFFE-0000-C000-000000111146}']
    property Name: WideString readonly dispid 769;
    property Caption: WideString dispid 770;
    property Visible: WordBool dispid 771;
    property Id: SYSINT readonly dispid 772;
    property Object_: IDispatch readonly dispid 773;
    property ObjectContext: WideString dispid 774;
    procedure Active; dispid 775;
    property Enable: WordBool dispid 776;
    procedure Delete; dispid 777;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// Interface: TaskPanes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002B801-FFFE-0000-C000-000000111146}
// *********************************************************************//
  TaskPanes = interface(IKsoDispObj)
    ['{0002B801-FFFE-0000-C000-000000111146}']
    function Get_Item(Index: OleVariant): _TaskPane; safecall;
    function Get_Count: SYSINT; safecall;
    function AddByProgID(const Name: WideString; const ProgId: WideString): _TaskPane; safecall;
    function AddByURL(const Name: WideString; const URL: WideString): _TaskPane; safecall;
    property Item[Index: OleVariant]: _TaskPane read Get_Item; default;
    property Count: SYSINT read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  TaskPanesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002B801-FFFE-0000-C000-000000111146}
// *********************************************************************//
  TaskPanesDisp = dispinterface
    ['{0002B801-FFFE-0000-C000-000000111146}']
    property Item[Index: OleVariant]: _TaskPane readonly dispid 0; default;
    property Count: SYSINT readonly dispid 779;
    function AddByProgID(const Name: WideString; const ProgId: WideString): _TaskPane; dispid 780;
    function AddByURL(const Name: WideString; const URL: WideString): _TaskPane; dispid 781;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// DispIntf:  _ITaskPaneEvent
// Flags:     (4096) Dispatchable
// GUID:      {1D4270E9-E494-47B5-B7A0-2280BCD2A352}
// *********************************************************************//
  _ITaskPaneEvent = dispinterface
    ['{1D4270E9-E494-47B5-B7A0-2280BCD2A352}']
    procedure OnBeforeDelete(const pTaskPane: _TaskPane; var Cancel: WordBool); dispid 1;
    procedure OnAfterDelete(nTaskPaneID: SYSINT); dispid 2;
    procedure OnCaptionChanged(const pTaskPane: _TaskPane); dispid 3;
  end;

// *********************************************************************//
// Interface: PluginPlatform
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020F01-FFFE-0000-C000-000000111146}
// *********************************************************************//
  PluginPlatform = interface(IKsoDispObj)
    ['{00020F01-FFFE-0000-C000-000000111146}']
    function Get_Service(const SrvID: WideString): IDispatch; safecall;
    property Service[const SrvID: WideString]: IDispatch read Get_Service;
  end;

// *********************************************************************//
// DispIntf:  PluginPlatformDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020F01-FFFE-0000-C000-000000111146}
// *********************************************************************//
  PluginPlatformDisp = dispinterface
    ['{00020F01-FFFE-0000-C000-000000111146}']
    property Service[const SrvID: WideString]: IDispatch readonly dispid 1;
    property Application: IDispatch readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    procedure zimp_DispObj_Reserved6; dispid 268435206;
  end;

// *********************************************************************//
// The Class CoCommandBarComboBox provides a Create and CreateRemote method to          
// create instances of the default interface _CommandBarComboBox exposed by              
// the CoClass CommandBarComboBox. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoCommandBarComboBox = class
    class function Create: _CommandBarComboBox;
    class function CreateRemote(const MachineName: string): _CommandBarComboBox;
  end;

// *********************************************************************//
// The Class CoCommandBarButton provides a Create and CreateRemote method to          
// create instances of the default interface _CommandBarButton exposed by              
// the CoClass CommandBarButton. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoCommandBarButton = class
    class function Create: _CommandBarButton;
    class function CreateRemote(const MachineName: string): _CommandBarButton;
  end;

// *********************************************************************//
// The Class CoCommandBars provides a Create and CreateRemote method to          
// create instances of the default interface _CommandBars exposed by              
// the CoClass CommandBars. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoCommandBars = class
    class function Create: _CommandBars;
    class function CreateRemote(const MachineName: string): _CommandBars;
  end;

// *********************************************************************//
// The Class CoTaskPane provides a Create and CreateRemote method to          
// create instances of the default interface _TaskPane exposed by              
// the CoClass TaskPane. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoTaskPane = class
    class function Create: _TaskPane;
    class function CreateRemote(const MachineName: string): _TaskPane;
  end;

implementation

uses ComObj;

class function CoCommandBarComboBox.Create: _CommandBarComboBox;
begin
  Result := CreateComObject(CLASS_CommandBarComboBox) as _CommandBarComboBox;
end;

class function CoCommandBarComboBox.CreateRemote(const MachineName: string): _CommandBarComboBox;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_CommandBarComboBox) as _CommandBarComboBox;
end;

class function CoCommandBarButton.Create: _CommandBarButton;
begin
  Result := CreateComObject(CLASS_CommandBarButton) as _CommandBarButton;
end;

class function CoCommandBarButton.CreateRemote(const MachineName: string): _CommandBarButton;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_CommandBarButton) as _CommandBarButton;
end;

class function CoCommandBars.Create: _CommandBars;
begin
  Result := CreateComObject(CLASS_CommandBars) as _CommandBars;
end;

class function CoCommandBars.CreateRemote(const MachineName: string): _CommandBars;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_CommandBars) as _CommandBars;
end;

class function CoTaskPane.Create: _TaskPane;
begin
  Result := CreateComObject(CLASS_TaskPane) as _TaskPane;
end;

class function CoTaskPane.CreateRemote(const MachineName: string): _TaskPane;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_TaskPane) as _TaskPane;
end;

end.
