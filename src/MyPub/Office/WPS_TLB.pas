unit WPS_TLB;

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
// Type Lib: C:\Program Files\Kingsoft\WPS Office Personal\office6\wpscore.dll (1)
// LIBID: {00020905-0000-4B30-A977-D214852036FE}
// LCID: 0
// Helpfile: C:\Program Files\Kingsoft\WPS Office Personal\office6\wpscore.dll\wpsapi.chm
// HelpString: Kingsoft WPS 2.0 Object Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\system32\stdole2.tlb)
//   (2) v1.0 KSO, (C:\Program Files\Kingsoft\WPS Office Personal\office6\kso10.dll)
// Errors:
//   Hint: Symbol 'OLEControl' renamed to 'OLECtrl'
//   Hint: Parameter 'Object' of _Application.OrganizerCopy changed to 'Object_'
//   Hint: Parameter 'Object' of _Application.OrganizerDelete changed to 'Object_'
//   Hint: Parameter 'Object' of _Application.OrganizerRename changed to 'Object_'
//   Hint: Parameter 'To' of _Application.PrintOut changed to 'To_'
//   Hint: Symbol 'System' renamed to 'System_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'End' of _Document.Range changed to 'End_'
//   Hint: Parameter 'To' of _Document.PrintOut changed to 'To_'
//   Hint: Parameter 'Type' of _Document.Protect changed to 'Type_'
//   Hint: Member 'End' of 'Range' changed to 'End_'
//   Hint: Parameter 'Type' of Range.Information changed to 'Type_'
//   Hint: Member 'Case' of 'Range' changed to 'Case_'
//   Hint: Parameter 'End' of Range.SetRange changed to 'End_'
//   Hint: Parameter 'Unit' of Range.Next changed to 'Unit_'
//   Hint: Parameter 'Unit' of Range.Previous changed to 'Unit_'
//   Hint: Parameter 'Unit' of Range.StartOf changed to 'Unit_'
//   Hint: Parameter 'Unit' of Range.EndOf changed to 'Unit_'
//   Hint: Parameter 'Unit' of Range.Move changed to 'Unit_'
//   Hint: Parameter 'Unit' of Range.MoveStart changed to 'Unit_'
//   Hint: Parameter 'Unit' of Range.MoveEnd changed to 'Unit_'
//   Hint: Parameter 'Type' of Range.InsertBreak changed to 'Type_'
//   Hint: Parameter 'Unit' of Range.Delete changed to 'Unit_'
//   Hint: Parameter 'Unit' of Range.Expand changed to 'Unit_'
//   Hint: Member 'GoTo' of 'Range' changed to 'GoTo_'
//   Hint: Parameter 'Label' of Range.InsertCaption changed to 'Label_'
//   Hint: Parameter 'Raise' of Range.PhoneticGuide changed to 'Raise_'
//   Hint: Parameter 'Type' of Shapes.AddCallout changed to 'Type_'
//   Hint: Parameter 'Type' of Shapes.AddConnector changed to 'Type_'
//   Hint: Parameter 'Type' of Shapes.AddShape changed to 'Type_'
//   Hint: Parameter 'Type' of Shapes.AddDiagram changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of Shape.Type changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of ShapeRange.Type changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Member 'Object' of 'OLEFormat' changed to 'Object_'
//   Hint: Parameter 'Type' of CanvasShapes.AddCallout changed to 'Type_'
//   Hint: Parameter 'Type' of CanvasShapes.AddConnector changed to 'Type_'
//   Hint: Parameter 'Label' of CanvasShapes.AddLabel changed to 'Label_'
//   Hint: Parameter 'Type' of CanvasShapes.AddShape changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of Fields.Add changed to 'Type_'
//   Hint: Member 'End' of 'Bookmark' changed to 'End_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of FormFields.Add changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of TextInput.EditType changed to 'Type_'
//   Hint: Parameter 'Type' of Styles.Add changed to 'Type_'
//   Hint: Parameter 'Type' of MailMerge.UseAddressBook changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'To' of Window.PrintOut changed to 'To_'
//   Hint: Member 'End' of 'Selection' changed to 'End_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Parameter 'Type' of Selection.Information changed to 'Type_'
//   Hint: Parameter 'End' of Selection.SetRange changed to 'End_'
//   Hint: Parameter 'Unit' of Selection.Next changed to 'Unit_'
//   Hint: Parameter 'Unit' of Selection.Previous changed to 'Unit_'
//   Hint: Parameter 'Unit' of Selection.StartOf changed to 'Unit_'
//   Hint: Parameter 'Unit' of Selection.EndOf changed to 'Unit_'
//   Hint: Parameter 'Unit' of Selection.Move changed to 'Unit_'
//   Hint: Parameter 'Unit' of Selection.MoveStart changed to 'Unit_'
//   Hint: Parameter 'Unit' of Selection.MoveEnd changed to 'Unit_'
//   Hint: Parameter 'Type' of Selection.InsertBreak changed to 'Type_'
//   Hint: Parameter 'Unit' of Selection.Delete changed to 'Unit_'
//   Hint: Parameter 'Unit' of Selection.Expand changed to 'Unit_'
//   Hint: Member 'GoTo' of 'Selection' changed to 'GoTo_'
//   Hint: Parameter 'Unit' of Selection.MoveLeft changed to 'Unit_'
//   Hint: Parameter 'Unit' of Selection.MoveRight changed to 'Unit_'
//   Hint: Parameter 'Unit' of Selection.MoveUp changed to 'Unit_'
//   Hint: Parameter 'Unit' of Selection.MoveDown changed to 'Unit_'
//   Hint: Parameter 'Unit' of Selection.HomeKey changed to 'Unit_'
//   Hint: Parameter 'Unit' of Selection.EndKey changed to 'Unit_'
//   Hint: Parameter 'Label' of Selection.InsertCaption changed to 'Label_'
//   Hint: Member 'Case' of 'Selection' changed to 'Case_'
//   Hint: Parameter 'Type' of Selection.ListType changed to 'Type_'
//   Hint: Parameter 'Type' of Selection.PasteAndFormat changed to 'Type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Symbol 'ClassName' renamed to '_className'
//   Hint: Symbol 'Type' renamed to 'type_'
//   Hint: Symbol 'Type' renamed to 'type_'
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}
interface

uses Windows, ActiveX, Classes, Graphics, KSO_TLB, OleServer, StdVCL, Variants;
  


// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  WPSMajorVersion = 2;
  WPSMinorVersion = 0;

  LIBID_WPS: TGUID = '{00020905-0000-4B30-A977-D214852036FE}';

  IID__IKsoDispObj: TGUID = '{000C0300-FFFF-0000-C000-000000000046}';
  IID__Application: TGUID = '{00020970-0000-4B30-A977-D214852036FE}';
  IID_Documents: TGUID = '{0002096C-0000-4B30-A977-D214852036FE}';
  IID__Document: TGUID = '{0002096B-0000-4B30-A977-D214852036FE}';
  IID_Range: TGUID = '{0002095E-0000-4B30-A977-D214852036FE}';
  IID__Font: TGUID = '{00020952-0000-4B30-A977-D214852036FE}';
  IID_Borders: TGUID = '{0002093C-0000-4B30-A977-D214852036FE}';
  IID_Border: TGUID = '{0002093B-0000-4B30-A977-D214852036FE}';
  IID_Shading: TGUID = '{0002093A-0000-4B30-A977-D214852036FE}';
  IID_Tables: TGUID = '{0002094D-0000-4B30-A977-D214852036FE}';
  IID_Table: TGUID = '{00020951-0000-4B30-A977-D214852036FE}';
  IID_Columns: TGUID = '{0002094B-0000-4B30-A977-D214852036FE}';
  IID_Column: TGUID = '{0002094F-0000-4B30-A977-D214852036FE}';
  IID_Cells: TGUID = '{0002094A-0000-4B30-A977-D214852036FE}';
  IID_Cell: TGUID = '{0002094E-0000-4B30-A977-D214852036FE}';
  IID_Row: TGUID = '{00020950-0000-4B30-A977-D214852036FE}';
  IID_Rows: TGUID = '{0002094C-0000-4B30-A977-D214852036FE}';
  IID_Footnotes: TGUID = '{00020942-0000-4B30-A977-D214852036FE}';
  IID_Footnote: TGUID = '{0002093F-0000-4B30-A977-D214852036FE}';
  IID_FootnoteOptions: TGUID = '{00025A24-0000-4B30-A977-D214852036FE}';
  IID_Endnotes: TGUID = '{00020941-0000-4B30-A977-D214852036FE}';
  IID_Endnote: TGUID = '{0002093E-0000-4B30-A977-D214852036FE}';
  IID_EndnoteOptions: TGUID = '{00023168-0000-4B30-A977-D214852036FE}';
  IID_Comments: TGUID = '{00020940-0000-4B30-A977-D214852036FE}';
  IID_Comment: TGUID = '{0002093D-0000-4B30-A977-D214852036FE}';
  IID_Sections: TGUID = '{0002095A-0000-4B30-A977-D214852036FE}';
  IID_Section: TGUID = '{00020959-0000-4B30-A977-D214852036FE}';
  IID_PageSetup: TGUID = '{00020971-0000-4B30-A977-D214852036FE}';
  IID_TextColumns: TGUID = '{00020973-0000-4B30-A977-D214852036FE}';
  IID_TextColumn: TGUID = '{00020974-0000-4B30-A977-D214852036FE}';
  IID_HeadersFooters: TGUID = '{00020984-0000-4B30-A977-D214852036FE}';
  IID_HeaderFooter: TGUID = '{00020985-0000-4B30-A977-D214852036FE}';
  IID_PageNumbers: TGUID = '{00020986-0000-4B30-A977-D214852036FE}';
  IID_PageNumber: TGUID = '{00020987-0000-4B30-A977-D214852036FE}';
  IID_Shapes: TGUID = '{0002099F-0000-4B30-A977-D214852036FE}';
  IID_Shape: TGUID = '{000209A0-0000-4B30-A977-D214852036FE}';
  IID_ShapeRange: TGUID = '{000209B5-0000-4B30-A977-D214852036FE}';
  IID_WrapFormat: TGUID = '{000209C3-0000-4B30-A977-D214852036FE}';
  IID_Frame: TGUID = '{0002092A-0000-4B30-A977-D214852036FE}';
  IID_InlineShape: TGUID = '{000209A8-0000-4B30-A977-D214852036FE}';
  IID_Field: TGUID = '{0002092F-0000-4B30-A977-D214852036FE}';
  IID_FillFormat: TGUID = '{000290C8-0000-4B30-A977-D214852036FE}';
  IID_ColorFormat: TGUID = '{000290C6-0000-4B30-A977-D214852036FE}';
  IID_PictureFormat: TGUID = '{000250C8-0000-4B30-A977-D214852036FE}';
  IID_TextEffectFormat: TGUID = '{000290CF-0000-4B30-A977-D214852036FE}';
  IID_OLEFormat: TGUID = '{00020322-0000-4B30-A977-D214852036FE}';
  IID_TextFrame: TGUID = '{000209B2-0000-4B30-A977-D214852036FE}';
  IID_Adjustments: TGUID = '{000209C4-0000-4B30-A977-D214852036FE}';
  IID_CalloutFormat: TGUID = '{000209C5-0000-4B30-A977-D214852036FE}';
  IID_ConnectorFormat: TGUID = '{000290C7-0000-4B30-A977-D214852036FE}';
  IID_GroupShapes: TGUID = '{000290B6-0000-4B30-A977-D214852036FE}';
  IID_LineFormat: TGUID = '{000290CA-0000-4B30-A977-D214852036FE}';
  IID_ShapeNodes: TGUID = '{000290CE-0000-4B30-A977-D214852036FE}';
  IID_ShapeNode: TGUID = '{000290CD-0000-4B30-A977-D214852036FE}';
  IID_ShadowFormat: TGUID = '{000290CC-0000-4B30-A977-D214852036FE}';
  IID_ThreeDFormat: TGUID = '{000290D0-0000-4B30-A977-D214852036FE}';
  IID_Diagram: TGUID = '{000290EC-0000-4B30-A977-D214852036FE}';
  IID_DiagramNodes: TGUID = '{000290EB-0000-4B30-A977-D214852036FE}';
  IID_DiagramNode: TGUID = '{000290E9-0000-4B30-A977-D214852036FE}';
  IID_DiagramNodeChildren: TGUID = '{000290EA-0000-4B30-A977-D214852036FE}';
  IID_CanvasShapes: TGUID = '{00029073-0000-4B30-A977-D214852036FE}';
  IID_FreeformBuilder: TGUID = '{000290C9-0000-4B30-A977-D214852036FE}';
  IID_Paragraphs: TGUID = '{00020958-0000-4B30-A977-D214852036FE}';
  IID_Paragraph: TGUID = '{00020957-0000-4B30-A977-D214852036FE}';
  IID__ParagraphFormat: TGUID = '{00020953-0000-4B30-A977-D214852036FE}';
  IID_TabStops: TGUID = '{00020955-0000-4B30-A977-D214852036FE}';
  IID_TabStop: TGUID = '{00020954-0000-4B30-A977-D214852036FE}';
  IID_DropCap: TGUID = '{00020956-0000-4B30-A977-D214852036FE}';
  IID_Style: TGUID = '{0002092C-0000-4B30-A977-D214852036FE}';
  IID_ListTemplate: TGUID = '{0002098F-0000-4B30-A977-D214852036FE}';
  IID_ListLevels: TGUID = '{0002098E-0000-4B30-A977-D214852036FE}';
  IID_ListLevel: TGUID = '{0002098D-0000-4B30-A977-D214852036FE}';
  IID_Fields: TGUID = '{00020930-0000-4B30-A977-D214852036FE}';
  IID_Frames: TGUID = '{0002092B-0000-4B30-A977-D214852036FE}';
  IID_ListFormat: TGUID = '{000209C0-0000-4B30-A977-D214852036FE}';
  IID_List: TGUID = '{00020992-0000-4B30-A977-D214852036FE}';
  IID_ListParagraphs: TGUID = '{00020991-0000-4B30-A977-D214852036FE}';
  IID_Bookmarks: TGUID = '{00020967-0000-4B30-A977-D214852036FE}';
  IID_Bookmark: TGUID = '{00020968-0000-4B30-A977-D214852036FE}';
  IID_Hyperlinks: TGUID = '{0002099C-0000-4B30-A977-D214852036FE}';
  IID_Hyperlink: TGUID = '{0002099D-0000-4B30-A977-D214852036FE}';
  IID_InlineShapes: TGUID = '{000209A9-0000-4B30-A977-D214852036FE}';
  IID_WordStat: TGUID = '{000219A6-0000-4B30-A977-D214852036FE}';
  IID_Revisions: TGUID = '{00020980-0000-4B30-A977-D214852036FE}';
  IID_Revision: TGUID = '{00020981-0000-4B30-A977-D214852036FE}';
  IID_SpellingSuggestions: TGUID = '{000209AA-0000-4B30-A977-D214852036FE}';
  IID_SpellingSuggestion: TGUID = '{000209AB-0000-4B30-A977-D214852036FE}';
  IID_ProofreadingErrors: TGUID = '{000209BB-0000-4B30-A977-D214852036FE}';
  IID_FormFields: TGUID = '{00020929-0000-4B30-A977-D214852036FE}';
  IID_FormField: TGUID = '{00020928-0000-4B30-A977-D214852036FE}';
  IID_TextInput: TGUID = '{00020927-0000-4B30-A977-D214852036FE}';
  IID_CheckBox: TGUID = '{00020926-0000-4B30-A977-D214852036FE}';
  IID_DropDown: TGUID = '{00020925-0000-4B30-A977-D214852036FE}';
  IID_ListEntries: TGUID = '{00020924-0000-4B30-A977-D214852036FE}';
  IID_ListEntry: TGUID = '{00020923-0000-4B30-A977-D214852036FE}';
  IID_Editors: TGUID = '{0002E08C-0000-4B30-A977-D214852036FE}';
  IID_Editor: TGUID = '{00027D72-0000-4B30-A977-D214852036FE}';
  IID_Styles: TGUID = '{0002092D-0000-4B30-A977-D214852036FE}';
  IID_TablesOfFigures: TGUID = '{00020922-0000-4B30-A977-D214852036FE}';
  IID_TableOfFigures: TGUID = '{00020921-0000-4B30-A977-D214852036FE}';
  IID_MailMerge: TGUID = '{00020920-0000-4B30-A977-D214852036FE}';
  IID_MailMergeDataSource: TGUID = '{0002091D-0000-4B30-A977-D214852036FE}';
  IID_MailMergeFieldNames: TGUID = '{0002091C-0000-4B30-A977-D214852036FE}';
  IID_MailMergeFieldName: TGUID = '{0002091B-0000-4B30-A977-D214852036FE}';
  IID_MailMergeDataFields: TGUID = '{0002091A-0000-4B30-A977-D214852036FE}';
  IID_MailMergeDataField: TGUID = '{00020919-0000-4B30-A977-D214852036FE}';
  IID_MappedDataFields: TGUID = '{00026814-0000-4B30-A977-D214852036FE}';
  IID_MappedDataField: TGUID = '{00021669-0000-4B30-A977-D214852036FE}';
  IID_MailMergeFields: TGUID = '{0002091F-0000-4B30-A977-D214852036FE}';
  IID_MailMergeField: TGUID = '{0002091E-0000-4B30-A977-D214852036FE}';
  IID_TablesOfContents: TGUID = '{00020914-0000-4B30-A977-D214852036FE}';
  IID_TableOfContents: TGUID = '{00020913-0000-4B30-A977-D214852036FE}';
  IID_Windows: TGUID = '{00020961-0000-4B30-A977-D214852036FE}';
  IID_Window: TGUID = '{00020962-0000-4B30-A977-D214852036FE}';
  IID_Pane: TGUID = '{00020960-0000-4B30-A977-D214852036FE}';
  IID_Selection: TGUID = '{00020975-0000-4B30-A977-D214852036FE}';
  IID_Find: TGUID = '{000209B0-0000-4B30-A977-D214852036FE}';
  IID_Replacement: TGUID = '{000209B1-0000-4B30-A977-D214852036FE}';
  IID_Zooms: TGUID = '{000209A7-0000-4B30-A977-D214852036FE}';
  IID_Zoom: TGUID = '{000209A6-0000-4B30-A977-D214852036FE}';
  IID_View: TGUID = '{000209A5-0000-4B30-A977-D214852036FE}';
  IID_Reviewers: TGUID = '{00020C9A-0000-4B30-A977-D214852036FE}';
  IID_Reviewer: TGUID = '{000204AE-0000-4B30-A977-D214852036FE}';
  IID_Pages: TGUID = '{00027402-0000-4B30-A977-D214852036FE}';
  IID_Page: TGUID = '{000240A9-0000-4B30-A977-D214852036FE}';
  IID_Breaks: TGUID = '{00029309-0000-4B30-A977-D214852036FE}';
  IID_Break: TGUID = '{00025BF1-0000-4B30-A977-D214852036FE}';
  IID_Panes: TGUID = '{0002095F-0000-4B30-A977-D214852036FE}';
  IID_ListTemplates: TGUID = '{00020990-0000-4B30-A977-D214852036FE}';
  IID_Lists: TGUID = '{00020993-0000-4B30-A977-D214852036FE}';
  IID_ExtraColors: TGUID = '{00027778-0000-4B30-A977-D214852036FE}';
  IID_Variables: TGUID = '{00020965-0000-4B30-A977-D214852036FE}';
  IID_Variable: TGUID = '{00020966-0000-4B30-A977-D214852036FE}';
  IID__WebOptions: TGUID = '{00020AA1-0000-4B30-A977-D214852036FE}';
  IID_ListGalleries: TGUID = '{00020995-0000-4B30-A977-D214852036FE}';
  IID_ListGallery: TGUID = '{00020994-0000-0000-C000-000000000046}';
  IID_FileConverters: TGUID = '{0002099A-0000-4B30-A977-D214852036FE}';
  IID_FileConverter: TGUID = '{00020999-0000-4B30-A977-D214852036FE}';
  IID_CaptionLabels: TGUID = '{00020978-0000-4B30-A977-D214852036FE}';
  IID_CaptionLabel: TGUID = '{00020979-0000-4B30-A977-D214852036FE}';
  IID_Options: TGUID = '{000209B7-0000-4B30-A977-D214852036FE}';
  IID_RecentFiles: TGUID = '{00020963-0000-4B30-A977-D214852036FE}';
  IID_RecentFile: TGUID = '{00020964-0000-4B30-A977-D214852036FE}';
  IID_Template: TGUID = '{0002096A-0000-4B30-A977-D214852036FE}';
  IID_Templates: TGUID = '{000209A2-0000-4B30-A977-D214852036FE}';
  IID_Browser: TGUID = '{0002092E-0000-4B30-A977-D214852036FE}';
  IID_System: TGUID = '{00020935-0000-4B30-A977-D214852036FE}';
  IID_PdfExportOptions: TGUID = '{00021000-0000-4B30-A977-D214852036FE}';
  IID_Dictionaries: TGUID = '{000209AC-0000-4B30-A977-D214852036FE}';
  IID_Dictionary: TGUID = '{000209AD-0000-4B30-A977-D214852036FE}';
  IID_Dialogs: TGUID = '{00020910-0000-4B30-A977-D214852036FE}';
  IID_Dialog: TGUID = '{000209B8-0000-4B30-A977-D214852036FE}';
  IID_AutoCorrect: TGUID = '{00020948-0000-4B30-A977-D214852036FE}';
  IID_FontNames: TGUID = '{0002096F-0000-4B30-A977-D214852036FE}';
  IID_AutoCaptions: TGUID = '{0002097A-0000-4B30-A977-D214852036FE}';
  IID_AutoCaption: TGUID = '{0002097B-0000-4B30-A977-D214852036FE}';
  DIID_ApplicationEvents: TGUID = '{000209FE-0000-4B30-A977-D214852036FE}';
  DIID_DocumentEvents: TGUID = '{000209F6-0000-4B30-A977-D214852036FE}';
  IID__OLEControl: TGUID = '{000209A4-0000-0000-C000-000000000046}';
  DIID_OCXEvents: TGUID = '{000209F3-0000-0000-C000-000000000046}';
  CLASS_OLECtrl: TGUID = '{000209F2-0000-0000-C000-000000000046}';
  CLASS_Application: TGUID = '{000209FF-0000-4B30-A977-D214852036FE}';
  CLASS_Document: TGUID = '{00020906-0000-4B30-A977-D214852036FE}';
  CLASS_Font: TGUID = '{000209F5-0000-4B30-A977-D214852036FE}';
  CLASS_ParagraphFormat: TGUID = '{000209F4-0000-4B30-A977-D214852036FE}';
  CLASS_WebOptions: TGUID = '{E4756489-8DE4-4994-8701-4D596F30F005}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum WpsColor
type
  WpsColor = TOleEnum;
const
  wpsColorAutomatic = $FF000000;
  wpsColorBlack = $00000000;
  wpsColorBlue = $00FF0000;
  wpsColorTurquoise = $00FFFF00;
  wpsColorBrightGreen = $0000FF00;
  wpsColorPink = $00FF00FF;
  wpsColorRed = $000000FF;
  wpsColorYellow = $0000FFFF;
  wpsColorWhite = $00FFFFFF;
  wpsColorDarkBlue = $00800000;
  wpsColorTeal = $00808000;
  wpsColorGreen = $00008000;
  wpsColorViolet = $00800080;
  wpsColorDarkRed = $00000080;
  wpsColorDarkYellow = $00008080;
  wpsColorBrown = $00003399;
  wpsColorOliveGreen = $00003333;
  wpsColorDarkGreen = $00003300;
  wpsColorDarkTeal = $00663300;
  wpsColorIndigo = $00993333;
  wpsColorOrange = $000066FF;
  wpsColorBlueGray = $00996666;
  wpsColorLightOrange = $000099FF;
  wpsColorLime = $0000CC99;
  wpsColorSeaGreen = $00669933;
  wpsColorAqua = $00CCCC33;
  wpsColorLightBlue = $00FF6633;
  wpsColorGold = $0000CCFF;
  wpsColorSkyBlue = $00FFCC00;
  wpsColorPlum = $00663399;
  wpsColorRose = $00CC99FF;
  wpsColorTan = $0099CCFF;
  wpsColorLightYellow = $0099FFFF;
  wpsColorLightGreen = $00CCFFCC;
  wpsColorLightTurquoise = $00FFFFCC;
  wpsColorPaleBlue = $00FFCC99;
  wpsColorLavender = $00FF99CC;
  wpsColorGray05 = $00F3F3F3;
  wpsColorGray10 = $00E6E6E6;
  wpsColorGray125 = $00E0E0E0;
  wpsColorGray15 = $00D9D9D9;
  wpsColorGray20 = $00CCCCCC;
  wpsColorGray25 = $00C0C0C0;
  wpsColorGray30 = $00B3B3B3;
  wpsColorGray35 = $00A6A6A6;
  wpsColorGray375 = $00A0A0A0;
  wpsColorGray40 = $00999999;
  wpsColorGray45 = $008C8C8C;
  wpsColorGray50 = $00808080;
  wpsColorGray55 = $00737373;
  wpsColorGray60 = $00666666;
  wpsColorGray625 = $00606060;
  wpsColorGray65 = $00595959;
  wpsColorGray70 = $004C4C4C;
  wpsColorGray75 = $00404040;
  wpsColorGray80 = $00333333;
  wpsColorGray85 = $00262626;
  wpsColorGray875 = $00202020;
  wpsColorGray90 = $00191919;
  wpsColorGray95 = $000C0C0C;

// Constants for enum WpsUnderline
type
  WpsUnderline = TOleEnum;
const
  wpsUnderlineNone = $00000000;
  wpsUnderlineSingle = $00000001;
  wpsUnderlineWords = $00000002;
  wpsUnderlineDouble = $00000003;
  wpsUnderlineDotted = $00000004;
  wpsUnderlineThick = $00000006;
  wpsUnderlineDash = $00000007;
  wpsUnderlineDotDash = $00000009;
  wpsUnderlineDotDotDash = $0000000A;
  wpsUnderlineWavy = $0000000B;
  wpsUnderlineWavyHeavy = $0000001B;
  wpsUnderlineDottedHeavy = $00000014;
  wpsUnderlineDashHeavy = $00000017;
  wpsUnderlineDotDashHeavy = $00000019;
  wpsUnderlineDotDotDashHeavy = $0000001A;
  wpsUnderlineDashLong = $00000027;
  wpsUnderlineDashLongHeavy = $00000037;
  wpsUnderlineWavyDouble = $0000002B;

// Constants for enum WpsLineStyle
type
  WpsLineStyle = TOleEnum;
const
  wpsLineStyleNone = $00000000;
  wpsLineStyleSingle = $00000001;
  wpsLineStyleDot = $00000002;
  wpsLineStyleDashSmallGap = $00000003;
  wpsLineStyleDashLargeGap = $00000004;
  wpsLineStyleDashDot = $00000005;
  wpsLineStyleDashDotDot = $00000006;
  wpsLineStyleDouble = $00000007;
  wpsLineStyleTriple = $00000008;
  wpsLineStyleThinThickSmallGap = $00000009;
  wpsLineStyleThickThinSmallGap = $0000000A;
  wpsLineStyleThinThickThinSmallGap = $0000000B;
  wpsLineStyleThinThickMedGap = $0000000C;
  wpsLineStyleThickThinMedGap = $0000000D;
  wpsLineStyleThinThickThinMedGap = $0000000E;
  wpsLineStyleThinThickLargeGap = $0000000F;
  wpsLineStyleThickThinLargeGap = $00000010;
  wpsLineStyleThinThickThinLargeGap = $00000011;
  wpsLineStyleSingleWavy = $00000012;
  wpsLineStyleDoubleWavy = $00000013;
  wpsLineStyleDashDotStroked = $00000014;
  wpsLineStyleEmboss3D = $00000015;
  wpsLineStyleEngrave3D = $00000016;
  wpsLineStyleOutset = $00000017;
  wpsLineStyleInset = $00000018;

// Constants for enum WpsLineWidth
type
  WpsLineWidth = TOleEnum;
const
  wpsLineWidth025pt = $00000002;
  wpsLineWidth050pt = $00000004;
  wpsLineWidth075pt = $00000006;
  wpsLineWidth100pt = $00000008;
  wpsLineWidth150pt = $0000000C;
  wpsLineWidth225pt = $00000012;
  wpsLineWidth300pt = $00000018;
  wpsLineWidth450pt = $00000024;
  wpsLineWidth600pt = $00000030;

// Constants for enum WpsBorderDistanceFrom
type
  WpsBorderDistanceFrom = TOleEnum;
const
  wpsBorderDistanceFromText = $00000000;
  wpsBorderDistanceFromPageEdge = $00000001;

// Constants for enum WpsBorderType
type
  WpsBorderType = TOleEnum;
const
  wpsBorderTop = $FFFFFFFF;
  wpsBorderLeft = $FFFFFFFE;
  wpsBorderBottom = $FFFFFFFD;
  wpsBorderRight = $FFFFFFFC;
  wpsBorderHorizontal = $FFFFFFFB;
  wpsBorderVertical = $FFFFFFFA;
  wpsBorderDiagonalDown = $FFFFFFF9;
  wpsBorderDiagonalUp = $FFFFFFF8;

// Constants for enum WpsTextureIndex
type
  WpsTextureIndex = TOleEnum;
const
  wpsTextureNone = $00000000;
  wpsTexture2Pt5Percent = $00000019;
  wpsTexture5Percent = $00000032;
  wpsTexture7Pt5Percent = $0000004B;
  wpsTexture10Percent = $00000064;
  wpsTexture12Pt5Percent = $0000007D;
  wpsTexture15Percent = $00000096;
  wpsTexture17Pt5Percent = $000000AF;
  wpsTexture20Percent = $000000C8;
  wpsTexture22Pt5Percent = $000000E1;
  wpsTexture25Percent = $000000FA;
  wpsTexture27Pt5Percent = $00000113;
  wpsTexture30Percent = $0000012C;
  wpsTexture32Pt5Percent = $00000145;
  wpsTexture35Percent = $0000015E;
  wpsTexture37Pt5Percent = $00000177;
  wpsTexture40Percent = $00000190;
  wpsTexture42Pt5Percent = $000001A9;
  wpsTexture45Percent = $000001C2;
  wpsTexture47Pt5Percent = $000001DB;
  wpsTexture50Percent = $000001F4;
  wpsTexture52Pt5Percent = $0000020D;
  wpsTexture55Percent = $00000226;
  wpsTexture57Pt5Percent = $0000023F;
  wpsTexture60Percent = $00000258;
  wpsTexture62Pt5Percent = $00000271;
  wpsTexture65Percent = $0000028A;
  wpsTexture67Pt5Percent = $000002A3;
  wpsTexture70Percent = $000002BC;
  wpsTexture72Pt5Percent = $000002D5;
  wpsTexture75Percent = $000002EE;
  wpsTexture77Pt5Percent = $00000307;
  wpsTexture80Percent = $00000320;
  wpsTexture82Pt5Percent = $00000339;
  wpsTexture85Percent = $00000352;
  wpsTexture87Pt5Percent = $0000036B;
  wpsTexture90Percent = $00000384;
  wpsTexture92Pt5Percent = $0000039D;
  wpsTexture95Percent = $000003B6;
  wpsTexture97Pt5Percent = $000003CF;
  wpsTextureSolid = $000003E8;
  wpsTextureDarkHorizontal = $FFFFFFFF;
  wpsTextureDarkVertical = $FFFFFFFE;
  wpsTextureDarkDiagonalDown = $FFFFFFFD;
  wpsTextureDarkDiagonalUp = $FFFFFFFC;
  wpsTextureDarkCross = $FFFFFFFB;
  wpsTextureDarkDiagonalCross = $FFFFFFFA;
  wpsTextureHorizontal = $FFFFFFF9;
  wpsTextureVertical = $FFFFFFF8;
  wpsTextureDiagonalDown = $FFFFFFF7;
  wpsTextureDiagonalUp = $FFFFFFF6;
  wpsTextureCross = $FFFFFFF5;
  wpsTextureDiagonalCross = $FFFFFFF4;

// Constants for enum WpsEmphasisMark
type
  WpsEmphasisMark = TOleEnum;
const
  wpsEmphasisMarkNone = $00000000;
  wpsEmphasisMarkOverSolidCircle = $00000001;
  wpsEmphasisMarkOverComma = $00000002;
  wpsEmphasisMarkOverWhiteCircle = $00000003;
  wpsEmphasisMarkUnderSolidCircle = $00000004;

// Constants for enum WpsStoryType
type
  WpsStoryType = TOleEnum;
const
  wpsMainTextStory = $00000001;
  wpsFootnotesStory = $00000002;
  wpsEndnotesStory = $00000003;
  wpsCommentsStory = $00000004;
  wpsTextFrameStory = $00000005;
  wpsEvenPagesHeaderStory = $00000006;
  wpsPrimaryHeaderStory = $00000007;
  wpsEvenPagesFooterStory = $00000008;
  wpsPrimaryFooterStory = $00000009;
  wpsFirstPageHeaderStory = $0000000A;
  wpsFirstPageFooterStory = $0000000B;

// Constants for enum WpsRowHeightRule
type
  WpsRowHeightRule = TOleEnum;
const
  wpsRowHeightAuto = $00000000;
  wpsRowHeightAtLeast = $00000001;
  wpsRowHeightExactly = $00000002;

// Constants for enum WpsCellVerticalAlignment
type
  WpsCellVerticalAlignment = TOleEnum;
const
  wpsCellAlignVerticalTop = $00000000;
  wpsCellAlignVerticalCenter = $00000001;
  wpsCellAlignVerticalBottom = $00000003;

// Constants for enum WpsRowAlignment
type
  WpsRowAlignment = TOleEnum;
const
  wpsAlignRowLeft = $00000000;
  wpsAlignRowCenter = $00000001;
  wpsAlignRowRight = $00000002;

// Constants for enum WpsRulerStyle
type
  WpsRulerStyle = TOleEnum;
const
  wpsAdjustNone = $00000000;
  wpsAdjustProportional = $00000001;
  wpsAdjustFirstColumn = $00000002;
  wpsAdjustSameWidth = $00000003;

// Constants for enum WpsPreferredWidthType
type
  WpsPreferredWidthType = TOleEnum;
const
  wpsPreferredWidthAuto = $00000001;
  wpsPreferredWidthPercent = $00000002;
  wpsPreferredWidthPoints = $00000003;

// Constants for enum WpsDiagonalCellType
type
  WpsDiagonalCellType = TOleEnum;
const
  wpsDiagonalCellType_None = $00000000;
  wpsDiagonalCellType_1Row1ColumnNoBody = $00000001;
  WpsDiagonalCellType_1Row1Column = $00000002;
  WpsDiagonalCellType_1Row1ColumnReverse = $00000003;
  WpsDiagonalCellType_1Row2ColumnNoBody = $00000004;
  WpsDiagonalCellType_2Row1ColumnNoBody = $00000005;
  WpsDiagonalCellType_1Row2Column = $00000006;
  WpsDiagonalCellType_2Row1Column = $00000007;
  WpsDiagonalCellType_2Row2ColumnNoBody = $00000008;

// Constants for enum WpsTextOrientation
type
  WpsTextOrientation = TOleEnum;
const
  wpsTextOrientationHorizontal = $00000000;
  wpsTextOrientationVerticalFarEast = $00000001;
  wpsTextOrientationUpward = $00000002;
  wpsTextOrientationDownward = $00000003;
  wpsTextOrientationHorizontalRotatedFarEast = $00000004;
  wpsTextOrientationVertL2R = $0000000A;
  wpsTextOrientationVertL2RTrans = $0000000B;

// Constants for enum WpsRelativeHorizontalPosition
type
  WpsRelativeHorizontalPosition = TOleEnum;
const
  wpsRelativeHorizontalPositionMargin = $00000000;
  wpsRelativeHorizontalPositionPage = $00000001;
  wpsRelativeHorizontalPositionColumn = $00000002;
  wpsRelativeHorizontalPositionCharacter = $00000003;

// Constants for enum WpsRelativeVerticalPosition
type
  WpsRelativeVerticalPosition = TOleEnum;
const
  wpsRelativeVerticalPositionMargin = $00000000;
  wpsRelativeVerticalPositionPage = $00000001;
  wpsRelativeVerticalPositionParagraph = $00000002;
  wpsRelativeVerticalPositionLine = $00000003;

// Constants for enum WpsTableDirection
type
  WpsTableDirection = TOleEnum;
const
  wpsTableDirectionRtl = $00000000;
  wpsTableDirectionLtr = $00000001;

// Constants for enum WpsAutoFitBehavior
type
  WpsAutoFitBehavior = TOleEnum;
const
  wpsAutoFitFixed = $00000000;
  wpsAutoFitContent = $00000001;
  wpsAutoFitWindow = $00000002;

// Constants for enum WpsFootnoteLocation
type
  WpsFootnoteLocation = TOleEnum;
const
  wpsBottomOfPage = $00000000;
  wpsBeneathText = $00000001;

// Constants for enum WpsNoteNumberStyle
type
  WpsNoteNumberStyle = TOleEnum;
const
  wpsNoteNumberStyleArabic = $00000000;
  wpsNoteNumberStyleUppercaseRoman = $00000001;
  wpsNoteNumberStyleLowercaseRoman = $00000002;
  wpsNoteNumberStyleUppercaseLetter = $00000003;
  wpsNoteNumberStyleLowercaseLetter = $00000004;
  wpsNoteNumberStyleSymbol = $00000009;
  wpsNoteNumberStyleArabicFullWidth = $0000000E;
  wpsNoteNumberStyleKanji = $0000000A;
  wpsNoteNumberStyleKanjiDigit = $0000000B;
  wpsNoteNumberStyleKanjiTraditional = $00000010;
  wpsNoteNumberStyleNumberInCircle = $00000012;
  wpsNoteNumberStyleHanjaRead = $00000029;
  wpsNoteNumberStyleHanjaReadDigit = $0000002A;
  wpsNoteNumberStyleTradChinNum1 = $00000021;
  wpsNoteNumberStyleTradChinNum2 = $00000022;
  wpsNoteNumberStyleSimpChinNum1 = $00000025;
  wpsNoteNumberStyleSimpChinNum2 = $00000026;
  wpsNoteNumberStyleHebrewLetter1 = $0000002D;
  wpsNoteNumberStyleArabicLetter1 = $0000002E;
  wpsNoteNumberStyleHebrewLetter2 = $0000002F;
  wpsNoteNumberStyleArabicLetter2 = $00000030;
  wpsNoteNumberStyleHindiLetter1 = $00000031;
  wpsNoteNumberStyleHindiLetter2 = $00000032;
  wpsNoteNumberStyleHindiArabic = $00000033;
  wpsNoteNumberStyleHindiCardinalText = $00000034;
  wpsNoteNumberStyleThaiLetter = $00000035;
  wpsNoteNumberStyleThaiArabic = $00000036;
  wpsNoteNumberStyleThaiCardinalText = $00000037;
  wpsNoteNumberStyleVietCardinalText = $00000038;

// Constants for enum WpsNumberingRule
type
  WpsNumberingRule = TOleEnum;
const
  wpsRestartContinuous = $00000000;
  wpsRestartSection = $00000001;
  wpsRestartPage = $00000002;

// Constants for enum WpsEndnoteLocation
type
  WpsEndnoteLocation = TOleEnum;
const
  wpsEndOfSection = $00000000;
  wpsEndOfDocument = $00000001;

// Constants for enum WpsOrientation
type
  WpsOrientation = TOleEnum;
const
  wpsOrientPortrait = $00000000;
  wpsOrientLandscape = $00000001;

// Constants for enum WpsPaperTray
type
  WpsPaperTray = TOleEnum;
const
  wpsPrinterDefaultBin = $00000000;
  wpsPrinterUpperBin = $00000001;
  wpsPrinterOnlyBin = $00000001;
  wpsPrinterLowerBin = $00000002;
  wpsPrinterMiddleBin = $00000003;
  wpsPrinterManualFeed = $00000004;
  wpsPrinterEnvelopeFeed = $00000005;
  wpsPrinterManualEnvelopeFeed = $00000006;
  wpsPrinterAutomaticSheetFeed = $00000007;
  wpsPrinterTractorFeed = $00000008;
  wpsPrinterSmallFormatBin = $00000009;
  wpsPrinterLargeFormatBin = $0000000A;
  wpsPrinterLargeCapacityBin = $0000000B;
  wpsPrinterPaperCassette = $0000000E;
  wpsPrinterFormSource = $0000000F;

// Constants for enum WpsSectionStart
type
  WpsSectionStart = TOleEnum;
const
  wpsSectionContinuous = $00000000;
  wpsSectionNewColumn = $00000001;
  wpsSectionNewPage = $00000002;
  wpsSectionEvenPage = $00000003;
  wpsSectionOddPage = $00000004;

// Constants for enum WpsFlowDirection
type
  WpsFlowDirection = TOleEnum;
const
  wpsFlowLtr = $00000000;
  wpsFlowRtl = $00000001;

// Constants for enum WpsPaperSize
type
  WpsPaperSize = TOleEnum;
const
  wpsPaper10x14 = $00000000;
  wpsPaper11x17 = $00000001;
  wpsPaperLetter = $00000002;
  wpsPaperLetterSmall = $00000003;
  wpsPaperLegal = $00000004;
  wpsPaperExecutive = $00000005;
  wpsPaperA3 = $00000006;
  wpsPaperA4 = $00000007;
  wpsPaperA4Small = $00000008;
  wpsPaperA5 = $00000009;
  wpsPaperB4 = $0000000A;
  wpsPaperB5 = $0000000B;
  wpsPaperCSheet = $0000000C;
  wpsPaperFanfoldLegalGerman = $0000000D;
  wpsPaperFanfoldStdGerman = $0000000E;
  wpsPaperFanfoldUS = $0000000F;
  wpsPaperFolio = $00000010;
  wpsPaperLedger = $00000011;
  wpsPaperNote = $00000012;
  wpsPaperQuarto = $00000013;
  wpsPaperStatement = $00000014;
  wpsPaperTabloid = $00000015;
  wpsPaperEnvelope9 = $00000016;
  wpsPaperEnvelope10 = $00000017;
  wpsPaperEnvelope11 = $00000018;
  wpsPaperEnvelope14 = $00000019;
  wpsPaperEnvelopeB4 = $0000001A;
  wpsPaperEnvelopeB5 = $0000001B;
  wpsPaperEnvelopeB6 = $0000001C;
  wpsPaperEnvelopeC3 = $0000001D;
  wpsPaperEnvelopeC4 = $0000001E;
  wpsPaperEnvelopeC5 = $0000001F;
  wpsPaperEnvelopeC6 = $00000020;
  wpsPaperEnvelopeC65 = $00000021;
  wpsPaperEnvelopeDL = $00000022;
  wpsPaperEnvelopeItaly = $00000023;
  wpsPaperEnvelopeMonarch = $00000024;
  wpsPaperEnvelopePersonal = $00000025;
  wpsPaperCustom = $00000026;

// Constants for enum WpsLayoutMode
type
  WpsLayoutMode = TOleEnum;
const
  wpsLayoutModeDefault = $00000000;
  wpsLayoutModeGrid = $00000001;
  wpsLayoutModeLineGrid = $00000002;
  wpsLayoutModeGenko = $00000003;

// Constants for enum WpsGutterStyle
type
  WpsGutterStyle = TOleEnum;
const
  wpsGutterPosLeft = $00000000;
  wpsGutterPosTop = $00000001;
  wpsGutterPosRight = $00000002;

// Constants for enum WpsGenkoGridStyle
type
  WpsGenkoGridStyle = TOleEnum;
const
  wpsGenkoGrid10x20 = $00000000;
  wpsGenkoGrid15x20 = $00000001;
  wpsGenkoGrid20x20 = $00000002;
  wpsGenkoGrid20x25 = $00000003;
  wpsGenkoGrid20x20Classical = $00000004;
  wpsGenkoGrid20x25Classical = $00000005;
  wpsGenkoGrid24x25Classical = $00000006;
  wpsGenkoGridModel = $00000007;

// Constants for enum WpsGenkoGrid
type
  WpsGenkoGrid = TOleEnum;
const
  wpsGenkoGrid_Grid = $00000000;
  wpsGenkoGrid_Underline = $00000001;
  wpsGenkoGrid_Border = $00000002;

// Constants for enum WpsHeaderFooterIndex
type
  WpsHeaderFooterIndex = TOleEnum;
const
  wpsHeaderFooterPrimary = $00000001;
  wpsHeaderFooterFirstPage = $00000002;
  wpsHeaderFooterEvenPages = $00000003;

// Constants for enum WpsPageNumberStyle
type
  WpsPageNumberStyle = TOleEnum;
const
  wpsPageNumberStyleArabic = $00000000;
  wpsPageNumberStyleUppercaseRoman = $00000001;
  wpsPageNumberStyleLowercaseRoman = $00000002;
  wpsPageNumberStyleUppercaseLetter = $00000003;
  wpsPageNumberStyleLowercaseLetter = $00000004;
  wpsPageNumberStyleArabicFullWidth = $0000000E;
  wpsPageNumberStyleKanji = $0000000A;
  wpsPageNumberStyleKanjiDigit = $0000000B;
  wpsPageNumberStyleKanjiTraditional = $00000010;
  wpsPageNumberStyleNumberInCircle = $00000012;
  wpsPageNumberStyleHanjaRead = $00000029;
  wpsPageNumberStyleHanjaReadDigit = $0000002A;
  wpsPageNumberStyleTradChinNum1 = $00000021;
  wpsPageNumberStyleTradChinNum2 = $00000022;
  wpsPageNumberStyleSimpChinNum1 = $00000025;
  wpsPageNumberStyleSimpChinNum2 = $00000026;
  wpsPageNumberStyleHebrewLetter1 = $0000002D;
  wpsPageNumberStyleArabicLetter1 = $0000002E;
  wpsPageNumberStyleHebrewLetter2 = $0000002F;
  wpsPageNumberStyleArabicLetter2 = $00000030;
  wpsPageNumberStyleHindiLetter1 = $00000031;
  wpsPageNumberStyleHindiLetter2 = $00000032;
  wpsPageNumberStyleHindiArabic = $00000033;
  wpsPageNumberStyleHindiCardinalText = $00000034;
  wpsPageNumberStyleThaiLetter = $00000035;
  wpsPageNumberStyleThaiArabic = $00000036;
  wpsPageNumberStyleThaiCardinalText = $00000037;
  wpsPageNumberStyleVietCardinalText = $00000038;
  wpsPageNumberStyleNumberInDash = $00000039;

// Constants for enum WpsSeparatorType
type
  WpsSeparatorType = TOleEnum;
const
  wpsSeparatorHyphen = $00000000;
  wpsSeparatorPeriod = $00000001;
  wpsSeparatorColon = $00000002;
  wpsSeparatorEmDash = $00000003;
  wpsSeparatorEnDash = $00000004;

// Constants for enum WpsPageNumberAlignment
type
  WpsPageNumberAlignment = TOleEnum;
const
  wpsAlignPageNumberLeft = $00000000;
  wpsAlignPageNumberCenter = $00000001;
  wpsAlignPageNumberRight = $00000002;
  wpsAlignPageNumberInside = $00000003;
  wpsAlignPageNumberOutside = $00000004;

// Constants for enum WpsWrapType
type
  WpsWrapType = TOleEnum;
const
  wpsWrapSquare = $00000000;
  wpsWrapTight = $00000001;
  wpsWrapThrough = $00000002;
  wpsWrapNone = $00000003;
  wpsWrapTopBottom = $00000004;
  wpsWrapInline = $00000007;

// Constants for enum WpsWrapSideType
type
  WpsWrapSideType = TOleEnum;
const
  wpsWrapBoth = $00000000;
  wpsWrapLeft = $00000001;
  wpsWrapRight = $00000002;
  wpsWrapLargest = $00000003;

// Constants for enum WpsFrameSizeRule
type
  WpsFrameSizeRule = TOleEnum;
const
  wpsFrameAuto = $00000000;
  wpsFrameAtLeast = $00000001;
  wpsFrameExact = $00000002;

// Constants for enum WpsFieldType
type
  WpsFieldType = TOleEnum;
const
  wpsFieldEmpty = $FFFFFFFF;
  wpsFieldRef = $00000003;
  wpsFieldIndexEntry = $00000004;
  wpsFieldFootnoteRef = $00000005;
  wpsFieldSet = $00000006;
  wpsFieldIf = $00000007;
  wpsFieldIndex = $00000008;
  wpsFieldTOCEntry = $00000009;
  wpsFieldStyleRef = $0000000A;
  wpsFieldRefDoc = $0000000B;
  wpsFieldSequence = $0000000C;
  wpsFieldTOC = $0000000D;
  wpsFieldInfo = $0000000E;
  wpsFieldTitle = $0000000F;
  wpsFieldSubject = $00000010;
  wpsFieldAuthor = $00000011;
  wpsFieldKeyWord = $00000012;
  wpsFieldComments = $00000013;
  wpsFieldLastSavedBy = $00000014;
  wpsFieldCreateDate = $00000015;
  wpsFieldSaveDate = $00000016;
  wpsFieldPrintDate = $00000017;
  wpsFieldRevisionNum = $00000018;
  wpsFieldEditTime = $00000019;
  wpsFieldNumPages = $0000001A;
  wpsFieldNumWords = $0000001B;
  wpsFieldNumChars = $0000001C;
  wpsFieldFileName = $0000001D;
  wpsFieldTemplate = $0000001E;
  wpsFieldDate = $0000001F;
  wpsFieldTime = $00000020;
  wpsFieldPage = $00000021;
  wpsFieldExpression = $00000022;
  wpsFieldQuote = $00000023;
  wpsFieldInclude = $00000024;
  wpsFieldPageRef = $00000025;
  wpsFieldAsk = $00000026;
  wpsFieldFillIn = $00000027;
  wpsFieldData = $00000028;
  wpsFieldNext = $00000029;
  wpsFieldNextIf = $0000002A;
  wpsFieldSkipIf = $0000002B;
  wpsFieldMergeRec = $0000002C;
  wpsFieldDDE = $0000002D;
  wpsFieldDDEAuto = $0000002E;
  wpsFieldGlossary = $0000002F;
  wpsFieldPrint = $00000030;
  wpsFieldFormula = $00000031;
  wpsFieldGoToButton = $00000032;
  wpsFieldMacroButton = $00000033;
  wpsFieldAutoNumOutline = $00000034;
  wpsFieldAutoNumLegal = $00000035;
  wpsFieldAutoNum = $00000036;
  wpsFieldImport = $00000037;
  wpsFieldLink = $00000038;
  wpsFieldSymbol = $00000039;
  wpsFieldEmbed = $0000003A;
  wpsFieldMergeField = $0000003B;
  wpsFieldUserName = $0000003C;
  wpsFieldUserInitials = $0000003D;
  wpsFieldUserAddress = $0000003E;
  wpsFieldBarCode = $0000003F;
  wpsFieldDocVariable = $00000040;
  wpsFieldSection = $00000041;
  wpsFieldSectionPages = $00000042;
  wpsFieldIncludePicture = $00000043;
  wpsFieldIncludeText = $00000044;
  wpsFieldFileSize = $00000045;
  wpsFieldFormTextInput = $00000046;
  wpsFieldFormCheckBox = $00000047;
  wpsFieldNoteRef = $00000048;
  wpsFieldTOA = $00000049;
  wpsFieldTOAEntry = $0000004A;
  wpsFieldMergeSeq = $0000004B;
  wpsFieldPrivate = $0000004D;
  wpsFieldDatabase = $0000004E;
  wpsFieldAutoText = $0000004F;
  wpsFieldCompare = $00000050;
  wpsFieldAddin = $00000051;
  wpsFieldSubscriber = $00000052;
  wpsFieldFormDropDown = $00000053;
  wpsFieldAdvance = $00000054;
  wpsFieldDocProperty = $00000055;
  wpsFieldOCX = $00000057;
  wpsFieldHyperlink = $00000058;
  wpsFieldAutoTextList = $00000059;
  wpsFieldListNum = $0000005A;
  wpsFieldHTMLActiveX = $0000005B;
  wpsFieldBidiOutline = $0000005C;
  wpsFieldAddressBlock = $0000005D;
  wpsFieldGreetingLine = $0000005E;
  wpsFieldShape = $0000005F;

// Constants for enum WpsFieldKind
type
  WpsFieldKind = TOleEnum;
const
  wpsFieldKindNone = $00000000;
  wpsFieldKindHot = $00000001;
  wpsFieldKindWarm = $00000002;
  wpsFieldKindCold = $00000003;

// Constants for enum WpsInlineShapeType
type
  WpsInlineShapeType = TOleEnum;
const
  wpsInlineShapeEmbeddedOLEObject = $00000001;
  wpsInlineShapeLinkedOLEObject = $00000002;
  wpsInlineShapePicture = $00000003;
  wpsInlineShapeLinkedPicture = $00000004;
  wpsInlineShapeOLEControlObject = $00000005;
  wpsInlineShapeHorizontalLine = $00000006;
  wpsInlineShapePictureHorizontalLine = $00000007;
  wpsInlineShapeLinkedPictureHorizontalLine = $00000008;
  wpsInlineShapePictureBullet = $00000009;
  wpsInlineShapeScriptAnchor = $0000000A;
  wpsInlineShapeOWSAnchor = $0000000B;

// Constants for enum WpsParagraphAlignment
type
  WpsParagraphAlignment = TOleEnum;
const
  wpsAlignParagraphLeft = $00000000;
  wpsAlignParagraphCenter = $00000001;
  wpsAlignParagraphRight = $00000002;
  wpsAlignParagraphJustify = $00000003;
  wpsAlignParagraphDistribute = $00000004;
  wpsAlignParagraphJustifyMed = $00000005;
  wpsAlignParagraphJustifyHi = $00000007;
  wpsAlignParagraphJustifyLow = $00000008;
  wpsAlignParagraphThaiJustify = $00000009;

// Constants for enum WpsLineSpacing
type
  WpsLineSpacing = TOleEnum;
const
  wpsLineSpaceSingle = $00000000;
  wpsLineSpace1pt5 = $00000001;
  wpsLineSpaceDouble = $00000002;
  wpsLineSpaceAtLeast = $00000003;
  wpsLineSpaceExactly = $00000004;
  wpsLineSpaceMultiple = $00000005;

// Constants for enum WpsBaselineAlignment
type
  WpsBaselineAlignment = TOleEnum;
const
  wpsBaselineAlignTop = $00000000;
  wpsBaselineAlignCenter = $00000001;
  wpsBaselineAlignBaseline = $00000002;
  wpsBaselineAlignFarEast50 = $00000003;
  wpsBaselineAlignAuto = $00000004;

// Constants for enum WpsTabAlignment
type
  WpsTabAlignment = TOleEnum;
const
  wpsAlignTabLeft = $00000000;
  wpsAlignTabCenter = $00000001;
  wpsAlignTabRight = $00000002;
  wpsAlignTabDecimal = $00000003;
  wpsAlignTabBar = $00000004;
  wpsAlignTabList = $00000006;

// Constants for enum WpsTabLeader
type
  WpsTabLeader = TOleEnum;
const
  wpsTabLeaderSpaces = $00000000;
  wpsTabLeaderDots = $00000001;
  wpsTabLeaderDashes = $00000002;
  wpsTabLeaderLines = $00000003;
  wpsTabLeaderHeavy = $00000004;
  wpsTabLeaderMiddleDot = $00000005;

// Constants for enum WpsOutlineLevel
type
  WpsOutlineLevel = TOleEnum;
const
  wpsOutlineLevel1 = $00000001;
  wpsOutlineLevel2 = $00000002;
  wpsOutlineLevel3 = $00000003;
  wpsOutlineLevel4 = $00000004;
  wpsOutlineLevel5 = $00000005;
  wpsOutlineLevel6 = $00000006;
  wpsOutlineLevel7 = $00000007;
  wpsOutlineLevel8 = $00000008;
  wpsOutlineLevel9 = $00000009;
  wpsOutlineLevelBodyText = $0000000A;

// Constants for enum WpsParagraphReadingOrder
type
  WpsParagraphReadingOrder = TOleEnum;
const
  wpsReadingOrderLtr = $00000000;
  wpsReadingOrderRtl = $00000001;

// Constants for enum WpsDropPosition
type
  WpsDropPosition = TOleEnum;
const
  wpsDropNone = $00000000;
  wpsDropNormal = $00000001;
  wpsDropMargin = $00000002;

// Constants for enum WpsStyleType
type
  WpsStyleType = TOleEnum;
const
  wpsStyleTypeParagraph = $00000001;
  wpsStyleTypeCharacter = $00000002;
  wpsStyleTypeTable = $00000003;
  wpsStyleTypeList = $00000004;

// Constants for enum WpsTrailingCharacter
type
  WpsTrailingCharacter = TOleEnum;
const
  wpsTrailingTab = $00000000;
  wpsTrailingSpace = $00000001;
  wpsTrailingNone = $00000002;

// Constants for enum WpsListNumberStyle
type
  WpsListNumberStyle = TOleEnum;
const
  wpsListNumberStyleArabic = $00000000;
  wpsListNumberStyleUppercaseRoman = $00000001;
  wpsListNumberStyleLowercaseRoman = $00000002;
  wpsListNumberStyleUppercaseLetter = $00000003;
  wpsListNumberStyleLowercaseLetter = $00000004;
  wpsListNumberStyleOrdinal = $00000005;
  wpsListNumberStyleCardinalText = $00000006;
  wpsListNumberStyleOrdinalText = $00000007;
  wpsListNumberStyleKanji = $0000000A;
  wpsListNumberStyleKanjiDigit = $0000000B;
  wpsListNumberStyleAiueoHalfWidth = $0000000C;
  wpsListNumberStyleIrohaHalfWidth = $0000000D;
  wpsListNumberStyleArabicFullWidth = $0000000E;
  wpsListNumberStyleKanjiTraditional = $00000010;
  wpsListNumberStyleKanjiTraditional2 = $00000011;
  wpsListNumberStyleNumberInCircle = $00000012;
  wpsListNumberStyleAiueo = $00000014;
  wpsListNumberStyleIroha = $00000015;
  wpsListNumberStyleArabicLZ = $00000016;
  wpsListNumberStyleBullet = $00000017;
  wpsListNumberStyleGanada = $00000018;
  wpsListNumberStyleChosung = $00000019;
  wpsListNumberStyleGBNum1 = $0000001A;
  wpsListNumberStyleGBNum2 = $0000001B;
  wpsListNumberStyleGBNum3 = $0000001C;
  wpsListNumberStyleGBNum4 = $0000001D;
  wpsListNumberStyleZodiac1 = $0000001E;
  wpsListNumberStyleZodiac2 = $0000001F;
  wpsListNumberStyleZodiac3 = $00000020;
  wpsListNumberStyleTradChinNum1 = $00000021;
  wpsListNumberStyleTradChinNum2 = $00000022;
  wpsListNumberStyleTradChinNum3 = $00000023;
  wpsListNumberStyleTradChinNum4 = $00000024;
  wpsListNumberStyleSimpChinNum1 = $00000025;
  wpsListNumberStyleSimpChinNum2 = $00000026;
  wpsListNumberStyleSimpChinNum3 = $00000027;
  wpsListNumberStyleSimpChinNum4 = $00000028;
  wpsListNumberStyleHanjaRead = $00000029;
  wpsListNumberStyleHanjaReadDigit = $0000002A;
  wpsListNumberStyleHangul = $0000002B;
  wpsListNumberStyleHanja = $0000002C;
  wpsListNumberStyleHebrew1 = $0000002D;
  wpsListNumberStyleArabic1 = $0000002E;
  wpsListNumberStyleHebrew2 = $0000002F;
  wpsListNumberStyleArabic2 = $00000030;
  wpsListNumberStyleHindiLetter1 = $00000031;
  wpsListNumberStyleHindiLetter2 = $00000032;
  wpsListNumberStyleHindiArabic = $00000033;
  wpsListNumberStyleHindiCardinalText = $00000034;
  wpsListNumberStyleThaiLetter = $00000035;
  wpsListNumberStyleThaiArabic = $00000036;
  wpsListNumberStyleThaiCardinalText = $00000037;
  wpsListNumberStyleVietCardinalText = $00000038;
  wpsListNumberStyleLowercaseRussian = $0000003A;
  wpsListNumberStyleUppercaseRussian = $0000003B;
  wpsListNumberStylePictureBullet = $000000F9;
  wpsListNumberStyleLegal = $000000FD;
  wpsListNumberStyleLegalLZ = $000000FE;
  wpsListNumberStyleNone = $000000FF;

// Constants for enum WpsListLevelAlignment
type
  WpsListLevelAlignment = TOleEnum;
const
  wpsListLevelAlignLeft = $00000000;
  wpsListLevelAlignCenter = $00000001;
  wpsListLevelAlignRight = $00000002;

// Constants for enum WpsNumberType
type
  WpsNumberType = TOleEnum;
const
  wpsNumberParagraph = $00000001;
  wpsNumberListNum = $00000002;
  wpsNumberAllNumbers = $00000003;

// Constants for enum WpsContinue
type
  WpsContinue = TOleEnum;
const
  wpsContinueDisabled = $00000000;
  wpsResetList = $00000001;
  wpsContinueList = $00000002;

// Constants for enum WpsListType
type
  WpsListType = TOleEnum;
const
  wpsListNoNumbering = $00000000;
  wpsListListNumOnly = $00000001;
  wpsListBullet = $00000002;
  wpsListSimpleNumbering = $00000003;
  wpsListOutlineNumbering = $00000004;
  wpsListMixedNumbering = $00000005;
  wpsListPictureBullet = $00000006;

// Constants for enum WpsListApplyTo
type
  WpsListApplyTo = TOleEnum;
const
  wpsListApplyToWholeList = $00000000;
  wpsListApplyToThisPointForward = $00000001;
  wpsListApplyToSelection = $00000002;

// Constants for enum WpsBookmarkSortBy
type
  WpsBookmarkSortBy = TOleEnum;
const
  wpsSortByName = $00000000;
  wpsSortByLocation = $00000001;

// Constants for enum WpsInformation
type
  WpsInformation = TOleEnum;
const
  wpsActiveEndAdjustedPageNumber = $00000001;
  wpsActiveEndSectionNumber = $00000002;
  wpsActiveEndPageNumber = $00000003;
  wpsNumberOfPagesInDocument = $00000004;
  wpsHorizontalPositionRelativeToPage = $00000005;
  wpsVerticalPositionRelativeToPage = $00000006;
  wpsHorizontalPositionRelativeToTextBoundary = $00000007;
  wpsVerticalPositionRelativeToTextBoundary = $00000008;
  wpsFirstCharacterColumnNumber = $00000009;
  wpsFirstCharacterLineNumber = $0000000A;
  wpsFrameIsSelected = $0000000B;
  wpsWithInTable = $0000000C;
  wpsStartOfRangeRowNumber = $0000000D;
  wpsEndOfRangeRowNumber = $0000000E;
  wpsMaximumNumberOfRows = $0000000F;
  wpsStartOfRangeColumnNumber = $00000010;
  wpsEndOfRangeColumnNumber = $00000011;
  wpsMaximumNumberOfColumns = $00000012;
  wpsZoomPercentage = $00000013;
  wpsSelectionMode = $00000014;
  wpsCapsLock = $00000015;
  wpsNumLock = $00000016;
  wpsOverType = $00000017;
  wpsRevisionMarking = $00000018;
  wpsInFootnoteEndnotePane = $00000019;
  wpsInCommentPane = $0000001A;
  wpsInHeaderFooter = $0000001C;
  wpsAtEndOfRowMarker = $0000001F;
  wpsReferenceOfType = $00000020;
  wpsHeaderFooterType = $00000021;
  wpsInMasterDocument = $00000022;
  wpsInFootnote = $00000023;
  wpsInEndnote = $00000024;
  wpsInWordMail = $00000025;
  wpsInClipboard = $00000026;
  wpsInFirstTextColumnOfPage = $00000027;

// Constants for enum WpsCharacterCase
type
  WpsCharacterCase = TOleEnum;
const
  wpsNextCase = $FFFFFFFF;
  wpsLowerCase = $00000000;
  wpsUpperCase = $00000001;
  wpsTitleWord = $00000002;
  wpsTitleSentence = $00000004;
  wpsToggleCase = $00000005;
  wpsHalfWidth = $00000006;
  wpsFullWidth = $00000007;
  wpsKatakana = $00000008;
  wpsHiragana = $00000009;

// Constants for enum WpsCollapseDirection
type
  WpsCollapseDirection = TOleEnum;
const
  wpsCollapseStart = $00000001;
  wpsCollapseEnd = $00000000;

// Constants for enum WpsUnits
type
  WpsUnits = TOleEnum;
const
  wpsCharacter = $00000001;
  wpsWord = $00000002;
  wpsSentence = $00000003;
  wpsParagraph = $00000004;
  wpsLine = $00000005;
  wpsStory = $00000006;
  wpsScreen = $00000007;
  wpsSection = $00000008;
  wpsColumn = $00000009;
  wpsRow = $0000000A;
  wpsWindow = $0000000B;
  wpsCell = $0000000C;
  wpsCharacterFormatting = $0000000D;
  wpsParagraphFormatting = $0000000E;
  wpsTable = $0000000F;
  wpsItem = $00000010;
  wpsFormatting = $00000014;
  __wpsReverseCharacter = $FFFFFFFF;

// Constants for enum WpsMovementType
type
  WpsMovementType = TOleEnum;
const
  wpsMove = $00000000;
  wpsExtend = $00000001;

// Constants for enum WpsBreakType
type
  WpsBreakType = TOleEnum;
const
  wpsParagraphBreak = $00000001;
  wpsSectionBreakNextPage = $00000002;
  wpsSectionBreakContinuous = $00000003;
  wpsSectionBreakEvenPage = $00000004;
  wpsSectionBreakOddPage = $00000005;
  wpsLineBreak = $00000006;
  wpsPageBreak = $00000007;
  wpsColumnBreak = $00000008;
  wpsLineBreakClearLeft = $00000009;
  wpsLineBreakClearRight = $0000000A;
  wpsTextWrappingBreak = $0000000B;

// Constants for enum WpsGoToItem
type
  WpsGoToItem = TOleEnum;
const
  wpsGoToBookmark = $FFFFFFFF;
  wpsGoToSection = $00000000;
  wpsGoToPage = $00000001;
  wpsGoToTable = $00000002;
  wpsGoToLine = $00000003;
  wpsGoToFootnote = $00000004;
  wpsGoToEndnote = $00000005;
  wpsGoToComment = $00000006;
  wpsGoToField = $00000007;
  wpsGoToGraphic = $00000008;
  wpsGoToObject = $00000009;
  wpsGoToEquation = $0000000A;
  wpsGoToHeading = $0000000B;
  wpsGoToPercent = $0000000C;
  wpsGoToSpellingError = $0000000D;
  wpsGoToGrammaticalError = $0000000E;
  wpsGoToProofreadingError = $0000000F;

// Constants for enum WpsGoToDirection
type
  WpsGoToDirection = TOleEnum;
const
  wpsGoToFirst = $00000004;
  wpsGoToLast = $FFFFFFFF;
  wpsGoToNext = $00000002;
  wpsGoToRelative = $00000002;
  wpsGoToPrevious = $00000003;
  wpsGoToAbsolute = $00000001;

// Constants for enum WpsDateLanguage
type
  WpsDateLanguage = TOleEnum;
const
  wpsDateLanguageBidi = $0000000A;
  wpsateLanguageLatin = $00000409;

// Constants for enum WpsCalendarTypeBi
type
  WpsCalendarTypeBi = TOleEnum;
const
  wpsCalendarTypeBidi = $00000063;
  wpsCalendarTypeGregorian = $00000064;

// Constants for enum WpsCaptionPosition
type
  WpsCaptionPosition = TOleEnum;
const
  wpsCaptionPositionAbove = $00000000;
  wpsCaptionPositionBelow = $00000001;

// Constants for enum WpsEncloseStyle
type
  WpsEncloseStyle = TOleEnum;
const
  wpsEncloseStyleNone = $00000000;
  wpsEncloseStyleSmall = $00000001;
  wpsEncloseStyleLarge = $00000002;

// Constants for enum WpsEnclosureType
type
  WpsEnclosureType = TOleEnum;
const
  wpsEnclosureCircle = $00000000;
  wpsEnclosureSquare = $00000001;
  wpsEnclosureTriangle = $00000002;
  wpsEnclosureDiamond = $00000003;

// Constants for enum WpsPhoneticGuideAlignmentType
type
  WpsPhoneticGuideAlignmentType = TOleEnum;
const
  wpsPhoneticGuideAlignmentCenter = $00000000;
  wpsPhoneticGuideAlignmentZeroOneZero = $00000001;
  wpsPhoneticGuideAlignmentOneTwoOne = $00000002;
  wpsPhoneticGuideAlignmentLeft = $00000003;
  wpsPhoneticGuideAlignmentRight = $00000004;
  wpsPhoneticGuideAlignmentRightVertical = $00000005;

// Constants for enum WpsRevisionType
type
  WpsRevisionType = TOleEnum;
const
  wpsNoRevision = $00000000;
  wpsRevisionInsert = $00000001;
  wpsRevisionDelete = $00000002;
  wpsRevisionProperty = $00000003;
  wpsRevisionParagraphNumber = $00000004;
  wpsRevisionDisplayField = $00000005;
  wpsRevisionReconcile = $00000006;
  wpsRevisionConflict = $00000007;
  wpsRevisionStyle = $00000008;
  wpsRevisionReplace = $00000009;
  wpsRevisionParagraphProperty = $0000000A;
  wpsRevisionTableProperty = $0000000B;
  wpsRevisionSectionProperty = $0000000C;
  wpsRevisionStyleDefinition = $0000000D;

// Constants for enum WpsTCSCConverterDirection
type
  WpsTCSCConverterDirection = TOleEnum;
const
  wpsTCSCConverterDirectionSCTC = $00000000;
  wpsTCSCConverterDirectionTCSC = $00000001;
  wpsTCSCConverterDirectionAuto = $00000002;

// Constants for enum WpsLanguageID
type
  WpsLanguageID = TOleEnum;
const
  wpsLanguageNone = $00000000;
  wpsNoProofing = $00000400;
  wpsAfrikaans = $00000436;
  wpsAlbanian = $0000041C;
  wpsAmharic = $0000045E;
  wpsArabicAlgeria = $00001401;
  wpsArabicBahrain = $00003C01;
  wpsArabicEgypt = $00000C01;
  wpsArabicIraq = $00000801;
  wpsArabicJordan = $00002C01;
  wpsArabicKuwait = $00003401;
  wpsArabicLebanon = $00003001;
  wpsArabicLibya = $00001001;
  wpsArabicMorocco = $00001801;
  wpsArabicOman = $00002001;
  wpsArabicQatar = $00004001;
  wpsArabic = $00000401;
  wpsArabicSyria = $00002801;
  wpsArabicTunisia = $00001C01;
  wpsArabicUAE = $00003801;
  wpsArabicYemen = $00002401;
  wpsArmenian = $0000042B;
  wpsAssamese = $0000044D;
  wpsAzeriCyrillic = $0000082C;
  wpsAzeriLatin = $0000042C;
  wpsBasque = $0000042D;
  wpsByelorussian = $00000423;
  wpsBengali = $00000445;
  wpsBulgarian = $00000402;
  wpsBurmese = $00000455;
  wpsCatalan = $00000403;
  wpsCherokee = $0000045C;
  wpsChineseHongKongSAR = $00000C04;
  wpsChineseMacaoSAR = $00001404;
  wpsSimplifiedChinese = $00000804;
  wpsChineseSingapore = $00001004;
  wpsTraditionalChinese = $00000404;
  wpsCroatian = $0000041A;
  wpsCzech = $00000405;
  wpsDanish = $00000406;
  wpsDivehi = $00000465;
  wpsBelgianDutch = $00000813;
  wpsDutch = $00000413;
  wpsDzongkhaBhutan = $00000851;
  wpsEdo = $00000466;
  wpsEnglishAUS = $00000C09;
  wpsEnglishBelize = $00002809;
  wpsEnglishCanadian = $00001009;
  wpsEnglishCaribbean = $00002409;
  wpsEnglishIreland = $00001809;
  wpsEnglishJamaica = $00002009;
  wpsEnglishNewZealand = $00001409;
  wpsEnglishPhilippines = $00003409;
  wpsEnglishSouthAfrica = $00001C09;
  wpsEnglishTrinidadTobago = $00002C09;
  wpsEnglishUK = $00000809;
  wpsEnglishUS = $00000409;
  wpsEnglishZimbabwe = $00003009;
  wpsEnglishIndonesia = $00003809;
  wpsEstonian = $00000425;
  wpsFaeroese = $00000438;
  wpsFarsi = $00000429;
  wpsFilipino = $00000464;
  wpsFinnish = $0000040B;
  wpsFulfulde = $00000467;
  wpsBelgianFrench = $0000080C;
  wpsFrenchCameroon = $00002C0C;
  wpsFrenchCanadian = $00000C0C;
  wpsFrenchCotedIvoire = $0000300C;
  wpsFrench = $0000040C;
  wpsFrenchLuxembourg = $0000140C;
  wpsFrenchMali = $0000340C;
  wpsFrenchMonaco = $0000180C;
  wpsFrenchReunion = $0000200C;
  wpsFrenchSenegal = $0000280C;
  wpsFrenchMorocco = $0000380C;
  wpsFrenchHaiti = $00003C0C;
  wpsSwissFrench = $0000100C;
  wpsFrenchWestIndies = $00001C0C;
  wpsFrenchZaire = $0000240C;
  wpsFrisianNetherlands = $00000462;
  wpsGaelicIreland = $0000083C;
  wpsGaelicScotland = $0000043C;
  wpsGalician = $00000456;
  wpsGeorgian = $00000437;
  wpsGermanAustria = $00000C07;
  wpsGerman = $00000407;
  wpsGermanLiechtenstein = $00001407;
  wpsGermanLuxembourg = $00001007;
  wpsSwissGerman = $00000807;
  wpsGreek = $00000408;
  wpsGuarani = $00000474;
  wpsGujarati = $00000447;
  wpsHausa = $00000468;
  wpsHawaiian = $00000475;
  wpsHebrew = $0000040D;
  wpsHindi = $00000439;
  wpsHungarian = $0000040E;
  wpsIbibio = $00000469;
  wpsIcelandic = $0000040F;
  wpsIgbo = $00000470;
  wpsIndonesian = $00000421;
  wpsInuktitut = $0000045D;
  wpsItalian = $00000410;
  wpsSwissItalian = $00000810;
  wpsJapanese = $00000411;
  wpsKannada = $0000044B;
  wpsKanuri = $00000471;
  wpsKashmiri = $00000460;
  wpsKazakh = $0000043F;
  wpsKhmer = $00000453;
  wpsKirghiz = $00000440;
  wpsKonkani = $00000457;
  wpsKorean = $00000412;
  wpsKyrgyz = $00000440;
  wpsLao = $00000454;
  wpsLatin = $00000476;
  wpsLatvian = $00000426;
  wpsLithuanian = $00000427;
  wpsMacedonian = $0000042F;
  wpsMalaysian = $0000043E;
  wpsMalayBruneiDarussalam = $0000083E;
  wpsMalayalam = $0000044C;
  wpsMaltese = $0000043A;
  wpsManipuri = $00000458;
  wpsMarathi = $0000044E;
  wpsMongolian = $00000450;
  wpsNepali = $00000461;
  wpsNorwegianBokmol = $00000414;
  wpsNorwegianNynorsk = $00000814;
  wpsOriya = $00000448;
  wpsOromo = $00000472;
  wpsPashto = $00000463;
  wpsPolish = $00000415;
  wpsBrazilianPortuguese = $00000416;
  wpsPortuguese = $00000816;
  wpsPunjabi = $00000446;
  wpsRhaetoRomanic = $00000417;
  wpsRomanianMoldova = $00000818;
  wpsRomanian = $00000418;
  wpsRussianMoldova = $00000819;
  wpsRussian = $00000419;
  wpsSamiLappish = $0000043B;
  wpsSanskrit = $0000044F;
  wpsSerbianCyrillic = $00000C1A;
  wpsSerbianLatin = $0000081A;
  wpsSinhalese = $0000045B;
  wpsSindhi = $00000459;
  wpsSindhiPakistan = $00000859;
  wpsSlovak = $0000041B;
  wpsSlovenian = $00000424;
  wpsSomali = $00000477;
  wpsSorbian = $0000042E;
  wpsSpanishArgentina = $00002C0A;
  wpsSpanishBolivia = $0000400A;
  wpsSpanishChile = $0000340A;
  wpsSpanishColombia = $0000240A;
  wpsSpanishCostaRica = $0000140A;
  wpsSpanishDominicanRepublic = $00001C0A;
  wpsSpanishEcuador = $0000300A;
  wpsSpanishElSalvador = $0000440A;
  wpsSpanishGuatemala = $0000100A;
  wpsSpanishHonduras = $0000480A;
  wpsMexicanSpanish = $0000080A;
  wpsSpanishNicaragua = $00004C0A;
  wpsSpanishPanama = $0000180A;
  wpsSpanishParaguay = $00003C0A;
  wpsSpanishPeru = $0000280A;
  wpsSpanishPuertoRico = $0000500A;
  wpsSpanishModernSort = $00000C0A;
  wpsSpanish = $0000040A;
  wpsSpanishUruguay = $0000380A;
  wpsSpanishVenezuela = $0000200A;
  wpsSesotho = $00000430;
  wpsSutu = $00000430;
  wpsSwahili = $00000441;
  wpsSwedishFinland = $0000081D;
  wpsSwedish = $0000041D;
  wpsSyriac = $0000045A;
  wpsTajik = $00000428;
  wpsTamazight = $0000045F;
  wpsTamazightLatin = $0000085F;
  wpsTamil = $00000449;
  wpsTatar = $00000444;
  wpsTelugu = $0000044A;
  wpsThai = $0000041E;
  wpsTibetan = $00000451;
  wpsTigrignaEthiopic = $00000473;
  wpsTigrignaEritrea = $00000873;
  wpsTsonga = $00000431;
  wpsTswana = $00000432;
  wpsTurkish = $0000041F;
  wpsTurkmen = $00000442;
  wpsUkrainian = $00000422;
  wpsUrdu = $00000420;
  wpsUzbekCyrillic = $00000843;
  wpsUzbekLatin = $00000443;
  wpsVenda = $00000433;
  wpsVietnamese = $0000042A;
  wpsWelsh = $00000452;
  wpsXhosa = $00000434;
  wpsYi = $00000478;
  wpsYiddish = $0000043D;
  wpsYoruba = $0000046A;
  wpsZulu = $00000435;

// Constants for enum WpsTwoLinesInOne
type
  WpsTwoLinesInOne = TOleEnum;
const
  wpsTwoLinesInOneNone = $00000000;
  wpsTwoLinesInOneNoBrackets = $00000001;
  wpsTwoLinesInOneParentheses = $00000002;
  wpsTwoLinesInOneSquareBrackets = $00000003;
  wpsTwoLinesInOneAngleBrackets = $00000004;
  wpsTwoLinesInOneCurlyBrackets = $00000005;

// Constants for enum WpsTextFormFieldType
type
  WpsTextFormFieldType = TOleEnum;
const
  wpsRegularText = $00000000;
  wpsNumberText = $00000001;
  wpsDateText = $00000002;
  wpsCurrentDateText = $00000003;
  wpsCurrentTimeText = $00000004;
  wpsCalculationText = $00000005;

// Constants for enum WpsRelocate
type
  WpsRelocate = TOleEnum;
const
  wpsRelocateUp = $00000000;
  wpsRelocateDown = $00000001;

// Constants for enum WpsDocumentType
type
  WpsDocumentType = TOleEnum;
const
  wpsTypeDocument = $00000000;
  wpsTypeTemplate = $00000001;
  wpsTypeFrameset = $00000002;

// Constants for enum WpsTofFormat
type
  WpsTofFormat = TOleEnum;
const
  wpsTOFTemplate = $00000000;
  wpsTOFClassic = $00000001;
  wpsTOFDistinctive = $00000002;
  wpsTOFCentered = $00000003;
  wpsTOFFormal = $00000004;
  wpsTOFSimple = $00000005;

// Constants for enum WpsMailMergeMainDocType
type
  WpsMailMergeMainDocType = TOleEnum;
const
  wpsNotAMergeDocument = $FFFFFFFF;
  wpsFormLetters = $00000000;
  wpsMailingLabels = $00000001;
  wpsEnvelopes = $00000002;
  wpsCatalog = $00000003;
  wpsEMail = $00000004;
  wpsFax = $00000005;
  wpsDirectory = $00000003;

// Constants for enum WpsMailMergeState
type
  WpsMailMergeState = TOleEnum;
const
  wpsNormalDocument = $00000000;
  wpsMainDocumentOnly = $00000001;
  wpsMainAndDataSource = $00000002;
  wpsMainAndHeader = $00000003;
  wpsMainAndSourceAndHeader = $00000004;
  wpsDataSource = $00000005;

// Constants for enum WpsMailMergeDestination
type
  WpsMailMergeDestination = TOleEnum;
const
  wpsSendToNewDocument = $00000000;
  wpsSendToPrinter = $00000001;
  wpsSendToEmail = $00000002;
  wpsSendToFax = $00000003;
  wpsSendToNewDocuments = $00000004;

// Constants for enum WpsMailMergeDataSource
type
  WpsMailMergeDataSource = TOleEnum;
const
  wpsNoMergeInfo = $FFFFFFFF;
  wpsMergeInfoFromWord = $00000000;
  wpsMergeInfoFromAccessDDE = $00000001;
  wpsMergeInfoFromExcelDDE = $00000002;
  wpsMergeInfoFromMSQueryDDE = $00000003;
  wpsMergeInfoFromODBC = $00000004;
  wpsMergeInfoFromODSO = $00000005;

// Constants for enum WpsMailMergeActiveRecord
type
  WpsMailMergeActiveRecord = TOleEnum;
const
  wpsNoActiveRecord = $FFFFFFFF;
  wpsNextRecord = $FFFFFFFE;
  wpsPreviousRecord = $FFFFFFFD;
  wpsFirstRecord = $FFFFFFFC;
  wpsLastRecord = $FFFFFFFB;
  wpsFirstDataSourceRecord = $FFFFFFFA;
  wpsLastDataSourceRecord = $FFFFFFF9;
  wpsNextDataSourceRecord = $FFFFFFF8;
  wpsPreviousDataSourceRecord = $FFFFFFF7;

// Constants for enum WpsMappedDataFields
type
  WpsMappedDataFields = TOleEnum;
const
  wpsUniqueIdentifier = $00000001;
  wpsCourtesyTitle = $00000002;
  wpsFirstName = $00000003;
  wpsMiddleName = $00000004;
  wpsLastName = $00000005;
  wpsSuffix = $00000006;
  wpsNickname = $00000007;
  wpsJobTitle = $00000008;
  wpsCompany = $00000009;
  wpsAddress1 = $0000000A;
  wpsAddress2 = $0000000B;
  wpsCity = $0000000C;
  wpsState = $0000000D;
  wpsPostalCode = $0000000E;
  wpsCountryRegion = $0000000F;
  wpsBusinessPhone = $00000010;
  wpsBusinessFax = $00000011;
  wpsHomePhone = $00000012;
  wpsHomeFax = $00000013;
  wpsEmailAddress = $00000014;
  wpsWebPageURL = $00000015;
  wpsSpouseCourtesyTitle = $00000016;
  wpsSpouseFirstName = $00000017;
  wpsSpouseMiddleName = $00000018;
  wpsSpouseLastName = $00000019;
  wpsSpouseNickname = $0000001A;
  wpsRubyFirstName = $0000001B;
  wpsRubyLastName = $0000001C;
  wpsAddress3 = $0000001D;
  wpsDepartment = $0000001E;

// Constants for enum WpsMailMergeComparison
type
  WpsMailMergeComparison = TOleEnum;
const
  wpsMergeIfEqual = $00000000;
  wpsMergeIfNotEqual = $00000001;
  wpsMergeIfLessThan = $00000002;
  wpsMergeIfGreaterThan = $00000003;
  wpsMergeIfLessThanOrEqual = $00000004;
  wpsMergeIfGreaterThanOrEqual = $00000005;
  wpsMergeIfIsBlank = $00000006;
  wpsMergeIfIsNotBlank = $00000007;

// Constants for enum WpsMailMergeMailFormat
type
  WpsMailMergeMailFormat = TOleEnum;
const
  wpsMailFormatPlainText = $00000000;
  wpsMailFormatHTML = $00000001;

// Constants for enum WpsTocFormat
type
  WpsTocFormat = TOleEnum;
const
  wpsTOCTemplate = $00000000;
  wpsTOCClassic = $00000001;
  wpsTOCDistinctive = $00000002;
  wpsTOCFancy = $00000003;
  wpsTOCModern = $00000004;
  wpsTOCFormal = $00000005;
  wpsTOCSimple = $00000006;

// Constants for enum WpsSelectionType
type
  WpsSelectionType = TOleEnum;
const
  wpsNoSelection = $00000000;
  wpsSelectionIP = $00000001;
  wpsSelectionNormal = $00000002;
  wpsSelectionFrame = $00000003;
  wpsSelectionColumn = $00000004;
  wpsSelectionRow = $00000005;
  wpsSelectionBlock = $00000006;
  wpsSelectionInlineShape = $00000007;
  wpsSelectionShape = $00000008;

// Constants for enum WpsSelectionFlags
type
  WpsSelectionFlags = TOleEnum;
const
  wpsSelStartActive = $00000001;
  wpsSelAtEOL = $00000002;
  wpsSelOvertype = $00000004;
  wpsSelActive = $00000008;
  wpsSelReplace = $00000010;

// Constants for enum WpsOLEPlacement
type
  WpsOLEPlacement = TOleEnum;
const
  wpsInLine = $00000000;
  wpsFloatOverText = $00000001;

// Constants for enum WpsFindWrap
type
  WpsFindWrap = TOleEnum;
const
  wpsFindStop = $00000000;
  wpsFindContinue = $00000001;
  wpsFindAsk = $00000002;
  wpsFindHighlightMainDoc = $00000003;
  wpsFindHighlightCurSelection = $00000004;

// Constants for enum WpsFindScope
type
  WpsFindScope = TOleEnum;
const
  wpsFindScope_WholeDocument = $00000000;
  wpsFindScope_Selection = $00000001;
  wpsFindScope_MainText = $00000002;
  wpsFindScope_HeaderFooters = $00000003;
  wpsFindScope_Footnotes = $00000004;
  wpsFindScope_Endnotes = $00000005;
  wpsFindScope_TextFrames = $00000006;
  wpsFindScope_HeaderFooterTextFrames = $00000007;
  wpsFindScope_Comments = $00000008;

// Constants for enum WpsListGalleryType
type
  WpsListGalleryType = TOleEnum;
const
  wpsBulletGallery = $00000001;
  wpsNumberGallery = $00000002;
  wpsOutlineNumberGallery = $00000003;
  wpsCustomGallery = $00000004;
  wpsTemGallery = $00000005;

// Constants for enum WpsRecoveryType
type
  WpsRecoveryType = TOleEnum;
const
  wpsPasteDefault = $00000000;
  wpsSingleCellText = $00000005;
  wpsSingleCellTable = $00000006;
  wpsListContinueNumbering = $00000007;
  wpsListRestartNumbering = $00000008;
  wpsTableInsertAsRows = $0000000B;
  wpsTableAppendTable = $0000000A;
  wpsTableOriginalFormatting = $0000000C;
  wpsChartPicture = $0000000D;
  wpsChart = $0000000E;
  wpsChartLinked = $0000000F;
  wpsFormatOriginalFormatting = $00000010;
  wpsFormatSurroundingFormattingWithEmphasis = $00000014;
  wpsFormatPlainText = $00000016;
  wpsTableOverwriteCells = $00000017;
  wpsListCombineWithExistingList = $00000018;
  wpsListDontMerge = $00000019;

// Constants for enum WpsReferenceKind
type
  WpsReferenceKind = TOleEnum;
const
  wpsContentText = $FFFFFFFF;
  wpsNumberRelativeContext = $FFFFFFFE;
  wpsNumberNoContext = $FFFFFFFD;
  wpsNumberFullContext = $FFFFFFFC;
  wpsEntireCaption = $00000002;
  wpsOnlyLabelAndNumber = $00000003;
  wpsOnlyCaptionText = $00000004;
  wpsFootnoteNumber = $00000005;
  wpsEndnoteNumber = $00000006;
  wpsPageNumber = $00000007;
  wpsPosition = $0000000F;
  wpsFootnoteNumberFormatted = $00000010;
  wpsEndnoteNumberFormatted = $00000011;

// Constants for enum WpsViewType
type
  WpsViewType = TOleEnum;
const
  wpsNormalView = $00000001;
  wpsOutlineView = $00000002;
  wpsPrintView = $00000003;
  wpsPrintPreview = $00000004;
  wpsMasterView = $00000005;
  wpsWebView = $00000006;

// Constants for enum WpsPageFit
type
  WpsPageFit = TOleEnum;
const
  wpsPageFitNone = $00000000;
  wpsPageFitFullPage = $00000001;
  wpsPageFitBestFit = $00000002;
  wpsPageFitTextFit = $00000003;

// Constants for enum WpsFieldShading
type
  WpsFieldShading = TOleEnum;
const
  wpsFieldShadingNever = $00000000;
  wpsFieldShadingAlways = $00000001;
  wpsFieldShadingWhenSelected = $00000002;

// Constants for enum WpsSeekView
type
  WpsSeekView = TOleEnum;
const
  wpsSeekMainDocument = $00000000;
  wpsSeekPrimaryHeader = $00000001;
  wpsSeekFirstPageHeader = $00000002;
  wpsSeekEvenPagesHeader = $00000003;
  wpsSeekPrimaryFooter = $00000004;
  wpsSeekFirstPageFooter = $00000005;
  wpsSeekEvenPagesFooter = $00000006;
  wpsSeekFootnotes = $00000007;
  wpsSeekEndnotes = $00000008;
  wpsSeekCurrentPageHeader = $00000009;
  wpsSeekCurrentPageFooter = $0000000A;

// Constants for enum WpsRevisionsView
type
  WpsRevisionsView = TOleEnum;
const
  wpsRevisionsViewFinal = $00000000;
  wpsRevisionsViewOriginal = $00000001;

// Constants for enum WpsRevisionsMode
type
  WpsRevisionsMode = TOleEnum;
const
  wpsBalloonRevisions = $00000000;
  wpsInLineRevisions = $00000001;
  wpsMixedRevisions = $00000002;

// Constants for enum WpsRevisionsBalloonWidthType
type
  WpsRevisionsBalloonWidthType = TOleEnum;
const
  wpsBalloonWidthPercent = $00000000;
  wpsBalloonWidthPoints = $00000001;

// Constants for enum WpsRevisionsBalloonMargin
type
  WpsRevisionsBalloonMargin = TOleEnum;
const
  wpsLeftMargin = $00000000;
  wpsRightMargin = $00000001;

// Constants for enum WpsPagesOrientation
type
  WpsPagesOrientation = TOleEnum;
const
  wpsPagesOrientation_Vertical = $00000000;
  wpsPagesOrientation_Horizontal = $00000001;

// Constants for enum WpsWindowState
type
  WpsWindowState = TOleEnum;
const
  wpsWindowStateNormal = $00000000;
  wpsWindowStateMaximize = $00000001;
  wpsWindowStateMinimize = $00000002;

// Constants for enum WpsWindowType
type
  WpsWindowType = TOleEnum;
const
  wpsWindowDocument = $00000000;
  wpsWindowTemplate = $00000001;

// Constants for enum WpsSaveOptions
type
  WpsSaveOptions = TOleEnum;
const
  wpsDoNotSaveChanges = $00000000;
  wpsSaveChanges = $FFFFFFFF;
  wpsPromptToSaveChanges = $FFFFFFFE;

// Constants for enum WpsPrintOutRange
type
  WpsPrintOutRange = TOleEnum;
const
  wpsPrintAllDocument = $00000000;
  wpsPrintSelection = $00000001;
  wpsPrintCurrentPage = $00000002;
  wpsPrintFromTo = $00000003;
  wpsPrintRangeOfPages = $00000004;

// Constants for enum WpsPrintOutItem
type
  WpsPrintOutItem = TOleEnum;
const
  wpsPrintDocumentContent = $00000000;
  wpsPrintProperties = $00000001;
  wpsPrintComments = $00000002;
  wpsPrintMarkup = $00000002;
  wpsPrintStyles = $00000003;
  wpsPrintAutoTextEntries = $00000004;
  wpsPrintKeyAssignments = $00000005;
  wpsPrintEnvelope = $00000006;
  wpsPrintDocumentWithMarkup = $00000007;

// Constants for enum WpsPrintOutPages
type
  WpsPrintOutPages = TOleEnum;
const
  wpsPrintAllPages = $00000000;
  wpsPrintOddPagesOnly = $00000001;
  wpsPrintEvenPagesOnly = $00000002;

// Constants for enum WpsPaperOrder
type
  WpsPaperOrder = TOleEnum;
const
  wpsPrinterOverThenDown = $00000000;
  wpsPrinterDownThenOver = $00000001;
  wpsPrinterRepeat = $00000002;

// Constants for enum WpsDocumentKind
type
  WpsDocumentKind = TOleEnum;
const
  wpsDocumentNotSpecified = $00000000;
  wpsDocumentLetter = $00000001;
  wpsDocumentEmail = $00000002;

// Constants for enum WpsJustificationMode
type
  WpsJustificationMode = TOleEnum;
const
  wpsJustificationModeExpand = $00000000;
  wpsJustificationModeCompress = $00000001;
  wpsJustificationModeCompressKana = $00000002;

// Constants for enum WpsFarEastLineBreakLevel
type
  WpsFarEastLineBreakLevel = TOleEnum;
const
  wpsFarEastLineBreakLevelNormal = $00000000;
  wpsFarEastLineBreakLevelStrict = $00000001;
  wpsFarEastLineBreakLevelCustom = $00000002;

// Constants for enum WpsShowFilter
type
  WpsShowFilter = TOleEnum;
const
  wpsShowFilterStylesAvailable = $00000000;
  wpsShowFilterStylesInUse = $00000001;
  wpsShowFilterStylesAll = $00000002;
  wpsShowFilterFormattingInUse = $00000003;
  wpsShowFilterFormattingAvailable = $00000004;

// Constants for enum WpsProtectionType
type
  WpsProtectionType = TOleEnum;
const
  wpsNoProtection = $FFFFFFFF;
  wpsAllowOnlyRevisions = $00000000;
  wpsAllowOnlyComments = $00000001;
  wpsAllowOnlyFormFields = $00000002;
  wpsAllowOnlyReading = $00000003;

// Constants for enum WpsEncoding
type
  WpsEncoding = TOleEnum;
const
  msoEncodingThai = $0000036A;
  msoEncodingJapaneseShiftJIS = $000003A4;
  msoEncodingSimplifiedChineseGBK = $000003A8;
  msoEncodingKorean = $000003B5;
  msoEncodingTraditionalChineseBig5 = $000003B6;
  msoEncodingUnicodeLittleEndian = $000004B0;
  msoEncodingUnicodeBigEndian = $000004B1;
  msoEncodingCentralEuropean = $000004E2;
  msoEncodingCyrillic = $000004E3;
  msoEncodingWestern = $000004E4;
  msoEncodingGreek = $000004E5;
  msoEncodingTurkish = $000004E6;
  msoEncodingHebrew = $000004E7;
  msoEncodingArabic = $000004E8;
  msoEncodingBaltic = $000004E9;
  msoEncodingVietnamese = $000004EA;
  msoEncodingAutoDetect = $0000C351;
  msoEncodingJapaneseAutoDetect = $0000C6F4;
  msoEncodingSimplifiedChineseAutoDetect = $0000C6F8;
  msoEncodingKoreanAutoDetect = $0000C705;
  msoEncodingTraditionalChineseAutoDetect = $0000C706;
  msoEncodingCyrillicAutoDetect = $0000C833;
  msoEncodingGreekAutoDetect = $0000C835;
  msoEncodingArabicAutoDetect = $0000C838;
  msoEncodingISO88591Latin1 = $00006FAF;
  msoEncodingISO88592CentralEurope = $00006FB0;
  msoEncodingISO88593Latin3 = $00006FB1;
  msoEncodingISO88594Baltic = $00006FB2;
  msoEncodingISO88595Cyrillic = $00006FB3;
  msoEncodingISO88596Arabic = $00006FB4;
  msoEncodingISO88597Greek = $00006FB5;
  msoEncodingISO88598Hebrew = $00006FB6;
  msoEncodingISO88599Turkish = $00006FB7;
  msoEncodingISO885915Latin9 = $00006FBD;
  msoEncodingISO2022JPNoHalfwidthKatakana = $0000C42C;
  msoEncodingISO2022JPJISX02021984 = $0000C42D;
  msoEncodingISO2022JPJISX02011989 = $0000C42E;
  msoEncodingISO2022KR = $0000C431;
  msoEncodingISO2022CNTraditionalChinese = $0000C433;
  msoEncodingISO2022CNSimplifiedChinese = $0000C435;
  msoEncodingMacRoman = $00002710;
  msoEncodingMacJapanese = $00002711;
  msoEncodingMacTraditionalChineseBig5 = $00002712;
  msoEncodingMacKorean = $00002713;
  msoEncodingMacArabic = $00002714;
  msoEncodingMacHebrew = $00002715;
  msoEncodingMacGreek1 = $00002716;
  msoEncodingMacCyrillic = $00002717;
  msoEncodingMacSimplifiedChineseGB2312 = $00002718;
  msoEncodingMacRomania = $0000271A;
  msoEncodingMacUkraine = $00002721;
  msoEncodingMacLatin2 = $0000272D;
  msoEncodingMacIcelandic = $0000275F;
  msoEncodingMacTurkish = $00002761;
  msoEncodingMacCroatia = $00002762;
  msoEncodingEBCDICUSCanada = $00000025;
  msoEncodingEBCDICInternational = $000001F4;
  msoEncodingEBCDICMultilingualROECELatin2 = $00000366;
  msoEncodingEBCDICGreekModern = $0000036B;
  msoEncodingEBCDICTurkishLatin5 = $00000402;
  msoEncodingEBCDICGermany = $00004F31;
  msoEncodingEBCDICDenmarkNorway = $00004F35;
  msoEncodingEBCDICFinlandSweden = $00004F36;
  msoEncodingEBCDICItaly = $00004F38;
  msoEncodingEBCDICLatinAmericaSpain = $00004F3C;
  msoEncodingEBCDICUnitedKingdom = $00004F3D;
  msoEncodingEBCDICJapaneseKatakanaExtended = $00004F42;
  msoEncodingEBCDICFrance = $00004F49;
  msoEncodingEBCDICArabic = $00004FC4;
  msoEncodingEBCDICGreek = $00004FC7;
  msoEncodingEBCDICHebrew = $00004FC8;
  msoEncodingEBCDICKoreanExtended = $00005161;
  msoEncodingEBCDICThai = $00005166;
  msoEncodingEBCDICIcelandic = $00005187;
  msoEncodingEBCDICTurkish = $000051A9;
  msoEncodingEBCDICRussian = $00005190;
  msoEncodingEBCDICSerbianBulgarian = $00005221;
  msoEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = $0000C6F2;
  msoEncodingEBCDICUSCanadaAndJapanese = $0000C6F3;
  msoEncodingEBCDICKoreanExtendedAndKorean = $0000C6F5;
  msoEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = $0000C6F7;
  msoEncodingEBCDICUSCanadaAndTraditionalChinese = $0000C6F9;
  msoEncodingEBCDICJapaneseLatinExtendedAndJapanese = $0000C6FB;
  msoEncodingOEMUnitedStates = $000001B5;
  msoEncodingOEMGreek437G = $000002E1;
  msoEncodingOEMBaltic = $00000307;
  msoEncodingOEMMultilingualLatinI = $00000352;
  msoEncodingOEMMultilingualLatinII = $00000354;
  msoEncodingOEMCyrillic = $00000357;
  msoEncodingOEMTurkish = $00000359;
  msoEncodingOEMPortuguese = $0000035C;
  msoEncodingOEMIcelandic = $0000035D;
  msoEncodingOEMHebrew = $0000035E;
  msoEncodingOEMCanadianFrench = $0000035F;
  msoEncodingOEMArabic = $00000360;
  msoEncodingOEMNordic = $00000361;
  msoEncodingOEMCyrillicII = $00000362;
  msoEncodingOEMModernGreek = $00000365;
  msoEncodingEUCJapanese = $0000CADC;
  msoEncodingEUCChineseSimplifiedChinese = $0000CAE0;
  msoEncodingEUCKorean = $0000CAED;
  msoEncodingEUCTaiwaneseTraditionalChinese = $0000CAEE;
  msoEncodingISCIIDevanagari = $0000DEAA;
  msoEncodingISCIIBengali = $0000DEAB;
  msoEncodingISCIITamil = $0000DEAC;
  msoEncodingISCIITelugu = $0000DEAD;
  msoEncodingISCIIAssamese = $0000DEAE;
  msoEncodingISCIIOriya = $0000DEAF;
  msoEncodingISCIIKannada = $0000DEB0;
  msoEncodingISCIIMalayalam = $0000DEB1;
  msoEncodingISCIIGujarati = $0000DEB2;
  msoEncodingISCIIPunjabi = $0000DEB3;
  msoEncodingArabicASMO = $000002C4;
  msoEncodingArabicTransparentASMO = $000002D0;
  msoEncodingKoreanJohab = $00000551;
  msoEncodingTaiwanCNS = $00004E20;
  msoEncodingTaiwanTCA = $00004E21;
  msoEncodingTaiwanEten = $00004E22;
  msoEncodingTaiwanIBM5550 = $00004E23;
  msoEncodingTaiwanTeleText = $00004E24;
  msoEncodingTaiwanWang = $00004E25;
  msoEncodingIA5IRV = $00004E89;
  msoEncodingIA5German = $00004E8A;
  msoEncodingIA5Swedish = $00004E8B;
  msoEncodingIA5Norwegian = $00004E8C;
  msoEncodingUSASCII = $00004E9F;
  msoEncodingT61 = $00004F25;
  msoEncodingISO6937NonSpacingAccent = $00004F2D;
  msoEncodingKOI8R = $00005182;
  msoEncodingExtAlphaLowercase = $00005223;
  msoEncodingKOI8U = $0000556A;
  msoEncodingEuropa3 = $00007149;
  msoEncodingHZGBSimplifiedChinese = $0000CEC8;
  msoEncodingUTF7 = $0000FDE8;
  msoEncodingUTF8 = $0000FDE9;

// Constants for enum WpsLineEndingType
type
  WpsLineEndingType = TOleEnum;
const
  wpsCRLF = $00000000;
  wpsCROnly = $00000001;
  wpsLFOnly = $00000002;
  wpsLFCR = $00000003;
  wpsLSPS = $00000004;

// Constants for enum WpsPdfStyleLevel
type
  WpsPdfStyleLevel = TOleEnum;
const
  wpsPdfDontExport = $FFFFFFFE;
  wpsPdfFollowNearest = $FFFFFFFF;
  wpsPdfRootBookmark = $00000000;
  wpsPdf1stBookmark = $00000000;
  wpsPdf2ndBookmark = $00000001;
  wpsPdf3rdBookmark = $00000002;
  wpsPdf4thBookmark = $00000003;
  wpsPdf5thBookmark = $00000004;
  wpsPdf6thBookmark = $00000005;
  wpsPdf7thBookmark = $00000006;
  wpsPdf8thBookmark = $00000007;
  wpsPdf9thBookmark = $00000008;
  wpsPdfLeafBookmark = $00000009;

// Constants for enum WpsCompatibility
type
  WpsCompatibility = TOleEnum;
const
  wpsNoTabHangIndent = $00000001;
  wpsNoSpaceRaiseLower = $00000002;
  wpsPrintColBlack = $00000003;
  wpsWrapTrailSpaces = $00000004;
  wpsNoColumnBalance = $00000005;
  wpsConvMailMergeEsc = $00000006;
  wpsSuppressSpBfAfterPgBrk = $00000007;
  wpsSuppressTopSpacing = $00000008;
  wpsOrigWordTableRules = $00000009;
  wpsTransparentMetafiles = $0000000A;
  wpsShowBreaksInFrames = $0000000B;
  wpsSwapBordersFacingPages = $0000000C;
  wpsLeaveBackslashAlone = $0000000D;
  wpsExpandShiftReturn = $0000000E;
  wpsDontULTrailSpace = $0000000F;
  wpsDontBalanceSingleByteDoubleByteWidth = $00000010;
  wpsSuppressTopSpacingMac5 = $00000011;
  wpsSpacingInWholePoints = $00000012;
  wpsPrintBodyTextBeforeHeader = $00000013;
  wpsNoLeading = $00000014;
  wpsNoSpaceForUL = $00000015;
  wpsMWSmallCaps = $00000016;
  wpsNoExtraLineSpacing = $00000017;
  wpsTruncateFontHeight = $00000018;
  wpsSubFontBySize = $00000019;
  wpsUsePrinterMetrics = $0000001A;
  wpsWW6BorderRules = $0000001B;
  wpsExactOnTop = $0000001C;
  wpsSuppressBottomSpacing = $0000001D;
  wpsWPSpaceWidth = $0000001E;
  wpsWPJustification = $0000001F;
  wpsLineWrapLikeWord6 = $00000020;
  wpsShapeLayoutLikeWW8 = $00000021;
  wpsFootnoteLayoutLikeWW8 = $00000022;
  wpsDontUseHTMLParagraphAutoSpacing = $00000023;
  wpsDontAdjustLineHeightInTable = $00000024;
  wpsForgetLastTabAlignment = $00000025;
  wpsAutospaceLikeWW7 = $00000026;
  wpsAlignTablesRowByRow = $00000027;
  wpsLayoutRawTableWidth = $00000028;
  wpsLayoutTableRowsApart = $00000029;
  wpsUseWord97LineBreakingRules = $0000002A;
  wpsDontBreakWrappedTables = $0000002B;
  wpsDontSnapTextToGridInTableWithObjects = $0000002C;
  wpsSelectFieldWithFirstOrLastCharacter = $0000002D;
  wpsApplyBreakingRules = $0000002E;
  wpsDontWrapTextWithPunctuation = $0000002F;
  wpsDontUseAsianBreakRulesInGrid = $00000030;

// Constants for enum WpsOrganizerObject
type
  WpsOrganizerObject = TOleEnum;
const
  wpsOrganizerObjectStyles = $00000000;
  wpsOrganizerObjectAutoText = $00000001;
  wpsOrganizerObjectCommandBars = $00000002;
  wpsOrganizerObjectProjectItems = $00000003;

// Constants for enum WpsCaptionLabelID
type
  WpsCaptionLabelID = TOleEnum;
const
  wpsCaptionFigure = $FFFFFFFF;
  wpsCaptionTable = $FFFFFFFE;
  wpsCaptionEquation = $FFFFFFFD;

// Constants for enum WpsCaptionNumberStyle
type
  WpsCaptionNumberStyle = TOleEnum;
const
  wpsCaptionNumberStyleArabic = $00000000;
  wpsCaptionNumberStyleUppercaseRoman = $00000001;
  wpsCaptionNumberStyleLowercaseRoman = $00000002;
  wpsCaptionNumberStyleUppercaseLetter = $00000003;
  wpsCaptionNumberStyleLowercaseLetter = $00000004;
  wpsCaptionNumberStyleArabicFullWidth = $0000000E;
  wpsCaptionNumberStyleKanji = $0000000A;
  wpsCaptionNumberStyleKanjiDigit = $0000000B;
  wpsCaptionNumberStyleKanjiTraditional = $00000010;
  wpsCaptionNumberStyleNumberInCircle = $00000012;
  wpsCaptionNumberStyleGanada = $00000018;
  wpsCaptionNumberStyleChosung = $00000019;
  wpsCaptionNumberStyleZodiac1 = $0000001E;
  wpsCaptionNumberStyleZodiac2 = $0000001F;
  wpsCaptionNumberStyleHanjaRead = $00000029;
  wpsCaptionNumberStyleHanjaReadDigit = $0000002A;
  wpsCaptionNumberStyleTradChinNum2 = $00000022;
  wpsCaptionNumberStyleTradChinNum3 = $00000023;
  wpsCaptionNumberStyleSimpChinNum2 = $00000026;
  wpsCaptionNumberStyleSimpChinNum3 = $00000027;
  wpsCaptionNumberStyleHebrewLetter1 = $0000002D;
  wpsCaptionNumberStyleArabicLetter1 = $0000002E;
  wpsCaptionNumberStyleHebrewLetter2 = $0000002F;
  wpsCaptionNumberStyleArabicLetter2 = $00000030;
  wpsCaptionNumberStyleHindiLetter1 = $00000031;
  wpsCaptionNumberStyleHindiLetter2 = $00000032;
  wpsCaptionNumberStyleHindiArabic = $00000033;
  wpsCaptionNumberStyleHindiCardinalText = $00000034;
  wpsCaptionNumberStyleThaiLetter = $00000035;
  wpsCaptionNumberStyleThaiArabic = $00000036;
  wpsCaptionNumberStyleThaiCardinalText = $00000037;
  wpsCaptionNumberStyleVietCardinalText = $00000038;

// Constants for enum WpsMeasurementUnits
type
  WpsMeasurementUnits = TOleEnum;
const
  wpsInches = $00000000;
  wpsCentimeters = $00000001;
  wpsMillimeters = $00000002;
  wpsPoints = $00000003;
  wpsPicas = $00000004;

// Constants for enum WpsDefaultFilePath
type
  WpsDefaultFilePath = TOleEnum;
const
  wpsDocumentsPath = $00000000;
  wpsPicturesPath = $00000001;
  wpsUserTemplatesPath = $00000002;
  wpsWorkgroupTemplatesPath = $00000003;
  wpsUserOptionsPath = $00000004;
  wpsAutoRecoverPath = $00000005;
  wpsToolsPath = $00000006;
  wpsTutorialPath = $00000007;
  wpsStartupPath = $00000008;
  wpsProgramPath = $00000009;
  wpsGraphicsFiltersPath = $0000000A;
  wpsTextConvertersPath = $0000000B;
  wpsProofingToolsPath = $0000000C;
  wpsTempFilePath = $0000000D;
  wpsCurrentFolderPath = $0000000E;
  wpsStyleGalleryPath = $0000000F;
  wpsBorderArtPath = $00000013;
  wpsDefaultTemplateFile = $000000C8;

// Constants for enum WpsOpenFormat
type
  WpsOpenFormat = TOleEnum;
const
  wpsOpenFormatAuto = $00000000;
  wpsOpenFormatDocument = $00000001;
  wpsOpenFormatTemplate = $00000002;
  wpsOpenFormatRTF = $00000003;
  wpsOpenFormatText = $00000004;
  wpsOpenFormatUnicodeText = $00000005;
  wpsOpenFormatEncodedText = $00000005;
  wpsOpenFormatAllWord = $00000006;
  wpsOpenFormatWebPages = $00000007;

// Constants for enum WpsPrintHiddenTextMode
type
  WpsPrintHiddenTextMode = TOleEnum;
const
  wpsDontPrintHiddenText = $00000000;
  wpsPrintHiddenText = $00000001;
  wpsPrintSpaceHiddenText = $00000002;

// Constants for enum WpsInsertedTextMark
type
  WpsInsertedTextMark = TOleEnum;
const
  wpsInsertedTextMarkNone = $00000000;
  wpsInsertedTextMarkColorOnly = $00000001;
  wpsInsertedTextMarkBold = $00000002;
  wpsInsertedTextMarkItalic = $00000003;
  wpsInsertedTextMarkUnderline = $00000004;
  wpsInsertedTextMarkDoubleUnderline = $00000005;
  wpsInsertedTextMarkStrikeThrough = $00000006;

// Constants for enum WpsDeletedTextMark
type
  WpsDeletedTextMark = TOleEnum;
const
  wpsDeletedTextMarkNone = $00000000;
  wpsDeletedTextMarkColorOnly = $00000001;
  wpsDeletedTextMarkBold = $00000002;
  wpsDeletedTextMarkItalic = $00000003;
  wpsDeletedTextMarkUnderline = $00000004;
  wpsDeletedTextMarkDoubleUnderline = $00000005;
  wpsDeletedTextMarkStrikeThrough = $00000006;
  wpsDeletedTextMarkHidden = $00000007;
  wpsDeletedTextMarkCaret = $00000008;
  wpsDeletedTextMarkPound = $00000009;

// Constants for enum WpsRevisedLinesMark
type
  WpsRevisedLinesMark = TOleEnum;
const
  wpsRevisedLinesMarkNone = $00000000;
  wpsRevisedLinesMarkLeftBorder = $00000001;
  wpsRevisedLinesMarkRightBorder = $00000002;
  wpsRevisedLinesMarkOutsideBorder = $00000003;

// Constants for enum WpsRevisedPropertiesMark
type
  WpsRevisedPropertiesMark = TOleEnum;
const
  wpsRevisedPropertiesMarkNone = $00000000;
  wpsRevisedPropertiesMarkBold = $00000001;
  wpsRevisedPropertiesMarkItalic = $00000002;
  wpsRevisedPropertiesMarkUnderline = $00000003;
  wpsRevisedPropertiesMarkDoubleUnderline = $00000004;
  wpsRevisedPropertiesMarkColorOnly = $00000005;
  wpsRevisedPropertiesMarkStrikeThrough = $00000006;

// Constants for enum WpsRevisionsBalloonPrintOrientation
type
  WpsRevisionsBalloonPrintOrientation = TOleEnum;
const
  wpsBalloonPrintOrientationAuto = $00000000;
  wpsBalloonPrintOrientationPreserve = $00000001;
  wpsBalloonPrintOrientationForceLandscape = $00000002;

// Constants for enum WpsRevisionsBalloonTitle
type
  WpsRevisionsBalloonTitle = TOleEnum;
const
  wpsRevisionsBalloonTitleSimple = $00000000;
  wpsRevisionsBalloonTitleDetail = $00000001;
  wpsRevisionsBalloonTitleFullName = $00000002;

// Constants for enum WpsDocumentTabPosition
type
  WpsDocumentTabPosition = TOleEnum;
const
  wpsDocumentTabTop = $00000000;
  wpsDocumentTabBottom = $00000001;

// Constants for enum WpsDocumentViewDirection
type
  WpsDocumentViewDirection = TOleEnum;
const
  wpsDocumentViewLtr = $00000000;
  wpsDocumentViewRtl = $00000001;

// Constants for enum WpsTemplateType
type
  WpsTemplateType = TOleEnum;
const
  wpsNormalTemplate = $00000000;
  wpsGlobalTemplate = $00000001;
  wpsAttachedTemplate = $00000002;

// Constants for enum WpsKeyCategory
type
  WpsKeyCategory = TOleEnum;
const
  wpsKeyCategoryNil = $FFFFFFFF;
  wpsKeyCategoryDisable = $00000000;
  wpsKeyCategoryCommand = $00000001;
  wpsKeyCategoryMacro = $00000002;
  wpsKeyCategoryFont = $00000003;
  wpsKeyCategoryAutoText = $00000004;
  wpsKeyCategoryStyle = $00000005;
  wpsKeyCategorySymbol = $00000006;
  wpsKeyCategoryPrefix = $00000007;

// Constants for enum WpsKey
type
  WpsKey = TOleEnum;
const
  wpsNoKey = $000000FF;
  wpsKeyShift = $00000100;
  wpsKeyControl = $00000200;
  wpsKeyCommand = $00000200;
  wpsKeyAlt = $00000400;
  wpsKeyOption = $00000400;
  wpsKeyA = $00000041;
  wpsKeyB = $00000042;
  wpsKeyC = $00000043;
  wpsKeyD = $00000044;
  wpsKeyE = $00000045;
  wpsKeyF = $00000046;
  wpsKeyG = $00000047;
  wpsKeyH = $00000048;
  wpsKeyI = $00000049;
  wpsKeyJ = $0000004A;
  wpsKeyK = $0000004B;
  wpsKeyL = $0000004C;
  wpsKeyM = $0000004D;
  wpsKeyN = $0000004E;
  wpsKeyO = $0000004F;
  wpsKeyP = $00000050;
  wpsKeyQ = $00000051;
  wpsKeyR = $00000052;
  wpsKeyS = $00000053;
  wpsKeyT = $00000054;
  wpsKeyU = $00000055;
  wpsKeyV = $00000056;
  wpsKeyW = $00000057;
  wpsKeyX = $00000058;
  wpsKeyY = $00000059;
  wpsKeyZ = $0000005A;
  wpsKey0 = $00000030;
  wpsKey1 = $00000031;
  wpsKey2 = $00000032;
  wpsKey3 = $00000033;
  wpsKey4 = $00000034;
  wpsKey5 = $00000035;
  wpsKey6 = $00000036;
  wpsKey7 = $00000037;
  wpsKey8 = $00000038;
  wpsKey9 = $00000039;
  wpsKeyBackspace = $00000008;
  wpsKeyTab = $00000009;
  wpsKeyNumeric5Special = $0000000C;
  wpsKeyReturn = $0000000D;
  wpsKeyPause = $00000013;
  wpsKeyEsc = $0000001B;
  wpsKeySpacebar = $00000020;
  wpsKeyPageUp = $00000021;
  wpsKeyPageDown = $00000022;
  wpsKeyEnd = $00000023;
  wpsKeyHome = $00000024;
  wpsKeyInsert = $0000002D;
  wpsKeyDelete = $0000002E;
  wpsKeyNumeric0 = $00000060;
  wpsKeyNumeric1 = $00000061;
  wpsKeyNumeric2 = $00000062;
  wpsKeyNumeric3 = $00000063;
  wpsKeyNumeric4 = $00000064;
  wpsKeyNumeric5 = $00000065;
  wpsKeyNumeric6 = $00000066;
  wpsKeyNumeric7 = $00000067;
  wpsKeyNumeric8 = $00000068;
  wpsKeyNumeric9 = $00000069;
  wpsKeyNumericMultiply = $0000006A;
  wpsKeyNumericAdd = $0000006B;
  wpsKeyNumericSubtract = $0000006D;
  wpsKeyNumericDecimal = $0000006E;
  wpsKeyNumericDivide = $0000006F;
  wpsKeyF1 = $00000070;
  wpsKeyF2 = $00000071;
  wpsKeyF3 = $00000072;
  wpsKeyF4 = $00000073;
  wpsKeyF5 = $00000074;
  wpsKeyF6 = $00000075;
  wpsKeyF7 = $00000076;
  wpsKeyF8 = $00000077;
  wpsKeyF9 = $00000078;
  wpsKeyF10 = $00000079;
  wpsKeyF11 = $0000007A;
  wpsKeyF12 = $0000007B;
  wpsKeyF13 = $0000007C;
  wpsKeyF14 = $0000007D;
  wpsKeyF15 = $0000007E;
  wpsKeyF16 = $0000007F;
  wpsKeyScrollLock = $00000091;
  wpsKeySemiColon = $000000BA;
  wpsKeyEquals = $000000BB;
  wpsKeyComma = $000000BC;
  wpsKeyHyphen = $000000BD;
  wpsKeyPeriod = $000000BE;
  wpsKeySlash = $000000BF;
  wpsKeyBackSingleQuote = $000000C0;
  wpsKeyOpenSquareBrace = $000000DB;
  wpsKeyBackSlash = $000000DC;
  wpsKeyCloseSquareBrace = $000000DD;
  wpsKeySingleQuote = $000000DE;

// Constants for enum WpsBrowseTarget
type
  WpsBrowseTarget = TOleEnum;
const
  wpsBrowsePage = $00000001;
  wpsBrowseSection = $00000002;
  wpsBrowseComment = $00000003;
  wpsBrowseFind = $00000004;
  wpsBrowseBookmark = $00000005;
  wpsBrowseGoTo = $00000006;

// Constants for enum WpsCountry
type
  WpsCountry = TOleEnum;
const
  wpsUS = $00000001;
  wpsCanada = $00000002;
  wpsLatinAmerica = $00000003;
  wpsNetherlands = $0000001F;
  wpsFrance = $00000021;
  wpsSpain = $00000022;
  wpsItaly = $00000027;
  wpsUK = $0000002C;
  wpsDenmark = $0000002D;
  wpsSweden = $0000002E;
  wpsNorway = $0000002F;
  wpsGermany = $00000031;
  wpsPeru = $00000033;
  wpsMexico = $00000034;
  wpsArgentina = $00000036;
  wpsBrazil = $00000037;
  wpsChile = $00000038;
  wpsVenezuela = $0000003A;
  wpsJapan = $00000051;
  wpsTaiwan = $00000376;
  wpsChina = $00000056;
  wpsKorea = $00000052;
  wpsFinland = $00000166;
  wpsIceland = $00000162;

// Constants for enum WpsCursorType
type
  WpsCursorType = TOleEnum;
const
  wpsCursorWait = $00000000;
  wpsCursorIBeam = $00000001;
  wpsCursorNormal = $00000002;
  wpsCursorNorthwestArrow = $00000003;

// Constants for enum WpsPdfCommentsMode
type
  WpsPdfCommentsMode = TOleEnum;
const
  wpsPdfIgnoreComments = $00000001;
  wpsPdfPrintCommentsOnly = $00000002;
  wpsPdfCompatibleComments = $00000003;

// Constants for enum WpsPdfEditRight
type
  WpsPdfEditRight = TOleEnum;
const
  wpsPdfAssemble = $00000001;
  wpsPdfFillForm = $00000002;
  wpsPdfAnnotate = $00000003;
  wpsPdfModify = $00000004;

// Constants for enum WpsPdfPrintRight
type
  WpsPdfPrintRight = TOleEnum;
const
  wpsPdfNotAllowPrint = $00000001;
  wpsPdfPrintLowQulityOnly = $00000002;
  wpsPdfFreePrint = $00000003;

// Constants for enum WpsPdfCopyRight
type
  WpsPdfCopyRight = TOleEnum;
const
  wpsPdfNotAllowCopy = $00000000;
  wpsPdfFreeCopy = $0000FFFF;

// Constants for enum WpsDialog
type
  WpsDialog = TOleEnum;
const
  wpsDialogWordArtGallery = $00005001;
  wpsDialogEditWordArtText = $00005002;
  wpsDialogInsertSymbol = $00005003;
  wpsDialogHyperlinkDlg = $00005004;
  wpsDialogNewTemplate = $00005006;
  wpsDialogFillFormat = $00005008;
  wpsDialogInsertOLEObject = $00005009;
  wpsDialogPasteSpecial = $00005010;
  wpsDialogDocProperties = $00005011;
  wpsDialogCOMAddins = $00005014;
  wpsDialogAbout = $00005000;
  wpsDialogFontDlg = $00004000;
  wpsDialogDropCap = $00004001;
  wpsDialogTextDirection = $00004002;
  wpsDialogPrint = $00004003;
  wpsDialogPageSetup = $00004005;
  wpsDialogBordersAndShading = $00004006;
  wpsDialogParagraph = $00004007;
  wpsDialogFindReplace = $00004008;
  wpsDialogTabs = $00004009;
  wpsDialogBookmarks = $0000400A;
  wpsDialogZoom = $0000400B;
  wpsDialogFootAndEndnote = $0000400C;
  wpsDialogInsertCells = $0000400D;
  wpsDialogDeleteCells = $0000400E;
  wpsDialogSplitCells = $0000400F;
  wpsDialogTableProperties = $00004010;
  wpsDialogColumns = $00004011;
  wpsDialogCaption = $00004012;
  wpsDialogNewStyle = $00004013;
  wpsDialogModifyStyle = $00004014;
  wpsDialogInsertDateTime = $00004016;
  wpsDialogChangeCase = $00004017;
  wpsDialogContents = $00004018;
  wpsDialogPageNumbers = $00004019;
  wpsDialogInsertTable = $0000401B;
  wpsDialogBulletsNumbering = $0000401C;
  wpsDialogDrawingGrid = $00004020;
  wpsDialogField = $00004021;
  wpsDialogWordCount = $00004022;
  wpsDialogMultidiagonalCell = $00004023;
  wpsDialogOptions = $00004024;
  wpsDialogEncloseCharacters = $00004025;
  wpsDialogCombineCharacters = $00004026;
  wpsDialogPageNumberFormat = $00004028;
  wpsDialogInsertTableRows = $0000402A;
  wpsDialogGenkoSetting = $0000402C;
  wpsDialogProtectDocument = $0000402D;
  wpsDialogExportToPDF = $00004035;
  wpsDialogChineseTranslation = $00004036;
  wpsDialogCheckSpelling = $00004037;
  wpsDialogOpenFile = $00004500;
  wpsDialogSaveAs = $00004501;
  wpsDialogInsertPictureFromFile = $00004502;
  wpsDialogInsertFile = $00004503;

// Constants for enum WpsDialogTab
type
  WpsDialogTab = TOleEnum;
const
  wpsDialogReserved = $00000000;

// Constants for enum WpsAlertLevel
type
  WpsAlertLevel = TOleEnum;
const
  wpsAlertsNone = $00000000;
  wpsAlertsMessageBox = $FFFFFFFE;
  wpsAlertsAll = $FFFFFFFF;

// Constants for enum WpsShapePosition
type
  WpsShapePosition = TOleEnum;
const
  wpsShapeBottom = $FFF0BDC3;
  wpsShapeCenter = $FFF0BDC5;
  wpsShapeInside = $FFF0BDC6;
  wpsShapeLeft = $FFF0BDC2;
  wpsShapeOutside = $FFF0BDC7;
  wpsShapeRight = $FFF0BDC4;
  wpsShapeTop = $FFF0BDC1;

// Constants for enum WpsMailSystem
type
  WpsMailSystem = TOleEnum;
const
  wpsNoMailSystem = $00000000;
  wpsMAPI = $00000001;
  wpsPowerTalk = $00000002;
  wpsMAPIandPowerTalk = $00000003;

// Constants for enum WpsIMEMode
type
  WpsIMEMode = TOleEnum;
const
  wpsIMEModeNoControl = $00000000;
  wpsIMEModeOn = $00000001;
  wpsIMEModeOff = $00000002;
  wpsIMEModeHiragana = $00000004;
  wpsIMEModeKatakana = $00000005;
  wpsIMEModeKatakanaHalf = $00000006;
  wpsIMEModeAlphaFull = $00000007;
  wpsIMEModeAlpha = $00000008;
  wpsIMEModeHangulFull = $00000009;
  wpsIMEModeHangul = $0000000A;

// Constants for enum WpsIndexFilter
type
  WpsIndexFilter = TOleEnum;
const
  wpsIndexFilterNone = $00000000;
  wpsIndexFilterAiueo = $00000001;
  wpsIndexFilterAkasatana = $00000002;
  wpsIndexFilterChosung = $00000003;
  wpsIndexFilterLow = $00000004;
  wpsIndexFilterMedium = $00000005;
  wpsIndexFilterFull = $00000006;

// Constants for enum WpsIndexSortBy
type
  WpsIndexSortBy = TOleEnum;
const
  wpsIndexSortByStroke = $00000000;
  wpsIndexSortBySyllable = $00000001;

// Constants for enum WpsMultipleWordConversionsMode
type
  WpsMultipleWordConversionsMode = TOleEnum;
const
  wpsHangulToHanja = $00000000;
  wpsHanjaToHangul = $00000001;

// Constants for enum WpsInternationalIndex
type
  WpsInternationalIndex = TOleEnum;
const
  wpsListSeparator = $00000011;
  wpsDecimalSeparator = $00000012;
  wpsThousandsSeparator = $00000013;
  wpsCurrencyCode = $00000014;
  wps24HourClock = $00000015;
  wpsInternationalAM = $00000016;
  wpsInternationalPM = $00000017;
  wpsTimeSeparator = $00000018;
  wpsDateSeparator = $00000019;
  wpsProductLanguageID = $0000001A;

// Constants for enum WpsAutoMacros
type
  WpsAutoMacros = TOleEnum;
const
  wpsAutoExec = $00000000;
  wpsAutoNew = $00000001;
  wpsAutoOpen = $00000002;
  wpsAutoClose = $00000003;
  wpsAutoExit = $00000004;

// Constants for enum WpsHeadingSeparator
type
  WpsHeadingSeparator = TOleEnum;
const
  wpsHeadingSeparatorNone = $00000000;
  wpsHeadingSeparatorBlankLine = $00000001;
  wpsHeadingSeparatorLetter = $00000002;
  wpsHeadingSeparatorLetterLow = $00000003;
  wpsHeadingSeparatorLetterFull = $00000004;

// Constants for enum WpsFramePosition
type
  WpsFramePosition = TOleEnum;
const
  wpsFrameTop = $FFF0BDC1;
  wpsFrameLeft = $FFF0BDC2;
  wpsFrameBottom = $FFF0BDC3;
  wpsFrameRight = $FFF0BDC4;
  wpsFrameCenter = $FFF0BDC5;
  wpsFrameInside = $FFF0BDC6;
  wpsFrameOutside = $FFF0BDC7;

// Constants for enum WpsAnimation
type
  WpsAnimation = TOleEnum;
const
  wpsAnimationNone = $00000000;
  wpsAnimationLasVegasLights = $00000001;
  wpsAnimationBlinkingBackground = $00000002;
  wpsAnimationSparkleText = $00000003;
  wpsAnimationMarchingBlackAnts = $00000004;
  wpsAnimationMarchingRedAnts = $00000005;
  wpsAnimationShimmer = $00000006;

// Constants for enum WpsSummaryMode
type
  WpsSummaryMode = TOleEnum;
const
  wpsSummaryModeHighlight = $00000000;
  wpsSummaryModeHideAllButSummary = $00000001;
  wpsSummaryModeInsert = $00000002;
  wpsSummaryModeCreateNew = $00000003;

// Constants for enum WpsSummaryLength
type
  WpsSummaryLength = TOleEnum;
const
  wps10Sentences = $FFFFFFFE;
  wps20Sentences = $FFFFFFFD;
  wps100Words = $FFFFFFFC;
  wps500Words = $FFFFFFFB;
  wps10Percent = $FFFFFFFA;
  wps25Percent = $FFFFFFF9;
  wps50Percent = $FFFFFFF8;
  wps75Percent = $FFFFFFF7;

// Constants for enum WpsInsertCells
type
  WpsInsertCells = TOleEnum;
const
  wpsInsertCellsShiftRight = $00000000;
  wpsInsertCellsShiftDown = $00000001;
  wpsInsertCellsEntireRow = $00000002;
  wpsInsertCellsEntireColumn = $00000003;

// Constants for enum WpsDeleteCells
type
  WpsDeleteCells = TOleEnum;
const
  wpsDeleteCellsShiftLeft = $00000000;
  wpsDeleteCellsShiftUp = $00000001;
  wpsDeleteCellsEntireRow = $00000002;
  wpsDeleteCellsEntireColumn = $00000003;

// Constants for enum WpsListApplyBehavior
type
  WpsListApplyBehavior = TOleEnum;
const
  wpsListBehaviorApplyBuiltin = $00000000;
  wpsListBehaviorApplyCustom = $00000001;

// Constants for enum WpsDefaultListBehavior
type
  WpsDefaultListBehavior = TOleEnum;
const
  wps2001ListBehavior = $00010000;
  wps2002ListBehavior = $00020000;
  wps2003ListBehavior = $00030000;
  wpsv6ListBehavior = $00040000;

// Constants for enum WpsEnableCancelKey
type
  WpsEnableCancelKey = TOleEnum;
const
  wpsCancelDisabled = $00000000;
  wpsCancelInterrupt = $00000001;

// Constants for enum WpsVerticalAlignment
type
  WpsVerticalAlignment = TOleEnum;
const
  wpsAlignVerticalTop = $00000000;
  wpsAlignVerticalCenter = $00000001;
  wpsAlignVerticalJustify = $00000002;
  wpsAlignVerticalBottom = $00000003;

// Constants for enum WpsStatistic
type
  WpsStatistic = TOleEnum;
const
  wpsStatisticWords = $00000000;
  wpsStatisticLines = $00000001;
  wpsStatisticPages = $00000002;
  wpsStatisticCharacters = $00000003;
  wpsStatisticParagraphs = $00000004;
  wpsStatisticCharactersWithSpaces = $00000005;
  wpsStatisticFarEastCharacters = $00000006;

// Constants for enum WpsBuiltInProperty
type
  WpsBuiltInProperty = TOleEnum;
const
  wpsPropertyTitle = $00000001;
  wpsPropertySubject = $00000002;
  wpsPropertyAuthor = $00000003;
  wpsPropertyKeywords = $00000004;
  wpsPropertyComments = $00000005;
  wpsPropertyTemplate = $00000006;
  wpsPropertyLastAuthor = $00000007;
  wpsPropertyRevision = $00000008;
  wpsPropertyAppName = $00000009;
  wpsPropertyTimeLastPrinted = $0000000A;
  wpsPropertyTimeCreated = $0000000B;
  wpsPropertyTimeLastSaved = $0000000C;
  wpsPropertyVBATotalEdit = $0000000D;
  wpsPropertyPages = $0000000E;
  wpsPropertyWords = $0000000F;
  wpsPropertyCharacters = $00000010;
  wpsPropertySecurity = $00000011;
  wpsPropertyCategory = $00000012;
  wpsPropertyFormat = $00000013;
  wpsPropertyManager = $00000014;
  wpsPropertyCompany = $00000015;
  wpsPropertyBytes = $00000016;
  wpsPropertyLines = $00000017;
  wpsPropertyParas = $00000018;
  wpsPropertySlides = $00000019;
  wpsPropertyNotes = $0000001A;
  wpsPropertyHiddenSlides = $0000001B;
  wpsPropertyMMClips = $0000001C;
  wpsPropertyHyperlinkBase = $0000001D;
  wpsPropertyCharsWSpaces = $0000001E;

// Constants for enum WpsSaveFormat
type
  WpsSaveFormat = TOleEnum;
const
  wpsFormatDocument = $00000000;
  wpsFormatTemplate = $00000001;
  wpsFormatText = $00000002;
  wpsFormatTextLineBreaks = $00000003;
  wpsFormatDOSText = $00000004;
  wpsFormatDOSTextLineBreaks = $00000005;
  wpsFormatRTF = $00000006;
  wpsFormatUnicodeText = $00000007;
  wpsFormatEncodedText = $00000007;
  wpsFormatHTML = $00000008;
  wpsFormatWebArchive = $00000009;
  wpsFormatFilteredHTML = $0000000A;

// Constants for enum WpsToaFormat
type
  WpsToaFormat = TOleEnum;
const
  wpsTOATemplate = $00000000;
  wpsTOAClassic = $00000001;
  wpsTOADistinctive = $00000002;
  wpsTOAFormal = $00000003;
  wpsTOASimple = $00000004;

// Constants for enum WpsSortSeparator
type
  WpsSortSeparator = TOleEnum;
const
  wpsSortSeparateByTabs = $00000000;
  wpsSortSeparateByCommas = $00000001;
  wpsSortSeparateByDefaultTableSeparator = $00000002;

// Constants for enum WpsTableFieldSeparator
type
  WpsTableFieldSeparator = TOleEnum;
const
  wpsSeparateByParagraphs = $00000000;
  wpsSeparateByTabs = $00000001;
  wpsSeparateByCommas = $00000002;
  wpsSeparateByDefaultListSeparator = $00000003;
  wpsSeparateBySpace = $00000004;

// Constants for enum WpsSortFieldType
type
  WpsSortFieldType = TOleEnum;
const
  wpsSortFieldAlphanumeric = $00000000;
  wpsSortFieldNumeric = $00000001;
  wpsSortFieldDate = $00000002;
  wpsSortFieldSyllable = $00000003;
  wpsSortFieldJapanJIS = $00000004;
  wpsSortFieldStroke = $00000005;
  wpsSortFieldKoreaKS = $00000006;

// Constants for enum WpsSortOrder
type
  WpsSortOrder = TOleEnum;
const
  wpsSortOrderAscending = $00000000;
  wpsSortOrderDescending = $00000001;

// Constants for enum WpsTableFormat
type
  WpsTableFormat = TOleEnum;
const
  wpsTableFormatNone = $00000000;
  wpsTableFormatSimple1 = $00000001;
  wpsTableFormatSimple2 = $00000002;
  wpsTableFormatSimple3 = $00000003;
  wpsTableFormatClassic1 = $00000004;
  wpsTableFormatClassic2 = $00000005;
  wpsTableFormatClassic3 = $00000006;
  wpsTableFormatClassic4 = $00000007;
  wpsTableFormatColorful1 = $00000008;
  wpsTableFormatColorful2 = $00000009;
  wpsTableFormatColorful3 = $0000000A;
  wpsTableFormatColumns1 = $0000000B;
  wpsTableFormatColumns2 = $0000000C;
  wpsTableFormatColumns3 = $0000000D;
  wpsTableFormatColumns4 = $0000000E;
  wpsTableFormatColumns5 = $0000000F;
  wpsTableFormatGrid1 = $00000010;
  wpsTableFormatGrid2 = $00000011;
  wpsTableFormatGrid3 = $00000012;
  wpsTableFormatGrid4 = $00000013;
  wpsTableFormatGrid5 = $00000014;
  wpsTableFormatGrid6 = $00000015;
  wpsTableFormatGrid7 = $00000016;
  wpsTableFormatGrid8 = $00000017;
  wpsTableFormatList1 = $00000018;
  wpsTableFormatList2 = $00000019;
  wpsTableFormatList3 = $0000001A;
  wpsTableFormatList4 = $0000001B;
  wpsTableFormatList5 = $0000001C;
  wpsTableFormatList6 = $0000001D;
  wpsTableFormatList7 = $0000001E;
  wpsTableFormatList8 = $0000001F;
  wpsTableFormat3DEffects1 = $00000020;
  wpsTableFormat3DEffects2 = $00000021;
  wpsTableFormat3DEffects3 = $00000022;
  wpsTableFormatContemporary = $00000023;
  wpsTableFormatElegant = $00000024;
  wpsTableFormatProfessional = $00000025;
  wpsTableFormatSubtle1 = $00000026;
  wpsTableFormatSubtle2 = $00000027;
  wpsTableFormatWeb1 = $00000028;
  wpsTableFormatWeb2 = $00000029;
  wpsTableFormatWeb3 = $0000002A;

// Constants for enum WpsTableFormatApply
type
  WpsTableFormatApply = TOleEnum;
const
  wpsTableFormatApplyBorders = $00000001;
  wpsTableFormatApplyShading = $00000002;
  wpsTableFormatApplyFont = $00000004;
  wpsTableFormatApplyColor = $00000008;
  wpsTableFormatApplyAutoFit = $00000010;
  wpsTableFormatApplyHeadingRows = $00000020;
  wpsTableFormatApplyLastRow = $00000040;
  wpsTableFormatApplyFirstColumn = $00000080;
  wpsTableFormatApplyLastColumn = $00000100;

// Constants for enum WpsBuiltinStyle
type
  WpsBuiltinStyle = TOleEnum;
const
  wpsStyleNormal = $FFFFFFFF;
  wpsStyleEnvelopeAddress = $FFFFFFDB;
  wpsStyleEnvelopeReturn = $FFFFFFDA;
  wpsStyleBodyText = $FFFFFFBD;
  wpsStyleHeading1 = $FFFFFFFE;
  wpsStyleHeading2 = $FFFFFFFD;
  wpsStyleHeading3 = $FFFFFFFC;
  wpsStyleHeading4 = $FFFFFFFB;
  wpsStyleHeading5 = $FFFFFFFA;
  wpsStyleHeading6 = $FFFFFFF9;
  wpsStyleHeading7 = $FFFFFFF8;
  wpsStyleHeading8 = $FFFFFFF7;
  wpsStyleHeading9 = $FFFFFFF6;
  wpsStyleIndex1 = $FFFFFFF5;
  wpsStyleIndex2 = $FFFFFFF4;
  wpsStyleIndex3 = $FFFFFFF3;
  wpsStyleIndex4 = $FFFFFFF2;
  wpsStyleIndex5 = $FFFFFFF1;
  wpsStyleIndex6 = $FFFFFFF0;
  wpsStyleIndex7 = $FFFFFFEF;
  wpsStyleIndex8 = $FFFFFFEE;
  wpsStyleIndex9 = $FFFFFFED;
  wpsStyleTOC1 = $FFFFFFEC;
  wpsStyleTOC2 = $FFFFFFEB;
  wpsStyleTOC3 = $FFFFFFEA;
  wpsStyleTOC4 = $FFFFFFE9;
  wpsStyleTOC5 = $FFFFFFE8;
  wpsStyleTOC6 = $FFFFFFE7;
  wpsStyleTOC7 = $FFFFFFE6;
  wpsStyleTOC8 = $FFFFFFE5;
  wpsStyleTOC9 = $FFFFFFE4;
  wpsStyleNormalIndent = $FFFFFFE3;
  wpsStyleFootnoteText = $FFFFFFE2;
  wpsStyleCommentText = $FFFFFFE1;
  wpsStyleHeader = $FFFFFFE0;
  wpsStyleFooter = $FFFFFFDF;
  wpsStyleIndexHeading = $FFFFFFDE;
  wpsStyleCaption = $FFFFFFDD;
  wpsStyleTableOfFigures = $FFFFFFDC;
  wpsStyleFootnoteReference = $FFFFFFD9;
  wpsStyleCommentReference = $FFFFFFD8;
  wpsStyleLineNumber = $FFFFFFD7;
  wpsStylePageNumber = $FFFFFFD6;
  wpsStyleEndnoteReference = $FFFFFFD5;
  wpsStyleEndnoteText = $FFFFFFD4;
  wpsStyleTableOfAuthorities = $FFFFFFD3;
  wpsStyleMacroText = $FFFFFFD2;
  wpsStyleTOAHeading = $FFFFFFD1;
  wpsStyleList = $FFFFFFD0;
  wpsStyleListBullet = $FFFFFFCF;
  wpsStyleListNumber = $FFFFFFCE;
  wpsStyleList2 = $FFFFFFCD;
  wpsStyleList3 = $FFFFFFCC;
  wpsStyleList4 = $FFFFFFCB;
  wpsStyleList5 = $FFFFFFCA;
  wpsStyleListBullet2 = $FFFFFFC9;
  wpsStyleListBullet3 = $FFFFFFC8;
  wpsStyleListBullet4 = $FFFFFFC7;
  wpsStyleListBullet5 = $FFFFFFC6;
  wpsStyleListNumber2 = $FFFFFFC5;
  wpsStyleListNumber3 = $FFFFFFC4;
  wpsStyleListNumber4 = $FFFFFFC3;
  wpsStyleListNumber5 = $FFFFFFC2;
  wpsStyleTitle = $FFFFFFC1;
  wpsStyleClosing = $FFFFFFC0;
  wpsStyleSignature = $FFFFFFBF;
  wpsStyleDefaultParagraphFont = $FFFFFFBE;
  wpsStyleBodyTextIndent = $FFFFFFBC;
  wpsStyleListContinue = $FFFFFFBB;
  wpsStyleListContinue2 = $FFFFFFBA;
  wpsStyleListContinue3 = $FFFFFFB9;
  wpsStyleListContinue4 = $FFFFFFB8;
  wpsStyleListContinue5 = $FFFFFFB7;
  wpsStyleMessageHeader = $FFFFFFB6;
  wpsStyleSubtitle = $FFFFFFB5;
  wpsStyleSalutation = $FFFFFFB4;
  wpsStyleDate = $FFFFFFB3;
  wpsStyleBodyTextFirstIndent = $FFFFFFB2;
  wpsStyleBodyTextFirstIndent2 = $FFFFFFB1;
  wpsStyleNoteHeading = $FFFFFFB0;
  wpsStyleBodyText2 = $FFFFFFAF;
  wpsStyleBodyText3 = $FFFFFFAE;
  wpsStyleBodyTextIndent2 = $FFFFFFAD;
  wpsStyleBodyTextIndent3 = $FFFFFFAC;
  wpsStyleBlockQuotation = $FFFFFFAB;
  wpsStyleHyperlink = $FFFFFFAA;
  wpsStyleHyperlinkFollowed = $FFFFFFA9;
  wpsStyleStrong = $FFFFFFA8;
  wpsStyleEmphasis = $FFFFFFA7;
  wpsStyleNavPane = $FFFFFFA6;
  wpsStylePlainText = $FFFFFFA5;
  wpsStyleHtmlNormal = $FFFFFFA1;
  wpsStyleHtmlAcronym = $FFFFFFA0;
  wpsStyleHtmlAddress = $FFFFFF9F;
  wpsStyleHtmlCite = $FFFFFF9E;
  wpsStyleHtmlCode = $FFFFFF9D;
  wpsStyleHtmlDfn = $FFFFFF9C;
  wpsStyleHtmlKbd = $FFFFFF9B;
  wpsStyleHtmlPre = $FFFFFF9A;
  wpsStyleHtmlSamp = $FFFFFF99;
  wpsStyleHtmlTt = $FFFFFF98;
  wpsStyleHtmlVar = $FFFFFF97;
  wpsStyleNormalTable = $FFFFFF96;

// Constants for enum WpsChevronConvertRule
type
  WpsChevronConvertRule = TOleEnum;
const
  wpsNeverConvert = $00000000;
  wpsAlwaysConvert = $00000001;
  wpsAskToNotConvert = $00000002;
  wpsAskToConvert = $00000003;

// Constants for enum WpsMailMergeDefaultRecord
type
  WpsMailMergeDefaultRecord = TOleEnum;
const
  wpsDefaultFirstRecord = $00000001;
  wpsDefaultLastRecord = $FFFFFFF0;

// Constants for enum WpsPictureLinkType
type
  WpsPictureLinkType = TOleEnum;
const
  wpsLinkNone = $00000000;
  wpsLinkDataInDoc = $00000001;
  wpsLinkDataOnDisk = $00000002;

// Constants for enum WpsLinkType
type
  WpsLinkType = TOleEnum;
const
  wpsLinkTypeOLE = $00000000;
  wpsLinkTypePicture = $00000001;
  wpsLinkTypeText = $00000002;
  wpsLinkTypeReference = $00000003;
  wpsLinkTypeInclude = $00000004;
  wpsLinkTypeImport = $00000005;
  wpsLinkTypeDDE = $00000006;
  wpsLinkTypeDDEAuto = $00000007;

// Constants for enum WpsSpecialPane
type
  WpsSpecialPane = TOleEnum;
const
  wpsPaneNone = $00000000;
  wpsPanePrimaryHeader = $00000001;
  wpsPaneFirstPageHeader = $00000002;
  wpsPaneEvenPagesHeader = $00000003;
  wpsPanePrimaryFooter = $00000004;
  wpsPaneFirstPageFooter = $00000005;
  wpsPaneEvenPagesFooter = $00000006;
  wpsPaneFootnotes = $00000007;
  wpsPaneEndnotes = $00000008;
  wpsPaneFootnoteContinuationNotice = $00000009;
  wpsPaneFootnoteContinuationSeparator = $0000000A;
  wpsPaneFootnoteSeparator = $0000000B;
  wpsPaneEndnoteContinuationNotice = $0000000C;
  wpsPaneEndnoteContinuationSeparator = $0000000D;
  wpsPaneEndnoteSeparator = $0000000E;
  wpsPaneComments = $0000000F;
  wpsPaneCurrentPageHeader = $00000010;
  wpsPaneCurrentPageFooter = $00000011;
  wpsPaneRevisions = $00000012;

// Constants for enum WpsReferenceType
type
  WpsReferenceType = TOleEnum;
const
  wpsRefTypeNumberedItem = $00000000;
  wpsRefTypeHeading = $00000001;
  wpsRefTypeBookmark = $00000002;
  wpsRefTypeFootnote = $00000003;
  wpsRefTypeEndnote = $00000004;

// Constants for enum WpsIndexFormat
type
  WpsIndexFormat = TOleEnum;
const
  wpsIndexTemplate = $00000000;
  wpsIndexClassic = $00000001;
  wpsIndexFancy = $00000002;
  wpsIndexModern = $00000003;
  wpsIndexBulleted = $00000004;
  wpsIndexFormal = $00000005;
  wpsIndexSimple = $00000006;

// Constants for enum WpsIndexType
type
  WpsIndexType = TOleEnum;
const
  wpsIndexIndent = $00000000;
  wpsIndexRunin = $00000001;

// Constants for enum WpsRevisionsWrap
type
  WpsRevisionsWrap = TOleEnum;
const
  wpsWrapNever = $00000000;
  wpsWrapAlways = $00000001;
  wpsWrapAsk = $00000002;

// Constants for enum WpsRoutingSlipDelivery
type
  WpsRoutingSlipDelivery = TOleEnum;
const
  wpsOneAfterAnother = $00000000;
  wpsAllAtOnce = $00000001;

// Constants for enum WpsRoutingSlipStatus
type
  WpsRoutingSlipStatus = TOleEnum;
const
  wpsNotYetRouted = $00000000;
  wpsRouteInProgress = $00000001;
  wpsRouteComplete = $00000002;

// Constants for enum WpsOriginalFormat
type
  WpsOriginalFormat = TOleEnum;
const
  wpsWordDocument = $00000000;
  wpsOriginalDocumentFormat = $00000001;
  wpsPromptUser = $00000002;
  wpsWpsDocument = $00000003;

// Constants for enum WpsCustomLabelPageSize
type
  WpsCustomLabelPageSize = TOleEnum;
const
  wpsCustomLabelLetter = $00000000;
  wpsCustomLabelLetterLS = $00000001;
  wpsCustomLabelA4 = $00000002;
  wpsCustomLabelA4LS = $00000003;
  wpsCustomLabelA5 = $00000004;
  wpsCustomLabelA5LS = $00000005;
  wpsCustomLabelB5 = $00000006;
  wpsCustomLabelMini = $00000007;
  wpsCustomLabelFanfold = $00000008;
  wpsCustomLabelVertHalfSheet = $00000009;
  wpsCustomLabelVertHalfSheetLS = $0000000A;
  wpsCustomLabelHigaki = $0000000B;
  wpsCustomLabelHigakiLS = $0000000C;
  wpsCustomLabelB4JIS = $0000000D;

// Constants for enum WpsPartOfSpeech
type
  WpsPartOfSpeech = TOleEnum;
const
  wpsAdjective = $00000000;
  wpsNoun = $00000001;
  wpsAdverb = $00000002;
  wpsVerb = $00000003;
  wpsPronoun = $00000004;
  wpsConjunction = $00000005;
  wpsPreposition = $00000006;
  wpsInterjection = $00000007;
  wpsIdiom = $00000008;
  wpsOther = $00000009;

// Constants for enum WpsSubscriberFormats
type
  WpsSubscriberFormats = TOleEnum;
const
  wpsSubscriberBestFormat = $00000000;
  wpsSubscriberRTF = $00000001;
  wpsSubscriberText = $00000002;
  wpsSubscriberPict = $00000004;

// Constants for enum WpsEditionType
type
  WpsEditionType = TOleEnum;
const
  wpsPublisher = $00000000;
  wpsSubscriber = $00000001;

// Constants for enum WpsEditionOption
type
  WpsEditionOption = TOleEnum;
const
  wpsCancelPublisher = $00000000;
  wpsSendPublisher = $00000001;
  wpsSelectPublisher = $00000002;
  wpsAutomaticUpdate = $00000003;
  wpsManualUpdate = $00000004;
  wpsChangeAttributes = $00000005;
  wpsUpdateSubscriber = $00000006;
  wpsOpenSource = $00000007;

// Constants for enum WpsTablePosition
type
  WpsTablePosition = TOleEnum;
const
  wpsTableTop = $FFF0BDC1;
  wpsTableLeft = $FFF0BDC2;
  wpsTableBottom = $FFF0BDC3;
  wpsTableRight = $FFF0BDC4;
  wpsTableCenter = $FFF0BDC5;
  wpsTableInside = $FFF0BDC6;
  wpsTableOutside = $FFF0BDC7;

// Constants for enum WpsHelpType
type
  WpsHelpType = TOleEnum;
const
  wpsHelp = $00000000;
  wpsHelpAbout = $00000001;
  wpsHelpActiveWindow = $00000002;
  wpsHelpContents = $00000003;
  wpsHelpExamplesAndDemos = $00000004;
  wpsHelpIndex = $00000005;
  wpsHelpKeyboard = $00000006;
  wpsHelpPSSHelp = $00000007;
  wpsHelpQuickPreview = $00000008;
  wpsHelpSearch = $00000009;
  wpsHelpUsingHelp = $0000000A;
  wpsHelpIchitaro = $0000000B;
  wpsHelpPE2 = $0000000C;
  wpsHelpHWP = $0000000D;

// Constants for enum WpsOLEType
type
  WpsOLEType = TOleEnum;
const
  wpsOLELink = $00000000;
  wpsOLEEmbed = $00000001;
  wpsOLEControl = $00000002;

// Constants for enum WpsOLEVerb
type
  WpsOLEVerb = TOleEnum;
const
  wpsOLEVerbPrimary = $00000000;
  wpsOLEVerbShow = $FFFFFFFF;
  wpsOLEVerbOpen = $FFFFFFFE;
  wpsOLEVerbHide = $FFFFFFFD;
  wpsOLEVerbUIActivate = $FFFFFFFC;
  wpsOLEVerbInPlaceActivate = $FFFFFFFB;
  wpsOLEVerbDiscardUndoState = $FFFFFFFA;

// Constants for enum WpsEnvelopeOrientation
type
  WpsEnvelopeOrientation = TOleEnum;
const
  wpsLeftPortrait = $00000000;
  wpsCenterPortrait = $00000001;
  wpsRightPortrait = $00000002;
  wpsLeftLandscape = $00000003;
  wpsCenterLandscape = $00000004;
  wpsRightLandscape = $00000005;
  wpsLeftClockwise = $00000006;
  wpsCenterClockwise = $00000007;
  wpsRightClockwise = $00000008;

// Constants for enum WpsLetterStyle
type
  WpsLetterStyle = TOleEnum;
const
  wpsFullBlock = $00000000;
  wpsModifiedBlock = $00000001;
  wpsSemiBlock = $00000002;

// Constants for enum WpsLetterheadLocation
type
  WpsLetterheadLocation = TOleEnum;
const
  wpsLetterTop = $00000000;
  wpsLetterBottom = $00000001;
  wpsLetterLeft = $00000002;
  wpsLetterRight = $00000003;

// Constants for enum WpsSalutationType
type
  WpsSalutationType = TOleEnum;
const
  wpsSalutationInformal = $00000000;
  wpsSalutationFormal = $00000001;
  wpsSalutationBusiness = $00000002;
  wpsSalutationOther = $00000003;

// Constants for enum WpsSalutationGender
type
  WpsSalutationGender = TOleEnum;
const
  wpsGenderFemale = $00000000;
  wpsGenderMale = $00000001;
  wpsGenderNeutral = $00000002;
  wpsGenderUnknown = $00000003;

// Constants for enum WpsConstants
type
  WpsConstants = TOleEnum;
const
  wpsUndefined = $0098967F;
  wpsToggle = $0098967E;
  wpsForward = $3FFFFFFF;
  wpsBackward = $C0000001;
  wpsAutoPosition = $00000000;
  wpsFirst = $00000001;
  wpsCreatorCode = $4B575053;

// Constants for enum WpsPasteDataType
type
  WpsPasteDataType = TOleEnum;
const
  wpsPasteOLEObject = $00000000;
  wpsPasteRTF = $00000001;
  wpsPasteText = $00000002;
  wpsPasteMetafilePicture = $00000003;
  wpsPasteBitmap = $00000004;
  wpsPasteDeviceIndependentBitmap = $00000005;
  wpsPasteHyperlink = $00000007;
  wpsPasteShape = $00000008;
  wpsPasteEnhancedMetafile = $00000009;
  wpsPasteHTML = $0000000A;

// Constants for enum WpsDictionaryType
type
  WpsDictionaryType = TOleEnum;
const
  wpsSpelling = $00000000;
  wpsGrammar = $00000001;
  wpsThesaurus = $00000002;
  wpsHyphenation = $00000003;
  wpsSpellingComplete = $00000004;
  wpsSpellingCustom = $00000005;
  wpsSpellingLegal = $00000006;
  wpsSpellingMedical = $00000007;
  wpsHangulHanjaConversion = $00000008;
  wpsHangulHanjaConversionCustom = $00000009;

// Constants for enum WpsSpellingWordType
type
  WpsSpellingWordType = TOleEnum;
const
  wpsSpellword = $00000000;
  wpsWildcard = $00000001;
  wpsAnagram = $00000002;

// Constants for enum WpsSpellingErrorType
type
  WpsSpellingErrorType = TOleEnum;
const
  wpsSpellingCorrect = $00000000;
  wpsSpellingNotInDictionary = $00000001;
  wpsSpellingCapitalization = $00000002;

// Constants for enum WpsProofreadingErrorType
type
  WpsProofreadingErrorType = TOleEnum;
const
  wpsSpellingError = $00000000;
  wpsGrammaticalError = $00000001;

// Constants for enum WpsArrangeStyle
type
  WpsArrangeStyle = TOleEnum;
const
  wpsTiled = $00000000;
  wpsIcons = $00000001;

// Constants for enum WpsAutoVersions
type
  WpsAutoVersions = TOleEnum;
const
  wpsAutoVersionOff = $00000000;
  wpsAutoVersionOnClose = $00000001;

// Constants for enum WpsFindMatch
type
  WpsFindMatch = TOleEnum;
const
  wpsMatchParagraphMark = $0001000F;
  wpsMatchTabCharacter = $00000009;
  wpsMatchCommentMark = $00000005;
  wpsMatchAnyCharacter = $0001003F;
  wpsMatchAnyDigit = $0001001F;
  wpsMatchAnyLetter = $0001002F;
  wpsMatchCaretCharacter = $0000000B;
  wpsMatchColumnBreak = $0000000E;
  wpsMatchEmDash = $00002014;
  wpsMatchEnDash = $00002013;
  wpsMatchEndnoteMark = $00010013;
  wpsMatchField = $00000013;
  wpsMatchFootnoteMark = $00010012;
  wpsMatchGraphic = $00000001;
  wpsMatchManualLineBreak = $0001000F;
  wpsMatchManualPageBreak = $0001001C;
  wpsMatchNonbreakingHyphen = $0000001E;
  wpsMatchNonbreakingSpace = $000000A0;
  wpsMatchOptionalHyphen = $0000001F;
  wpsMatchSectionBreak = $0001002C;
  wpsMatchWhiteSpace = $00010077;

// Constants for enum WpsPageBorderArt
type
  WpsPageBorderArt = TOleEnum;
const
  wpsArtApples = $00000001;
  wpsArtMapleMuffins = $00000002;
  wpsArtCakeSlice = $00000003;
  wpsArtCandyCorn = $00000004;
  wpsArtIceCreamCones = $00000005;
  wpsArtChampagneBottle = $00000006;
  wpsArtPartyGlass = $00000007;
  wpsArtChristmasTree = $00000008;
  wpsArtTrees = $00000009;
  wpsArtPalmsColor = $0000000A;
  wpsArtBalloons3Colors = $0000000B;
  wpsArtBalloonsHotAir = $0000000C;
  wpsArtPartyFavor = $0000000D;
  wpsArtConfettiStreamers = $0000000E;
  wpsArtHearts = $0000000F;
  wpsArtHeartBalloon = $00000010;
  wpsArtStars3D = $00000011;
  wpsArtStarsShadowed = $00000012;
  wpsArtStars = $00000013;
  wpsArtSun = $00000014;
  wpsArtEarth2 = $00000015;
  wpsArtEarth1 = $00000016;
  wpsArtPeopleHats = $00000017;
  wpsArtSombrero = $00000018;
  wpsArtPencils = $00000019;
  wpsArtPackages = $0000001A;
  wpsArtClocks = $0000001B;
  wpsArtFirecrackers = $0000001C;
  wpsArtRings = $0000001D;
  wpsArtMapPins = $0000001E;
  wpsArtConfetti = $0000001F;
  wpsArtCreaturesButterfly = $00000020;
  wpsArtCreaturesLadyBug = $00000021;
  wpsArtCreaturesFish = $00000022;
  wpsArtBirdsFlight = $00000023;
  wpsArtScaredCat = $00000024;
  wpsArtBats = $00000025;
  wpsArtFlowersRoses = $00000026;
  wpsArtFlowersRedRose = $00000027;
  wpsArtPoinsettias = $00000028;
  wpsArtHolly = $00000029;
  wpsArtFlowersTiny = $0000002A;
  wpsArtFlowersPansy = $0000002B;
  wpsArtFlowersModern2 = $0000002C;
  wpsArtFlowersModern1 = $0000002D;
  wpsArtWhiteFlowers = $0000002E;
  wpsArtVine = $0000002F;
  wpsArtFlowersDaisies = $00000030;
  wpsArtFlowersBlockPrint = $00000031;
  wpsArtDecoArchColor = $00000032;
  wpsArtFans = $00000033;
  wpsArtFilm = $00000034;
  wpsArtLightning1 = $00000035;
  wpsArtCompass = $00000036;
  wpsArtDoubleD = $00000037;
  wpsArtClassicalWave = $00000038;
  wpsArtShadowedSquares = $00000039;
  wpsArtTwistedLines1 = $0000003A;
  wpsArtWaveline = $0000003B;
  wpsArtQuadrants = $0000003C;
  wpsArtCheckedBarColor = $0000003D;
  wpsArtSwirligig = $0000003E;
  wpsArtPushPinNote1 = $0000003F;
  wpsArtPushPinNote2 = $00000040;
  wpsArtPumpkin1 = $00000041;
  wpsArtEggsBlack = $00000042;
  wpsArtCup = $00000043;
  wpsArtHeartGray = $00000044;
  wpsArtGingerbreadMan = $00000045;
  wpsArtBabyPacifier = $00000046;
  wpsArtBabyRattle = $00000047;
  wpsArtCabins = $00000048;
  wpsArtHouseFunky = $00000049;
  wpsArtStarsBlack = $0000004A;
  wpsArtSnowflakes = $0000004B;
  wpsArtSnowflakeFancy = $0000004C;
  wpsArtSkyrocket = $0000004D;
  wpsArtSeattle = $0000004E;
  wpsArtMusicNotes = $0000004F;
  wpsArtPalmsBlack = $00000050;
  wpsArtMapleLeaf = $00000051;
  wpsArtPaperClips = $00000052;
  wpsArtShorebirdTracks = $00000053;
  wpsArtPeople = $00000054;
  wpsArtPeopleWaving = $00000055;
  wpsArtEclipsingSquares2 = $00000056;
  wpsArtHypnotic = $00000057;
  wpsArtDiamondsGray = $00000058;
  wpsArtDecoArch = $00000059;
  wpsArtDecoBlocks = $0000005A;
  wpsArtCirclesLines = $0000005B;
  wpsArtPapyrus = $0000005C;
  wpsArtWoodwork = $0000005D;
  wpsArtWeavingBraid = $0000005E;
  wpsArtWeavingRibbon = $0000005F;
  wpsArtWeavingAngles = $00000060;
  wpsArtArchedScallops = $00000061;
  wpsArtSafari = $00000062;
  wpsArtCelticKnotwork = $00000063;
  wpsArtCrazyMaze = $00000064;
  wpsArtEclipsingSquares1 = $00000065;
  wpsArtBirds = $00000066;
  wpsArtFlowersTeacup = $00000067;
  wpsArtNorthwest = $00000068;
  wpsArtSouthwest = $00000069;
  wpsArtTribal6 = $0000006A;
  wpsArtTribal4 = $0000006B;
  wpsArtTribal3 = $0000006C;
  wpsArtTribal2 = $0000006D;
  wpsArtTribal5 = $0000006E;
  wpsArtXIllusions = $0000006F;
  wpsArtZanyTriangles = $00000070;
  wpsArtPyramids = $00000071;
  wpsArtPyramidsAbove = $00000072;
  wpsArtConfettiGrays = $00000073;
  wpsArtConfettiOutline = $00000074;
  wpsArtConfettiWhite = $00000075;
  wpsArtMosaic = $00000076;
  wpsArtLightning2 = $00000077;
  wpsArtHeebieJeebies = $00000078;
  wpsArtLightBulb = $00000079;
  wpsArtGradient = $0000007A;
  wpsArtTriangleParty = $0000007B;
  wpsArtTwistedLines2 = $0000007C;
  wpsArtMoons = $0000007D;
  wpsArtOvals = $0000007E;
  wpsArtDoubleDiamonds = $0000007F;
  wpsArtChainLink = $00000080;
  wpsArtTriangles = $00000081;
  wpsArtTribal1 = $00000082;
  wpsArtMarqueeToothed = $00000083;
  wpsArtSharksTeeth = $00000084;
  wpsArtSawtooth = $00000085;
  wpsArtSawtoothGray = $00000086;
  wpsArtPostageStamp = $00000087;
  wpsArtWeavingStrips = $00000088;
  wpsArtZigZag = $00000089;
  wpsArtCrossStitch = $0000008A;
  wpsArtGems = $0000008B;
  wpsArtCirclesRectangles = $0000008C;
  wpsArtCornerTriangles = $0000008D;
  wpsArtCreaturesInsects = $0000008E;
  wpsArtZigZagStitch = $0000008F;
  wpsArtCheckered = $00000090;
  wpsArtCheckedBarBlack = $00000091;
  wpsArtMarquee = $00000092;
  wpsArtBasicWhiteDots = $00000093;
  wpsArtBasicWideMidline = $00000094;
  wpsArtBasicWideOutline = $00000095;
  wpsArtBasicWideInline = $00000096;
  wpsArtBasicThinLines = $00000097;
  wpsArtBasicWhiteDashes = $00000098;
  wpsArtBasicWhiteSquares = $00000099;
  wpsArtBasicBlackSquares = $0000009A;
  wpsArtBasicBlackDashes = $0000009B;
  wpsArtBasicBlackDots = $0000009C;
  wpsArtStarsTop = $0000009D;
  wpsArtCertificateBanner = $0000009E;
  wpsArtHandmade1 = $0000009F;
  wpsArtHandmade2 = $000000A0;
  wpsArtTornPaper = $000000A1;
  wpsArtTornPaperBlack = $000000A2;
  wpsArtCouponCutoutDashes = $000000A3;
  wpsArtCouponCutoutDots = $000000A4;

// Constants for enum WpsReplace
type
  WpsReplace = TOleEnum;
const
  wpsReplaceNone = $00000000;
  wpsReplaceOne = $00000001;
  wpsReplaceAll = $00000002;

// Constants for enum WpsFontBias
type
  WpsFontBias = TOleEnum;
const
  wpsFontBiasDontCare = $000000FF;
  wpsFontBiasDefault = $00000000;
  wpsFontBiasFareast = $00000001;

// Constants for enum WpsDocumentMedium
type
  WpsDocumentMedium = TOleEnum;
const
  wpsEmailMessage = $00000000;
  wpsDocument = $00000001;
  wpsWebPage = $00000002;

// Constants for enum WpsFarEastLineBreakLanguageID
type
  WpsFarEastLineBreakLanguageID = TOleEnum;
const
  wpsLineBreakJapanese = $00000411;
  wpsLineBreakKorean = $00000412;
  wpsLineBreakSimplifiedChinese = $00000804;
  wpsLineBreakTraditionalChinese = $00000404;

// Constants for enum WpsDisableFeaturesIntroducedAfter
type
  WpsDisableFeaturesIntroducedAfter = TOleEnum;
const
  wps70 = $00000000;
  wps70FE = $00000001;
  wps80 = $00000002;

// Constants for enum WpsTaskPanes
type
  WpsTaskPanes = TOleEnum;
const
  wpsTaskPaneDocumentActions = $00000000;
  wpsTaskPaneClipArt = $00000001;
  wpsTaskPaneFormatting = $00000002;
  wpsTaskPaneBackupManage = $00000003;

// Constants for enum WpsPasteType
type
  WpsPasteType = TOleEnum;
const
  wpsPaste_Default = $00000000;
  wpsPaste_Text = $00000001;
  wpsPaste_FormatText = $00000002;
  wpsPaste_MatchingFormat = $00000003;

// Constants for enum WpsEditorType
type
  WpsEditorType = TOleEnum;
const
  wpsEditorEveryone = $FFFFFFFF;
  wpsEditorOwners = $FFFFFFFC;
  wpsEditorEditors = $FFFFFFFB;
  wpsEditorCurrent = $FFFFFFFA;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0001
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0001 = TOleEnum;
const
  wdNoMailSystem = $00000000;
  wdMAPI = $00000001;
  wdPowerTalk = $00000002;
  wdMAPIandPowerTalk = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0002
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0002 = TOleEnum;
const
  wdNormalTemplate = $00000000;
  wdGlobalTemplate = $00000001;
  wdAttachedTemplate = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0003
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0003 = TOleEnum;
const
  wdContinueDisabled = $00000000;
  wdResetList = $00000001;
  wdContinueList = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0004
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0004 = TOleEnum;
const
  wdIMEModeNoControl = $00000000;
  wdIMEModeOn = $00000001;
  wdIMEModeOff = $00000002;
  wdIMEModeHiragana = $00000004;
  wdIMEModeKatakana = $00000005;
  wdIMEModeKatakanaHalf = $00000006;
  wdIMEModeAlphaFull = $00000007;
  wdIMEModeAlpha = $00000008;
  wdIMEModeHangulFull = $00000009;
  wdIMEModeHangul = $0000000A;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0005
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0005 = TOleEnum;
const
  wdBaselineAlignTop = $00000000;
  wdBaselineAlignCenter = $00000001;
  wdBaselineAlignBaseline = $00000002;
  wdBaselineAlignFarEast50 = $00000003;
  wdBaselineAlignAuto = $00000004;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0006
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0006 = TOleEnum;
const
  wdIndexFilterNone = $00000000;
  wdIndexFilterAiueo = $00000001;
  wdIndexFilterAkasatana = $00000002;
  wdIndexFilterChosung = $00000003;
  wdIndexFilterLow = $00000004;
  wdIndexFilterMedium = $00000005;
  wdIndexFilterFull = $00000006;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0007
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0007 = TOleEnum;
const
  wdIndexSortByStroke = $00000000;
  wdIndexSortBySyllable = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0008
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0008 = TOleEnum;
const
  wdJustificationModeExpand = $00000000;
  wdJustificationModeCompress = $00000001;
  wdJustificationModeCompressKana = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0009
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0009 = TOleEnum;
const
  wdFarEastLineBreakLevelNormal = $00000000;
  wdFarEastLineBreakLevelStrict = $00000001;
  wdFarEastLineBreakLevelCustom = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0010
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0010 = TOleEnum;
const
  wdHangulToHanja = $00000000;
  wdHanjaToHangul = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0011
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0011 = TOleEnum;
const
  wdAuto = $00000000;
  wdBlack = $00000001;
  wdBlue = $00000002;
  wdTurquoise = $00000003;
  wdBrightGreen = $00000004;
  wdPink = $00000005;
  wdRed = $00000006;
  wdYellow = $00000007;
  wdWhite = $00000008;
  wdDarkBlue = $00000009;
  wdTeal = $0000000A;
  wdGreen = $0000000B;
  wdViolet = $0000000C;
  wdDarkRed = $0000000D;
  wdDarkYellow = $0000000E;
  wdGray50 = $0000000F;
  wdGray25 = $00000010;
  wdByAuthor = $FFFFFFFF;
  wdNoHighlight = $00000000;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0012
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0012 = TOleEnum;
const
  wdTextureNone = $00000000;
  wdTexture2Pt5Percent = $00000019;
  wdTexture5Percent = $00000032;
  wdTexture7Pt5Percent = $0000004B;
  wdTexture10Percent = $00000064;
  wdTexture12Pt5Percent = $0000007D;
  wdTexture15Percent = $00000096;
  wdTexture17Pt5Percent = $000000AF;
  wdTexture20Percent = $000000C8;
  wdTexture22Pt5Percent = $000000E1;
  wdTexture25Percent = $000000FA;
  wdTexture27Pt5Percent = $00000113;
  wdTexture30Percent = $0000012C;
  wdTexture32Pt5Percent = $00000145;
  wdTexture35Percent = $0000015E;
  wdTexture37Pt5Percent = $00000177;
  wdTexture40Percent = $00000190;
  wdTexture42Pt5Percent = $000001A9;
  wdTexture45Percent = $000001C2;
  wdTexture47Pt5Percent = $000001DB;
  wdTexture50Percent = $000001F4;
  wdTexture52Pt5Percent = $0000020D;
  wdTexture55Percent = $00000226;
  wdTexture57Pt5Percent = $0000023F;
  wdTexture60Percent = $00000258;
  wdTexture62Pt5Percent = $00000271;
  wdTexture65Percent = $0000028A;
  wdTexture67Pt5Percent = $000002A3;
  wdTexture70Percent = $000002BC;
  wdTexture72Pt5Percent = $000002D5;
  wdTexture75Percent = $000002EE;
  wdTexture77Pt5Percent = $00000307;
  wdTexture80Percent = $00000320;
  wdTexture82Pt5Percent = $00000339;
  wdTexture85Percent = $00000352;
  wdTexture87Pt5Percent = $0000036B;
  wdTexture90Percent = $00000384;
  wdTexture92Pt5Percent = $0000039D;
  wdTexture95Percent = $000003B6;
  wdTexture97Pt5Percent = $000003CF;
  wdTextureSolid = $000003E8;
  wdTextureDarkHorizontal = $FFFFFFFF;
  wdTextureDarkVertical = $FFFFFFFE;
  wdTextureDarkDiagonalDown = $FFFFFFFD;
  wdTextureDarkDiagonalUp = $FFFFFFFC;
  wdTextureDarkCross = $FFFFFFFB;
  wdTextureDarkDiagonalCross = $FFFFFFFA;
  wdTextureHorizontal = $FFFFFFF9;
  wdTextureVertical = $FFFFFFF8;
  wdTextureDiagonalDown = $FFFFFFF7;
  wdTextureDiagonalUp = $FFFFFFF6;
  wdTextureCross = $FFFFFFF5;
  wdTextureDiagonalCross = $FFFFFFF4;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0013
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0013 = TOleEnum;
const
  wdUnderlineNone = $00000000;
  wdUnderlineSingle = $00000001;
  wdUnderlineWords = $00000002;
  wdUnderlineDouble = $00000003;
  wdUnderlineDotted = $00000004;
  wdUnderlineThick = $00000006;
  wdUnderlineDash = $00000007;
  wdUnderlineDotDash = $00000009;
  wdUnderlineDotDotDash = $0000000A;
  wdUnderlineWavy = $0000000B;
  wdUnderlineWavyHeavy = $0000001B;
  wdUnderlineDottedHeavy = $00000014;
  wdUnderlineDashHeavy = $00000017;
  wdUnderlineDotDashHeavy = $00000019;
  wdUnderlineDotDotDashHeavy = $0000001A;
  wdUnderlineDashLong = $00000027;
  wdUnderlineDashLongHeavy = $00000037;
  wdUnderlineWavyDouble = $0000002B;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0014
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0014 = TOleEnum;
const
  wdEmphasisMarkNone = $00000000;
  wdEmphasisMarkOverSolidCircle = $00000001;
  wdEmphasisMarkOverComma = $00000002;
  wdEmphasisMarkOverWhiteCircle = $00000003;
  wdEmphasisMarkUnderSolidCircle = $00000004;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0015
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0015 = TOleEnum;
const
  wdListSeparator = $00000011;
  wdDecimalSeparator = $00000012;
  wdThousandsSeparator = $00000013;
  wdCurrencyCode = $00000014;
  wd24HourClock = $00000015;
  wdInternationalAM = $00000016;
  wdInternationalPM = $00000017;
  wdTimeSeparator = $00000018;
  wdDateSeparator = $00000019;
  wdProductLanguageID = $0000001A;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0016
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0016 = TOleEnum;
const
  wdAutoExec = $00000000;
  wdAutoNew = $00000001;
  wdAutoOpen = $00000002;
  wdAutoClose = $00000003;
  wdAutoExit = $00000004;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0017
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0017 = TOleEnum;
const
  wdCaptionPositionAbove = $00000000;
  wdCaptionPositionBelow = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0018
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0018 = TOleEnum;
const
  wdUS = $00000001;
  wdCanada = $00000002;
  wdLatinAmerica = $00000003;
  wdNetherlands = $0000001F;
  wdFrance = $00000021;
  wdSpain = $00000022;
  wdItaly = $00000027;
  wdUK = $0000002C;
  wdDenmark = $0000002D;
  wdSweden = $0000002E;
  wdNorway = $0000002F;
  wdGermany = $00000031;
  wdPeru = $00000033;
  wdMexico = $00000034;
  wdArgentina = $00000036;
  wdBrazil = $00000037;
  wdChile = $00000038;
  wdVenezuela = $0000003A;
  wdJapan = $00000051;
  wdTaiwan = $00000376;
  wdChina = $00000056;
  wdKorea = $00000052;
  wdFinland = $00000166;
  wdIceland = $00000162;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0019
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0019 = TOleEnum;
const
  wdHeadingSeparatorNone = $00000000;
  wdHeadingSeparatorBlankLine = $00000001;
  wdHeadingSeparatorLetter = $00000002;
  wdHeadingSeparatorLetterLow = $00000003;
  wdHeadingSeparatorLetterFull = $00000004;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0020
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0020 = TOleEnum;
const
  wdSeparatorHyphen = $00000000;
  wdSeparatorPeriod = $00000001;
  wdSeparatorColon = $00000002;
  wdSeparatorEmDash = $00000003;
  wdSeparatorEnDash = $00000004;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0021
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0021 = TOleEnum;
const
  wdAlignPageNumberLeft = $00000000;
  wdAlignPageNumberCenter = $00000001;
  wdAlignPageNumberRight = $00000002;
  wdAlignPageNumberInside = $00000003;
  wdAlignPageNumberOutside = $00000004;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0022
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0022 = TOleEnum;
const
  wdBorderTop = $FFFFFFFF;
  wdBorderLeft = $FFFFFFFE;
  wdBorderBottom = $FFFFFFFD;
  wdBorderRight = $FFFFFFFC;
  wdBorderHorizontal = $FFFFFFFB;
  wdBorderVertical = $FFFFFFFA;
  wdBorderDiagonalDown = $FFFFFFF9;
  wdBorderDiagonalUp = $FFFFFFF8;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0023
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0023 = TOleEnum;
const
  wdFrameTop = $FFF0BDC1;
  wdFrameLeft = $FFF0BDC2;
  wdFrameBottom = $FFF0BDC3;
  wdFrameRight = $FFF0BDC4;
  wdFrameCenter = $FFF0BDC5;
  wdFrameInside = $FFF0BDC6;
  wdFrameOutside = $FFF0BDC7;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0024
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0024 = TOleEnum;
const
  wdAnimationNone = $00000000;
  wdAnimationLasVegasLights = $00000001;
  wdAnimationBlinkingBackground = $00000002;
  wdAnimationSparkleText = $00000003;
  wdAnimationMarchingBlackAnts = $00000004;
  wdAnimationMarchingRedAnts = $00000005;
  wdAnimationShimmer = $00000006;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0025
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0025 = TOleEnum;
const
  wdNextCase = $FFFFFFFF;
  wdLowerCase = $00000000;
  wdUpperCase = $00000001;
  wdTitleWord = $00000002;
  wdTitleSentence = $00000004;
  wdToggleCase = $00000005;
  wdHalfWidth = $00000006;
  wdFullWidth = $00000007;
  wdKatakana = $00000008;
  wdHiragana = $00000009;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0026
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0026 = TOleEnum;
const
  wdSummaryModeHighlight = $00000000;
  wdSummaryModeHideAllButSummary = $00000001;
  wdSummaryModeInsert = $00000002;
  wdSummaryModeCreateNew = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0027
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0027 = TOleEnum;
const
  wd10Sentences = $FFFFFFFE;
  wd20Sentences = $FFFFFFFD;
  wd100Words = $FFFFFFFC;
  wd500Words = $FFFFFFFB;
  wd10Percent = $FFFFFFFA;
  wd25Percent = $FFFFFFF9;
  wd50Percent = $FFFFFFF8;
  wd75Percent = $FFFFFFF7;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0028
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0028 = TOleEnum;
const
  wdStyleTypeParagraph = $00000001;
  wdStyleTypeCharacter = $00000002;
  wdStyleTypeTable = $00000003;
  wdStyleTypeList = $00000004;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0029
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0029 = TOleEnum;
const
  wdCharacter = $00000001;
  wdWord = $00000002;
  wdSentence = $00000003;
  wdParagraph = $00000004;
  wdLine = $00000005;
  wdStory = $00000006;
  wdScreen = $00000007;
  wdSection = $00000008;
  wdColumn = $00000009;
  wdRow = $0000000A;
  wdWindow = $0000000B;
  wdCell = $0000000C;
  wdCharacterFormatting = $0000000D;
  wdParagraphFormatting = $0000000E;
  wdTable = $0000000F;
  wdItem = $00000010;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0030
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0030 = TOleEnum;
const
  wdGoToBookmark = $FFFFFFFF;
  wdGoToSection = $00000000;
  wdGoToPage = $00000001;
  wdGoToTable = $00000002;
  wdGoToLine = $00000003;
  wdGoToFootnote = $00000004;
  wdGoToEndnote = $00000005;
  wdGoToComment = $00000006;
  wdGoToField = $00000007;
  wdGoToGraphic = $00000008;
  wdGoToObject = $00000009;
  wdGoToEquation = $0000000A;
  wdGoToHeading = $0000000B;
  wdGoToPercent = $0000000C;
  wdGoToSpellingError = $0000000D;
  wdGoToGrammaticalError = $0000000E;
  wdGoToProofreadingError = $0000000F;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0031
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0031 = TOleEnum;
const
  wdGoToFirst = $00000001;
  wdGoToLast = $FFFFFFFF;
  wdGoToNext = $00000002;
  wdGoToRelative = $00000002;
  wdGoToPrevious = $00000003;
  wdGoToAbsolute = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0032
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0032 = TOleEnum;
const
  wdCollapseStart = $00000001;
  wdCollapseEnd = $00000000;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0033
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0033 = TOleEnum;
const
  wdRowHeightAuto = $00000000;
  wdRowHeightAtLeast = $00000001;
  wdRowHeightExactly = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0034
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0034 = TOleEnum;
const
  wdFrameAuto = $00000000;
  wdFrameAtLeast = $00000001;
  wdFrameExact = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0035
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0035 = TOleEnum;
const
  wdInsertCellsShiftRight = $00000000;
  wdInsertCellsShiftDown = $00000001;
  wdInsertCellsEntireRow = $00000002;
  wdInsertCellsEntireColumn = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0036
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0036 = TOleEnum;
const
  wdDeleteCellsShiftLeft = $00000000;
  wdDeleteCellsShiftUp = $00000001;
  wdDeleteCellsEntireRow = $00000002;
  wdDeleteCellsEntireColumn = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0037
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0037 = TOleEnum;
const
  wdListApplyToWholeList = $00000000;
  wdListApplyToThisPointForward = $00000001;
  wdListApplyToSelection = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0038
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0038 = TOleEnum;
const
  wdAlertsNone = $00000000;
  wdAlertsMessageBox = $FFFFFFFE;
  wdAlertsAll = $FFFFFFFF;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0039
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0039 = TOleEnum;
const
  wdCursorWait = $00000000;
  wdCursorIBeam = $00000001;
  wdCursorNormal = $00000002;
  wdCursorNorthwestArrow = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0040
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0040 = TOleEnum;
const
  wdCancelDisabled = $00000000;
  wdCancelInterrupt = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0041
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0041 = TOleEnum;
const
  wdAdjustNone = $00000000;
  wdAdjustProportional = $00000001;
  wdAdjustFirstColumn = $00000002;
  wdAdjustSameWidth = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0042
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0042 = TOleEnum;
const
  wdAlignParagraphLeft = $00000000;
  wdAlignParagraphCenter = $00000001;
  wdAlignParagraphRight = $00000002;
  wdAlignParagraphJustify = $00000003;
  wdAlignParagraphDistribute = $00000004;
  wdAlignParagraphJustifyMed = $00000005;
  wdAlignParagraphJustifyHi = $00000007;
  wdAlignParagraphJustifyLow = $00000008;
  wdAlignParagraphThaiJustify = $00000009;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0043
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0043 = TOleEnum;
const
  wdListLevelAlignLeft = $00000000;
  wdListLevelAlignCenter = $00000001;
  wdListLevelAlignRight = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0044
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0044 = TOleEnum;
const
  wdAlignRowLeft = $00000000;
  wdAlignRowCenter = $00000001;
  wdAlignRowRight = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0045
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0045 = TOleEnum;
const
  wdAlignTabLeft = $00000000;
  wdAlignTabCenter = $00000001;
  wdAlignTabRight = $00000002;
  wdAlignTabDecimal = $00000003;
  wdAlignTabBar = $00000004;
  wdAlignTabList = $00000006;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0046
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0046 = TOleEnum;
const
  wdAlignVerticalTop = $00000000;
  wdAlignVerticalCenter = $00000001;
  wdAlignVerticalJustify = $00000002;
  wdAlignVerticalBottom = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0047
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0047 = TOleEnum;
const
  wdCellAlignVerticalTop = $00000000;
  wdCellAlignVerticalCenter = $00000001;
  wdCellAlignVerticalBottom = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0048
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0048 = TOleEnum;
const
  wdTrailingTab = $00000000;
  wdTrailingSpace = $00000001;
  wdTrailingNone = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0049
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0049 = TOleEnum;
const
  wdBulletGallery = $00000001;
  wdNumberGallery = $00000002;
  wdOutlineNumberGallery = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0050
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0050 = TOleEnum;
const
  wdListNumberStyleArabic = $00000000;
  wdListNumberStyleUppercaseRoman = $00000001;
  wdListNumberStyleLowercaseRoman = $00000002;
  wdListNumberStyleUppercaseLetter = $00000003;
  wdListNumberStyleLowercaseLetter = $00000004;
  wdListNumberStyleOrdinal = $00000005;
  wdListNumberStyleCardinalText = $00000006;
  wdListNumberStyleOrdinalText = $00000007;
  wdListNumberStyleKanji = $0000000A;
  wdListNumberStyleKanjiDigit = $0000000B;
  wdListNumberStyleAiueoHalfWidth = $0000000C;
  wdListNumberStyleIrohaHalfWidth = $0000000D;
  wdListNumberStyleArabicFullWidth = $0000000E;
  wdListNumberStyleKanjiTraditional = $00000010;
  wdListNumberStyleKanjiTraditional2 = $00000011;
  wdListNumberStyleNumberInCircle = $00000012;
  wdListNumberStyleAiueo = $00000014;
  wdListNumberStyleIroha = $00000015;
  wdListNumberStyleArabicLZ = $00000016;
  wdListNumberStyleBullet = $00000017;
  wdListNumberStyleGanada = $00000018;
  wdListNumberStyleChosung = $00000019;
  wdListNumberStyleGBNum1 = $0000001A;
  wdListNumberStyleGBNum2 = $0000001B;
  wdListNumberStyleGBNum3 = $0000001C;
  wdListNumberStyleGBNum4 = $0000001D;
  wdListNumberStyleZodiac1 = $0000001E;
  wdListNumberStyleZodiac2 = $0000001F;
  wdListNumberStyleZodiac3 = $00000020;
  wdListNumberStyleTradChinNum1 = $00000021;
  wdListNumberStyleTradChinNum2 = $00000022;
  wdListNumberStyleTradChinNum3 = $00000023;
  wdListNumberStyleTradChinNum4 = $00000024;
  wdListNumberStyleSimpChinNum1 = $00000025;
  wdListNumberStyleSimpChinNum2 = $00000026;
  wdListNumberStyleSimpChinNum3 = $00000027;
  wdListNumberStyleSimpChinNum4 = $00000028;
  wdListNumberStyleHanjaRead = $00000029;
  wdListNumberStyleHanjaReadDigit = $0000002A;
  wdListNumberStyleHangul = $0000002B;
  wdListNumberStyleHanja = $0000002C;
  wdListNumberStyleHebrew1 = $0000002D;
  wdListNumberStyleArabic1 = $0000002E;
  wdListNumberStyleHebrew2 = $0000002F;
  wdListNumberStyleArabic2 = $00000030;
  wdListNumberStyleHindiLetter1 = $00000031;
  wdListNumberStyleHindiLetter2 = $00000032;
  wdListNumberStyleHindiArabic = $00000033;
  wdListNumberStyleHindiCardinalText = $00000034;
  wdListNumberStyleThaiLetter = $00000035;
  wdListNumberStyleThaiArabic = $00000036;
  wdListNumberStyleThaiCardinalText = $00000037;
  wdListNumberStyleVietCardinalText = $00000038;
  wdListNumberStyleLowercaseRussian = $0000003A;
  wdListNumberStyleUppercaseRussian = $0000003B;
  wdListNumberStylePictureBullet = $000000F9;
  wdListNumberStyleLegal = $000000FD;
  wdListNumberStyleLegalLZ = $000000FE;
  wdListNumberStyleNone = $000000FF;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0051
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0051 = TOleEnum;
const
  wdNoteNumberStyleArabic = $00000000;
  wdNoteNumberStyleUppercaseRoman = $00000001;
  wdNoteNumberStyleLowercaseRoman = $00000002;
  wdNoteNumberStyleUppercaseLetter = $00000003;
  wdNoteNumberStyleLowercaseLetter = $00000004;
  wdNoteNumberStyleSymbol = $00000009;
  wdNoteNumberStyleArabicFullWidth = $0000000E;
  wdNoteNumberStyleKanji = $0000000A;
  wdNoteNumberStyleKanjiDigit = $0000000B;
  wdNoteNumberStyleKanjiTraditional = $00000010;
  wdNoteNumberStyleNumberInCircle = $00000012;
  wdNoteNumberStyleHanjaRead = $00000029;
  wdNoteNumberStyleHanjaReadDigit = $0000002A;
  wdNoteNumberStyleTradChinNum1 = $00000021;
  wdNoteNumberStyleTradChinNum2 = $00000022;
  wdNoteNumberStyleSimpChinNum1 = $00000025;
  wdNoteNumberStyleSimpChinNum2 = $00000026;
  wdNoteNumberStyleHebrewLetter1 = $0000002D;
  wdNoteNumberStyleArabicLetter1 = $0000002E;
  wdNoteNumberStyleHebrewLetter2 = $0000002F;
  wdNoteNumberStyleArabicLetter2 = $00000030;
  wdNoteNumberStyleHindiLetter1 = $00000031;
  wdNoteNumberStyleHindiLetter2 = $00000032;
  wdNoteNumberStyleHindiArabic = $00000033;
  wdNoteNumberStyleHindiCardinalText = $00000034;
  wdNoteNumberStyleThaiLetter = $00000035;
  wdNoteNumberStyleThaiArabic = $00000036;
  wdNoteNumberStyleThaiCardinalText = $00000037;
  wdNoteNumberStyleVietCardinalText = $00000038;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0052
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0052 = TOleEnum;
const
  wdCaptionNumberStyleArabic = $00000000;
  wdCaptionNumberStyleUppercaseRoman = $00000001;
  wdCaptionNumberStyleLowercaseRoman = $00000002;
  wdCaptionNumberStyleUppercaseLetter = $00000003;
  wdCaptionNumberStyleLowercaseLetter = $00000004;
  wdCaptionNumberStyleArabicFullWidth = $0000000E;
  wdCaptionNumberStyleKanji = $0000000A;
  wdCaptionNumberStyleKanjiDigit = $0000000B;
  wdCaptionNumberStyleKanjiTraditional = $00000010;
  wdCaptionNumberStyleNumberInCircle = $00000012;
  wdCaptionNumberStyleGanada = $00000018;
  wdCaptionNumberStyleChosung = $00000019;
  wdCaptionNumberStyleZodiac1 = $0000001E;
  wdCaptionNumberStyleZodiac2 = $0000001F;
  wdCaptionNumberStyleHanjaRead = $00000029;
  wdCaptionNumberStyleHanjaReadDigit = $0000002A;
  wdCaptionNumberStyleTradChinNum2 = $00000022;
  wdCaptionNumberStyleTradChinNum3 = $00000023;
  wdCaptionNumberStyleSimpChinNum2 = $00000026;
  wdCaptionNumberStyleSimpChinNum3 = $00000027;
  wdCaptionNumberStyleHebrewLetter1 = $0000002D;
  wdCaptionNumberStyleArabicLetter1 = $0000002E;
  wdCaptionNumberStyleHebrewLetter2 = $0000002F;
  wdCaptionNumberStyleArabicLetter2 = $00000030;
  wdCaptionNumberStyleHindiLetter1 = $00000031;
  wdCaptionNumberStyleHindiLetter2 = $00000032;
  wdCaptionNumberStyleHindiArabic = $00000033;
  wdCaptionNumberStyleHindiCardinalText = $00000034;
  wdCaptionNumberStyleThaiLetter = $00000035;
  wdCaptionNumberStyleThaiArabic = $00000036;
  wdCaptionNumberStyleThaiCardinalText = $00000037;
  wdCaptionNumberStyleVietCardinalText = $00000038;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0053
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0053 = TOleEnum;
const
  wdPageNumberStyleArabic = $00000000;
  wdPageNumberStyleUppercaseRoman = $00000001;
  wdPageNumberStyleLowercaseRoman = $00000002;
  wdPageNumberStyleUppercaseLetter = $00000003;
  wdPageNumberStyleLowercaseLetter = $00000004;
  wdPageNumberStyleArabicFullWidth = $0000000E;
  wdPageNumberStyleKanji = $0000000A;
  wdPageNumberStyleKanjiDigit = $0000000B;
  wdPageNumberStyleKanjiTraditional = $00000010;
  wdPageNumberStyleNumberInCircle = $00000012;
  wdPageNumberStyleHanjaRead = $00000029;
  wdPageNumberStyleHanjaReadDigit = $0000002A;
  wdPageNumberStyleTradChinNum1 = $00000021;
  wdPageNumberStyleTradChinNum2 = $00000022;
  wdPageNumberStyleSimpChinNum1 = $00000025;
  wdPageNumberStyleSimpChinNum2 = $00000026;
  wdPageNumberStyleHebrewLetter1 = $0000002D;
  wdPageNumberStyleArabicLetter1 = $0000002E;
  wdPageNumberStyleHebrewLetter2 = $0000002F;
  wdPageNumberStyleArabicLetter2 = $00000030;
  wdPageNumberStyleHindiLetter1 = $00000031;
  wdPageNumberStyleHindiLetter2 = $00000032;
  wdPageNumberStyleHindiArabic = $00000033;
  wdPageNumberStyleHindiCardinalText = $00000034;
  wdPageNumberStyleThaiLetter = $00000035;
  wdPageNumberStyleThaiArabic = $00000036;
  wdPageNumberStyleThaiCardinalText = $00000037;
  wdPageNumberStyleVietCardinalText = $00000038;
  wdPageNumberStyleNumberInDash = $00000039;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0054
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0054 = TOleEnum;
const
  wdStatisticWords = $00000000;
  wdStatisticLines = $00000001;
  wdStatisticPages = $00000002;
  wdStatisticCharacters = $00000003;
  wdStatisticParagraphs = $00000004;
  wdStatisticCharactersWithSpaces = $00000005;
  wdStatisticFarEastCharacters = $00000006;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0055
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0055 = TOleEnum;
const
  wdPropertyTitle = $00000001;
  wdPropertySubject = $00000002;
  wdPropertyAuthor = $00000003;
  wdPropertyKeywords = $00000004;
  wdPropertyComments = $00000005;
  wdPropertyTemplate = $00000006;
  wdPropertyLastAuthor = $00000007;
  wdPropertyRevision = $00000008;
  wdPropertyAppName = $00000009;
  wdPropertyTimeLastPrinted = $0000000A;
  wdPropertyTimeCreated = $0000000B;
  wdPropertyTimeLastSaved = $0000000C;
  wdPropertyVBATotalEdit = $0000000D;
  wdPropertyPages = $0000000E;
  wdPropertyWords = $0000000F;
  wdPropertyCharacters = $00000010;
  wdPropertySecurity = $00000011;
  wdPropertyCategory = $00000012;
  wdPropertyFormat = $00000013;
  wdPropertyManager = $00000014;
  wdPropertyCompany = $00000015;
  wdPropertyBytes = $00000016;
  wdPropertyLines = $00000017;
  wdPropertyParas = $00000018;
  wdPropertySlides = $00000019;
  wdPropertyNotes = $0000001A;
  wdPropertyHiddenSlides = $0000001B;
  wdPropertyMMClips = $0000001C;
  wdPropertyHyperlinkBase = $0000001D;
  wdPropertyCharsWSpaces = $0000001E;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0056
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0056 = TOleEnum;
const
  wdLineSpaceSingle = $00000000;
  wdLineSpace1pt5 = $00000001;
  wdLineSpaceDouble = $00000002;
  wdLineSpaceAtLeast = $00000003;
  wdLineSpaceExactly = $00000004;
  wdLineSpaceMultiple = $00000005;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0057
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0057 = TOleEnum;
const
  wdNumberParagraph = $00000001;
  wdNumberListNum = $00000002;
  wdNumberAllNumbers = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0058
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0058 = TOleEnum;
const
  wdListNoNumbering = $00000000;
  wdListListNumOnly = $00000001;
  wdListBullet = $00000002;
  wdListSimpleNumbering = $00000003;
  wdListOutlineNumbering = $00000004;
  wdListMixedNumbering = $00000005;
  wdListPictureBullet = $00000006;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0059
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0059 = TOleEnum;
const
  wdMainTextStory = $00000001;
  wdFootnotesStory = $00000002;
  wdEndnotesStory = $00000003;
  wdCommentsStory = $00000004;
  wdTextFrameStory = $00000005;
  wdEvenPagesHeaderStory = $00000006;
  wdPrimaryHeaderStory = $00000007;
  wdEvenPagesFooterStory = $00000008;
  wdPrimaryFooterStory = $00000009;
  wdFirstPageHeaderStory = $0000000A;
  wdFirstPageFooterStory = $0000000B;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0060
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0060 = TOleEnum;
const
  wdFormatDocument = $00000000;
  wdFormatTemplate = $00000001;
  wdFormatText = $00000002;
  wdFormatTextLineBreaks = $00000003;
  wdFormatDOSText = $00000004;
  wdFormatDOSTextLineBreaks = $00000005;
  wdFormatRTF = $00000006;
  wdFormatUnicodeText = $00000007;
  wdFormatEncodedText = $00000007;
  wdFormatHTML = $00000008;
  wdFormatWebArchive = $00000009;
  wdFormatFilteredHTML = $0000000A;
  wdFormatXML = $0000000B;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0061
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0061 = TOleEnum;
const
  wdOpenFormatAuto = $00000000;
  wdOpenFormatDocument = $00000001;
  wdOpenFormatTemplate = $00000002;
  wdOpenFormatRTF = $00000003;
  wdOpenFormatText = $00000004;
  wdOpenFormatUnicodeText = $00000005;
  wdOpenFormatEncodedText = $00000005;
  wdOpenFormatAllWord = $00000006;
  wdOpenFormatWebPages = $00000007;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0062
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0062 = TOleEnum;
const
  wdHeaderFooterPrimary = $00000001;
  wdHeaderFooterFirstPage = $00000002;
  wdHeaderFooterEvenPages = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0063
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0063 = TOleEnum;
const
  wdTOCTemplate = $00000000;
  wdTOCClassic = $00000001;
  wdTOCDistinctive = $00000002;
  wdTOCFancy = $00000003;
  wdTOCModern = $00000004;
  wdTOCFormal = $00000005;
  wdTOCSimple = $00000006;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0064
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0064 = TOleEnum;
const
  wdTOFTemplate = $00000000;
  wdTOFClassic = $00000001;
  wdTOFDistinctive = $00000002;
  wdTOFCentered = $00000003;
  wdTOFFormal = $00000004;
  wdTOFSimple = $00000005;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0065
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0065 = TOleEnum;
const
  wdTOATemplate = $00000000;
  wdTOAClassic = $00000001;
  wdTOADistinctive = $00000002;
  wdTOAFormal = $00000003;
  wdTOASimple = $00000004;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0066
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0066 = TOleEnum;
const
  wdLineStyleNone = $00000000;
  wdLineStyleSingle = $00000001;
  wdLineStyleDot = $00000002;
  wdLineStyleDashSmallGap = $00000003;
  wdLineStyleDashLargeGap = $00000004;
  wdLineStyleDashDot = $00000005;
  wdLineStyleDashDotDot = $00000006;
  wdLineStyleDouble = $00000007;
  wdLineStyleTriple = $00000008;
  wdLineStyleThinThickSmallGap = $00000009;
  wdLineStyleThickThinSmallGap = $0000000A;
  wdLineStyleThinThickThinSmallGap = $0000000B;
  wdLineStyleThinThickMedGap = $0000000C;
  wdLineStyleThickThinMedGap = $0000000D;
  wdLineStyleThinThickThinMedGap = $0000000E;
  wdLineStyleThinThickLargeGap = $0000000F;
  wdLineStyleThickThinLargeGap = $00000010;
  wdLineStyleThinThickThinLargeGap = $00000011;
  wdLineStyleSingleWavy = $00000012;
  wdLineStyleDoubleWavy = $00000013;
  wdLineStyleDashDotStroked = $00000014;
  wdLineStyleEmboss3D = $00000015;
  wdLineStyleEngrave3D = $00000016;
  wdLineStyleOutset = $00000017;
  wdLineStyleInset = $00000018;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0067
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0067 = TOleEnum;
const
  wdLineWidth025pt = $00000002;
  wdLineWidth050pt = $00000004;
  wdLineWidth075pt = $00000006;
  wdLineWidth100pt = $00000008;
  wdLineWidth150pt = $0000000C;
  wdLineWidth225pt = $00000012;
  wdLineWidth300pt = $00000018;
  wdLineWidth450pt = $00000024;
  wdLineWidth600pt = $00000030;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0068
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0068 = TOleEnum;
const
  wdParagraphBreak = $00000001;
  wdSectionBreakNextPage = $00000002;
  wdSectionBreakContinuous = $00000003;
  wdSectionBreakEvenPage = $00000004;
  wdSectionBreakOddPage = $00000005;
  wdLineBreak = $00000006;
  wdPageBreak = $00000007;
  wdColumnBreak = $00000008;
  wdLineBreakClearLeft = $00000009;
  wdLineBreakClearRight = $0000000A;
  wdTextWrappingBreak = $0000000B;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0069
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0069 = TOleEnum;
const
  wdTabLeaderSpaces = $00000000;
  wdTabLeaderDots = $00000001;
  wdTabLeaderDashes = $00000002;
  wdTabLeaderLines = $00000003;
  wdTabLeaderHeavy = $00000004;
  wdTabLeaderMiddleDot = $00000005;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0070
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0070 = TOleEnum;
const
  wdInches = $00000000;
  wdCentimeters = $00000001;
  wdMillimeters = $00000002;
  wdPoints = $00000003;
  wdPicas = $00000004;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0071
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0071 = TOleEnum;
const
  wdDropNone = $00000000;
  wdDropNormal = $00000001;
  wdDropMargin = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0072
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0072 = TOleEnum;
const
  wdRestartContinuous = $00000000;
  wdRestartSection = $00000001;
  wdRestartPage = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0073
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0073 = TOleEnum;
const
  wdBottomOfPage = $00000000;
  wdBeneathText = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0074
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0074 = TOleEnum;
const
  wdEndOfSection = $00000000;
  wdEndOfDocument = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0075
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0075 = TOleEnum;
const
  wdSortSeparateByTabs = $00000000;
  wdSortSeparateByCommas = $00000001;
  wdSortSeparateByDefaultTableSeparator = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0076
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0076 = TOleEnum;
const
  wdSeparateByParagraphs = $00000000;
  wdSeparateByTabs = $00000001;
  wdSeparateByCommas = $00000002;
  wdSeparateByDefaultListSeparator = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0077
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0077 = TOleEnum;
const
  wdSortFieldAlphanumeric = $00000000;
  wdSortFieldNumeric = $00000001;
  wdSortFieldDate = $00000002;
  wdSortFieldSyllable = $00000003;
  wdSortFieldJapanJIS = $00000004;
  wdSortFieldStroke = $00000005;
  wdSortFieldKoreaKS = $00000006;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0078
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0078 = TOleEnum;
const
  wdSortOrderAscending = $00000000;
  wdSortOrderDescending = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0079
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0079 = TOleEnum;
const
  wdTableFormatNone = $00000000;
  wdTableFormatSimple1 = $00000001;
  wdTableFormatSimple2 = $00000002;
  wdTableFormatSimple3 = $00000003;
  wdTableFormatClassic1 = $00000004;
  wdTableFormatClassic2 = $00000005;
  wdTableFormatClassic3 = $00000006;
  wdTableFormatClassic4 = $00000007;
  wdTableFormatColorful1 = $00000008;
  wdTableFormatColorful2 = $00000009;
  wdTableFormatColorful3 = $0000000A;
  wdTableFormatColumns1 = $0000000B;
  wdTableFormatColumns2 = $0000000C;
  wdTableFormatColumns3 = $0000000D;
  wdTableFormatColumns4 = $0000000E;
  wdTableFormatColumns5 = $0000000F;
  wdTableFormatGrid1 = $00000010;
  wdTableFormatGrid2 = $00000011;
  wdTableFormatGrid3 = $00000012;
  wdTableFormatGrid4 = $00000013;
  wdTableFormatGrid5 = $00000014;
  wdTableFormatGrid6 = $00000015;
  wdTableFormatGrid7 = $00000016;
  wdTableFormatGrid8 = $00000017;
  wdTableFormatList1 = $00000018;
  wdTableFormatList2 = $00000019;
  wdTableFormatList3 = $0000001A;
  wdTableFormatList4 = $0000001B;
  wdTableFormatList5 = $0000001C;
  wdTableFormatList6 = $0000001D;
  wdTableFormatList7 = $0000001E;
  wdTableFormatList8 = $0000001F;
  wdTableFormat3DEffects1 = $00000020;
  wdTableFormat3DEffects2 = $00000021;
  wdTableFormat3DEffects3 = $00000022;
  wdTableFormatContemporary = $00000023;
  wdTableFormatElegant = $00000024;
  wdTableFormatProfessional = $00000025;
  wdTableFormatSubtle1 = $00000026;
  wdTableFormatSubtle2 = $00000027;
  wdTableFormatWeb1 = $00000028;
  wdTableFormatWeb2 = $00000029;
  wdTableFormatWeb3 = $0000002A;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0080
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0080 = TOleEnum;
const
  wdTableFormatApplyBorders = $00000001;
  wdTableFormatApplyShading = $00000002;
  wdTableFormatApplyFont = $00000004;
  wdTableFormatApplyColor = $00000008;
  wdTableFormatApplyAutoFit = $00000010;
  wdTableFormatApplyHeadingRows = $00000020;
  wdTableFormatApplyLastRow = $00000040;
  wdTableFormatApplyFirstColumn = $00000080;
  wdTableFormatApplyLastColumn = $00000100;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0081
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0081 = TOleEnum;
const
  wdLanguageNone = $00000000;
  wdNoProofing = $00000400;
  wdAfrikaans = $00000436;
  wdAlbanian = $0000041C;
  wdAmharic = $0000045E;
  wdArabicAlgeria = $00001401;
  wdArabicBahrain = $00003C01;
  wdArabicEgypt = $00000C01;
  wdArabicIraq = $00000801;
  wdArabicJordan = $00002C01;
  wdArabicKuwait = $00003401;
  wdArabicLebanon = $00003001;
  wdArabicLibya = $00001001;
  wdArabicMorocco = $00001801;
  wdArabicOman = $00002001;
  wdArabicQatar = $00004001;
  wdArabic = $00000401;
  wdArabicSyria = $00002801;
  wdArabicTunisia = $00001C01;
  wdArabicUAE = $00003801;
  wdArabicYemen = $00002401;
  wdArmenian = $0000042B;
  wdAssamese = $0000044D;
  wdAzeriCyrillic = $0000082C;
  wdAzeriLatin = $0000042C;
  wdBasque = $0000042D;
  wdByelorussian = $00000423;
  wdBengali = $00000445;
  wdBulgarian = $00000402;
  wdBurmese = $00000455;
  wdCatalan = $00000403;
  wdCherokee = $0000045C;
  wdChineseHongKongSAR = $00000C04;
  wdChineseMacaoSAR = $00001404;
  wdSimplifiedChinese = $00000804;
  wdChineseSingapore = $00001004;
  wdTraditionalChinese = $00000404;
  wdCroatian = $0000041A;
  wdCzech = $00000405;
  wdDanish = $00000406;
  wdDivehi = $00000465;
  wdBelgianDutch = $00000813;
  wdDutch = $00000413;
  wdDzongkhaBhutan = $00000851;
  wdEdo = $00000466;
  wdEnglishAUS = $00000C09;
  wdEnglishBelize = $00002809;
  wdEnglishCanadian = $00001009;
  wdEnglishCaribbean = $00002409;
  wdEnglishIreland = $00001809;
  wdEnglishJamaica = $00002009;
  wdEnglishNewZealand = $00001409;
  wdEnglishPhilippines = $00003409;
  wdEnglishSouthAfrica = $00001C09;
  wdEnglishTrinidadTobago = $00002C09;
  wdEnglishUK = $00000809;
  wdEnglishUS = $00000409;
  wdEnglishZimbabwe = $00003009;
  wdEnglishIndonesia = $00003809;
  wdEstonian = $00000425;
  wdFaeroese = $00000438;
  wdFarsi = $00000429;
  wdFilipino = $00000464;
  wdFinnish = $0000040B;
  wdFulfulde = $00000467;
  wdBelgianFrench = $0000080C;
  wdFrenchCameroon = $00002C0C;
  wdFrenchCanadian = $00000C0C;
  wdFrenchCotedIvoire = $0000300C;
  wdFrench = $0000040C;
  wdFrenchLuxembourg = $0000140C;
  wdFrenchMali = $0000340C;
  wdFrenchMonaco = $0000180C;
  wdFrenchReunion = $0000200C;
  wdFrenchSenegal = $0000280C;
  wdFrenchMorocco = $0000380C;
  wdFrenchHaiti = $00003C0C;
  wdSwissFrench = $0000100C;
  wdFrenchWestIndies = $00001C0C;
  wdFrenchZaire = $0000240C;
  wdFrisianNetherlands = $00000462;
  wdGaelicIreland = $0000083C;
  wdGaelicScotland = $0000043C;
  wdGalician = $00000456;
  wdGeorgian = $00000437;
  wdGermanAustria = $00000C07;
  wdGerman = $00000407;
  wdGermanLiechtenstein = $00001407;
  wdGermanLuxembourg = $00001007;
  wdSwissGerman = $00000807;
  wdGreek = $00000408;
  wdGuarani = $00000474;
  wdGujarati = $00000447;
  wdHausa = $00000468;
  wdHawaiian = $00000475;
  wdHebrew = $0000040D;
  wdHindi = $00000439;
  wdHungarian = $0000040E;
  wdIbibio = $00000469;
  wdIcelandic = $0000040F;
  wdIgbo = $00000470;
  wdIndonesian = $00000421;
  wdInuktitut = $0000045D;
  wdItalian = $00000410;
  wdSwissItalian = $00000810;
  wdJapanese = $00000411;
  wdKannada = $0000044B;
  wdKanuri = $00000471;
  wdKashmiri = $00000460;
  wdKazakh = $0000043F;
  wdKhmer = $00000453;
  wdKirghiz = $00000440;
  wdKonkani = $00000457;
  wdKorean = $00000412;
  wdKyrgyz = $00000440;
  wdLao = $00000454;
  wdLatin = $00000476;
  wdLatvian = $00000426;
  wdLithuanian = $00000427;
  wdMacedonian = $0000042F;
  wdMalaysian = $0000043E;
  wdMalayBruneiDarussalam = $0000083E;
  wdMalayalam = $0000044C;
  wdMaltese = $0000043A;
  wdManipuri = $00000458;
  wdMarathi = $0000044E;
  wdMongolian = $00000450;
  wdNepali = $00000461;
  wdNorwegianBokmol = $00000414;
  wdNorwegianNynorsk = $00000814;
  wdOriya = $00000448;
  wdOromo = $00000472;
  wdPashto = $00000463;
  wdPolish = $00000415;
  wdBrazilianPortuguese = $00000416;
  wdPortuguese = $00000816;
  wdPunjabi = $00000446;
  wdRhaetoRomanic = $00000417;
  wdRomanianMoldova = $00000818;
  wdRomanian = $00000418;
  wdRussianMoldova = $00000819;
  wdRussian = $00000419;
  wdSamiLappish = $0000043B;
  wdSanskrit = $0000044F;
  wdSerbianCyrillic = $00000C1A;
  wdSerbianLatin = $0000081A;
  wdSinhalese = $0000045B;
  wdSindhi = $00000459;
  wdSindhiPakistan = $00000859;
  wdSlovak = $0000041B;
  wdSlovenian = $00000424;
  wdSomali = $00000477;
  wdSorbian = $0000042E;
  wdSpanishArgentina = $00002C0A;
  wdSpanishBolivia = $0000400A;
  wdSpanishChile = $0000340A;
  wdSpanishColombia = $0000240A;
  wdSpanishCostaRica = $0000140A;
  wdSpanishDominicanRepublic = $00001C0A;
  wdSpanishEcuador = $0000300A;
  wdSpanishElSalvador = $0000440A;
  wdSpanishGuatemala = $0000100A;
  wdSpanishHonduras = $0000480A;
  wdMexicanSpanish = $0000080A;
  wdSpanishNicaragua = $00004C0A;
  wdSpanishPanama = $0000180A;
  wdSpanishParaguay = $00003C0A;
  wdSpanishPeru = $0000280A;
  wdSpanishPuertoRico = $0000500A;
  wdSpanishModernSort = $00000C0A;
  wdSpanish = $0000040A;
  wdSpanishUruguay = $0000380A;
  wdSpanishVenezuela = $0000200A;
  wdSesotho = $00000430;
  wdSutu = $00000430;
  wdSwahili = $00000441;
  wdSwedishFinland = $0000081D;
  wdSwedish = $0000041D;
  wdSyriac = $0000045A;
  wdTajik = $00000428;
  wdTamazight = $0000045F;
  wdTamazightLatin = $0000085F;
  wdTamil = $00000449;
  wdTatar = $00000444;
  wdTelugu = $0000044A;
  wdThai = $0000041E;
  wdTibetan = $00000451;
  wdTigrignaEthiopic = $00000473;
  wdTigrignaEritrea = $00000873;
  wdTsonga = $00000431;
  wdTswana = $00000432;
  wdTurkish = $0000041F;
  wdTurkmen = $00000442;
  wdUkrainian = $00000422;
  wdUrdu = $00000420;
  wdUzbekCyrillic = $00000843;
  wdUzbekLatin = $00000443;
  wdVenda = $00000433;
  wdVietnamese = $0000042A;
  wdWelsh = $00000452;
  wdXhosa = $00000434;
  wdYi = $00000478;
  wdYiddish = $0000043D;
  wdYoruba = $0000046A;
  wdZulu = $00000435;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0082
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0082 = TOleEnum;
const
  wdFieldEmpty = $FFFFFFFF;
  wdFieldRef = $00000003;
  wdFieldIndexEntry = $00000004;
  wdFieldFootnoteRef = $00000005;
  wdFieldSet = $00000006;
  wdFieldIf = $00000007;
  wdFieldIndex = $00000008;
  wdFieldTOCEntry = $00000009;
  wdFieldStyleRef = $0000000A;
  wdFieldRefDoc = $0000000B;
  wdFieldSequence = $0000000C;
  wdFieldTOC = $0000000D;
  wdFieldInfo = $0000000E;
  wdFieldTitle = $0000000F;
  wdFieldSubject = $00000010;
  wdFieldAuthor = $00000011;
  wdFieldKeyWord = $00000012;
  wdFieldComments = $00000013;
  wdFieldLastSavedBy = $00000014;
  wdFieldCreateDate = $00000015;
  wdFieldSaveDate = $00000016;
  wdFieldPrintDate = $00000017;
  wdFieldRevisionNum = $00000018;
  wdFieldEditTime = $00000019;
  wdFieldNumPages = $0000001A;
  wdFieldNumWords = $0000001B;
  wdFieldNumChars = $0000001C;
  wdFieldFileName = $0000001D;
  wdFieldTemplate = $0000001E;
  wdFieldDate = $0000001F;
  wdFieldTime = $00000020;
  wdFieldPage = $00000021;
  wdFieldExpression = $00000022;
  wdFieldQuote = $00000023;
  wdFieldInclude = $00000024;
  wdFieldPageRef = $00000025;
  wdFieldAsk = $00000026;
  wdFieldFillIn = $00000027;
  wdFieldData = $00000028;
  wdFieldNext = $00000029;
  wdFieldNextIf = $0000002A;
  wdFieldSkipIf = $0000002B;
  wdFieldMergeRec = $0000002C;
  wdFieldDDE = $0000002D;
  wdFieldDDEAuto = $0000002E;
  wdFieldGlossary = $0000002F;
  wdFieldPrint = $00000030;
  wdFieldFormula = $00000031;
  wdFieldGoToButton = $00000032;
  wdFieldMacroButton = $00000033;
  wdFieldAutoNumOutline = $00000034;
  wdFieldAutoNumLegal = $00000035;
  wdFieldAutoNum = $00000036;
  wdFieldImport = $00000037;
  wdFieldLink = $00000038;
  wdFieldSymbol = $00000039;
  wdFieldEmbed = $0000003A;
  wdFieldMergeField = $0000003B;
  wdFieldUserName = $0000003C;
  wdFieldUserInitials = $0000003D;
  wdFieldUserAddress = $0000003E;
  wdFieldBarCode = $0000003F;
  wdFieldDocVariable = $00000040;
  wdFieldSection = $00000041;
  wdFieldSectionPages = $00000042;
  wdFieldIncludePicture = $00000043;
  wdFieldIncludeText = $00000044;
  wdFieldFileSize = $00000045;
  wdFieldFormTextInput = $00000046;
  wdFieldFormCheckBox = $00000047;
  wdFieldNoteRef = $00000048;
  wdFieldTOA = $00000049;
  wdFieldTOAEntry = $0000004A;
  wdFieldMergeSeq = $0000004B;
  wdFieldPrivate = $0000004D;
  wdFieldDatabase = $0000004E;
  wdFieldAutoText = $0000004F;
  wdFieldCompare = $00000050;
  wdFieldAddin = $00000051;
  wdFieldSubscriber = $00000052;
  wdFieldFormDropDown = $00000053;
  wdFieldAdvance = $00000054;
  wdFieldDocProperty = $00000055;
  wdFieldOCX = $00000057;
  wdFieldHyperlink = $00000058;
  wdFieldAutoTextList = $00000059;
  wdFieldListNum = $0000005A;
  wdFieldHTMLActiveX = $0000005B;
  wdFieldBidiOutline = $0000005C;
  wdFieldAddressBlock = $0000005D;
  wdFieldGreetingLine = $0000005E;
  wdFieldShape = $0000005F;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0083
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0083 = TOleEnum;
const
  wdStyleNormal = $FFFFFFFF;
  wdStyleEnvelopeAddress = $FFFFFFDB;
  wdStyleEnvelopeReturn = $FFFFFFDA;
  wdStyleBodyText = $FFFFFFBD;
  wdStyleHeading1 = $FFFFFFFE;
  wdStyleHeading2 = $FFFFFFFD;
  wdStyleHeading3 = $FFFFFFFC;
  wdStyleHeading4 = $FFFFFFFB;
  wdStyleHeading5 = $FFFFFFFA;
  wdStyleHeading6 = $FFFFFFF9;
  wdStyleHeading7 = $FFFFFFF8;
  wdStyleHeading8 = $FFFFFFF7;
  wdStyleHeading9 = $FFFFFFF6;
  wdStyleIndex1 = $FFFFFFF5;
  wdStyleIndex2 = $FFFFFFF4;
  wdStyleIndex3 = $FFFFFFF3;
  wdStyleIndex4 = $FFFFFFF2;
  wdStyleIndex5 = $FFFFFFF1;
  wdStyleIndex6 = $FFFFFFF0;
  wdStyleIndex7 = $FFFFFFEF;
  wdStyleIndex8 = $FFFFFFEE;
  wdStyleIndex9 = $FFFFFFED;
  wdStyleTOC1 = $FFFFFFEC;
  wdStyleTOC2 = $FFFFFFEB;
  wdStyleTOC3 = $FFFFFFEA;
  wdStyleTOC4 = $FFFFFFE9;
  wdStyleTOC5 = $FFFFFFE8;
  wdStyleTOC6 = $FFFFFFE7;
  wdStyleTOC7 = $FFFFFFE6;
  wdStyleTOC8 = $FFFFFFE5;
  wdStyleTOC9 = $FFFFFFE4;
  wdStyleNormalIndent = $FFFFFFE3;
  wdStyleFootnoteText = $FFFFFFE2;
  wdStyleCommentText = $FFFFFFE1;
  wdStyleHeader = $FFFFFFE0;
  wdStyleFooter = $FFFFFFDF;
  wdStyleIndexHeading = $FFFFFFDE;
  wdStyleCaption = $FFFFFFDD;
  wdStyleTableOfFigures = $FFFFFFDC;
  wdStyleFootnoteReference = $FFFFFFD9;
  wdStyleCommentReference = $FFFFFFD8;
  wdStyleLineNumber = $FFFFFFD7;
  wdStylePageNumber = $FFFFFFD6;
  wdStyleEndnoteReference = $FFFFFFD5;
  wdStyleEndnoteText = $FFFFFFD4;
  wdStyleTableOfAuthorities = $FFFFFFD3;
  wdStyleMacroText = $FFFFFFD2;
  wdStyleTOAHeading = $FFFFFFD1;
  wdStyleList = $FFFFFFD0;
  wdStyleListBullet = $FFFFFFCF;
  wdStyleListNumber = $FFFFFFCE;
  wdStyleList2 = $FFFFFFCD;
  wdStyleList3 = $FFFFFFCC;
  wdStyleList4 = $FFFFFFCB;
  wdStyleList5 = $FFFFFFCA;
  wdStyleListBullet2 = $FFFFFFC9;
  wdStyleListBullet3 = $FFFFFFC8;
  wdStyleListBullet4 = $FFFFFFC7;
  wdStyleListBullet5 = $FFFFFFC6;
  wdStyleListNumber2 = $FFFFFFC5;
  wdStyleListNumber3 = $FFFFFFC4;
  wdStyleListNumber4 = $FFFFFFC3;
  wdStyleListNumber5 = $FFFFFFC2;
  wdStyleTitle = $FFFFFFC1;
  wdStyleClosing = $FFFFFFC0;
  wdStyleSignature = $FFFFFFBF;
  wdStyleDefaultParagraphFont = $FFFFFFBE;
  wdStyleBodyTextIndent = $FFFFFFBC;
  wdStyleListContinue = $FFFFFFBB;
  wdStyleListContinue2 = $FFFFFFBA;
  wdStyleListContinue3 = $FFFFFFB9;
  wdStyleListContinue4 = $FFFFFFB8;
  wdStyleListContinue5 = $FFFFFFB7;
  wdStyleMessageHeader = $FFFFFFB6;
  wdStyleSubtitle = $FFFFFFB5;
  wdStyleSalutation = $FFFFFFB4;
  wdStyleDate = $FFFFFFB3;
  wdStyleBodyTextFirstIndent = $FFFFFFB2;
  wdStyleBodyTextFirstIndent2 = $FFFFFFB1;
  wdStyleNoteHeading = $FFFFFFB0;
  wdStyleBodyText2 = $FFFFFFAF;
  wdStyleBodyText3 = $FFFFFFAE;
  wdStyleBodyTextIndent2 = $FFFFFFAD;
  wdStyleBodyTextIndent3 = $FFFFFFAC;
  wdStyleBlockQuotation = $FFFFFFAB;
  wdStyleHyperlink = $FFFFFFAA;
  wdStyleHyperlinkFollowed = $FFFFFFA9;
  wdStyleStrong = $FFFFFFA8;
  wdStyleEmphasis = $FFFFFFA7;
  wdStyleNavPane = $FFFFFFA6;
  wdStylePlainText = $FFFFFFA5;
  wdStyleHtmlNormal = $FFFFFFA1;
  wdStyleHtmlAcronym = $FFFFFFA0;
  wdStyleHtmlAddress = $FFFFFF9F;
  wdStyleHtmlCite = $FFFFFF9E;
  wdStyleHtmlCode = $FFFFFF9D;
  wdStyleHtmlDfn = $FFFFFF9C;
  wdStyleHtmlKbd = $FFFFFF9B;
  wdStyleHtmlPre = $FFFFFF9A;
  wdStyleHtmlSamp = $FFFFFF99;
  wdStyleHtmlTt = $FFFFFF98;
  wdStyleHtmlVar = $FFFFFF97;
  wdStyleNormalTable = $FFFFFF96;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0084
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0084 = TOleEnum;
const
  wdDialogToolsOptionsTabView = $000000CC;
  wdDialogToolsOptionsTabGeneral = $000000CB;
  wdDialogToolsOptionsTabEdit = $000000E0;
  wdDialogToolsOptionsTabPrint = $000000D0;
  wdDialogToolsOptionsTabSave = $000000D1;
  wdDialogToolsOptionsTabProofread = $000000D3;
  wdDialogToolsOptionsTabTrackChanges = $00000182;
  wdDialogToolsOptionsTabUserInfo = $000000D5;
  wdDialogToolsOptionsTabCompatibility = $0000020D;
  wdDialogToolsOptionsTabTypography = $000002E3;
  wdDialogToolsOptionsTabFileLocations = $000000E1;
  wdDialogToolsOptionsTabFuzzy = $00000316;
  wdDialogToolsOptionsTabHangulHanjaConversion = $00000312;
  wdDialogToolsOptionsTabBidi = $00000405;
  wdDialogToolsOptionsTabSecurity = $00000551;
  wdDialogFilePageSetupTabMargins = $000249F0;
  wdDialogFilePageSetupTabPaper = $000249F1;
  wdDialogFilePageSetupTabLayout = $000249F3;
  wdDialogFilePageSetupTabCharsLines = $000249F4;
  wdDialogInsertSymbolTabSymbols = $00030D40;
  wdDialogInsertSymbolTabSpecialCharacters = $00030D41;
  wdDialogNoteOptionsTabAllFootnotes = $000493E0;
  wdDialogNoteOptionsTabAllEndnotes = $000493E1;
  wdDialogInsertIndexAndTablesTabIndex = $00061A80;
  wdDialogInsertIndexAndTablesTabTableOfContents = $00061A81;
  wdDialogInsertIndexAndTablesTabTableOfFigures = $00061A82;
  wdDialogInsertIndexAndTablesTabTableOfAuthorities = $00061A83;
  wdDialogOrganizerTabStyles = $0007A120;
  wdDialogOrganizerTabAutoText = $0007A121;
  wdDialogOrganizerTabCommandBars = $0007A122;
  wdDialogOrganizerTabMacros = $0007A123;
  wdDialogFormatFontTabFont = $000927C0;
  wdDialogFormatFontTabCharacterSpacing = $000927C1;
  wdDialogFormatFontTabAnimation = $000927C2;
  wdDialogFormatBordersAndShadingTabBorders = $000AAE60;
  wdDialogFormatBordersAndShadingTabPageBorder = $000AAE61;
  wdDialogFormatBordersAndShadingTabShading = $000AAE62;
  wdDialogToolsEnvelopesAndLabelsTabEnvelopes = $000C3500;
  wdDialogToolsEnvelopesAndLabelsTabLabels = $000C3501;
  wdDialogFormatParagraphTabIndentsAndSpacing = $000F4240;
  wdDialogFormatParagraphTabTextFlow = $000F4241;
  wdDialogFormatParagraphTabTeisai = $000F4242;
  wdDialogFormatDrawingObjectTabColorsAndLines = $00124F80;
  wdDialogFormatDrawingObjectTabSize = $00124F81;
  wdDialogFormatDrawingObjectTabPosition = $00124F82;
  wdDialogFormatDrawingObjectTabWrapping = $00124F83;
  wdDialogFormatDrawingObjectTabPicture = $00124F84;
  wdDialogFormatDrawingObjectTabTextbox = $00124F85;
  wdDialogFormatDrawingObjectTabWeb = $00124F86;
  wdDialogFormatDrawingObjectTabHR = $00124F87;
  wdDialogToolsAutoCorrectExceptionsTabFirstLetter = $00155CC0;
  wdDialogToolsAutoCorrectExceptionsTabInitialCaps = $00155CC1;
  wdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet = $00155CC2;
  wdDialogToolsAutoCorrectExceptionsTabIac = $00155CC3;
  wdDialogFormatBulletsAndNumberingTabBulleted = $0016E360;
  wdDialogFormatBulletsAndNumberingTabNumbered = $0016E361;
  wdDialogFormatBulletsAndNumberingTabOutlineNumbered = $0016E362;
  wdDialogLetterWizardTabLetterFormat = $00186A00;
  wdDialogLetterWizardTabRecipientInfo = $00186A01;
  wdDialogLetterWizardTabOtherElements = $00186A02;
  wdDialogLetterWizardTabSenderInfo = $00186A03;
  wdDialogToolsAutoManagerTabAutoCorrect = $0019F0A0;
  wdDialogToolsAutoManagerTabAutoFormatAsYouType = $0019F0A1;
  wdDialogToolsAutoManagerTabAutoText = $0019F0A2;
  wdDialogToolsAutoManagerTabAutoFormat = $0019F0A3;
  wdDialogToolsAutoManagerTabSmartTags = $0019F0A4;
  wdDialogTablePropertiesTabTable = $001B7740;
  wdDialogTablePropertiesTabRow = $001B7741;
  wdDialogTablePropertiesTabColumn = $001B7742;
  wdDialogTablePropertiesTabCell = $001B7743;
  wdDialogEmailOptionsTabSignature = $001CFDE0;
  wdDialogEmailOptionsTabStationary = $001CFDE1;
  wdDialogEmailOptionsTabQuoting = $001CFDE2;
  wdDialogWebOptionsBrowsers = $001E8480;
  wdDialogWebOptionsGeneral = $001E8480;
  wdDialogWebOptionsFiles = $001E8481;
  wdDialogWebOptionsPictures = $001E8482;
  wdDialogWebOptionsEncoding = $001E8483;
  wdDialogWebOptionsFonts = $001E8484;
  wdDialogToolsOptionsTabAcetate = $000004F2;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0085
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0085 = TOleEnum;
const
  wdDialogFilePageSetupTabPaperSize = $000249F1;
  wdDialogFilePageSetupTabPaperSource = $000249F2;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0086
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0086 = TOleEnum;
const
  wdDialogHelpAbout = $00000009;
  wdDialogHelpWordPerfectHelp = $0000000A;
  wdDialogDocumentStatistics = $0000004E;
  wdDialogFileNew = $0000004F;
  wdDialogFileOpen = $00000050;
  wdDialogMailMergeOpenDataSource = $00000051;
  wdDialogMailMergeOpenHeaderSource = $00000052;
  wdDialogFileSaveAs = $00000054;
  wdDialogFileSummaryInfo = $00000056;
  wdDialogToolsTemplates = $00000057;
  wdDialogFilePrint = $00000058;
  wdDialogFilePrintSetup = $00000061;
  wdDialogFileFind = $00000063;
  wdDialogFormatAddrFonts = $00000067;
  wdDialogEditPasteSpecial = $0000006F;
  wdDialogEditFind = $00000070;
  wdDialogEditReplace = $00000075;
  wdDialogEditStyle = $00000078;
  wdDialogEditLinks = $0000007C;
  wdDialogEditObject = $0000007D;
  wdDialogTableToText = $00000080;
  wdDialogTextToTable = $0000007F;
  wdDialogTableInsertTable = $00000081;
  wdDialogTableInsertCells = $00000082;
  wdDialogTableInsertRow = $00000083;
  wdDialogTableDeleteCells = $00000085;
  wdDialogTableSplitCells = $00000089;
  wdDialogTableRowHeight = $0000008E;
  wdDialogTableColumnWidth = $0000008F;
  wdDialogToolsCustomize = $00000098;
  wdDialogInsertBreak = $0000009F;
  wdDialogInsertSymbol = $000000A2;
  wdDialogInsertPicture = $000000A3;
  wdDialogInsertFile = $000000A4;
  wdDialogInsertDateTime = $000000A5;
  wdDialogInsertField = $000000A6;
  wdDialogInsertMergeField = $000000A7;
  wdDialogInsertBookmark = $000000A8;
  wdDialogMarkIndexEntry = $000000A9;
  wdDialogInsertIndex = $000000AA;
  wdDialogInsertTableOfContents = $000000AB;
  wdDialogInsertObject = $000000AC;
  wdDialogToolsCreateEnvelope = $000000AD;
  wdDialogFormatFont = $000000AE;
  wdDialogFormatParagraph = $000000AF;
  wdDialogFormatSectionLayout = $000000B0;
  wdDialogFormatColumns = $000000B1;
  wdDialogFileDocumentLayout = $000000B2;
  wdDialogFilePageSetup = $000000B2;
  wdDialogFormatTabs = $000000B3;
  wdDialogFormatStyle = $000000B4;
  wdDialogFormatDefineStyleFont = $000000B5;
  wdDialogFormatDefineStylePara = $000000B6;
  wdDialogFormatDefineStyleTabs = $000000B7;
  wdDialogFormatDefineStyleFrame = $000000B8;
  wdDialogFormatDefineStyleBorders = $000000B9;
  wdDialogFormatDefineStyleLang = $000000BA;
  wdDialogFormatPicture = $000000BB;
  wdDialogToolsLanguage = $000000BC;
  wdDialogFormatBordersAndShading = $000000BD;
  wdDialogFormatFrame = $000000BE;
  wdDialogToolsThesaurus = $000000C2;
  wdDialogToolsHyphenation = $000000C3;
  wdDialogToolsBulletsNumbers = $000000C4;
  wdDialogToolsHighlightChanges = $000000C5;
  wdDialogToolsRevisions = $000000C5;
  wdDialogToolsCompareDocuments = $000000C6;
  wdDialogTableSort = $000000C7;
  wdDialogToolsOptionsGeneral = $000000CB;
  wdDialogToolsOptionsView = $000000CC;
  wdDialogToolsAdvancedSettings = $000000CE;
  wdDialogToolsOptionsPrint = $000000D0;
  wdDialogToolsOptionsSave = $000000D1;
  wdDialogToolsOptionsSpellingAndGrammar = $000000D3;
  wdDialogToolsOptionsUserInfo = $000000D5;
  wdDialogToolsMacroRecord = $000000D6;
  wdDialogToolsMacro = $000000D7;
  wdDialogWindowActivate = $000000DC;
  wdDialogFormatRetAddrFonts = $000000DD;
  wdDialogOrganizer = $000000DE;
  wdDialogToolsOptionsEdit = $000000E0;
  wdDialogToolsOptionsFileLocations = $000000E1;
  wdDialogToolsWordCount = $000000E4;
  wdDialogControlRun = $000000EB;
  wdDialogInsertPageNumbers = $00000126;
  wdDialogFormatPageNumber = $0000012A;
  wdDialogCopyFile = $0000012C;
  wdDialogFormatChangeCase = $00000142;
  wdDialogUpdateTOC = $0000014B;
  wdDialogInsertDatabase = $00000155;
  wdDialogTableFormula = $0000015C;
  wdDialogFormFieldOptions = $00000161;
  wdDialogInsertCaption = $00000165;
  wdDialogInsertCaptionNumbering = $00000166;
  wdDialogInsertAutoCaption = $00000167;
  wdDialogFormFieldHelp = $00000169;
  wdDialogInsertCrossReference = $0000016F;
  wdDialogInsertFootnote = $00000172;
  wdDialogNoteOptions = $00000175;
  wdDialogToolsAutoCorrect = $0000017A;
  wdDialogToolsOptionsTrackChanges = $00000182;
  wdDialogConvertObject = $00000188;
  wdDialogInsertAddCaption = $00000192;
  wdDialogConnect = $000001A4;
  wdDialogToolsCustomizeKeyboard = $000001B0;
  wdDialogToolsCustomizeMenus = $000001B1;
  wdDialogToolsMergeDocuments = $000001B3;
  wdDialogMarkTableOfContentsEntry = $000001BA;
  wdDialogFileMacPageSetupGX = $000001BC;
  wdDialogFilePrintOneCopy = $000001BD;
  wdDialogEditFrame = $000001CA;
  wdDialogMarkCitation = $000001CF;
  wdDialogTableOfContentsOptions = $000001D6;
  wdDialogInsertTableOfAuthorities = $000001D7;
  wdDialogInsertTableOfFigures = $000001D8;
  wdDialogInsertIndexAndTables = $000001D9;
  wdDialogInsertFormField = $000001E3;
  wdDialogFormatDropCap = $000001E8;
  wdDialogToolsCreateLabels = $000001E9;
  wdDialogToolsProtectDocument = $000001F7;
  wdDialogFormatStyleGallery = $000001F9;
  wdDialogToolsAcceptRejectChanges = $000001FA;
  wdDialogHelpWordPerfectHelpOptions = $000001FF;
  wdDialogToolsUnprotectDocument = $00000209;
  wdDialogToolsOptionsCompatibility = $0000020D;
  wdDialogTableOfCaptionsOptions = $00000227;
  wdDialogTableAutoFormat = $00000233;
  wdDialogMailMergeFindRecord = $00000239;
  wdDialogReviewAfmtRevisions = $0000023A;
  wdDialogViewZoom = $00000241;
  wdDialogToolsProtectSection = $00000242;
  wdDialogFontSubstitution = $00000245;
  wdDialogInsertSubdocument = $00000247;
  wdDialogNewToolbar = $0000024A;
  wdDialogToolsEnvelopesAndLabels = $0000025F;
  wdDialogFormatCallout = $00000262;
  wdDialogTableFormatCell = $00000264;
  wdDialogToolsCustomizeMenuBar = $00000267;
  wdDialogFileRoutingSlip = $00000270;
  wdDialogEditTOACategory = $00000271;
  wdDialogToolsManageFields = $00000277;
  wdDialogDrawSnapToGrid = $00000279;
  wdDialogDrawAlign = $0000027A;
  wdDialogMailMergeCreateDataSource = $00000282;
  wdDialogMailMergeCreateHeaderSource = $00000283;
  wdDialogMailMerge = $000002A4;
  wdDialogMailMergeCheck = $000002A5;
  wdDialogMailMergeHelper = $000002A8;
  wdDialogMailMergeQueryOptions = $000002A9;
  wdDialogFileMacPageSetup = $000002AD;
  wdDialogListCommands = $000002D3;
  wdDialogEditCreatePublisher = $000002DC;
  wdDialogEditSubscribeTo = $000002DD;
  wdDialogEditPublishOptions = $000002DF;
  wdDialogEditSubscribeOptions = $000002E0;
  wdDialogFileMacCustomPageSetupGX = $000002E1;
  wdDialogToolsOptionsTypography = $000002E3;
  wdDialogToolsAutoCorrectExceptions = $000002FA;
  wdDialogToolsOptionsAutoFormatAsYouType = $0000030A;
  wdDialogMailMergeUseAddressBook = $0000030B;
  wdDialogToolsHangulHanjaConversion = $00000310;
  wdDialogToolsOptionsFuzzy = $00000316;
  wdDialogEditGoToOld = $0000032B;
  wdDialogInsertNumber = $0000032C;
  wdDialogLetterWizard = $00000335;
  wdDialogFormatBulletsAndNumbering = $00000338;
  wdDialogToolsSpellingAndGrammar = $0000033C;
  wdDialogToolsCreateDirectory = $00000341;
  wdDialogTableWrapping = $00000356;
  wdDialogFormatTheme = $00000357;
  wdDialogTableProperties = $0000035D;
  wdDialogEmailOptions = $0000035F;
  wdDialogCreateAutoText = $00000368;
  wdDialogToolsAutoSummarize = $0000036A;
  wdDialogToolsGrammarSettings = $00000375;
  wdDialogEditGoTo = $00000380;
  wdDialogWebOptions = $00000382;
  wdDialogInsertHyperlink = $0000039D;
  wdDialogToolsAutoManager = $00000393;
  wdDialogFileVersions = $000003B1;
  wdDialogToolsOptionsAutoFormat = $000003BF;
  wdDialogFormatDrawingObject = $000003C0;
  wdDialogToolsOptions = $000003CE;
  wdDialogFitText = $000003D7;
  wdDialogEditAutoText = $000003D9;
  wdDialogPhoneticGuide = $000003DA;
  wdDialogToolsDictionary = $000003DD;
  wdDialogFileSaveVersion = $000003EF;
  wdDialogToolsOptionsBidi = $00000405;
  wdDialogFrameSetProperties = $00000432;
  wdDialogTableTableOptions = $00000438;
  wdDialogTableCellOptions = $00000439;
  wdDialogIMESetDefault = $00000446;
  wdDialogTCSCTranslator = $00000484;
  wdDialogHorizontalInVertical = $00000488;
  wdDialogTwoLinesInOne = $00000489;
  wdDialogFormatEncloseCharacters = $0000048A;
  wdDialogConsistencyChecker = $00000461;
  wdDialogToolsOptionsSmartTag = $00000573;
  wdDialogFormatStylesCustom = $000004E0;
  wdDialogCSSLinks = $000004ED;
  wdDialogInsertWebComponent = $0000052C;
  wdDialogToolsOptionsEditCopyPaste = $0000054C;
  wdDialogToolsOptionsSecurity = $00000551;
  wdDialogSearch = $00000553;
  wdDialogShowRepairs = $00000565;
  wdDialogMailMergeInsertAsk = $00000FCF;
  wdDialogMailMergeInsertFillIn = $00000FD0;
  wdDialogMailMergeInsertIf = $00000FD1;
  wdDialogMailMergeInsertNextIf = $00000FD5;
  wdDialogMailMergeInsertSet = $00000FD6;
  wdDialogMailMergeInsertSkipIf = $00000FD7;
  wdDialogMailMergeFieldMapping = $00000518;
  wdDialogMailMergeInsertAddressBlock = $00000519;
  wdDialogMailMergeInsertGreetingLine = $0000051A;
  wdDialogMailMergeInsertFields = $0000051B;
  wdDialogMailMergeRecipients = $0000051C;
  wdDialogMailMergeFindRecipient = $0000052E;
  wdDialogMailMergeSetDocumentType = $0000053B;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0087
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0087 = TOleEnum;
const
  wdFieldKindNone = $00000000;
  wdFieldKindHot = $00000001;
  wdFieldKindWarm = $00000002;
  wdFieldKindCold = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0088
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0088 = TOleEnum;
const
  wdRegularText = $00000000;
  wdNumberText = $00000001;
  wdDateText = $00000002;
  wdCurrentDateText = $00000003;
  wdCurrentTimeText = $00000004;
  wdCalculationText = $00000005;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0089
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0089 = TOleEnum;
const
  wdNeverConvert = $00000000;
  wdAlwaysConvert = $00000001;
  wdAskToNotConvert = $00000002;
  wdAskToConvert = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0090
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0090 = TOleEnum;
const
  wdNotAMergeDocument = $FFFFFFFF;
  wdFormLetters = $00000000;
  wdMailingLabels = $00000001;
  wdEnvelopes = $00000002;
  wdCatalog = $00000003;
  wdEMail = $00000004;
  wdFax = $00000005;
  wdDirectory = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0091
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0091 = TOleEnum;
const
  wdNormalDocument = $00000000;
  wdMainDocumentOnly = $00000001;
  wdMainAndDataSource = $00000002;
  wdMainAndHeader = $00000003;
  wdMainAndSourceAndHeader = $00000004;
  wdDataSource = $00000005;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0092
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0092 = TOleEnum;
const
  wdSendToNewDocument = $00000000;
  wdSendToPrinter = $00000001;
  wdSendToEmail = $00000002;
  wdSendToFax = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0093
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0093 = TOleEnum;
const
  wdNoActiveRecord = $FFFFFFFF;
  wdNextRecord = $FFFFFFFE;
  wdPreviousRecord = $FFFFFFFD;
  wdFirstRecord = $FFFFFFFC;
  wdLastRecord = $FFFFFFFB;
  wdFirstDataSourceRecord = $FFFFFFFA;
  wdLastDataSourceRecord = $FFFFFFF9;
  wdNextDataSourceRecord = $FFFFFFF8;
  wdPreviousDataSourceRecord = $FFFFFFF7;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0094
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0094 = TOleEnum;
const
  wdDefaultFirstRecord = $00000001;
  wdDefaultLastRecord = $FFFFFFF0;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0095
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0095 = TOleEnum;
const
  wdNoMergeInfo = $FFFFFFFF;
  wdMergeInfoFromWord = $00000000;
  wdMergeInfoFromAccessDDE = $00000001;
  wdMergeInfoFromExcelDDE = $00000002;
  wdMergeInfoFromMSQueryDDE = $00000003;
  wdMergeInfoFromODBC = $00000004;
  wdMergeInfoFromODSO = $00000005;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0096
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0096 = TOleEnum;
const
  wdMergeIfEqual = $00000000;
  wdMergeIfNotEqual = $00000001;
  wdMergeIfLessThan = $00000002;
  wdMergeIfGreaterThan = $00000003;
  wdMergeIfLessThanOrEqual = $00000004;
  wdMergeIfGreaterThanOrEqual = $00000005;
  wdMergeIfIsBlank = $00000006;
  wdMergeIfIsNotBlank = $00000007;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0097
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0097 = TOleEnum;
const
  wdSortByName = $00000000;
  wdSortByLocation = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0098
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0098 = TOleEnum;
const
  wdWindowStateNormal = $00000000;
  wdWindowStateMaximize = $00000001;
  wdWindowStateMinimize = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0099
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0099 = TOleEnum;
const
  wdLinkNone = $00000000;
  wdLinkDataInDoc = $00000001;
  wdLinkDataOnDisk = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0100
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0100 = TOleEnum;
const
  wdLinkTypeOLE = $00000000;
  wdLinkTypePicture = $00000001;
  wdLinkTypeText = $00000002;
  wdLinkTypeReference = $00000003;
  wdLinkTypeInclude = $00000004;
  wdLinkTypeImport = $00000005;
  wdLinkTypeDDE = $00000006;
  wdLinkTypeDDEAuto = $00000007;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0101
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0101 = TOleEnum;
const
  wdWindowDocument = $00000000;
  wdWindowTemplate = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0102
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0102 = TOleEnum;
const
  wdNormalView = $00000001;
  wdOutlineView = $00000002;
  wdPrintView = $00000003;
  wdPrintPreview = $00000004;
  wdMasterView = $00000005;
  wdWebView = $00000006;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0103
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0103 = TOleEnum;
const
  wdSeekMainDocument = $00000000;
  wdSeekPrimaryHeader = $00000001;
  wdSeekFirstPageHeader = $00000002;
  wdSeekEvenPagesHeader = $00000003;
  wdSeekPrimaryFooter = $00000004;
  wdSeekFirstPageFooter = $00000005;
  wdSeekEvenPagesFooter = $00000006;
  wdSeekFootnotes = $00000007;
  wdSeekEndnotes = $00000008;
  wdSeekCurrentPageHeader = $00000009;
  wdSeekCurrentPageFooter = $0000000A;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0104
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0104 = TOleEnum;
const
  wdPaneNone = $00000000;
  wdPanePrimaryHeader = $00000001;
  wdPaneFirstPageHeader = $00000002;
  wdPaneEvenPagesHeader = $00000003;
  wdPanePrimaryFooter = $00000004;
  wdPaneFirstPageFooter = $00000005;
  wdPaneEvenPagesFooter = $00000006;
  wdPaneFootnotes = $00000007;
  wdPaneEndnotes = $00000008;
  wdPaneFootnoteContinuationNotice = $00000009;
  wdPaneFootnoteContinuationSeparator = $0000000A;
  wdPaneFootnoteSeparator = $0000000B;
  wdPaneEndnoteContinuationNotice = $0000000C;
  wdPaneEndnoteContinuationSeparator = $0000000D;
  wdPaneEndnoteSeparator = $0000000E;
  wdPaneComments = $0000000F;
  wdPaneCurrentPageHeader = $00000010;
  wdPaneCurrentPageFooter = $00000011;
  wdPaneRevisions = $00000012;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0105
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0105 = TOleEnum;
const
  wdPageFitNone = $00000000;
  wdPageFitFullPage = $00000001;
  wdPageFitBestFit = $00000002;
  wdPageFitTextFit = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0106
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0106 = TOleEnum;
const
  wdBrowsePage = $00000001;
  wdBrowseSection = $00000002;
  wdBrowseComment = $00000003;
  wdBrowseFootnote = $00000004;
  wdBrowseEndnote = $00000005;
  wdBrowseField = $00000006;
  wdBrowseTable = $00000007;
  wdBrowseGraphic = $00000008;
  wdBrowseHeading = $00000009;
  wdBrowseEdit = $0000000A;
  wdBrowseFind = $0000000B;
  wdBrowseGoTo = $0000000C;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0107
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0107 = TOleEnum;
const
  wdPrinterDefaultBin = $00000000;
  wdPrinterUpperBin = $00000001;
  wdPrinterOnlyBin = $00000001;
  wdPrinterLowerBin = $00000002;
  wdPrinterMiddleBin = $00000003;
  wdPrinterManualFeed = $00000004;
  wdPrinterEnvelopeFeed = $00000005;
  wdPrinterManualEnvelopeFeed = $00000006;
  wdPrinterAutomaticSheetFeed = $00000007;
  wdPrinterTractorFeed = $00000008;
  wdPrinterSmallFormatBin = $00000009;
  wdPrinterLargeFormatBin = $0000000A;
  wdPrinterLargeCapacityBin = $0000000B;
  wdPrinterPaperCassette = $0000000E;
  wdPrinterFormSource = $0000000F;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0108
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0108 = TOleEnum;
const
  wdOrientPortrait = $00000000;
  wdOrientLandscape = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0109
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0109 = TOleEnum;
const
  wdNoSelection = $00000000;
  wdSelectionIP = $00000001;
  wdSelectionNormal = $00000002;
  wdSelectionFrame = $00000003;
  wdSelectionColumn = $00000004;
  wdSelectionRow = $00000005;
  wdSelectionBlock = $00000006;
  wdSelectionInlineShape = $00000007;
  wdSelectionShape = $00000008;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0110
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0110 = TOleEnum;
const
  wdCaptionFigure = $FFFFFFFF;
  wdCaptionTable = $FFFFFFFE;
  wdCaptionEquation = $FFFFFFFD;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0111
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0111 = TOleEnum;
const
  wdRefTypeNumberedItem = $00000000;
  wdRefTypeHeading = $00000001;
  wdRefTypeBookmark = $00000002;
  wdRefTypeFootnote = $00000003;
  wdRefTypeEndnote = $00000004;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0112
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0112 = TOleEnum;
const
  wdContentText = $FFFFFFFF;
  wdNumberRelativeContext = $FFFFFFFE;
  wdNumberNoContext = $FFFFFFFD;
  wdNumberFullContext = $FFFFFFFC;
  wdEntireCaption = $00000002;
  wdOnlyLabelAndNumber = $00000003;
  wdOnlyCaptionText = $00000004;
  wdFootnoteNumber = $00000005;
  wdEndnoteNumber = $00000006;
  wdPageNumber = $00000007;
  wdPosition = $0000000F;
  wdFootnoteNumberFormatted = $00000010;
  wdEndnoteNumberFormatted = $00000011;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0113
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0113 = TOleEnum;
const
  wdIndexTemplate = $00000000;
  wdIndexClassic = $00000001;
  wdIndexFancy = $00000002;
  wdIndexModern = $00000003;
  wdIndexBulleted = $00000004;
  wdIndexFormal = $00000005;
  wdIndexSimple = $00000006;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0114
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0114 = TOleEnum;
const
  wdIndexIndent = $00000000;
  wdIndexRunin = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0115
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0115 = TOleEnum;
const
  wdWrapNever = $00000000;
  wdWrapAlways = $00000001;
  wdWrapAsk = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0116
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0116 = TOleEnum;
const
  wdNoRevision = $00000000;
  wdRevisionInsert = $00000001;
  wdRevisionDelete = $00000002;
  wdRevisionProperty = $00000003;
  wdRevisionParagraphNumber = $00000004;
  wdRevisionDisplayField = $00000005;
  wdRevisionReconcile = $00000006;
  wdRevisionConflict = $00000007;
  wdRevisionStyle = $00000008;
  wdRevisionReplace = $00000009;
  wdRevisionParagraphProperty = $0000000A;
  wdRevisionTableProperty = $0000000B;
  wdRevisionSectionProperty = $0000000C;
  wdRevisionStyleDefinition = $0000000D;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0117
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0117 = TOleEnum;
const
  wdOneAfterAnother = $00000000;
  wdAllAtOnce = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0118
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0118 = TOleEnum;
const
  wdNotYetRouted = $00000000;
  wdRouteInProgress = $00000001;
  wdRouteComplete = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0119
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0119 = TOleEnum;
const
  wdSectionContinuous = $00000000;
  wdSectionNewColumn = $00000001;
  wdSectionNewPage = $00000002;
  wdSectionEvenPage = $00000003;
  wdSectionOddPage = $00000004;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0120
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0120 = TOleEnum;
const
  wdDoNotSaveChanges = $00000000;
  wdSaveChanges = $FFFFFFFF;
  wdPromptToSaveChanges = $FFFFFFFE;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0121
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0121 = TOleEnum;
const
  wdDocumentNotSpecified = $00000000;
  wdDocumentLetter = $00000001;
  wdDocumentEmail = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0122
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0122 = TOleEnum;
const
  wdTypeDocument = $00000000;
  wdTypeTemplate = $00000001;
  wdTypeFrameset = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0123
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0123 = TOleEnum;
const
  wdWordDocument = $00000000;
  wdOriginalDocumentFormat = $00000001;
  wdPromptUser = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0124
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0124 = TOleEnum;
const
  wdRelocateUp = $00000000;
  wdRelocateDown = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0125
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0125 = TOleEnum;
const
  wdInsertedTextMarkNone = $00000000;
  wdInsertedTextMarkBold = $00000001;
  wdInsertedTextMarkItalic = $00000002;
  wdInsertedTextMarkUnderline = $00000003;
  wdInsertedTextMarkDoubleUnderline = $00000004;
  wdInsertedTextMarkColorOnly = $00000005;
  wdInsertedTextMarkStrikeThrough = $00000006;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0126
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0126 = TOleEnum;
const
  wdRevisedLinesMarkNone = $00000000;
  wdRevisedLinesMarkLeftBorder = $00000001;
  wdRevisedLinesMarkRightBorder = $00000002;
  wdRevisedLinesMarkOutsideBorder = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0127
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0127 = TOleEnum;
const
  wdDeletedTextMarkHidden = $00000000;
  wdDeletedTextMarkStrikeThrough = $00000001;
  wdDeletedTextMarkCaret = $00000002;
  wdDeletedTextMarkPound = $00000003;
  wdDeletedTextMarkNone = $00000004;
  wdDeletedTextMarkBold = $00000005;
  wdDeletedTextMarkItalic = $00000006;
  wdDeletedTextMarkUnderline = $00000007;
  wdDeletedTextMarkDoubleUnderline = $00000008;
  wdDeletedTextMarkColorOnly = $00000009;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0128
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0128 = TOleEnum;
const
  wdRevisedPropertiesMarkNone = $00000000;
  wdRevisedPropertiesMarkBold = $00000001;
  wdRevisedPropertiesMarkItalic = $00000002;
  wdRevisedPropertiesMarkUnderline = $00000003;
  wdRevisedPropertiesMarkDoubleUnderline = $00000004;
  wdRevisedPropertiesMarkColorOnly = $00000005;
  wdRevisedPropertiesMarkStrikeThrough = $00000006;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0129
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0129 = TOleEnum;
const
  wdFieldShadingNever = $00000000;
  wdFieldShadingAlways = $00000001;
  wdFieldShadingWhenSelected = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0130
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0130 = TOleEnum;
const
  wdDocumentsPath = $00000000;
  wdPicturesPath = $00000001;
  wdUserTemplatesPath = $00000002;
  wdWorkgroupTemplatesPath = $00000003;
  wdUserOptionsPath = $00000004;
  wdAutoRecoverPath = $00000005;
  wdToolsPath = $00000006;
  wdTutorialPath = $00000007;
  wdStartupPath = $00000008;
  wdProgramPath = $00000009;
  wdGraphicsFiltersPath = $0000000A;
  wdTextConvertersPath = $0000000B;
  wdProofingToolsPath = $0000000C;
  wdTempFilePath = $0000000D;
  wdCurrentFolderPath = $0000000E;
  wdStyleGalleryPath = $0000000F;
  wdBorderArtPath = $00000013;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0131
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0131 = TOleEnum;
const
  wdNoTabHangIndent = $00000001;
  wdNoSpaceRaiseLower = $00000002;
  wdPrintColBlack = $00000003;
  wdWrapTrailSpaces = $00000004;
  wdNoColumnBalance = $00000005;
  wdConvMailMergeEsc = $00000006;
  wdSuppressSpBfAfterPgBrk = $00000007;
  wdSuppressTopSpacing = $00000008;
  wdOrigWordTableRules = $00000009;
  wdTransparentMetafiles = $0000000A;
  wdShowBreaksInFrames = $0000000B;
  wdSwapBordersFacingPages = $0000000C;
  wdLeaveBackslashAlone = $0000000D;
  wdExpandShiftReturn = $0000000E;
  wdDontULTrailSpace = $0000000F;
  wdDontBalanceSingleByteDoubleByteWidth = $00000010;
  wdSuppressTopSpacingMac5 = $00000011;
  wdSpacingInWholePoints = $00000012;
  wdPrintBodyTextBeforeHeader = $00000013;
  wdNoLeading = $00000014;
  wdNoSpaceForUL = $00000015;
  wdMWSmallCaps = $00000016;
  wdNoExtraLineSpacing = $00000017;
  wdTruncateFontHeight = $00000018;
  wdSubFontBySize = $00000019;
  wdUsePrinterMetrics = $0000001A;
  wdWW6BorderRules = $0000001B;
  wdExactOnTop = $0000001C;
  wdSuppressBottomSpacing = $0000001D;
  wdWPSpaceWidth = $0000001E;
  wdWPJustification = $0000001F;
  wdLineWrapLikeWord6 = $00000020;
  wdShapeLayoutLikeWW8 = $00000021;
  wdFootnoteLayoutLikeWW8 = $00000022;
  wdDontUseHTMLParagraphAutoSpacing = $00000023;
  wdDontAdjustLineHeightInTable = $00000024;
  wdForgetLastTabAlignment = $00000025;
  wdAutospaceLikeWW7 = $00000026;
  wdAlignTablesRowByRow = $00000027;
  wdLayoutRawTableWidth = $00000028;
  wdLayoutTableRowsApart = $00000029;
  wdUseWord97LineBreakingRules = $0000002A;
  wdDontBreakWrappedTables = $0000002B;
  wdDontSnapTextToGridInTableWithObjects = $0000002C;
  wdSelectFieldWithFirstOrLastCharacter = $0000002D;
  wdApplyBreakingRules = $0000002E;
  wdDontWrapTextWithPunctuation = $0000002F;
  wdDontUseAsianBreakRulesInGrid = $00000030;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0132
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0132 = TOleEnum;
const
  wdPaper10x14 = $00000000;
  wdPaper11x17 = $00000001;
  wdPaperLetter = $00000002;
  wdPaperLetterSmall = $00000003;
  wdPaperLegal = $00000004;
  wdPaperExecutive = $00000005;
  wdPaperA3 = $00000006;
  wdPaperA4 = $00000007;
  wdPaperA4Small = $00000008;
  wdPaperA5 = $00000009;
  wdPaperB4 = $0000000A;
  wdPaperB5 = $0000000B;
  wdPaperCSheet = $0000000C;
  wdPaperDSheet = $0000000D;
  wdPaperESheet = $0000000E;
  wdPaperFanfoldLegalGerman = $0000000F;
  wdPaperFanfoldStdGerman = $00000010;
  wdPaperFanfoldUS = $00000011;
  wdPaperFolio = $00000012;
  wdPaperLedger = $00000013;
  wdPaperNote = $00000014;
  wdPaperQuarto = $00000015;
  wdPaperStatement = $00000016;
  wdPaperTabloid = $00000017;
  wdPaperEnvelope9 = $00000018;
  wdPaperEnvelope10 = $00000019;
  wdPaperEnvelope11 = $0000001A;
  wdPaperEnvelope12 = $0000001B;
  wdPaperEnvelope14 = $0000001C;
  wdPaperEnvelopeB4 = $0000001D;
  wdPaperEnvelopeB5 = $0000001E;
  wdPaperEnvelopeB6 = $0000001F;
  wdPaperEnvelopeC3 = $00000020;
  wdPaperEnvelopeC4 = $00000021;
  wdPaperEnvelopeC5 = $00000022;
  wdPaperEnvelopeC6 = $00000023;
  wdPaperEnvelopeC65 = $00000024;
  wdPaperEnvelopeDL = $00000025;
  wdPaperEnvelopeItaly = $00000026;
  wdPaperEnvelopeMonarch = $00000027;
  wdPaperEnvelopePersonal = $00000028;
  wdPaperCustom = $00000029;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0133
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0133 = TOleEnum;
const
  wdCustomLabelLetter = $00000000;
  wdCustomLabelLetterLS = $00000001;
  wdCustomLabelA4 = $00000002;
  wdCustomLabelA4LS = $00000003;
  wdCustomLabelA5 = $00000004;
  wdCustomLabelA5LS = $00000005;
  wdCustomLabelB5 = $00000006;
  wdCustomLabelMini = $00000007;
  wdCustomLabelFanfold = $00000008;
  wdCustomLabelVertHalfSheet = $00000009;
  wdCustomLabelVertHalfSheetLS = $0000000A;
  wdCustomLabelHigaki = $0000000B;
  wdCustomLabelHigakiLS = $0000000C;
  wdCustomLabelB4JIS = $0000000D;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0134
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0134 = TOleEnum;
const
  wdNoProtection = $FFFFFFFF;
  wdAllowOnlyRevisions = $00000000;
  wdAllowOnlyComments = $00000001;
  wdAllowOnlyFormFields = $00000002;
  wdAllowOnlyReading = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0135
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0135 = TOleEnum;
const
  wdAdjective = $00000000;
  wdNoun = $00000001;
  wdAdverb = $00000002;
  wdVerb = $00000003;
  wdPronoun = $00000004;
  wdConjunction = $00000005;
  wdPreposition = $00000006;
  wdInterjection = $00000007;
  wdIdiom = $00000008;
  wdOther = $00000009;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0136
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0136 = TOleEnum;
const
  wdSubscriberBestFormat = $00000000;
  wdSubscriberRTF = $00000001;
  wdSubscriberText = $00000002;
  wdSubscriberPict = $00000004;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0137
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0137 = TOleEnum;
const
  wdPublisher = $00000000;
  wdSubscriber = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0138
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0138 = TOleEnum;
const
  wdCancelPublisher = $00000000;
  wdSendPublisher = $00000001;
  wdSelectPublisher = $00000002;
  wdAutomaticUpdate = $00000003;
  wdManualUpdate = $00000004;
  wdChangeAttributes = $00000005;
  wdUpdateSubscriber = $00000006;
  wdOpenSource = $00000007;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0139
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0139 = TOleEnum;
const
  wdRelativeHorizontalPositionMargin = $00000000;
  wdRelativeHorizontalPositionPage = $00000001;
  wdRelativeHorizontalPositionColumn = $00000002;
  wdRelativeHorizontalPositionCharacter = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0140
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0140 = TOleEnum;
const
  wdRelativeVerticalPositionMargin = $00000000;
  wdRelativeVerticalPositionPage = $00000001;
  wdRelativeVerticalPositionParagraph = $00000002;
  wdRelativeVerticalPositionLine = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0141
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0141 = TOleEnum;
const
  wdHelp = $00000000;
  wdHelpAbout = $00000001;
  wdHelpActiveWindow = $00000002;
  wdHelpContents = $00000003;
  wdHelpExamplesAndDemos = $00000004;
  wdHelpIndex = $00000005;
  wdHelpKeyboard = $00000006;
  wdHelpPSSHelp = $00000007;
  wdHelpQuickPreview = $00000008;
  wdHelpSearch = $00000009;
  wdHelpUsingHelp = $0000000A;
  wdHelpIchitaro = $0000000B;
  wdHelpPE2 = $0000000C;
  wdHelpHWP = $0000000D;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0142
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0142 = TOleEnum;
const
  wdKeyCategoryNil = $FFFFFFFF;
  wdKeyCategoryDisable = $00000000;
  wdKeyCategoryCommand = $00000001;
  wdKeyCategoryMacro = $00000002;
  wdKeyCategoryFont = $00000003;
  wdKeyCategoryAutoText = $00000004;
  wdKeyCategoryStyle = $00000005;
  wdKeyCategorySymbol = $00000006;
  wdKeyCategoryPrefix = $00000007;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0143
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0143 = TOleEnum;
const
  wdNoKey = $000000FF;
  wdKeyShift = $00000100;
  wdKeyControl = $00000200;
  wdKeyCommand = $00000200;
  wdKeyAlt = $00000400;
  wdKeyOption = $00000400;
  wdKeyA = $00000041;
  wdKeyB = $00000042;
  wdKeyC = $00000043;
  wdKeyD = $00000044;
  wdKeyE = $00000045;
  wdKeyF = $00000046;
  wdKeyG = $00000047;
  wdKeyH = $00000048;
  wdKeyI = $00000049;
  wdKeyJ = $0000004A;
  wdKeyK = $0000004B;
  wdKeyL = $0000004C;
  wdKeyM = $0000004D;
  wdKeyN = $0000004E;
  wdKeyO = $0000004F;
  wdKeyP = $00000050;
  wdKeyQ = $00000051;
  wdKeyR = $00000052;
  wdKeyS = $00000053;
  wdKeyT = $00000054;
  wdKeyU = $00000055;
  wdKeyV = $00000056;
  wdKeyW = $00000057;
  wdKeyX = $00000058;
  wdKeyY = $00000059;
  wdKeyZ = $0000005A;
  wdKey0 = $00000030;
  wdKey1 = $00000031;
  wdKey2 = $00000032;
  wdKey3 = $00000033;
  wdKey4 = $00000034;
  wdKey5 = $00000035;
  wdKey6 = $00000036;
  wdKey7 = $00000037;
  wdKey8 = $00000038;
  wdKey9 = $00000039;
  wdKeyBackspace = $00000008;
  wdKeyTab = $00000009;
  wdKeyNumeric5Special = $0000000C;
  wdKeyReturn = $0000000D;
  wdKeyPause = $00000013;
  wdKeyEsc = $0000001B;
  wdKeySpacebar = $00000020;
  wdKeyPageUp = $00000021;
  wdKeyPageDown = $00000022;
  wdKeyEnd = $00000023;
  wdKeyHome = $00000024;
  wdKeyInsert = $0000002D;
  wdKeyDelete = $0000002E;
  wdKeyNumeric0 = $00000060;
  wdKeyNumeric1 = $00000061;
  wdKeyNumeric2 = $00000062;
  wdKeyNumeric3 = $00000063;
  wdKeyNumeric4 = $00000064;
  wdKeyNumeric5 = $00000065;
  wdKeyNumeric6 = $00000066;
  wdKeyNumeric7 = $00000067;
  wdKeyNumeric8 = $00000068;
  wdKeyNumeric9 = $00000069;
  wdKeyNumericMultiply = $0000006A;
  wdKeyNumericAdd = $0000006B;
  wdKeyNumericSubtract = $0000006D;
  wdKeyNumericDecimal = $0000006E;
  wdKeyNumericDivide = $0000006F;
  wdKeyF1 = $00000070;
  wdKeyF2 = $00000071;
  wdKeyF3 = $00000072;
  wdKeyF4 = $00000073;
  wdKeyF5 = $00000074;
  wdKeyF6 = $00000075;
  wdKeyF7 = $00000076;
  wdKeyF8 = $00000077;
  wdKeyF9 = $00000078;
  wdKeyF10 = $00000079;
  wdKeyF11 = $0000007A;
  wdKeyF12 = $0000007B;
  wdKeyF13 = $0000007C;
  wdKeyF14 = $0000007D;
  wdKeyF15 = $0000007E;
  wdKeyF16 = $0000007F;
  wdKeyScrollLock = $00000091;
  wdKeySemiColon = $000000BA;
  wdKeyEquals = $000000BB;
  wdKeyComma = $000000BC;
  wdKeyHyphen = $000000BD;
  wdKeyPeriod = $000000BE;
  wdKeySlash = $000000BF;
  wdKeyBackSingleQuote = $000000C0;
  wdKeyOpenSquareBrace = $000000DB;
  wdKeyBackSlash = $000000DC;
  wdKeyCloseSquareBrace = $000000DD;
  wdKeySingleQuote = $000000DE;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0144
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0144 = TOleEnum;
const
  wdOLELink = $00000000;
  wdOLEEmbed = $00000001;
  wdOLEControl = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0145
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0145 = TOleEnum;
const
  wdOLEVerbPrimary = $00000000;
  wdOLEVerbShow = $FFFFFFFF;
  wdOLEVerbOpen = $FFFFFFFE;
  wdOLEVerbHide = $FFFFFFFD;
  wdOLEVerbUIActivate = $FFFFFFFC;
  wdOLEVerbInPlaceActivate = $FFFFFFFB;
  wdOLEVerbDiscardUndoState = $FFFFFFFA;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0146
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0146 = TOleEnum;
const
  wdInLine = $00000000;
  wdFloatOverText = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0147
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0147 = TOleEnum;
const
  wdLeftPortrait = $00000000;
  wdCenterPortrait = $00000001;
  wdRightPortrait = $00000002;
  wdLeftLandscape = $00000003;
  wdCenterLandscape = $00000004;
  wdRightLandscape = $00000005;
  wdLeftClockwise = $00000006;
  wdCenterClockwise = $00000007;
  wdRightClockwise = $00000008;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0148
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0148 = TOleEnum;
const
  wdFullBlock = $00000000;
  wdModifiedBlock = $00000001;
  wdSemiBlock = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0149
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0149 = TOleEnum;
const
  wdLetterTop = $00000000;
  wdLetterBottom = $00000001;
  wdLetterLeft = $00000002;
  wdLetterRight = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0150
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0150 = TOleEnum;
const
  wdSalutationInformal = $00000000;
  wdSalutationFormal = $00000001;
  wdSalutationBusiness = $00000002;
  wdSalutationOther = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0151
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0151 = TOleEnum;
const
  wdGenderFemale = $00000000;
  wdGenderMale = $00000001;
  wdGenderNeutral = $00000002;
  wdGenderUnknown = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0152
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0152 = TOleEnum;
const
  wdMove = $00000000;
  wdExtend = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0153
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0153 = TOleEnum;
const
  wdUndefined = $0098967F;
  wdToggle = $0098967E;
  wdForward = $3FFFFFFF;
  wdBackward = $C0000001;
  wdAutoPosition = $00000000;
  wdFirst = $00000001;
  wdCreatorCode = $4D535744;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0154
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0154 = TOleEnum;
const
  wdPasteOLEObject = $00000000;
  wdPasteRTF = $00000001;
  wdPasteText = $00000002;
  wdPasteMetafilePicture = $00000003;
  wdPasteBitmap = $00000004;
  wdPasteDeviceIndependentBitmap = $00000005;
  wdPasteHyperlink = $00000007;
  wdPasteShape = $00000008;
  wdPasteEnhancedMetafile = $00000009;
  wdPasteHTML = $0000000A;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0155
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0155 = TOleEnum;
const
  wdPrintDocumentContent = $00000000;
  wdPrintProperties = $00000001;
  wdPrintComments = $00000002;
  wdPrintMarkup = $00000002;
  wdPrintStyles = $00000003;
  wdPrintAutoTextEntries = $00000004;
  wdPrintKeyAssignments = $00000005;
  wdPrintEnvelope = $00000006;
  wdPrintDocumentWithMarkup = $00000007;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0156
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0156 = TOleEnum;
const
  wdPrintAllPages = $00000000;
  wdPrintOddPagesOnly = $00000001;
  wdPrintEvenPagesOnly = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0157
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0157 = TOleEnum;
const
  wdPrintAllDocument = $00000000;
  wdPrintSelection = $00000001;
  wdPrintCurrentPage = $00000002;
  wdPrintFromTo = $00000003;
  wdPrintRangeOfPages = $00000004;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0158
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0158 = TOleEnum;
const
  wdSpelling = $00000000;
  wdGrammar = $00000001;
  wdThesaurus = $00000002;
  wdHyphenation = $00000003;
  wdSpellingComplete = $00000004;
  wdSpellingCustom = $00000005;
  wdSpellingLegal = $00000006;
  wdSpellingMedical = $00000007;
  wdHangulHanjaConversion = $00000008;
  wdHangulHanjaConversionCustom = $00000009;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0159
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0159 = TOleEnum;
const
  wdSpellword = $00000000;
  wdWildcard = $00000001;
  wdAnagram = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0160
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0160 = TOleEnum;
const
  wdSpellingCorrect = $00000000;
  wdSpellingNotInDictionary = $00000001;
  wdSpellingCapitalization = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0161
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0161 = TOleEnum;
const
  wdSpellingError = $00000000;
  wdGrammaticalError = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0162
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0162 = TOleEnum;
const
  wdInlineShapeEmbeddedOLEObject = $00000001;
  wdInlineShapeLinkedOLEObject = $00000002;
  wdInlineShapePicture = $00000003;
  wdInlineShapeLinkedPicture = $00000004;
  wdInlineShapeOLEControlObject = $00000005;
  wdInlineShapeHorizontalLine = $00000006;
  wdInlineShapePictureHorizontalLine = $00000007;
  wdInlineShapeLinkedPictureHorizontalLine = $00000008;
  wdInlineShapePictureBullet = $00000009;
  wdInlineShapeScriptAnchor = $0000000A;
  wdInlineShapeOWSAnchor = $0000000B;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0163
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0163 = TOleEnum;
const
  wdTiled = $00000000;
  wdIcons = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0164
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0164 = TOleEnum;
const
  wdSelStartActive = $00000001;
  wdSelAtEOL = $00000002;
  wdSelOvertype = $00000004;
  wdSelActive = $00000008;
  wdSelReplace = $00000010;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0165
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0165 = TOleEnum;
const
  wdAutoVersionOff = $00000000;
  wdAutoVersionOnClose = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0166
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0166 = TOleEnum;
const
  wdOrganizerObjectStyles = $00000000;
  wdOrganizerObjectAutoText = $00000001;
  wdOrganizerObjectCommandBars = $00000002;
  wdOrganizerObjectProjectItems = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0167
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0167 = TOleEnum;
const
  wdMatchParagraphMark = $0001000F;
  wdMatchTabCharacter = $00000009;
  wdMatchCommentMark = $00000005;
  wdMatchAnyCharacter = $0001003F;
  wdMatchAnyDigit = $0001001F;
  wdMatchAnyLetter = $0001002F;
  wdMatchCaretCharacter = $0000000B;
  wdMatchColumnBreak = $0000000E;
  wdMatchEmDash = $00002014;
  wdMatchEnDash = $00002013;
  wdMatchEndnoteMark = $00010013;
  wdMatchField = $00000013;
  wdMatchFootnoteMark = $00010012;
  wdMatchGraphic = $00000001;
  wdMatchManualLineBreak = $0001000F;
  wdMatchManualPageBreak = $0001001C;
  wdMatchNonbreakingHyphen = $0000001E;
  wdMatchNonbreakingSpace = $000000A0;
  wdMatchOptionalHyphen = $0000001F;
  wdMatchSectionBreak = $0001002C;
  wdMatchWhiteSpace = $00010077;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0168
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0168 = TOleEnum;
const
  wdFindStop = $00000000;
  wdFindContinue = $00000001;
  wdFindAsk = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0169
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0169 = TOleEnum;
const
  wdActiveEndAdjustedPageNumber = $00000001;
  wdActiveEndSectionNumber = $00000002;
  wdActiveEndPageNumber = $00000003;
  wdNumberOfPagesInDocument = $00000004;
  wdHorizontalPositionRelativeToPage = $00000005;
  wdVerticalPositionRelativeToPage = $00000006;
  wdHorizontalPositionRelativeToTextBoundary = $00000007;
  wdVerticalPositionRelativeToTextBoundary = $00000008;
  wdFirstCharacterColumnNumber = $00000009;
  wdFirstCharacterLineNumber = $0000000A;
  wdFrameIsSelected = $0000000B;
  wdWithInTable = $0000000C;
  wdStartOfRangeRowNumber = $0000000D;
  wdEndOfRangeRowNumber = $0000000E;
  wdMaximumNumberOfRows = $0000000F;
  wdStartOfRangeColumnNumber = $00000010;
  wdEndOfRangeColumnNumber = $00000011;
  wdMaximumNumberOfColumns = $00000012;
  wdZoomPercentage = $00000013;
  wdSelectionMode = $00000014;
  wdCapsLock = $00000015;
  wdNumLock = $00000016;
  wdOverType = $00000017;
  wdRevisionMarking = $00000018;
  wdInFootnoteEndnotePane = $00000019;
  wdInCommentPane = $0000001A;
  wdInHeaderFooter = $0000001C;
  wdAtEndOfRowMarker = $0000001F;
  wdReferenceOfType = $00000020;
  wdHeaderFooterType = $00000021;
  wdInMasterDocument = $00000022;
  wdInFootnote = $00000023;
  wdInEndnote = $00000024;
  wdInWordMail = $00000025;
  wdInClipboard = $00000026;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0170
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0170 = TOleEnum;
const
  wdWrapSquare = $00000000;
  wdWrapTight = $00000001;
  wdWrapThrough = $00000002;
  wdWrapNone = $00000003;
  wdWrapTopBottom = $00000004;
  wdWrapInline = $00000007;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0171
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0171 = TOleEnum;
const
  wdWrapBoth = $00000000;
  wdWrapLeft = $00000001;
  wdWrapRight = $00000002;
  wdWrapLargest = $00000003;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0172
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0172 = TOleEnum;
const
  wdOutlineLevel1 = $00000001;
  wdOutlineLevel2 = $00000002;
  wdOutlineLevel3 = $00000003;
  wdOutlineLevel4 = $00000004;
  wdOutlineLevel5 = $00000005;
  wdOutlineLevel6 = $00000006;
  wdOutlineLevel7 = $00000007;
  wdOutlineLevel8 = $00000008;
  wdOutlineLevel9 = $00000009;
  wdOutlineLevelBodyText = $0000000A;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0173
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0173 = TOleEnum;
const
  wdTextOrientationHorizontal = $00000000;
  wdTextOrientationUpward = $00000002;
  wdTextOrientationDownward = $00000003;
  wdTextOrientationVerticalFarEast = $00000001;
  wdTextOrientationHorizontalRotatedFarEast = $00000004;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0174
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0174 = TOleEnum;
const
  wdArtApples = $00000001;
  wdArtMapleMuffins = $00000002;
  wdArtCakeSlice = $00000003;
  wdArtCandyCorn = $00000004;
  wdArtIceCreamCones = $00000005;
  wdArtChampagneBottle = $00000006;
  wdArtPartyGlass = $00000007;
  wdArtChristmasTree = $00000008;
  wdArtTrees = $00000009;
  wdArtPalmsColor = $0000000A;
  wdArtBalloons3Colors = $0000000B;
  wdArtBalloonsHotAir = $0000000C;
  wdArtPartyFavor = $0000000D;
  wdArtConfettiStreamers = $0000000E;
  wdArtHearts = $0000000F;
  wdArtHeartBalloon = $00000010;
  wdArtStars3D = $00000011;
  wdArtStarsShadowed = $00000012;
  wdArtStars = $00000013;
  wdArtSun = $00000014;
  wdArtEarth2 = $00000015;
  wdArtEarth1 = $00000016;
  wdArtPeopleHats = $00000017;
  wdArtSombrero = $00000018;
  wdArtPencils = $00000019;
  wdArtPackages = $0000001A;
  wdArtClocks = $0000001B;
  wdArtFirecrackers = $0000001C;
  wdArtRings = $0000001D;
  wdArtMapPins = $0000001E;
  wdArtConfetti = $0000001F;
  wdArtCreaturesButterfly = $00000020;
  wdArtCreaturesLadyBug = $00000021;
  wdArtCreaturesFish = $00000022;
  wdArtBirdsFlight = $00000023;
  wdArtScaredCat = $00000024;
  wdArtBats = $00000025;
  wdArtFlowersRoses = $00000026;
  wdArtFlowersRedRose = $00000027;
  wdArtPoinsettias = $00000028;
  wdArtHolly = $00000029;
  wdArtFlowersTiny = $0000002A;
  wdArtFlowersPansy = $0000002B;
  wdArtFlowersModern2 = $0000002C;
  wdArtFlowersModern1 = $0000002D;
  wdArtWhiteFlowers = $0000002E;
  wdArtVine = $0000002F;
  wdArtFlowersDaisies = $00000030;
  wdArtFlowersBlockPrint = $00000031;
  wdArtDecoArchColor = $00000032;
  wdArtFans = $00000033;
  wdArtFilm = $00000034;
  wdArtLightning1 = $00000035;
  wdArtCompass = $00000036;
  wdArtDoubleD = $00000037;
  wdArtClassicalWave = $00000038;
  wdArtShadowedSquares = $00000039;
  wdArtTwistedLines1 = $0000003A;
  wdArtWaveline = $0000003B;
  wdArtQuadrants = $0000003C;
  wdArtCheckedBarColor = $0000003D;
  wdArtSwirligig = $0000003E;
  wdArtPushPinNote1 = $0000003F;
  wdArtPushPinNote2 = $00000040;
  wdArtPumpkin1 = $00000041;
  wdArtEggsBlack = $00000042;
  wdArtCup = $00000043;
  wdArtHeartGray = $00000044;
  wdArtGingerbreadMan = $00000045;
  wdArtBabyPacifier = $00000046;
  wdArtBabyRattle = $00000047;
  wdArtCabins = $00000048;
  wdArtHouseFunky = $00000049;
  wdArtStarsBlack = $0000004A;
  wdArtSnowflakes = $0000004B;
  wdArtSnowflakeFancy = $0000004C;
  wdArtSkyrocket = $0000004D;
  wdArtSeattle = $0000004E;
  wdArtMusicNotes = $0000004F;
  wdArtPalmsBlack = $00000050;
  wdArtMapleLeaf = $00000051;
  wdArtPaperClips = $00000052;
  wdArtShorebirdTracks = $00000053;
  wdArtPeople = $00000054;
  wdArtPeopleWaving = $00000055;
  wdArtEclipsingSquares2 = $00000056;
  wdArtHypnotic = $00000057;
  wdArtDiamondsGray = $00000058;
  wdArtDecoArch = $00000059;
  wdArtDecoBlocks = $0000005A;
  wdArtCirclesLines = $0000005B;
  wdArtPapyrus = $0000005C;
  wdArtWoodwork = $0000005D;
  wdArtWeavingBraid = $0000005E;
  wdArtWeavingRibbon = $0000005F;
  wdArtWeavingAngles = $00000060;
  wdArtArchedScallops = $00000061;
  wdArtSafari = $00000062;
  wdArtCelticKnotwork = $00000063;
  wdArtCrazyMaze = $00000064;
  wdArtEclipsingSquares1 = $00000065;
  wdArtBirds = $00000066;
  wdArtFlowersTeacup = $00000067;
  wdArtNorthwest = $00000068;
  wdArtSouthwest = $00000069;
  wdArtTribal6 = $0000006A;
  wdArtTribal4 = $0000006B;
  wdArtTribal3 = $0000006C;
  wdArtTribal2 = $0000006D;
  wdArtTribal5 = $0000006E;
  wdArtXIllusions = $0000006F;
  wdArtZanyTriangles = $00000070;
  wdArtPyramids = $00000071;
  wdArtPyramidsAbove = $00000072;
  wdArtConfettiGrays = $00000073;
  wdArtConfettiOutline = $00000074;
  wdArtConfettiWhite = $00000075;
  wdArtMosaic = $00000076;
  wdArtLightning2 = $00000077;
  wdArtHeebieJeebies = $00000078;
  wdArtLightBulb = $00000079;
  wdArtGradient = $0000007A;
  wdArtTriangleParty = $0000007B;
  wdArtTwistedLines2 = $0000007C;
  wdArtMoons = $0000007D;
  wdArtOvals = $0000007E;
  wdArtDoubleDiamonds = $0000007F;
  wdArtChainLink = $00000080;
  wdArtTriangles = $00000081;
  wdArtTribal1 = $00000082;
  wdArtMarqueeToothed = $00000083;
  wdArtSharksTeeth = $00000084;
  wdArtSawtooth = $00000085;
  wdArtSawtoothGray = $00000086;
  wdArtPostageStamp = $00000087;
  wdArtWeavingStrips = $00000088;
  wdArtZigZag = $00000089;
  wdArtCrossStitch = $0000008A;
  wdArtGems = $0000008B;
  wdArtCirclesRectangles = $0000008C;
  wdArtCornerTriangles = $0000008D;
  wdArtCreaturesInsects = $0000008E;
  wdArtZigZagStitch = $0000008F;
  wdArtCheckered = $00000090;
  wdArtCheckedBarBlack = $00000091;
  wdArtMarquee = $00000092;
  wdArtBasicWhiteDots = $00000093;
  wdArtBasicWideMidline = $00000094;
  wdArtBasicWideOutline = $00000095;
  wdArtBasicWideInline = $00000096;
  wdArtBasicThinLines = $00000097;
  wdArtBasicWhiteDashes = $00000098;
  wdArtBasicWhiteSquares = $00000099;
  wdArtBasicBlackSquares = $0000009A;
  wdArtBasicBlackDashes = $0000009B;
  wdArtBasicBlackDots = $0000009C;
  wdArtStarsTop = $0000009D;
  wdArtCertificateBanner = $0000009E;
  wdArtHandmade1 = $0000009F;
  wdArtHandmade2 = $000000A0;
  wdArtTornPaper = $000000A1;
  wdArtTornPaperBlack = $000000A2;
  wdArtCouponCutoutDashes = $000000A3;
  wdArtCouponCutoutDots = $000000A4;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0175
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0175 = TOleEnum;
const
  wdBorderDistanceFromText = $00000000;
  wdBorderDistanceFromPageEdge = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0176
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0176 = TOleEnum;
const
  wdReplaceNone = $00000000;
  wdReplaceOne = $00000001;
  wdReplaceAll = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0177
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0177 = TOleEnum;
const
  wdFontBiasDontCare = $000000FF;
  wdFontBiasDefault = $00000000;
  wdFontBiasFareast = $00000001;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0178
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0178 = TOleEnum;
const
  wdEmailMessage = $00000000;
  wdDocument = $00000001;
  wdWebPage = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0179
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0179 = TOleEnum;
const
  wdLineBreakJapanese = $00000411;
  wdLineBreakKorean = $00000412;
  wdLineBreakSimplifiedChinese = $00000804;
  wdLineBreakTraditionalChinese = $00000404;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0180
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0180 = TOleEnum;
const
  wd70 = $00000000;
  wd70FE = $00000001;
  wd80 = $00000002;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0181
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0181 = TOleEnum;
const
  wdCRLF = $00000000;
  wdCROnly = $00000001;
  wdLFOnly = $00000002;
  wdLFCR = $00000003;
  wdLSPS = $00000004;

// Constants for enum __MIDL___MIDL_itf_wpsapi_0000_0000_0182
type
  __MIDL___MIDL_itf_wpsapi_0000_0000_0182 = TOleEnum;
const
  wdShowFilterStylesAvailable = $00000000;
  wdShowFilterStylesInUse = $00000001;
  wdShowFilterStylesAll = $00000002;
  wdShowFilterFormattingInUse = $00000003;
  wdShowFilterFormattingAvailable = $00000004;

// Constants for enum WdShapePosition
type
  WdShapePosition = TOleEnum;
const
  wdShapeBottom = $FFF0BDC3;
  wdShapeCenter = $FFF0BDC5;
  wdShapeInside = $FFF0BDC6;
  wdShapeLeft = $FFF0BDC2;
  wdShapeOutside = $FFF0BDC7;
  wdShapeRight = $FFF0BDC4;
  wdShapeTop = $FFF0BDC1;

// Constants for enum WdDiagonalCellType
type
  WdDiagonalCellType = TOleEnum;
const
  wdDiagonalCellType_None = $00000000;
  wdDiagonalCellType_1Row1ColumnNoBody = $00000001;
  WdDiagonalCellType_1Row1Column = $00000002;
  WdDiagonalCellType_1Row1ColumnReverse = $00000003;
  WdDiagonalCellType_1Row2ColumnNoBody = $00000004;
  WdDiagonalCellType_2Row1ColumnNoBody = $00000005;
  WdDiagonalCellType_1Row2Column = $00000006;
  WdDiagonalCellType_2Row1Column = $00000007;
  WdDiagonalCellType_2Row2ColumnNoBody = $00000008;

// Constants for enum WdPagesOrientation
type
  WdPagesOrientation = TOleEnum;
const
  wdPagesOrientation_Vertical = $00000000;
  wdPagesOrientation_Horizontal = $00000001;

// Constants for enum WdPaperOrder
type
  WdPaperOrder = TOleEnum;
const
  wdPrinterOverThenDown = $00000000;
  wdPrinterDownThenOver = $00000001;
  wdPrinterRepeat = $00000002;

// Constants for enum WdRevisionsView
type
  WdRevisionsView = TOleEnum;
const
  wdRevisionsViewFinal = $00000000;
  wdRevisionsViewOriginal = $00000001;

// Constants for enum WdTablePosition
type
  WdTablePosition = TOleEnum;
const
  wdTableTop = $FFF0BDC1;
  wdTableLeft = $FFF0BDC2;
  wdTableBottom = $FFF0BDC3;
  wdTableRight = $FFF0BDC4;
  wdTableCenter = $FFF0BDC5;
  wdTableInside = $FFF0BDC6;
  wdTableOutside = $FFF0BDC7;

// Constants for enum WdGutterStyle
type
  WdGutterStyle = TOleEnum;
const
  wdGutterPosLeft = $00000000;
  wdGutterPosTop = $00000001;
  wdGutterPosRight = $00000002;

// Constants for enum WdAutoFitBehavior
type
  WdAutoFitBehavior = TOleEnum;
const
  wdAutoFitFixed = $00000000;
  wdAutoFitContent = $00000001;
  wdAutoFitWindow = $00000002;

// Constants for enum WdTableDirection
type
  WdTableDirection = TOleEnum;
const
  wdTableDirectionRtl = $00000000;
  wdTableDirectionLtr = $00000001;

// Constants for enum WdLayoutMode
type
  WdLayoutMode = TOleEnum;
const
  wdLayoutModeDefault = $00000000;
  wdLayoutModeGrid = $00000001;
  wdLayoutModeLineGrid = $00000002;
  wdLayoutModeGenko = $00000003;

// Constants for enum WdColor
type
  WdColor = TOleEnum;
const
  wdColorAutomatic = $FF000000;
  wdColorBlack = $00000000;
  wdColorBlue = $00FF0000;
  wdColorTurquoise = $00FFFF00;
  wdColorBrightGreen = $0000FF00;
  wdColorPink = $00FF00FF;
  wdColorRed = $000000FF;
  wdColorYellow = $0000FFFF;
  wdColorWhite = $00FFFFFF;
  wdColorDarkBlue = $00800000;
  wdColorTeal = $00808000;
  wdColorGreen = $00008000;
  wdColorViolet = $00800080;
  wdColorDarkRed = $00000080;
  wdColorDarkYellow = $00008080;
  wdColorBrown = $00003399;
  wdColorOliveGreen = $00003333;
  wdColorDarkGreen = $00003300;
  wdColorDarkTeal = $00663300;
  wdColorIndigo = $00993333;
  wdColorOrange = $000066FF;
  wdColorBlueGray = $00996666;
  wdColorLightOrange = $000099FF;
  wdColorLime = $0000CC99;
  wdColorSeaGreen = $00669933;
  wdColorAqua = $00CCCC33;
  wdColorLightBlue = $00FF6633;
  wdColorGold = $0000CCFF;
  wdColorSkyBlue = $00FFCC00;
  wdColorPlum = $00663399;
  wdColorRose = $00CC99FF;
  wdColorTan = $0099CCFF;
  wdColorLightYellow = $0099FFFF;
  wdColorLightGreen = $00CCFFCC;
  wdColorLightTurquoise = $00FFFFCC;
  wdColorPaleBlue = $00FFCC99;
  wdColorLavender = $00FF99CC;
  wdColorGray05 = $00F3F3F3;
  wdColorGray10 = $00E6E6E6;
  wdColorGray125 = $00E0E0E0;
  wdColorGray15 = $00D9D9D9;
  wdColorGray20 = $00CCCCCC;
  wdColorGray25 = $00C0C0C0;
  wdColorGray30 = $00B3B3B3;
  wdColorGray35 = $00A6A6A6;
  wdColorGray375 = $00A0A0A0;
  wdColorGray40 = $00999999;
  wdColorGray45 = $008C8C8C;
  wdColorGray50 = $00808080;
  wdColorGray55 = $00737373;
  wdColorGray60 = $00666666;
  wdColorGray625 = $00606060;
  wdColorGray65 = $00595959;
  wdColorGray70 = $004C4C4C;
  wdColorGray75 = $00404040;
  wdColorGray80 = $00333333;
  wdColorGray85 = $00262626;
  wdColorGray875 = $00202020;
  wdColorGray90 = $00191919;
  wdColorGray95 = $000C0C0C;

// Constants for enum WdPhoneticGuideAlignmentType
type
  WdPhoneticGuideAlignmentType = TOleEnum;
const
  wdPhoneticGuideAlignmentCenter = $00000000;
  wdPhoneticGuideAlignmentZeroOneZero = $00000001;
  wdPhoneticGuideAlignmentOneTwoOne = $00000002;
  wdPhoneticGuideAlignmentLeft = $00000003;
  wdPhoneticGuideAlignmentRight = $00000004;
  wdPhoneticGuideAlignmentRightVertical = $00000005;

// Constants for enum WdEncloseStyle
type
  WdEncloseStyle = TOleEnum;
const
  wdEncloseStyleNone = $00000000;
  wdEncloseStyleSmall = $00000001;
  wdEncloseStyleLarge = $00000002;

// Constants for enum WdEnclosureType
type
  WdEnclosureType = TOleEnum;
const
  wdEnclosureCircle = $00000000;
  wdEnclosureSquare = $00000001;
  wdEnclosureTriangle = $00000002;
  wdEnclosureDiamond = $00000003;

// Constants for enum WdDateLanguage
type
  WdDateLanguage = TOleEnum;
const
  wdDateLanguageBidi = $0000000A;
  wdateLanguageLatin = $00000409;

// Constants for enum WdCalendarTypeBi
type
  WdCalendarTypeBi = TOleEnum;
const
  wdCalendarTypeBidi = $00000063;
  wdCalendarTypeGregorian = $00000064;

// Constants for enum WdRevisionsMode
type
  WdRevisionsMode = TOleEnum;
const
  wdBalloonRevisions = $00000000;
  wdInLineRevisions = $00000001;
  wdMixedRevisions = $00000002;

// Constants for enum WdRevisionsBalloonWidthType
type
  WdRevisionsBalloonWidthType = TOleEnum;
const
  wdBalloonWidthPercent = $00000000;
  wdBalloonWidthPoints = $00000001;

// Constants for enum WdRevisionsBalloonPrintOrientation
type
  WdRevisionsBalloonPrintOrientation = TOleEnum;
const
  wdBalloonPrintOrientationAuto = $00000000;
  wdBalloonPrintOrientationPreserve = $00000001;
  wdBalloonPrintOrientationForceLandscape = $00000002;

// Constants for enum WdRevisionsBalloonMargin
type
  WdRevisionsBalloonMargin = TOleEnum;
const
  wdLeftMargin = $00000000;
  wdRightMargin = $00000001;

// Constants for enum WdFindScope
type
  WdFindScope = TOleEnum;
const
  wdFindScope_WholeDocument = $00000000;
  wdFindScope_Selection = $00000001;
  wdFindScope_MainText = $00000002;
  wdFindScope_HeaderFooters = $00000003;
  wdFindScope_Footnotes = $00000004;
  wdFindScope_Endnotes = $00000005;
  wdFindScope_TextFrames = $00000006;
  wdFindScope_HeaderFooterTextFrames = $00000007;
  wdFindScope_Comments = $00000008;

// Constants for enum WdRecoveryType
type
  WdRecoveryType = TOleEnum;
const
  wdPasteDefault = $00000000;
  wdSingleCellText = $00000005;
  wdSingleCellTable = $00000006;
  wdListContinueNumbering = $00000007;
  wdListRestartNumbering = $00000008;
  wdTableInsertAsRows = $0000000B;
  wdTableAppendTable = $0000000A;
  wdTableOriginalFormatting = $0000000C;
  wdChartPicture = $0000000D;
  wdChart = $0000000E;
  wdChartLinked = $0000000F;
  wdFormatOriginalFormatting = $00000010;
  wdFormatSurroundingFormattingWithEmphasis = $00000014;
  wdFormatPlainText = $00000016;
  wdTableOverwriteCells = $00000017;
  wdListCombineWithExistingList = $00000018;
  wdListDontMerge = $00000019;

// Constants for enum WdPrintHiddenTextMode
type
  WdPrintHiddenTextMode = TOleEnum;
const
  wdDontPrintHiddenText = $00000000;
  wdPrintHiddenText = $00000001;
  wdPrintSpaceHiddenText = $00000002;

// Constants for enum WdGenkoGrid
type
  WdGenkoGrid = TOleEnum;
const
  wdGenkoGrid_Grid = $00000000;
  wdGenkoGrid_Underline = $00000001;
  wdGenkoGrid_Border = $00000002;

// Constants for enum WdGenkoGridStyle
type
  WdGenkoGridStyle = TOleEnum;
const
  wdGenkoGrid10x20 = $00000000;
  wdGenkoGrid15x20 = $00000001;
  wdGenkoGrid20x20 = $00000002;
  wdGenkoGrid20x25 = $00000003;
  wdGenkoGrid20x20Classical = $00000004;
  wdGenkoGrid20x25Classical = $00000005;
  wdGenkoGrid24x25Classical = $00000006;
  wdGenkoGridModel = $00000007;

// Constants for enum WdRevisionsBalloonTitle
type
  WdRevisionsBalloonTitle = TOleEnum;
const
  wdRevisionsBalloonTitleSimple = $00000000;
  wdRevisionsBalloonTitleDetail = $00000001;

// Constants for enum WdMailMergeMailFormat
type
  WdMailMergeMailFormat = TOleEnum;
const
  wdMailFormatPlainText = $00000000;
  wdMailFormatHTML = $00000001;

// Constants for enum WdMappedDataFields
type
  WdMappedDataFields = TOleEnum;
const
  wdUniqueIdentifier = $00000001;
  wdCourtesyTitle = $00000002;
  wdFirstName = $00000003;
  wdMiddleName = $00000004;
  wdLastName = $00000005;
  wdSuffix = $00000006;
  wdNickname = $00000007;
  wdJobTitle = $00000008;
  wdCompany = $00000009;
  wdAddress1 = $0000000A;
  wdAddress2 = $0000000B;
  wdCity = $0000000C;
  wdState = $0000000D;
  wdPostalCode = $0000000E;
  wdCountryRegion = $0000000F;
  wdBusinessPhone = $00000010;
  wdBusinessFax = $00000011;
  wdHomePhone = $00000012;
  wdHomeFax = $00000013;
  wdEmailAddress = $00000014;
  wdWebPageURL = $00000015;
  wdSpouseCourtesyTitle = $00000016;
  wdSpouseFirstName = $00000017;
  wdSpouseMiddleName = $00000018;
  wdSpouseLastName = $00000019;
  wdSpouseNickname = $0000001A;
  wdRubyFirstName = $0000001B;
  wdRubyLastName = $0000001C;
  wdAddress3 = $0000001D;
  wdDepartment = $0000001E;

// Constants for enum WdPdfCommentsMode
type
  WdPdfCommentsMode = TOleEnum;
const
  wdPdfIgnoreComments = $00000001;
  wdPdfPrintCommentsOnly = $00000002;
  wdPdfCompatibleComments = $00000003;

// Constants for enum WdPdfPrintRight
type
  WdPdfPrintRight = TOleEnum;
const
  wdPdfNotAllowPrint = $00000001;
  wdPdfPrintLowQulityOnly = $00000002;
  wdPdfFreePrint = $00000003;

// Constants for enum WdPdfEditRight
type
  WdPdfEditRight = TOleEnum;
const
  wdPdfAssemble = $00000001;
  wdPdfFillForm = $00000002;
  wdPdfAnnotate = $00000003;
  wdPdfModify = $00000004;

// Constants for enum WdPdfCopyRight
type
  WdPdfCopyRight = TOleEnum;
const
  wdPdfNotAllowCopy = $00000000;
  wdPdfFreeCopy = $0000FFFF;

// Constants for enum WdPdfStyleLevel
type
  WdPdfStyleLevel = TOleEnum;
const
  wdPdfDontExport = $FFFFFFFE;
  wdPdfFollowNearest = $FFFFFFFF;
  wdPdfRootBookmark = $00000000;
  wdPdf1stBookmark = $00000000;
  wdPdf2ndBookmark = $00000001;
  wdPdf3rdBookmark = $00000002;
  wdPdf4thBookmark = $00000003;
  wdPdf5thBookmark = $00000004;
  wdPdf6thBookmark = $00000005;
  wdPdf7thBookmark = $00000006;
  wdPdf8thBookmark = $00000007;
  wdPdf9thBookmark = $00000008;
  wdPdfLeafBookmark = $00000009;

// Constants for enum WdTCSCConverterDirection
type
  WdTCSCConverterDirection = TOleEnum;
const
  wdTCSCConverterDirectionSCTC = $00000000;
  wdTCSCConverterDirectionTCSC = $00000001;
  wdTCSCConverterDirectionAuto = $00000002;

// Constants for enum WdDialogTab
type
  WdDialogTab = TOleEnum;
const
  wdDialogReserved = $00000000;

// Constants for enum WdTaskPanes
type
  WdTaskPanes = TOleEnum;
const
  wdTaskPaneDocumentActions = $00000000;
  wdTaskPaneClipArt = $00000001;
  wdTaskPaneFormatting = $00000002;
  wdTaskPaneBackupManage = $00000003;

// Constants for enum WdTwoLinesInOne
type
  WdTwoLinesInOne = TOleEnum;
const
  wdTwoLinesInOneNone = $00000000;
  wdTwoLinesInOneNoBrackets = $00000001;
  wdTwoLinesInOneParentheses = $00000002;
  wdTwoLinesInOneSquareBrackets = $00000003;
  wdTwoLinesInOneAngleBrackets = $00000004;
  wdTwoLinesInOneCurlyBrackets = $00000005;

// Constants for enum WdPasteType
type
  WdPasteType = TOleEnum;
const
  wdPaste_Default = $00000000;
  wdPaste_Text = $00000001;
  wdPaste_FormatText = $00000002;
  wdPaste_MatchingFormat = $00000003;

// Constants for enum WdDocumentViewDirection
type
  WdDocumentViewDirection = TOleEnum;
const
  wdDocumentViewLtr = $00000000;
  wdDocumentViewRtl = $00000001;

// Constants for enum WdFlowDirection
type
  WdFlowDirection = TOleEnum;
const
  wdFlowLtr = $00000000;
  wdFlowRtl = $00000001;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  _IKsoDispObj = interface;
  _IKsoDispObjDisp = dispinterface;
  _Application = interface;
  _ApplicationDisp = dispinterface;
  Documents = interface;
  DocumentsDisp = dispinterface;
  _Document = interface;
  _DocumentDisp = dispinterface;
  Range = interface;
  RangeDisp = dispinterface;
  _Font = interface;
  _FontDisp = dispinterface;
  Borders = interface;
  BordersDisp = dispinterface;
  Border = interface;
  BorderDisp = dispinterface;
  Shading = interface;
  ShadingDisp = dispinterface;
  Tables = interface;
  TablesDisp = dispinterface;
  Table = interface;
  TableDisp = dispinterface;
  Columns = interface;
  ColumnsDisp = dispinterface;
  Column = interface;
  ColumnDisp = dispinterface;
  Cells = interface;
  CellsDisp = dispinterface;
  Cell = interface;
  CellDisp = dispinterface;
  Row = interface;
  RowDisp = dispinterface;
  Rows = interface;
  RowsDisp = dispinterface;
  Footnotes = interface;
  FootnotesDisp = dispinterface;
  Footnote = interface;
  FootnoteDisp = dispinterface;
  FootnoteOptions = interface;
  FootnoteOptionsDisp = dispinterface;
  Endnotes = interface;
  EndnotesDisp = dispinterface;
  Endnote = interface;
  EndnoteDisp = dispinterface;
  EndnoteOptions = interface;
  EndnoteOptionsDisp = dispinterface;
  Comments = interface;
  CommentsDisp = dispinterface;
  Comment = interface;
  CommentDisp = dispinterface;
  Sections = interface;
  SectionsDisp = dispinterface;
  Section = interface;
  SectionDisp = dispinterface;
  PageSetup = interface;
  PageSetupDisp = dispinterface;
  TextColumns = interface;
  TextColumnsDisp = dispinterface;
  TextColumn = interface;
  TextColumnDisp = dispinterface;
  HeadersFooters = interface;
  HeadersFootersDisp = dispinterface;
  HeaderFooter = interface;
  HeaderFooterDisp = dispinterface;
  PageNumbers = interface;
  PageNumbersDisp = dispinterface;
  PageNumber = interface;
  PageNumberDisp = dispinterface;
  Shapes = interface;
  ShapesDisp = dispinterface;
  Shape = interface;
  ShapeDisp = dispinterface;
  ShapeRange = interface;
  ShapeRangeDisp = dispinterface;
  WrapFormat = interface;
  WrapFormatDisp = dispinterface;
  Frame = interface;
  FrameDisp = dispinterface;
  InlineShape = interface;
  InlineShapeDisp = dispinterface;
  Field = interface;
  FieldDisp = dispinterface;
  FillFormat = interface;
  FillFormatDisp = dispinterface;
  ColorFormat = interface;
  ColorFormatDisp = dispinterface;
  PictureFormat = interface;
  PictureFormatDisp = dispinterface;
  TextEffectFormat = interface;
  TextEffectFormatDisp = dispinterface;
  OLEFormat = interface;
  OLEFormatDisp = dispinterface;
  TextFrame = interface;
  TextFrameDisp = dispinterface;
  Adjustments = interface;
  AdjustmentsDisp = dispinterface;
  CalloutFormat = interface;
  CalloutFormatDisp = dispinterface;
  ConnectorFormat = interface;
  ConnectorFormatDisp = dispinterface;
  GroupShapes = interface;
  GroupShapesDisp = dispinterface;
  LineFormat = interface;
  LineFormatDisp = dispinterface;
  ShapeNodes = interface;
  ShapeNodesDisp = dispinterface;
  ShapeNode = interface;
  ShapeNodeDisp = dispinterface;
  ShadowFormat = interface;
  ShadowFormatDisp = dispinterface;
  ThreeDFormat = interface;
  ThreeDFormatDisp = dispinterface;
  Diagram = interface;
  DiagramDisp = dispinterface;
  DiagramNodes = interface;
  DiagramNodesDisp = dispinterface;
  DiagramNode = interface;
  DiagramNodeDisp = dispinterface;
  DiagramNodeChildren = interface;
  DiagramNodeChildrenDisp = dispinterface;
  CanvasShapes = interface;
  CanvasShapesDisp = dispinterface;
  FreeformBuilder = interface;
  FreeformBuilderDisp = dispinterface;
  Paragraphs = interface;
  ParagraphsDisp = dispinterface;
  Paragraph = interface;
  ParagraphDisp = dispinterface;
  _ParagraphFormat = interface;
  _ParagraphFormatDisp = dispinterface;
  TabStops = interface;
  TabStopsDisp = dispinterface;
  TabStop = interface;
  TabStopDisp = dispinterface;
  DropCap = interface;
  DropCapDisp = dispinterface;
  Style = interface;
  StyleDisp = dispinterface;
  ListTemplate = interface;
  ListTemplateDisp = dispinterface;
  ListLevels = interface;
  ListLevelsDisp = dispinterface;
  ListLevel = interface;
  ListLevelDisp = dispinterface;
  Fields = interface;
  FieldsDisp = dispinterface;
  Frames = interface;
  FramesDisp = dispinterface;
  ListFormat = interface;
  ListFormatDisp = dispinterface;
  List = interface;
  ListDisp = dispinterface;
  ListParagraphs = interface;
  ListParagraphsDisp = dispinterface;
  Bookmarks = interface;
  BookmarksDisp = dispinterface;
  Bookmark = interface;
  BookmarkDisp = dispinterface;
  Hyperlinks = interface;
  HyperlinksDisp = dispinterface;
  Hyperlink = interface;
  HyperlinkDisp = dispinterface;
  InlineShapes = interface;
  InlineShapesDisp = dispinterface;
  WordStat = interface;
  WordStatDisp = dispinterface;
  Revisions = interface;
  RevisionsDisp = dispinterface;
  Revision = interface;
  RevisionDisp = dispinterface;
  SpellingSuggestions = interface;
  SpellingSuggestionsDisp = dispinterface;
  SpellingSuggestion = interface;
  SpellingSuggestionDisp = dispinterface;
  ProofreadingErrors = interface;
  ProofreadingErrorsDisp = dispinterface;
  FormFields = interface;
  FormFieldsDisp = dispinterface;
  FormField = interface;
  FormFieldDisp = dispinterface;
  TextInput = interface;
  TextInputDisp = dispinterface;
  CheckBox = interface;
  CheckBoxDisp = dispinterface;
  DropDown = interface;
  DropDownDisp = dispinterface;
  ListEntries = interface;
  ListEntriesDisp = dispinterface;
  ListEntry = interface;
  ListEntryDisp = dispinterface;
  Editors = interface;
  EditorsDisp = dispinterface;
  Editor = interface;
  EditorDisp = dispinterface;
  Styles = interface;
  StylesDisp = dispinterface;
  TablesOfFigures = interface;
  TablesOfFiguresDisp = dispinterface;
  TableOfFigures = interface;
  TableOfFiguresDisp = dispinterface;
  MailMerge = interface;
  MailMergeDisp = dispinterface;
  MailMergeDataSource = interface;
  MailMergeDataSourceDisp = dispinterface;
  MailMergeFieldNames = interface;
  MailMergeFieldNamesDisp = dispinterface;
  MailMergeFieldName = interface;
  MailMergeFieldNameDisp = dispinterface;
  MailMergeDataFields = interface;
  MailMergeDataFieldsDisp = dispinterface;
  MailMergeDataField = interface;
  MailMergeDataFieldDisp = dispinterface;
  MappedDataFields = interface;
  MappedDataFieldsDisp = dispinterface;
  MappedDataField = interface;
  MappedDataFieldDisp = dispinterface;
  MailMergeFields = interface;
  MailMergeFieldsDisp = dispinterface;
  MailMergeField = interface;
  MailMergeFieldDisp = dispinterface;
  TablesOfContents = interface;
  TablesOfContentsDisp = dispinterface;
  TableOfContents = interface;
  TableOfContentsDisp = dispinterface;
  Windows = interface;
  WindowsDisp = dispinterface;
  Window = interface;
  WindowDisp = dispinterface;
  Pane = interface;
  PaneDisp = dispinterface;
  Selection = interface;
  SelectionDisp = dispinterface;
  Find = interface;
  FindDisp = dispinterface;
  Replacement = interface;
  ReplacementDisp = dispinterface;
  Zooms = interface;
  ZoomsDisp = dispinterface;
  Zoom = interface;
  ZoomDisp = dispinterface;
  View = interface;
  ViewDisp = dispinterface;
  Reviewers = interface;
  ReviewersDisp = dispinterface;
  Reviewer = interface;
  ReviewerDisp = dispinterface;
  Pages = interface;
  PagesDisp = dispinterface;
  Page = interface;
  PageDisp = dispinterface;
  Breaks = interface;
  BreaksDisp = dispinterface;
  Break = interface;
  BreakDisp = dispinterface;
  Panes = interface;
  PanesDisp = dispinterface;
  ListTemplates = interface;
  ListTemplatesDisp = dispinterface;
  Lists = interface;
  ListsDisp = dispinterface;
  ExtraColors = interface;
  ExtraColorsDisp = dispinterface;
  Variables = interface;
  VariablesDisp = dispinterface;
  Variable = interface;
  VariableDisp = dispinterface;
  _WebOptions = interface;
  _WebOptionsDisp = dispinterface;
  ListGalleries = interface;
  ListGalleriesDisp = dispinterface;
  ListGallery = interface;
  ListGalleryDisp = dispinterface;
  FileConverters = interface;
  FileConvertersDisp = dispinterface;
  FileConverter = interface;
  FileConverterDisp = dispinterface;
  CaptionLabels = interface;
  CaptionLabelsDisp = dispinterface;
  CaptionLabel = interface;
  CaptionLabelDisp = dispinterface;
  Options = interface;
  OptionsDisp = dispinterface;
  RecentFiles = interface;
  RecentFilesDisp = dispinterface;
  RecentFile = interface;
  RecentFileDisp = dispinterface;
  Template = interface;
  TemplateDisp = dispinterface;
  Templates = interface;
  TemplatesDisp = dispinterface;
  Browser = interface;
  BrowserDisp = dispinterface;
  System = interface;
  SystemDisp = dispinterface;
  PdfExportOptions = interface;
  PdfExportOptionsDisp = dispinterface;
  Dictionaries = interface;
  DictionariesDisp = dispinterface;
  Dictionary = interface;
  DictionaryDisp = dispinterface;
  Dialogs = interface;
  DialogsDisp = dispinterface;
  Dialog = interface;
  DialogDisp = dispinterface;
  AutoCorrect = interface;
  AutoCorrectDisp = dispinterface;
  FontNames = interface;
  FontNamesDisp = dispinterface;
  AutoCaptions = interface;
  AutoCaptionsDisp = dispinterface;
  AutoCaption = interface;
  AutoCaptionDisp = dispinterface;
  ApplicationEvents = dispinterface;
  DocumentEvents = dispinterface;
  _OLEControl = interface;
  _OLEControlDisp = dispinterface;
  OCXEvents = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  OLECtrl = _OLEControl;
  Application = _Application;
  Document = _Document;
  Font = _Font;
  ParagraphFormat = _ParagraphFormat;
  WebOptions = _WebOptions;


// *********************************************************************//
// Declaration of structures, unions and aliases.                         
// *********************************************************************//
  POleVariant1 = ^OleVariant; {*}
  PWordBool1 = ^WordBool; {*}
  PPUserType1 = ^FootnoteOptions; {*}
  PUserType1 = ^WpsSectionStart; {*}
  PInteger1 = ^Integer; {*}
  PUserType2 = ^WpsNumberType; {*}

  WdMailSystem = __MIDL___MIDL_itf_wpsapi_0000_0000_0001; 
  WdTemplateType = __MIDL___MIDL_itf_wpsapi_0000_0000_0002; 
  WdContinue = __MIDL___MIDL_itf_wpsapi_0000_0000_0003; 
  WdIMEMode = __MIDL___MIDL_itf_wpsapi_0000_0000_0004; 
  WdBaselineAlignment = __MIDL___MIDL_itf_wpsapi_0000_0000_0005; 
  WdIndexFilter = __MIDL___MIDL_itf_wpsapi_0000_0000_0006; 
  WdIndexSortBy = __MIDL___MIDL_itf_wpsapi_0000_0000_0007; 
  WdJustificationMode = __MIDL___MIDL_itf_wpsapi_0000_0000_0008; 
  WdFarEastLineBreakLevel = __MIDL___MIDL_itf_wpsapi_0000_0000_0009; 
  WdMultipleWordConversionsMode = __MIDL___MIDL_itf_wpsapi_0000_0000_0010; 
  WdColorIndex = __MIDL___MIDL_itf_wpsapi_0000_0000_0011; 
  WdTextureIndex = __MIDL___MIDL_itf_wpsapi_0000_0000_0012; 
  WdUnderline = __MIDL___MIDL_itf_wpsapi_0000_0000_0013; 
  WdEmphasisMark = __MIDL___MIDL_itf_wpsapi_0000_0000_0014; 
  WdInternationalIndex = __MIDL___MIDL_itf_wpsapi_0000_0000_0015; 
  WdAutoMacros = __MIDL___MIDL_itf_wpsapi_0000_0000_0016; 
  WdCaptionPosition = __MIDL___MIDL_itf_wpsapi_0000_0000_0017; 
  WdCountry = __MIDL___MIDL_itf_wpsapi_0000_0000_0018; 
  WdHeadingSeparator = __MIDL___MIDL_itf_wpsapi_0000_0000_0019; 
  WdSeparatorType = __MIDL___MIDL_itf_wpsapi_0000_0000_0020; 
  WdPageNumberAlignment = __MIDL___MIDL_itf_wpsapi_0000_0000_0021; 
  WdBorderType = __MIDL___MIDL_itf_wpsapi_0000_0000_0022; 
  WdFramePosition = __MIDL___MIDL_itf_wpsapi_0000_0000_0023; 
  WdAnimation = __MIDL___MIDL_itf_wpsapi_0000_0000_0024; 
  WdCharacterCase = __MIDL___MIDL_itf_wpsapi_0000_0000_0025; 
  WdSummaryMode = __MIDL___MIDL_itf_wpsapi_0000_0000_0026; 
  WdSummaryLength = __MIDL___MIDL_itf_wpsapi_0000_0000_0027; 
  WdStyleType = __MIDL___MIDL_itf_wpsapi_0000_0000_0028; 
  WdUnits = __MIDL___MIDL_itf_wpsapi_0000_0000_0029; 
  WdGoToItem = __MIDL___MIDL_itf_wpsapi_0000_0000_0030; 
  WdGoToDirection = __MIDL___MIDL_itf_wpsapi_0000_0000_0031; 
  WdCollapseDirection = __MIDL___MIDL_itf_wpsapi_0000_0000_0032; 
  WdRowHeightRule = __MIDL___MIDL_itf_wpsapi_0000_0000_0033; 
  WdFrameSizeRule = __MIDL___MIDL_itf_wpsapi_0000_0000_0034; 
  WdInsertCells = __MIDL___MIDL_itf_wpsapi_0000_0000_0035; 
  WdDeleteCells = __MIDL___MIDL_itf_wpsapi_0000_0000_0036; 
  WdListApplyTo = __MIDL___MIDL_itf_wpsapi_0000_0000_0037; 
  WdAlertLevel = __MIDL___MIDL_itf_wpsapi_0000_0000_0038; 
  WdCursorType = __MIDL___MIDL_itf_wpsapi_0000_0000_0039; 
  WdEnableCancelKey = __MIDL___MIDL_itf_wpsapi_0000_0000_0040; 
  WdRulerStyle = __MIDL___MIDL_itf_wpsapi_0000_0000_0041; 
  WdParagraphAlignment = __MIDL___MIDL_itf_wpsapi_0000_0000_0042; 
  WdListLevelAlignment = __MIDL___MIDL_itf_wpsapi_0000_0000_0043; 
  WdRowAlignment = __MIDL___MIDL_itf_wpsapi_0000_0000_0044; 
  WdTabAlignment = __MIDL___MIDL_itf_wpsapi_0000_0000_0045; 
  WdVerticalAlignment = __MIDL___MIDL_itf_wpsapi_0000_0000_0046; 
  WdCellVerticalAlignment = __MIDL___MIDL_itf_wpsapi_0000_0000_0047; 
  WdTrailingCharacter = __MIDL___MIDL_itf_wpsapi_0000_0000_0048; 
  WdListGalleryType = __MIDL___MIDL_itf_wpsapi_0000_0000_0049; 
  WdListNumberStyle = __MIDL___MIDL_itf_wpsapi_0000_0000_0050; 
  WdNoteNumberStyle = __MIDL___MIDL_itf_wpsapi_0000_0000_0051; 
  WdCaptionNumberStyle = __MIDL___MIDL_itf_wpsapi_0000_0000_0052; 
  WdPageNumberStyle = __MIDL___MIDL_itf_wpsapi_0000_0000_0053; 
  WdStatistic = __MIDL___MIDL_itf_wpsapi_0000_0000_0054; 
  WdBuiltInProperty = __MIDL___MIDL_itf_wpsapi_0000_0000_0055; 
  WdLineSpacing = __MIDL___MIDL_itf_wpsapi_0000_0000_0056; 
  WdNumberType = __MIDL___MIDL_itf_wpsapi_0000_0000_0057; 
  WdListType = __MIDL___MIDL_itf_wpsapi_0000_0000_0058; 
  WdStoryType = __MIDL___MIDL_itf_wpsapi_0000_0000_0059; 
  WdSaveFormat = __MIDL___MIDL_itf_wpsapi_0000_0000_0060; 
  WdOpenFormat = __MIDL___MIDL_itf_wpsapi_0000_0000_0061; 
  WdHeaderFooterIndex = __MIDL___MIDL_itf_wpsapi_0000_0000_0062; 
  WdTocFormat = __MIDL___MIDL_itf_wpsapi_0000_0000_0063; 
  WdTofFormat = __MIDL___MIDL_itf_wpsapi_0000_0000_0064; 
  WdToaFormat = __MIDL___MIDL_itf_wpsapi_0000_0000_0065; 
  WdLineStyle = __MIDL___MIDL_itf_wpsapi_0000_0000_0066; 
  WdLineWidth = __MIDL___MIDL_itf_wpsapi_0000_0000_0067; 
  WdBreakType = __MIDL___MIDL_itf_wpsapi_0000_0000_0068; 
  WdTabLeader = __MIDL___MIDL_itf_wpsapi_0000_0000_0069; 
  WdMeasurementUnits = __MIDL___MIDL_itf_wpsapi_0000_0000_0070; 
  WdDropPosition = __MIDL___MIDL_itf_wpsapi_0000_0000_0071; 
  WdNumberingRule = __MIDL___MIDL_itf_wpsapi_0000_0000_0072; 
  WdFootnoteLocation = __MIDL___MIDL_itf_wpsapi_0000_0000_0073; 
  WdEndnoteLocation = __MIDL___MIDL_itf_wpsapi_0000_0000_0074; 
  WdSortSeparator = __MIDL___MIDL_itf_wpsapi_0000_0000_0075; 
  WdTableFieldSeparator = __MIDL___MIDL_itf_wpsapi_0000_0000_0076; 
  WdSortFieldType = __MIDL___MIDL_itf_wpsapi_0000_0000_0077; 
  WdSortOrder = __MIDL___MIDL_itf_wpsapi_0000_0000_0078; 
  WdTableFormat = __MIDL___MIDL_itf_wpsapi_0000_0000_0079; 
  WdTableFormatApply = __MIDL___MIDL_itf_wpsapi_0000_0000_0080; 
  WdLanguageID = __MIDL___MIDL_itf_wpsapi_0000_0000_0081; 
  WdFieldType = __MIDL___MIDL_itf_wpsapi_0000_0000_0082; 
  WdBuiltinStyle = __MIDL___MIDL_itf_wpsapi_0000_0000_0083; 
  WdWordDialogTab = __MIDL___MIDL_itf_wpsapi_0000_0000_0084; 
  WdWordDialogTabHID = __MIDL___MIDL_itf_wpsapi_0000_0000_0085; 
  WdWordDialog = __MIDL___MIDL_itf_wpsapi_0000_0000_0086; 
  WdFieldKind = __MIDL___MIDL_itf_wpsapi_0000_0000_0087; 
  WdTextFormFieldType = __MIDL___MIDL_itf_wpsapi_0000_0000_0088; 
  WdChevronConvertRule = __MIDL___MIDL_itf_wpsapi_0000_0000_0089; 
  WdMailMergeMainDocType = __MIDL___MIDL_itf_wpsapi_0000_0000_0090; 
  WdMailMergeState = __MIDL___MIDL_itf_wpsapi_0000_0000_0091; 
  WdMailMergeDestination = __MIDL___MIDL_itf_wpsapi_0000_0000_0092; 
  WdMailMergeActiveRecord = __MIDL___MIDL_itf_wpsapi_0000_0000_0093; 
  WdMailMergeDefaultRecord = __MIDL___MIDL_itf_wpsapi_0000_0000_0094; 
  WdMailMergeDataSource = __MIDL___MIDL_itf_wpsapi_0000_0000_0095; 
  WdMailMergeComparison = __MIDL___MIDL_itf_wpsapi_0000_0000_0096; 
  WdBookmarkSortBy = __MIDL___MIDL_itf_wpsapi_0000_0000_0097; 
  WdWindowState = __MIDL___MIDL_itf_wpsapi_0000_0000_0098; 
  WdPictureLinkType = __MIDL___MIDL_itf_wpsapi_0000_0000_0099; 
  WdLinkType = __MIDL___MIDL_itf_wpsapi_0000_0000_0100; 
  WdWindowType = __MIDL___MIDL_itf_wpsapi_0000_0000_0101; 
  WdViewType = __MIDL___MIDL_itf_wpsapi_0000_0000_0102; 
  WdSeekView = __MIDL___MIDL_itf_wpsapi_0000_0000_0103; 
  WdSpecialPane = __MIDL___MIDL_itf_wpsapi_0000_0000_0104; 
  WdPageFit = __MIDL___MIDL_itf_wpsapi_0000_0000_0105; 
  WdBrowseTarget = __MIDL___MIDL_itf_wpsapi_0000_0000_0106; 
  WdPaperTray = __MIDL___MIDL_itf_wpsapi_0000_0000_0107; 
  WdOrientation = __MIDL___MIDL_itf_wpsapi_0000_0000_0108; 
  WdSelectionType = __MIDL___MIDL_itf_wpsapi_0000_0000_0109; 
  WdCaptionLabelID = __MIDL___MIDL_itf_wpsapi_0000_0000_0110; 
  WdReferenceType = __MIDL___MIDL_itf_wpsapi_0000_0000_0111; 
  WdReferenceKind = __MIDL___MIDL_itf_wpsapi_0000_0000_0112; 
  WdIndexFormat = __MIDL___MIDL_itf_wpsapi_0000_0000_0113; 
  WdIndexType = __MIDL___MIDL_itf_wpsapi_0000_0000_0114; 
  WdRevisionsWrap = __MIDL___MIDL_itf_wpsapi_0000_0000_0115; 
  WdRevisionType = __MIDL___MIDL_itf_wpsapi_0000_0000_0116; 
  WdRoutingSlipDelivery = __MIDL___MIDL_itf_wpsapi_0000_0000_0117; 
  WdRoutingSlipStatus = __MIDL___MIDL_itf_wpsapi_0000_0000_0118; 
  WdSectionStart = __MIDL___MIDL_itf_wpsapi_0000_0000_0119; 
  WdSaveOptions = __MIDL___MIDL_itf_wpsapi_0000_0000_0120; 
  WdDocumentKind = __MIDL___MIDL_itf_wpsapi_0000_0000_0121; 
  WdDocumentType = __MIDL___MIDL_itf_wpsapi_0000_0000_0122; 
  WdOriginalFormat = __MIDL___MIDL_itf_wpsapi_0000_0000_0123; 
  WdRelocate = __MIDL___MIDL_itf_wpsapi_0000_0000_0124; 
  WdInsertedTextMark = __MIDL___MIDL_itf_wpsapi_0000_0000_0125; 
  WdRevisedLinesMark = __MIDL___MIDL_itf_wpsapi_0000_0000_0126; 
  WdDeletedTextMark = __MIDL___MIDL_itf_wpsapi_0000_0000_0127; 
  WdRevisedPropertiesMark = __MIDL___MIDL_itf_wpsapi_0000_0000_0128; 
  WdFieldShading = __MIDL___MIDL_itf_wpsapi_0000_0000_0129; 
  WdDefaultFilePath = __MIDL___MIDL_itf_wpsapi_0000_0000_0130; 
  WdCompatibility = __MIDL___MIDL_itf_wpsapi_0000_0000_0131; 
  WdPaperSize = __MIDL___MIDL_itf_wpsapi_0000_0000_0132; 
  WdCustomLabelPageSize = __MIDL___MIDL_itf_wpsapi_0000_0000_0133; 
  WdProtectionType = __MIDL___MIDL_itf_wpsapi_0000_0000_0134; 
  WdPartOfSpeech = __MIDL___MIDL_itf_wpsapi_0000_0000_0135; 
  WdSubscriberFormats = __MIDL___MIDL_itf_wpsapi_0000_0000_0136; 
  WdEditionType = __MIDL___MIDL_itf_wpsapi_0000_0000_0137; 
  WdEditionOption = __MIDL___MIDL_itf_wpsapi_0000_0000_0138; 
  WdRelativeHorizontalPosition = __MIDL___MIDL_itf_wpsapi_0000_0000_0139; 
  WdRelativeVerticalPosition = __MIDL___MIDL_itf_wpsapi_0000_0000_0140; 
  WdHelpType = __MIDL___MIDL_itf_wpsapi_0000_0000_0141; 
  WdKeyCategory = __MIDL___MIDL_itf_wpsapi_0000_0000_0142; 
  WdKey = __MIDL___MIDL_itf_wpsapi_0000_0000_0143; 
  WdOLEType = __MIDL___MIDL_itf_wpsapi_0000_0000_0144; 
  WdOLEVerb = __MIDL___MIDL_itf_wpsapi_0000_0000_0145; 
  WdOLEPlacement = __MIDL___MIDL_itf_wpsapi_0000_0000_0146; 
  WdEnvelopeOrientation = __MIDL___MIDL_itf_wpsapi_0000_0000_0147; 
  WdLetterStyle = __MIDL___MIDL_itf_wpsapi_0000_0000_0148; 
  WdLetterheadLocation = __MIDL___MIDL_itf_wpsapi_0000_0000_0149; 
  WdSalutationType = __MIDL___MIDL_itf_wpsapi_0000_0000_0150; 
  WdSalutationGender = __MIDL___MIDL_itf_wpsapi_0000_0000_0151; 
  WdMovementType = __MIDL___MIDL_itf_wpsapi_0000_0000_0152; 
  WdConstants = __MIDL___MIDL_itf_wpsapi_0000_0000_0153; 
  WdPasteDataType = __MIDL___MIDL_itf_wpsapi_0000_0000_0154; 
  WdPrintOutItem = __MIDL___MIDL_itf_wpsapi_0000_0000_0155; 
  WdPrintOutPages = __MIDL___MIDL_itf_wpsapi_0000_0000_0156; 
  WdPrintOutRange = __MIDL___MIDL_itf_wpsapi_0000_0000_0157; 
  WdDictionaryType = __MIDL___MIDL_itf_wpsapi_0000_0000_0158; 
  WdSpellingWordType = __MIDL___MIDL_itf_wpsapi_0000_0000_0159; 
  WdSpellingErrorType = __MIDL___MIDL_itf_wpsapi_0000_0000_0160; 
  WdProofreadingErrorType = __MIDL___MIDL_itf_wpsapi_0000_0000_0161; 
  WdInlineShapeType = __MIDL___MIDL_itf_wpsapi_0000_0000_0162; 
  WdArrangeStyle = __MIDL___MIDL_itf_wpsapi_0000_0000_0163; 
  WdSelectionFlags = __MIDL___MIDL_itf_wpsapi_0000_0000_0164; 
  WdAutoVersions = __MIDL___MIDL_itf_wpsapi_0000_0000_0165; 
  WdOrganizerObject = __MIDL___MIDL_itf_wpsapi_0000_0000_0166; 
  WdFindMatch = __MIDL___MIDL_itf_wpsapi_0000_0000_0167; 
  WdFindWrap = __MIDL___MIDL_itf_wpsapi_0000_0000_0168; 
  WdInformation = __MIDL___MIDL_itf_wpsapi_0000_0000_0169; 
  WdWrapType = __MIDL___MIDL_itf_wpsapi_0000_0000_0170; 
  WdWrapSideType = __MIDL___MIDL_itf_wpsapi_0000_0000_0171; 
  WdOutlineLevel = __MIDL___MIDL_itf_wpsapi_0000_0000_0172; 
  WdTextOrientation = __MIDL___MIDL_itf_wpsapi_0000_0000_0173; 
  WdPageBorderArt = __MIDL___MIDL_itf_wpsapi_0000_0000_0174; 
  WdBorderDistanceFrom = __MIDL___MIDL_itf_wpsapi_0000_0000_0175; 
  WdReplace = __MIDL___MIDL_itf_wpsapi_0000_0000_0176; 
  WdFontBias = __MIDL___MIDL_itf_wpsapi_0000_0000_0177; 
  WdDocumentMedium = __MIDL___MIDL_itf_wpsapi_0000_0000_0178; 
  WdFarEastLineBreakLanguageID = __MIDL___MIDL_itf_wpsapi_0000_0000_0179; 
  WdDisableFeaturesIntroducedAfter = __MIDL___MIDL_itf_wpsapi_0000_0000_0180; 
  WdLineEndingType = __MIDL___MIDL_itf_wpsapi_0000_0000_0181; 
  WdShowFilter = __MIDL___MIDL_itf_wpsapi_0000_0000_0182; 

// *********************************************************************//
// Interface: _IKsoDispObj
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0300-FFFF-0000-C000-000000000046}
// *********************************************************************//
  _IKsoDispObj = interface(IDispatch)
    ['{000C0300-FFFF-0000-C000-000000000046}']
    function Get_Application: _Application; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    procedure zimp_DispObj_Reserved1; safecall;
    procedure zimp_DispObj_Reserved2; safecall;
    procedure zimp_DispObj_Reserved3; safecall;
    procedure zimp_DispObj_Reserved4; safecall;
    procedure zimp_DispObj_Reserved5; safecall;
    function Get_Description: WideString; safecall;
    property Application: _Application read Get_Application;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Description: WideString read Get_Description;
  end;

// *********************************************************************//
// DispIntf:  _IKsoDispObjDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000C0300-FFFF-0000-C000-000000000046}
// *********************************************************************//
  _IKsoDispObjDisp = dispinterface
    ['{000C0300-FFFF-0000-C000-000000000046}']
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: _Application
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020970-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  _Application = interface(_IKsoDispObj)
    ['{00020970-0000-4B30-A977-D214852036FE}']
    function Get_Name: WideString; safecall;
    function Get_Documents: Documents; safecall;
    function Get_Windows: Windows; safecall;
    function Get_ActiveDocument: _Document; safecall;
    function Get_ActiveWindow: Window; safecall;
    function Get_DisplayStatusBar: WordBool; safecall;
    procedure Set_DisplayStatusBar(prop: WordBool); safecall;
    function Get_ListGalleries: ListGalleries; safecall;
    function Get_Selection: Selection; safecall;
    function Get_FileConverters: FileConverters; safecall;
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(prop: WordBool); safecall;
    function Get_DisplayRecentFiles: WordBool; safecall;
    procedure Set_DisplayRecentFiles(prop: WordBool); safecall;
    function Get_DefaultSaveFormat: OleVariant; safecall;
    procedure Set_DefaultSaveFormat(prop: OleVariant); safecall;
    function Get_CommandBars: _CommandBars; safecall;
    procedure OrganizerCopy(const Source: WideString; const Destination: WideString; 
                            const Name: WideString; Object_: WpsOrganizerObject); safecall;
    procedure OrganizerDelete(const Source: WideString; const Name: WideString; 
                              Object_: WpsOrganizerObject); safecall;
    procedure OrganizerRename(const Source: WideString; const Name: WideString; 
                              const NewName: WideString; Object_: WpsOrganizerObject); safecall;
    function Get_ShowStartupDialog: WordBool; safecall;
    procedure Set_ShowStartupDialog(prop: WordBool); safecall;
    procedure Quit(var SaveChanges: OleVariant; var OriginalFormat: OleVariant; 
                   var RouteDocument: OleVariant); safecall;
    function Get_ShowVisualBasicEditor: WordBool; safecall;
    procedure Set_ShowVisualBasicEditor(prop: WordBool); safecall;
    procedure PrintOut(Background: WordBool; Append: WordBool; Range: WpsPrintOutRange; 
                       const OutputFileName: WideString; From: SYSINT; To_: SYSINT; 
                       Item: WpsPrintOutItem; Copies: SYSINT; const Pages: WideString; 
                       PageType: WpsPrintOutPages; PrintToFile: WordBool; Collate: WordBool; 
                       const FileName: WideString; var ActivePrinterMacGX: OleVariant; 
                       ManualDuplexPrint: WordBool; PrintZoomColumn: Integer; 
                       PrintZoomRow: Integer; PrintZoomPaperWidth: Integer; 
                       PrintZoomPaperHeight: Integer; FlipPrint: WordBool; PaperTray: WpsPaperTray; 
                       CutterLine: WordBool; PaperOrder: WpsPaperOrder); safecall;
    function Get_ActivePrinter: WideString; safecall;
    procedure Set_ActivePrinter(const prop: WideString); safecall;
    function Get_CaptionLabels: CaptionLabels; safecall;
    function InchesToPoints(Inches: Single): Single; safecall;
    function CentimetersToPoints(Centimeters: Single): Single; safecall;
    function MillimetersToPoints(Millimeters: Single): Single; safecall;
    function PicasToPoints(Picas: Single): Single; safecall;
    function LinesToPoints(Lines: Single): Single; safecall;
    function PointsToInches(Points: Single): Single; safecall;
    function PointsToCentimeters(Points: Single): Single; safecall;
    function PointsToMillimeters(Points: Single): Single; safecall;
    function PointsToPicas(Points: Single): Single; safecall;
    function PointsToLines(Points: Single): Single; safecall;
    function PointsToPixels(Points: Single; var fVertical: WordBool): Single; safecall;
    function PixelsToPoints(Pixels: Single; var fVertical: WordBool): Single; safecall;
    function Get_Options: Options; safecall;
    function Get_UserName: WideString; safecall;
    procedure Set_UserName(const prop: WideString); safecall;
    function Get_UserInitials: WideString; safecall;
    procedure Set_UserInitials(const prop: WideString); safecall;
    function Get_UserAddress: WideString; safecall;
    procedure Set_UserAddress(const prop: WideString); safecall;
    function Get_RecentFiles: RecentFiles; safecall;
    function Get_NormalTemplate: Template; safecall;
    function Get_Templates: Templates; safecall;
    function Get_KeyBindings: KeyBindings; safecall;
    function Get_KeysBoundTo(KeyCategory: WpsKeyCategory; const Command: WideString; 
                             var CommandParameter: OleVariant): KeysBoundTo; safecall;
    function Get_FindKey(KeyCode: Integer; var KeyCode2: OleVariant): KeyBinding; safecall;
    function BuildKeyCode(Arg1: WpsKey; var Arg2: OleVariant; var Arg3: OleVariant; 
                          var Arg4: OleVariant): Integer; safecall;
    function KeyString(KeyCode: Integer; var KeyCode2: OleVariant): WideString; safecall;
    function Get_Browser: Browser; safecall;
    procedure PreviousChangeOrComment; safecall;
    procedure NextChangeOrComment; safecall;
    function Get_UserControl: WordBool; safecall;
    procedure Set_UserControl(RHS: WordBool); safecall;
    function Get_COMAddIns: COMAddIns; safecall;
    function Run(const MacroName: WideString; var varg1: OleVariant; var varg2: OleVariant; 
                 var varg3: OleVariant; var varg4: OleVariant; var varg5: OleVariant; 
                 var varg6: OleVariant; var varg7: OleVariant; var varg8: OleVariant; 
                 var varg9: OleVariant; var varg10: OleVariant; var varg11: OleVariant; 
                 var varg12: OleVariant; var varg13: OleVariant; var varg14: OleVariant; 
                 var varg15: OleVariant; var varg16: OleVariant; var varg17: OleVariant; 
                 var varg18: OleVariant; var varg19: OleVariant; var varg20: OleVariant; 
                 var varg21: OleVariant; var varg22: OleVariant; var varg23: OleVariant; 
                 var varg24: OleVariant; var varg25: OleVariant; var varg26: OleVariant; 
                 var varg27: OleVariant; var varg28: OleVariant; var varg29: OleVariant; 
                 var varg30: OleVariant): OleVariant; safecall;
    function Get_Version: WideString; safecall;
    function Get_Build: WideString; safecall;
    function Get_WindowState: WpsWindowState; safecall;
    procedure Set_WindowState(prop: WpsWindowState); safecall;
    procedure Terminate(bForce: WordBool); safecall;
    function Get_VBE: IDispatch; safecall;
    function Get_System_: System; safecall;
    function Get_Left: Integer; safecall;
    procedure Set_Left(prop: Integer); safecall;
    function Get_Top: Integer; safecall;
    procedure Set_Top(prop: Integer); safecall;
    function Get_Width: Integer; safecall;
    procedure Set_Width(prop: Integer); safecall;
    function Get_Height: Integer; safecall;
    procedure Set_Height(prop: Integer); safecall;
    procedure Activate; safecall;
    function Get_DisplayScreenTips: WordBool; safecall;
    procedure Set_DisplayScreenTips(prop: WordBool); safecall;
    function Get_PdfExportOptions: PdfExportOptions; safecall;
    function Get_CustomDictionaries: Dictionaries; safecall;
    function CheckSpelling(const Word: WideString; var CustomDictionary: OleVariant; 
                           var IgnoreUppercase: OleVariant; var MainDictionary: OleVariant; 
                           var CustomDictionary2: OleVariant; var CustomDictionary3: OleVariant; 
                           var CustomDictionary4: OleVariant; var CustomDictionary5: OleVariant; 
                           var CustomDictionary6: OleVariant; var CustomDictionary7: OleVariant; 
                           var CustomDictionary8: OleVariant; var CustomDictionary9: OleVariant; 
                           var CustomDictionary10: OleVariant): WordBool; safecall;
    function GetSpellingSuggestions(const Word: WideString; var CustomDictionary: OleVariant; 
                                    var IgnoreUppercase: OleVariant; 
                                    var MainDictionary: OleVariant; var SuggestionMode: OleVariant; 
                                    var CustomDictionary2: OleVariant; 
                                    var CustomDictionary3: OleVariant; 
                                    var CustomDictionary4: OleVariant; 
                                    var CustomDictionary5: OleVariant; 
                                    var CustomDictionary6: OleVariant; 
                                    var CustomDictionary7: OleVariant; 
                                    var CustomDictionary8: OleVariant; 
                                    var CustomDictionary9: OleVariant; 
                                    var CustomDictionary10: OleVariant): SpellingSuggestions; safecall;
    procedure BeginCheckSpelling; safecall;
    procedure EndCheckSpelling; safecall;
    function Get_Dialogs: Dialogs; safecall;
    function Get_TaskPanes: TaskPanes; safecall;
    function Get_Caption: WideString; safecall;
    procedure Set_Caption(const prop: WideString); safecall;
    function Get_DisplayAlerts: WpsAlertLevel; safecall;
    procedure Set_DisplayAlerts(prop: WpsAlertLevel); safecall;
    function Get_AdvApiRoot: AdvApiRoot; safecall;
    procedure Set_EnableAppWindow(prop: WordBool); safecall;
    function Get_EnableAppWindow: WordBool; safecall;
    function Get_AutoCorrect: AutoCorrect; safecall;
    function Get_CapsLock: WordBool; safecall;
    function Get_NumLock: WordBool; safecall;
    function Get_PluginPlatform: PluginPlatform; safecall;
    function Get_FontNames: FontNames; safecall;
    function Get_AutoCircleNumber: WordBool; safecall;
    procedure Set_AutoCircleNumber(prop: WordBool); safecall;
    property Name: WideString read Get_Name;
    property Documents: Documents read Get_Documents;
    property Windows: Windows read Get_Windows;
    property ActiveDocument: _Document read Get_ActiveDocument;
    property ActiveWindow: Window read Get_ActiveWindow;
    property DisplayStatusBar: WordBool read Get_DisplayStatusBar write Set_DisplayStatusBar;
    property ListGalleries: ListGalleries read Get_ListGalleries;
    property Selection: Selection read Get_Selection;
    property FileConverters: FileConverters read Get_FileConverters;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property DisplayRecentFiles: WordBool read Get_DisplayRecentFiles write Set_DisplayRecentFiles;
    property DefaultSaveFormat: OleVariant read Get_DefaultSaveFormat write Set_DefaultSaveFormat;
    property CommandBars: _CommandBars read Get_CommandBars;
    property ShowStartupDialog: WordBool read Get_ShowStartupDialog write Set_ShowStartupDialog;
    property ShowVisualBasicEditor: WordBool read Get_ShowVisualBasicEditor write Set_ShowVisualBasicEditor;
    property ActivePrinter: WideString read Get_ActivePrinter write Set_ActivePrinter;
    property CaptionLabels: CaptionLabels read Get_CaptionLabels;
    property Options: Options read Get_Options;
    property UserName: WideString read Get_UserName write Set_UserName;
    property UserInitials: WideString read Get_UserInitials write Set_UserInitials;
    property UserAddress: WideString read Get_UserAddress write Set_UserAddress;
    property RecentFiles: RecentFiles read Get_RecentFiles;
    property NormalTemplate: Template read Get_NormalTemplate;
    property Templates: Templates read Get_Templates;
    property KeyBindings: KeyBindings read Get_KeyBindings;
    property KeysBoundTo[KeyCategory: WpsKeyCategory; const Command: WideString; 
                         var CommandParameter: OleVariant]: KeysBoundTo read Get_KeysBoundTo;
    property FindKey[KeyCode: Integer; var KeyCode2: OleVariant]: KeyBinding read Get_FindKey;
    property Browser: Browser read Get_Browser;
    property UserControl: WordBool read Get_UserControl write Set_UserControl;
    property COMAddIns: COMAddIns read Get_COMAddIns;
    property Version: WideString read Get_Version;
    property Build: WideString read Get_Build;
    property WindowState: WpsWindowState read Get_WindowState write Set_WindowState;
    property VBE: IDispatch read Get_VBE;
    property System_: System read Get_System_;
    property Left: Integer read Get_Left write Set_Left;
    property Top: Integer read Get_Top write Set_Top;
    property Width: Integer read Get_Width write Set_Width;
    property Height: Integer read Get_Height write Set_Height;
    property DisplayScreenTips: WordBool read Get_DisplayScreenTips write Set_DisplayScreenTips;
    property PdfExportOptions: PdfExportOptions read Get_PdfExportOptions;
    property CustomDictionaries: Dictionaries read Get_CustomDictionaries;
    property Dialogs: Dialogs read Get_Dialogs;
    property TaskPanes: TaskPanes read Get_TaskPanes;
    property Caption: WideString read Get_Caption write Set_Caption;
    property DisplayAlerts: WpsAlertLevel read Get_DisplayAlerts write Set_DisplayAlerts;
    property AdvApiRoot: AdvApiRoot read Get_AdvApiRoot;
    property EnableAppWindow: WordBool read Get_EnableAppWindow write Set_EnableAppWindow;
    property AutoCorrect: AutoCorrect read Get_AutoCorrect;
    property CapsLock: WordBool read Get_CapsLock;
    property NumLock: WordBool read Get_NumLock;
    property PluginPlatform: PluginPlatform read Get_PluginPlatform;
    property FontNames: FontNames read Get_FontNames;
    property AutoCircleNumber: WordBool read Get_AutoCircleNumber write Set_AutoCircleNumber;
  end;

// *********************************************************************//
// DispIntf:  _ApplicationDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020970-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  _ApplicationDisp = dispinterface
    ['{00020970-0000-4B30-A977-D214852036FE}']
    property Name: WideString readonly dispid 0;
    property Documents: Documents readonly dispid 6;
    property Windows: Windows readonly dispid 2;
    property ActiveDocument: _Document readonly dispid 3;
    property ActiveWindow: Window readonly dispid 4;
    property DisplayStatusBar: WordBool dispid 29;
    property ListGalleries: ListGalleries readonly dispid 65;
    property Selection: Selection readonly dispid 5;
    property FileConverters: FileConverters readonly dispid 17;
    property Visible: WordBool dispid 23;
    property DisplayRecentFiles: WordBool dispid 56;
    property DefaultSaveFormat: OleVariant dispid 64;
    property CommandBars: _CommandBars readonly dispid 57;
    procedure OrganizerCopy(const Source: WideString; const Destination: WideString; 
                            const Name: WideString; Object_: WpsOrganizerObject); dispid 318;
    procedure OrganizerDelete(const Source: WideString; const Name: WideString; 
                              Object_: WpsOrganizerObject); dispid 319;
    procedure OrganizerRename(const Source: WideString; const Name: WideString; 
                              const NewName: WideString; Object_: WpsOrganizerObject); dispid 320;
    property ShowStartupDialog: WordBool dispid 455;
    procedure Quit(var SaveChanges: OleVariant; var OriginalFormat: OleVariant; 
                   var RouteDocument: OleVariant); dispid 1105;
    property ShowVisualBasicEditor: WordBool dispid 106;
    procedure PrintOut(Background: WordBool; Append: WordBool; Range: WpsPrintOutRange; 
                       const OutputFileName: WideString; From: SYSINT; To_: SYSINT; 
                       Item: WpsPrintOutItem; Copies: SYSINT; const Pages: WideString; 
                       PageType: WpsPrintOutPages; PrintToFile: WordBool; Collate: WordBool; 
                       const FileName: WideString; var ActivePrinterMacGX: OleVariant; 
                       ManualDuplexPrint: WordBool; PrintZoomColumn: Integer; 
                       PrintZoomRow: Integer; PrintZoomPaperWidth: Integer; 
                       PrintZoomPaperHeight: Integer; FlipPrint: WordBool; PaperTray: WpsPaperTray; 
                       CutterLine: WordBool; PaperOrder: WpsPaperOrder); dispid 448;
    property ActivePrinter: WideString dispid 66;
    property CaptionLabels: CaptionLabels readonly dispid 20;
    function InchesToPoints(Inches: Single): Single; dispid 370;
    function CentimetersToPoints(Centimeters: Single): Single; dispid 371;
    function MillimetersToPoints(Millimeters: Single): Single; dispid 372;
    function PicasToPoints(Picas: Single): Single; dispid 373;
    function LinesToPoints(Lines: Single): Single; dispid 374;
    function PointsToInches(Points: Single): Single; dispid 380;
    function PointsToCentimeters(Points: Single): Single; dispid 381;
    function PointsToMillimeters(Points: Single): Single; dispid 382;
    function PointsToPicas(Points: Single): Single; dispid 383;
    function PointsToLines(Points: Single): Single; dispid 384;
    function PointsToPixels(Points: Single; var fVertical: WordBool): Single; dispid 387;
    function PixelsToPoints(Pixels: Single; var fVertical: WordBool): Single; dispid 388;
    property Options: Options readonly dispid 93;
    property UserName: WideString dispid 52;
    property UserInitials: WideString dispid 53;
    property UserAddress: WideString dispid 54;
    property RecentFiles: RecentFiles readonly dispid 7;
    property NormalTemplate: Template readonly dispid 8;
    property Templates: Templates readonly dispid 67;
    property KeyBindings: KeyBindings readonly dispid 69;
    property KeysBoundTo[KeyCategory: WpsKeyCategory; const Command: WideString; 
                         var CommandParameter: OleVariant]: KeysBoundTo readonly dispid 70;
    property FindKey[KeyCode: Integer; var KeyCode2: OleVariant]: KeyBinding readonly dispid 71;
    function BuildKeyCode(Arg1: WpsKey; var Arg2: OleVariant; var Arg3: OleVariant; 
                          var Arg4: OleVariant): Integer; dispid 316;
    function KeyString(KeyCode: Integer; var KeyCode2: OleVariant): WideString; dispid 317;
    property Browser: Browser readonly dispid 16;
    procedure PreviousChangeOrComment; dispid 645;
    procedure NextChangeOrComment; dispid 646;
    property UserControl: WordBool dispid 647;
    property COMAddIns: COMAddIns readonly dispid 111;
    function Run(const MacroName: WideString; var varg1: OleVariant; var varg2: OleVariant; 
                 var varg3: OleVariant; var varg4: OleVariant; var varg5: OleVariant; 
                 var varg6: OleVariant; var varg7: OleVariant; var varg8: OleVariant; 
                 var varg9: OleVariant; var varg10: OleVariant; var varg11: OleVariant; 
                 var varg12: OleVariant; var varg13: OleVariant; var varg14: OleVariant; 
                 var varg15: OleVariant; var varg16: OleVariant; var varg17: OleVariant; 
                 var varg18: OleVariant; var varg19: OleVariant; var varg20: OleVariant; 
                 var varg21: OleVariant; var varg22: OleVariant; var varg23: OleVariant; 
                 var varg24: OleVariant; var varg25: OleVariant; var varg26: OleVariant; 
                 var varg27: OleVariant; var varg28: OleVariant; var varg29: OleVariant; 
                 var varg30: OleVariant): OleVariant; dispid 445;
    property Version: WideString readonly dispid 24;
    property Build: WideString readonly dispid 47;
    property WindowState: WpsWindowState dispid 91;
    procedure Terminate(bForce: WordBool); dispid 65537;
    property VBE: IDispatch readonly dispid 61;
    property System_: System readonly dispid 9;
    property Left: Integer dispid 87;
    property Top: Integer dispid 88;
    property Width: Integer dispid 89;
    property Height: Integer dispid 90;
    procedure Activate; dispid 385;
    property DisplayScreenTips: WordBool dispid 99;
    property PdfExportOptions: PdfExportOptions readonly dispid 4097;
    property CustomDictionaries: Dictionaries readonly dispid 95;
    function CheckSpelling(const Word: WideString; var CustomDictionary: OleVariant; 
                           var IgnoreUppercase: OleVariant; var MainDictionary: OleVariant; 
                           var CustomDictionary2: OleVariant; var CustomDictionary3: OleVariant; 
                           var CustomDictionary4: OleVariant; var CustomDictionary5: OleVariant; 
                           var CustomDictionary6: OleVariant; var CustomDictionary7: OleVariant; 
                           var CustomDictionary8: OleVariant; var CustomDictionary9: OleVariant; 
                           var CustomDictionary10: OleVariant): WordBool; dispid 324;
    function GetSpellingSuggestions(const Word: WideString; var CustomDictionary: OleVariant; 
                                    var IgnoreUppercase: OleVariant; 
                                    var MainDictionary: OleVariant; var SuggestionMode: OleVariant; 
                                    var CustomDictionary2: OleVariant; 
                                    var CustomDictionary3: OleVariant; 
                                    var CustomDictionary4: OleVariant; 
                                    var CustomDictionary5: OleVariant; 
                                    var CustomDictionary6: OleVariant; 
                                    var CustomDictionary7: OleVariant; 
                                    var CustomDictionary8: OleVariant; 
                                    var CustomDictionary9: OleVariant; 
                                    var CustomDictionary10: OleVariant): SpellingSuggestions; dispid 327;
    procedure BeginCheckSpelling; dispid 4099;
    procedure EndCheckSpelling; dispid 4100;
    property Dialogs: Dialogs readonly dispid 19;
    property TaskPanes: TaskPanes readonly dispid 457;
    property Caption: WideString dispid 80;
    property DisplayAlerts: WpsAlertLevel dispid 94;
    property AdvApiRoot: AdvApiRoot readonly dispid 81;
    property EnableAppWindow: WordBool dispid 82;
    property AutoCorrect: AutoCorrect readonly dispid 10;
    property CapsLock: WordBool readonly dispid 48;
    property NumLock: WordBool readonly dispid 49;
    property PluginPlatform: PluginPlatform readonly dispid 328;
    property FontNames: FontNames readonly dispid 11;
    property AutoCircleNumber: WordBool dispid 329;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Documents
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096C-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Documents = interface(_IKsoDispObj)
    ['{0002096C-0000-4B30-A977-D214852036FE}']
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): _Document; safecall;
    procedure Close(var SaveChanges: OleVariant; var OriginalFormat: OleVariant; 
                    var RouteDocument: OleVariant); safecall;
    procedure Save(NoPrompt: WordBool; const OriginalFormat: WideString); safecall;
    function Add(var Template: OleVariant; NewTemplate: WordBool; DocumentType: Integer; 
                 Visible: WordBool): _Document; safecall;
    function Open(const FileName: WideString; ConfirmConversions: WordBool; ReadOnly: WordBool; 
                  AddToRecentFiles: WordBool; const PasswordDocument: WideString; 
                  const PasswordTemplate: WideString; Revert: WordBool; 
                  const WritePasswordDocument: WideString; const WritePasswordTemplate: WideString; 
                  Format: Integer; Encoding: Integer; Visible: WordBool; OpenAndRepair: WordBool; 
                  DocumentDirection: Integer; NoEncodingDialog: WordBool): _Document; safecall;
    function Get__NewEnum: IUnknown; safecall;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  DocumentsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096C-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DocumentsDisp = dispinterface
    ['{0002096C-0000-4B30-A977-D214852036FE}']
    property Count: Integer readonly dispid 2;
    function Item(var Index: OleVariant): _Document; dispid 0;
    procedure Close(var SaveChanges: OleVariant; var OriginalFormat: OleVariant; 
                    var RouteDocument: OleVariant); dispid 1105;
    procedure Save(NoPrompt: WordBool; const OriginalFormat: WideString); dispid 13;
    function Add(var Template: OleVariant; NewTemplate: WordBool; DocumentType: Integer; 
                 Visible: WordBool): _Document; dispid 14;
    function Open(const FileName: WideString; ConfirmConversions: WordBool; ReadOnly: WordBool; 
                  AddToRecentFiles: WordBool; const PasswordDocument: WideString; 
                  const PasswordTemplate: WideString; Revert: WordBool; 
                  const WritePasswordDocument: WideString; const WritePasswordTemplate: WideString; 
                  Format: Integer; Encoding: Integer; Visible: WordBool; OpenAndRepair: WordBool; 
                  DocumentDirection: Integer; NoEncodingDialog: WordBool): _Document; dispid 18;
    property _NewEnum: IUnknown readonly dispid -4;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: _Document
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096B-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  _Document = interface(_IKsoDispObj)
    ['{0002096B-0000-4B30-A977-D214852036FE}']
    function Undo(Times: Integer): WordBool; safecall;
    function Redo(Times: Integer): WordBool; safecall;
    function Get_Content: Range; safecall;
    function Get_BuiltInDocumentProperties: IDispatch; safecall;
    function Get_CustomDocumentProperties: IDispatch; safecall;
    function Get_Name: WideString; safecall;
    function Get_Path: WideString; safecall;
    function Get_Bookmarks: Bookmarks; safecall;
    function Get_Tables: Tables; safecall;
    function Get_Footnotes: Footnotes; safecall;
    function Get_Endnotes: Endnotes; safecall;
    function Get_Comments: Comments; safecall;
    function Get_type_: WpsDocumentType; safecall;
    function Get_Sections: Sections; safecall;
    function Get_Paragraphs: Paragraphs; safecall;
    function Get_Fields: Fields; safecall;
    function Get_Styles: Styles; safecall;
    function Get_Frames: Frames; safecall;
    function Get_TablesOfFigures: TablesOfFigures; safecall;
    function Get_MailMerge: MailMerge; safecall;
    function Get_FullName: WideString; safecall;
    function Get_TablesOfContents: TablesOfContents; safecall;
    function Get_PageSetup: PageSetup; safecall;
    procedure Set_PageSetup(const prop: PageSetup); safecall;
    function Get_Windows: Windows; safecall;
    function Get_Saved: WordBool; safecall;
    procedure Set_Saved(prop: WordBool); safecall;
    function Get_ActiveWindow: Window; safecall;
    function Get_Kind: WpsDocumentKind; safecall;
    procedure Set_Kind(prop: WpsDocumentKind); safecall;
    function Get_ReadOnly: WordBool; safecall;
    function Get_DefaultTabStop: Single; safecall;
    procedure Set_DefaultTabStop(prop: Single); safecall;
    function Get_Hyperlinks: Hyperlinks; safecall;
    function Get_Shapes: Shapes; safecall;
    function Get_ListTemplates: ListTemplates; safecall;
    function Get_Lists: Lists; safecall;
    function Get_UpdateStylesOnOpen: WordBool; safecall;
    procedure Set_UpdateStylesOnOpen(prop: WordBool); safecall;
    function Get_AttachedTemplate: OleVariant; safecall;
    procedure Set_AttachedTemplate(var prop: OleVariant); safecall;
    function Get_InlineShapes: InlineShapes; safecall;
    function Get_ListParagraphs: ListParagraphs; safecall;
    procedure Set_Password(const Param1: WideString); safecall;
    procedure Set_WritePassword(const Param1: WideString); safecall;
    function Get_HasPassword: WordBool; safecall;
    function Get_ReadOnlyRecommended: WordBool; safecall;
    procedure Set_ReadOnlyRecommended(prop: WordBool); safecall;
    function Get_UserControl: WordBool; safecall;
    procedure Set_UserControl(prop: WordBool); safecall;
    function Get_SnapToGrid: WordBool; safecall;
    procedure Set_SnapToGrid(prop: WordBool); safecall;
    function Get_SnapToShapes: WordBool; safecall;
    procedure Set_SnapToShapes(prop: WordBool); safecall;
    function Get_GridDistanceHorizontal: Single; safecall;
    procedure Set_GridDistanceHorizontal(prop: Single); safecall;
    function Get_GridDistanceVertical: Single; safecall;
    procedure Set_GridDistanceVertical(prop: Single); safecall;
    function Get_GridOriginHorizontal: Single; safecall;
    procedure Set_GridOriginHorizontal(prop: Single); safecall;
    function Get_GridOriginVertical: Single; safecall;
    procedure Set_GridOriginVertical(prop: Single); safecall;
    function Get_GridSpaceBetweenHorizontalLines: Integer; safecall;
    procedure Set_GridSpaceBetweenHorizontalLines(prop: Integer); safecall;
    function Get_GridSpaceBetweenVerticalLines: Integer; safecall;
    procedure Set_GridSpaceBetweenVerticalLines(prop: Integer); safecall;
    function Get_GridOriginFromMargin: WordBool; safecall;
    procedure Set_GridOriginFromMargin(prop: WordBool); safecall;
    function Get_KerningByAlgorithm: WordBool; safecall;
    procedure Set_KerningByAlgorithm(prop: WordBool); safecall;
    function Get_JustificationMode: WpsJustificationMode; safecall;
    procedure Set_JustificationMode(prop: WpsJustificationMode); safecall;
    function Get_FarEastLineBreakLevel: WpsFarEastLineBreakLevel; safecall;
    procedure Set_FarEastLineBreakLevel(prop: WpsFarEastLineBreakLevel); safecall;
    function Get_NoLineBreakBefore: WideString; safecall;
    procedure Set_NoLineBreakBefore(const prop: WideString); safecall;
    function Get_NoLineBreakAfter: WideString; safecall;
    procedure Set_NoLineBreakAfter(const prop: WideString); safecall;
    function Get_PasswordEncryptionProvider: WideString; safecall;
    function Get_PasswordEncryptionAlgorithm: WideString; safecall;
    function Get_PasswordEncryptionKeyLength: Integer; safecall;
    function Get_PasswordEncryptionFileProperties: WordBool; safecall;
    procedure SetPasswordEncryptionOptions(const PasswordEncryptionProvider: WideString; 
                                           const PasswordEncryptionAlgorithm: WideString; 
                                           PasswordEncryptionKeyLength: Integer; 
                                           var PasswordEncryptionFileProperties: OleVariant); safecall;
    function Get_FormattingShowFilter: WpsShowFilter; safecall;
    procedure Set_FormattingShowFilter(prop: WpsShowFilter); safecall;
    procedure Close(var SaveChanges: OleVariant; var OriginalFormat: OleVariant; 
                    var RouteDocument: OleVariant); safecall;
    procedure Repaginate; safecall;
    procedure FitToPages; safecall;
    procedure Select; safecall;
    procedure Save; safecall;
    procedure SendMail; safecall;
    function Range(Start: Integer; End_: Integer): Range; safecall;
    procedure Activate; safecall;
    procedure PrintPreview; safecall;
    procedure MakeCompatibilityDefault; safecall;
    procedure Unprotect(var Password: OleVariant); safecall;
    procedure CopyStylesFromTemplate(const Template: WideString); safecall;
    procedure UpdateStyles; safecall;
    procedure RemoveNumbers(var NumberType: OleVariant); safecall;
    procedure ConvertNumbersToText(var NumberType: OleVariant); safecall;
    function CountNumberedItems(var NumberType: OleVariant; var Level: OleVariant): Integer; safecall;
    procedure UpdateSummaryProperties; safecall;
    function GetCrossReferenceItems(var ReferenceType: OleVariant): OleVariant; safecall;
    procedure UndoClear; safecall;
    procedure ClosePrintPreview; safecall;
    procedure PrintOut(Background: WordBool; Append: WordBool; Range: WpsPrintOutRange; 
                       const OutputFileName: WideString; From: SYSINT; To_: SYSINT; 
                       Item: WpsPrintOutItem; Copies: SYSINT; const Pages: WideString; 
                       PageType: WpsPrintOutPages; PrintToFile: WordBool; Collate: WordBool; 
                       var ActivePrinterMacGX: OleVariant; ManualDuplexPrint: WordBool; 
                       PrintZoomColumn: Integer; PrintZoomRow: Integer; 
                       PrintZoomPaperWidth: Integer; PrintZoomPaperHeight: Integer; 
                       FlipPrint: WordBool; PaperTray: WpsPaperTray; CutterLine: WordBool; 
                       PaperOrder: WpsPaperOrder); safecall;
    function Get_DefaultTableStyle: OleVariant; safecall;
    procedure SetDefaultTableStyle(var Style: OleVariant; SetInTemplate: WordBool); safecall;
    procedure DeleteAllComments; safecall;
    procedure DeleteAllCommentsShown; safecall;
    procedure SaveAs(const FileName: WideString; var FileFormat: OleVariant; 
                     LockComments: WordBool; const Password: WideString; 
                     AddToRecentFiles: WordBool; const WritePassword: WideString; 
                     ReadOnlyRecommended: WordBool; EmbedTrueTypeFonts: WordBool; 
                     SaveNativePictureFormat: WordBool; SaveFormsData: WordBool; 
                     SaveAsAOCELetter: WordBool; Encoding: Integer; InsertLineBreaks: WordBool; 
                     AllowSubstitutions: WordBool; LineEnding: Integer; AddBiDiMarks: WordBool); safecall;
    function Get_RemoveDateAndTime: WordBool; safecall;
    procedure Set_RemoveDateAndTime(prop: WordBool); safecall;
    procedure Protect(Type_: WpsProtectionType; var NoReset: OleVariant; var Password: OleVariant; 
                      var UseIRM: OleVariant; var EnforceStyleLock: OleVariant); safecall;
    function Get_SaveFormat: WideString; safecall;
    function Get_ProtectionType: WpsProtectionType; safecall;
    function Get_TrackRevisions: WordBool; safecall;
    procedure Set_TrackRevisions(prop: WordBool); safecall;
    function Get_PrintRevisions: WordBool; safecall;
    procedure Set_PrintRevisions(prop: WordBool); safecall;
    function Get_ShowRevisions: WordBool; safecall;
    procedure Set_ShowRevisions(prop: WordBool); safecall;
    procedure AcceptAllRevisions; safecall;
    procedure RejectAllRevisions; safecall;
    procedure AcceptAllRevisionsShown; safecall;
    procedure RejectAllRevisionsShown; safecall;
    function Get_Revisions: Revisions; safecall;
    function Get_OpenEncoding: WpsEncoding; safecall;
    function Get_SaveEncoding: WpsEncoding; safecall;
    function Get_TextLineEnding: WpsLineEndingType; safecall;
    procedure Set_TextLineEnding(prop: WpsLineEndingType); safecall;
    function Get_ExtraColors: ExtraColors; safecall;
    function Get_ActiveView: View; safecall;
    function Get__Selection: Selection; safecall;
    function Get_WordStat: WordStat; safecall;
    function Get_Container: IDispatch; safecall;
    procedure BeginJob; safecall;
    procedure EndJob(var JobName: OleVariant; var bCommit: OleVariant); safecall;
    procedure ExportPdf(const PdfFilePath: WideString; const UserPassword: WideString; 
                        const MasterPassword: WideString); safecall;
    function Get_StyleLevel(const Style: Style): WpsPdfStyleLevel; safecall;
    procedure Set_StyleLevel(const Style: Style; prop: WpsPdfStyleLevel); safecall;
    function Get_Compatibility(Compatibility: WpsCompatibility): WordBool; safecall;
    procedure Set_Compatibility(Compatibility: WpsCompatibility; prop: WordBool); safecall;
    function Get_KRM: IDispatch; safecall;
    function Get_ESeal: IDispatch; safecall;
    function Get_ClickAndTypeParagraphStyle: WideString; safecall;
    procedure Set_ClickAndTypeParagraphStyle(const prop: WideString); safecall;
    function Get_UserMode: WordBool; safecall;
    procedure Set_UserMode(boolUserMode: WordBool); safecall;
    function Get_Variables: Variables; safecall;
    function Get_FormFields: FormFields; safecall;
    procedure ResetFormFields; safecall;
    function Get_ShowSpellingErrors: WordBool; safecall;
    procedure Set_ShowSpellingErrors(prop: WordBool); safecall;
    function Get_ShowSpellingIgnoredWords: WordBool; safecall;
    procedure Set_ShowSpellingIgnoredWords(prop: WordBool); safecall;
    function Get_WebOptions: _WebOptions; safecall;
    property Content: Range read Get_Content;
    property BuiltInDocumentProperties: IDispatch read Get_BuiltInDocumentProperties;
    property CustomDocumentProperties: IDispatch read Get_CustomDocumentProperties;
    property Name: WideString read Get_Name;
    property Path: WideString read Get_Path;
    property Bookmarks: Bookmarks read Get_Bookmarks;
    property Tables: Tables read Get_Tables;
    property Footnotes: Footnotes read Get_Footnotes;
    property Endnotes: Endnotes read Get_Endnotes;
    property Comments: Comments read Get_Comments;
    property type_: WpsDocumentType read Get_type_;
    property Sections: Sections read Get_Sections;
    property Paragraphs: Paragraphs read Get_Paragraphs;
    property Fields: Fields read Get_Fields;
    property Styles: Styles read Get_Styles;
    property Frames: Frames read Get_Frames;
    property TablesOfFigures: TablesOfFigures read Get_TablesOfFigures;
    property MailMerge: MailMerge read Get_MailMerge;
    property FullName: WideString read Get_FullName;
    property TablesOfContents: TablesOfContents read Get_TablesOfContents;
    property PageSetup: PageSetup read Get_PageSetup write Set_PageSetup;
    property Windows: Windows read Get_Windows;
    property Saved: WordBool read Get_Saved write Set_Saved;
    property ActiveWindow: Window read Get_ActiveWindow;
    property Kind: WpsDocumentKind read Get_Kind write Set_Kind;
    property ReadOnly: WordBool read Get_ReadOnly;
    property DefaultTabStop: Single read Get_DefaultTabStop write Set_DefaultTabStop;
    property Hyperlinks: Hyperlinks read Get_Hyperlinks;
    property Shapes: Shapes read Get_Shapes;
    property ListTemplates: ListTemplates read Get_ListTemplates;
    property Lists: Lists read Get_Lists;
    property UpdateStylesOnOpen: WordBool read Get_UpdateStylesOnOpen write Set_UpdateStylesOnOpen;
    property InlineShapes: InlineShapes read Get_InlineShapes;
    property ListParagraphs: ListParagraphs read Get_ListParagraphs;
    property Password: WideString write Set_Password;
    property WritePassword: WideString write Set_WritePassword;
    property HasPassword: WordBool read Get_HasPassword;
    property ReadOnlyRecommended: WordBool read Get_ReadOnlyRecommended write Set_ReadOnlyRecommended;
    property UserControl: WordBool read Get_UserControl write Set_UserControl;
    property SnapToGrid: WordBool read Get_SnapToGrid write Set_SnapToGrid;
    property SnapToShapes: WordBool read Get_SnapToShapes write Set_SnapToShapes;
    property GridDistanceHorizontal: Single read Get_GridDistanceHorizontal write Set_GridDistanceHorizontal;
    property GridDistanceVertical: Single read Get_GridDistanceVertical write Set_GridDistanceVertical;
    property GridOriginHorizontal: Single read Get_GridOriginHorizontal write Set_GridOriginHorizontal;
    property GridOriginVertical: Single read Get_GridOriginVertical write Set_GridOriginVertical;
    property GridSpaceBetweenHorizontalLines: Integer read Get_GridSpaceBetweenHorizontalLines write Set_GridSpaceBetweenHorizontalLines;
    property GridSpaceBetweenVerticalLines: Integer read Get_GridSpaceBetweenVerticalLines write Set_GridSpaceBetweenVerticalLines;
    property GridOriginFromMargin: WordBool read Get_GridOriginFromMargin write Set_GridOriginFromMargin;
    property KerningByAlgorithm: WordBool read Get_KerningByAlgorithm write Set_KerningByAlgorithm;
    property JustificationMode: WpsJustificationMode read Get_JustificationMode write Set_JustificationMode;
    property FarEastLineBreakLevel: WpsFarEastLineBreakLevel read Get_FarEastLineBreakLevel write Set_FarEastLineBreakLevel;
    property NoLineBreakBefore: WideString read Get_NoLineBreakBefore write Set_NoLineBreakBefore;
    property NoLineBreakAfter: WideString read Get_NoLineBreakAfter write Set_NoLineBreakAfter;
    property PasswordEncryptionProvider: WideString read Get_PasswordEncryptionProvider;
    property PasswordEncryptionAlgorithm: WideString read Get_PasswordEncryptionAlgorithm;
    property PasswordEncryptionKeyLength: Integer read Get_PasswordEncryptionKeyLength;
    property PasswordEncryptionFileProperties: WordBool read Get_PasswordEncryptionFileProperties;
    property FormattingShowFilter: WpsShowFilter read Get_FormattingShowFilter write Set_FormattingShowFilter;
    property DefaultTableStyle: OleVariant read Get_DefaultTableStyle;
    property RemoveDateAndTime: WordBool read Get_RemoveDateAndTime write Set_RemoveDateAndTime;
    property SaveFormat: WideString read Get_SaveFormat;
    property ProtectionType: WpsProtectionType read Get_ProtectionType;
    property TrackRevisions: WordBool read Get_TrackRevisions write Set_TrackRevisions;
    property PrintRevisions: WordBool read Get_PrintRevisions write Set_PrintRevisions;
    property ShowRevisions: WordBool read Get_ShowRevisions write Set_ShowRevisions;
    property Revisions: Revisions read Get_Revisions;
    property OpenEncoding: WpsEncoding read Get_OpenEncoding;
    property SaveEncoding: WpsEncoding read Get_SaveEncoding;
    property TextLineEnding: WpsLineEndingType read Get_TextLineEnding write Set_TextLineEnding;
    property ExtraColors: ExtraColors read Get_ExtraColors;
    property ActiveView: View read Get_ActiveView;
    property _Selection: Selection read Get__Selection;
    property WordStat: WordStat read Get_WordStat;
    property Container: IDispatch read Get_Container;
    property StyleLevel[const Style: Style]: WpsPdfStyleLevel read Get_StyleLevel write Set_StyleLevel;
    property Compatibility[Compatibility: WpsCompatibility]: WordBool read Get_Compatibility write Set_Compatibility;
    property KRM: IDispatch read Get_KRM;
    property ESeal: IDispatch read Get_ESeal;
    property ClickAndTypeParagraphStyle: WideString read Get_ClickAndTypeParagraphStyle write Set_ClickAndTypeParagraphStyle;
    property UserMode: WordBool read Get_UserMode write Set_UserMode;
    property Variables: Variables read Get_Variables;
    property FormFields: FormFields read Get_FormFields;
    property ShowSpellingErrors: WordBool read Get_ShowSpellingErrors write Set_ShowSpellingErrors;
    property ShowSpellingIgnoredWords: WordBool read Get_ShowSpellingIgnoredWords write Set_ShowSpellingIgnoredWords;
    property WebOptions: _WebOptions read Get_WebOptions;
  end;

// *********************************************************************//
// DispIntf:  _DocumentDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096B-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  _DocumentDisp = dispinterface
    ['{0002096B-0000-4B30-A977-D214852036FE}']
    function Undo(Times: Integer): WordBool; dispid 116;
    function Redo(Times: Integer): WordBool; dispid 117;
    property Content: Range readonly dispid 41;
    property BuiltInDocumentProperties: IDispatch readonly dispid 1;
    property CustomDocumentProperties: IDispatch readonly dispid 2;
    property Name: WideString readonly dispid 0;
    property Path: WideString readonly dispid 3;
    property Bookmarks: Bookmarks readonly dispid 4;
    property Tables: Tables readonly dispid 6;
    property Footnotes: Footnotes readonly dispid 7;
    property Endnotes: Endnotes readonly dispid 8;
    property Comments: Comments readonly dispid 9;
    property type_: WpsDocumentType readonly dispid 10;
    property Sections: Sections readonly dispid 15;
    property Paragraphs: Paragraphs readonly dispid 16;
    property Fields: Fields readonly dispid 20;
    property Styles: Styles readonly dispid 22;
    property Frames: Frames readonly dispid 23;
    property TablesOfFigures: TablesOfFigures readonly dispid 25;
    property MailMerge: MailMerge readonly dispid 27;
    property FullName: WideString readonly dispid 29;
    property TablesOfContents: TablesOfContents readonly dispid 31;
    property PageSetup: PageSetup dispid 1101;
    property Windows: Windows readonly dispid 34;
    property Saved: WordBool dispid 40;
    property ActiveWindow: Window readonly dispid 42;
    property Kind: WpsDocumentKind dispid 43;
    property ReadOnly: WordBool readonly dispid 44;
    property DefaultTabStop: Single dispid 48;
    property Hyperlinks: Hyperlinks readonly dispid 61;
    property Shapes: Shapes readonly dispid 62;
    property ListTemplates: ListTemplates readonly dispid 63;
    property Lists: Lists readonly dispid 64;
    property UpdateStylesOnOpen: WordBool dispid 66;
    function AttachedTemplate: OleVariant; dispid 67;
    property InlineShapes: InlineShapes readonly dispid 68;
    property ListParagraphs: ListParagraphs readonly dispid 84;
    property Password: WideString writeonly dispid 85;
    property WritePassword: WideString writeonly dispid 86;
    property HasPassword: WordBool readonly dispid 87;
    property ReadOnlyRecommended: WordBool dispid 52;
    property UserControl: WordBool dispid 92;
    property SnapToGrid: WordBool dispid 300;
    property SnapToShapes: WordBool dispid 301;
    property GridDistanceHorizontal: Single dispid 302;
    property GridDistanceVertical: Single dispid 303;
    property GridOriginHorizontal: Single dispid 304;
    property GridOriginVertical: Single dispid 305;
    property GridSpaceBetweenHorizontalLines: Integer dispid 306;
    property GridSpaceBetweenVerticalLines: Integer dispid 307;
    property GridOriginFromMargin: WordBool dispid 308;
    property KerningByAlgorithm: WordBool dispid 309;
    property JustificationMode: WpsJustificationMode dispid 310;
    property FarEastLineBreakLevel: WpsFarEastLineBreakLevel dispid 311;
    property NoLineBreakBefore: WideString dispid 312;
    property NoLineBreakAfter: WideString dispid 313;
    property PasswordEncryptionProvider: WideString readonly dispid 367;
    property PasswordEncryptionAlgorithm: WideString readonly dispid 368;
    property PasswordEncryptionKeyLength: Integer readonly dispid 369;
    property PasswordEncryptionFileProperties: WordBool readonly dispid 370;
    procedure SetPasswordEncryptionOptions(const PasswordEncryptionProvider: WideString; 
                                           const PasswordEncryptionAlgorithm: WideString; 
                                           PasswordEncryptionKeyLength: Integer; 
                                           var PasswordEncryptionFileProperties: OleVariant); dispid 361;
    property FormattingShowFilter: WpsShowFilter dispid 452;
    procedure Close(var SaveChanges: OleVariant; var OriginalFormat: OleVariant; 
                    var RouteDocument: OleVariant); dispid 1105;
    procedure Repaginate; dispid 103;
    procedure FitToPages; dispid 104;
    procedure Select; dispid 65535;
    procedure Save; dispid 108;
    procedure SendMail; dispid 110;
    function Range(Start: Integer; End_: Integer): Range; dispid 2000;
    procedure Activate; dispid 113;
    procedure PrintPreview; dispid 114;
    procedure MakeCompatibilityDefault; dispid 119;
    procedure Unprotect(var Password: OleVariant); dispid 121;
    procedure CopyStylesFromTemplate(const Template: WideString); dispid 126;
    procedure UpdateStyles; dispid 127;
    procedure RemoveNumbers(var NumberType: OleVariant); dispid 140;
    procedure ConvertNumbersToText(var NumberType: OleVariant); dispid 141;
    function CountNumberedItems(var NumberType: OleVariant; var Level: OleVariant): Integer; dispid 142;
    procedure UpdateSummaryProperties; dispid 146;
    function GetCrossReferenceItems(var ReferenceType: OleVariant): OleVariant; dispid 147;
    procedure UndoClear; dispid 254;
    procedure ClosePrintPreview; dispid 258;
    procedure PrintOut(Background: WordBool; Append: WordBool; Range: WpsPrintOutRange; 
                       const OutputFileName: WideString; From: SYSINT; To_: SYSINT; 
                       Item: WpsPrintOutItem; Copies: SYSINT; const Pages: WideString; 
                       PageType: WpsPrintOutPages; PrintToFile: WordBool; Collate: WordBool; 
                       var ActivePrinterMacGX: OleVariant; ManualDuplexPrint: WordBool; 
                       PrintZoomColumn: Integer; PrintZoomRow: Integer; 
                       PrintZoomPaperWidth: Integer; PrintZoomPaperHeight: Integer; 
                       FlipPrint: WordBool; PaperTray: WpsPaperTray; CutterLine: WordBool; 
                       PaperOrder: WpsPaperOrder); dispid 446;
    property DefaultTableStyle: OleVariant readonly dispid 365;
    procedure SetDefaultTableStyle(var Style: OleVariant; SetInTemplate: WordBool); dispid 366;
    procedure DeleteAllComments; dispid 371;
    procedure DeleteAllCommentsShown; dispid 374;
    procedure SaveAs(const FileName: WideString; var FileFormat: OleVariant; 
                     LockComments: WordBool; const Password: WideString; 
                     AddToRecentFiles: WordBool; const WritePassword: WideString; 
                     ReadOnlyRecommended: WordBool; EmbedTrueTypeFonts: WordBool; 
                     SaveNativePictureFormat: WordBool; SaveFormsData: WordBool; 
                     SaveAsAOCELetter: WordBool; Encoding: Integer; InsertLineBreaks: WordBool; 
                     AllowSubstitutions: WordBool; LineEnding: Integer; AddBiDiMarks: WordBool); dispid 376;
    property RemoveDateAndTime: WordBool dispid 484;
    procedure Protect(Type_: WpsProtectionType; var NoReset: OleVariant; var Password: OleVariant; 
                      var UseIRM: OleVariant; var EnforceStyleLock: OleVariant); dispid 467;
    property SaveFormat: WideString readonly dispid 59;
    property ProtectionType: WpsProtectionType readonly dispid 60;
    property TrackRevisions: WordBool dispid 314;
    property PrintRevisions: WordBool dispid 315;
    property ShowRevisions: WordBool dispid 316;
    procedure AcceptAllRevisions; dispid 317;
    procedure RejectAllRevisions; dispid 318;
    procedure AcceptAllRevisionsShown; dispid 372;
    procedure RejectAllRevisionsShown; dispid 373;
    property Revisions: Revisions readonly dispid 30;
    property OpenEncoding: WpsEncoding readonly dispid 332;
    property SaveEncoding: WpsEncoding readonly dispid 333;
    property TextLineEnding: WpsLineEndingType dispid 358;
    property ExtraColors: ExtraColors readonly dispid 509677568;
    property ActiveView: View readonly dispid 4097;
    property _Selection: Selection readonly dispid 4098;
    property WordStat: WordStat readonly dispid 4099;
    property Container: IDispatch readonly dispid 4100;
    procedure BeginJob; dispid 4101;
    procedure EndJob(var JobName: OleVariant; var bCommit: OleVariant); dispid 4102;
    procedure ExportPdf(const PdfFilePath: WideString; const UserPassword: WideString; 
                        const MasterPassword: WideString); dispid 4103;
    property StyleLevel[const Style: Style]: WpsPdfStyleLevel dispid 4104;
    property Compatibility[Compatibility: WpsCompatibility]: WordBool dispid 4105;
    property KRM: IDispatch readonly dispid 4112;
    property ESeal: IDispatch readonly dispid 4113;
    property ClickAndTypeParagraphStyle: WideString dispid 4114;
    property UserMode: WordBool dispid 4115;
    property Variables: Variables readonly dispid 26;
    property FormFields: FormFields readonly dispid 21;
    procedure ResetFormFields; dispid 375;
    property ShowSpellingErrors: WordBool dispid 73;
    property ShowSpellingIgnoredWords: WordBool dispid 4116;
    property WebOptions: _WebOptions readonly dispid 493;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Range
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095E-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Range = interface(_IKsoDispObj)
    ['{0002095E-0000-4B30-A977-D214852036FE}']
    function Get_Text: WideString; safecall;
    procedure Set_Text(const prop: WideString); safecall;
    function Get_FormattedText: Range; safecall;
    procedure Set_FormattedText(const prop: Range); safecall;
    function Get_Start: Integer; safecall;
    function Get_End_: Integer; safecall;
    function Get_Font: _Font; safecall;
    procedure Set_Font(const prop: _Font); safecall;
    function Get_Duplicate: Range; safecall;
    function Get_StoryType: WpsStoryType; safecall;
    function Get_Tables: Tables; safecall;
    function Get_Footnotes: Footnotes; safecall;
    function Get_Endnotes: Endnotes; safecall;
    function Get_Comments: Comments; safecall;
    function Get_Cells: Cells; safecall;
    function Get_Sections: Sections; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Paragraphs: Paragraphs; safecall;
    function Get_Fields: Fields; safecall;
    function Get_Frames: Frames; safecall;
    function Get_ParagraphFormat: _ParagraphFormat; safecall;
    procedure Set_ParagraphFormat(const prop: _ParagraphFormat); safecall;
    function Get_ListFormat: ListFormat; safecall;
    function Get_Bookmarks: Bookmarks; safecall;
    function Get_Bold: Integer; safecall;
    procedure Set_Bold(prop: Integer); safecall;
    function Get_Italic: Integer; safecall;
    procedure Set_Italic(prop: Integer); safecall;
    function Get_Underline: WpsUnderline; safecall;
    procedure Set_Underline(prop: WpsUnderline); safecall;
    function Get_EmphasisMark: WpsEmphasisMark; safecall;
    procedure Set_EmphasisMark(prop: WpsEmphasisMark); safecall;
    function Get_Style: OleVariant; safecall;
    procedure Set_Style(var prop: OleVariant); safecall;
    function Get_StoryLength: Integer; safecall;
    function Get_Hyperlinks: Hyperlinks; safecall;
    function Get_ListParagraphs: ListParagraphs; safecall;
    function Get_HighlightColor: WpsColor; safecall;
    procedure Set_HighlightColor(prop: WpsColor); safecall;
    function Get_Columns: Columns; safecall;
    function Get_Rows: Rows; safecall;
    function Get_BookmarkID: Integer; safecall;
    function Get_PreviousBookmarkID: Integer; safecall;
    function Get_Information(Type_: WpsInformation): OleVariant; safecall;
    function Get_PageSetup: PageSetup; safecall;
    procedure Set_PageSetup(const prop: PageSetup); safecall;
    function Get_ShapeRange: ShapeRange; safecall;
    function Get_Case_: WpsCharacterCase; safecall;
    procedure Set_Case_(prop: WpsCharacterCase); safecall;
    function Get_Orientation: WpsTextOrientation; safecall;
    procedure Set_Orientation(prop: WpsTextOrientation); safecall;
    function Get_InlineShapes: InlineShapes; safecall;
    function Get_NextStoryRange: Range; safecall;
    procedure Select; safecall;
    procedure SetRange(Start: Integer; End_: Integer); safecall;
    procedure Collapse(Direction: WpsCollapseDirection); safecall;
    procedure InsertBefore(const Text: WideString); safecall;
    procedure InsertAfter(const Text: WideString); safecall;
    function Next(Unit_: WpsUnits; Count: Integer): Range; safecall;
    function Previous(Unit_: WpsUnits; Count: Integer): Range; safecall;
    function StartOf(Unit_: WpsUnits; Extend: WpsMovementType): Integer; safecall;
    function EndOf(Unit_: WpsUnits; Extend: WpsMovementType): Integer; safecall;
    function Move(Unit_: WpsUnits; Count: Integer): Integer; safecall;
    function MoveStart(Unit_: WpsUnits; Count: Integer): Integer; safecall;
    function MoveEnd(Unit_: WpsUnits; Count: Integer): Integer; safecall;
    procedure Cut; safecall;
    procedure Copy; safecall;
    procedure Paste; safecall;
    procedure InsertBreak(Type_: WpsBreakType); safecall;
    function InStory(const Range: Range): WordBool; safecall;
    function InRange(const Range: Range): WordBool; safecall;
    function Delete(Unit_: WpsUnits; Count: Integer): Integer; safecall;
    procedure WholeStory; safecall;
    function Expand(Unit_: WpsUnits): Integer; safecall;
    procedure InsertParagraph; safecall;
    procedure InsertParagraphAfter; safecall;
    procedure InsertSymbol(CharacterNumber: Integer; const Font: WideString; Unicode: WordBool; 
                           var Bias: OleVariant); safecall;
    function IsEqual(const Range: Range): WordBool; safecall;
    function GoTo_(What: WpsGoToItem; Which: WpsGoToDirection; Count: Integer; 
                   const Name: WideString): Range; safecall;
    function GoToNext(What: WpsGoToItem): Range; safecall;
    function GoToPrevious(What: WpsGoToItem): Range; safecall;
    procedure InsertParagraphBefore; safecall;
    procedure InsertDateTime(const DateTimeFormat: WideString; InsertAsField: WordBool; 
                             InsertAsFullWidth: WordBool; DateLanguage: WpsDateLanguage; 
                             CalendarType: WpsCalendarTypeBi); safecall;
    function Get_Document: _Document; safecall;
    function Get_FootnoteOptions: FootnoteOptions; safecall;
    function Get_EndnoteOptions: EndnoteOptions; safecall;
    procedure InsertCaption(var Label_: OleVariant; const Title: WideString; 
                            const TitleAutoText: WideString; Position: WpsCaptionPosition; 
                            ExcludeLabel: WordBool); safecall;
    procedure ModifyEnclosure(Style: WpsEncloseStyle; Symbol: WpsEnclosureType; 
                              const EnclosedText: WideString); safecall;
    procedure PhoneticGuide(const Text: WideString; Alignment: WpsPhoneticGuideAlignmentType; 
                            Raise_: Integer; FontSize: Integer; const FontName: WideString); safecall;
    function Get_CombineCharacters: WordBool; safecall;
    procedure Set_CombineCharacters(prop: WordBool); safecall;
    function Get_DropCap: DropCap; safecall;
    procedure EnableDropCap(Position: WpsDropPosition; const FontName: WideString; 
                            LineToDrop: Integer; DistanceFromText: Single); safecall;
    function Get_SymbolFontName: WideString; safecall;
    function Get_Symbol: Integer; safecall;
    function Get_WordStat: WordStat; safecall;
    function Get_Revisions: Revisions; safecall;
    procedure TCSCConverter(WpsTCSCCDirection: WpsTCSCConverterDirection; CommonTerms: WordBool; 
                            UseVariants: WordBool); safecall;
    procedure ReplaceText(const ReplaceText: WideString); safecall;
    procedure CheckSpelling(var CustomDictionary: OleVariant; var IgnoreUppercase: OleVariant; 
                            var AlwaysSuggest: OleVariant; var CustomDictionary2: OleVariant; 
                            var CustomDictionary3: OleVariant; var CustomDictionary4: OleVariant; 
                            var CustomDictionary5: OleVariant; var CustomDictionary6: OleVariant; 
                            var CustomDictionary7: OleVariant; var CustomDictionary8: OleVariant; 
                            var CustomDictionary9: OleVariant; var CustomDictionary10: OleVariant); safecall;
    function GetSpellingSuggestions(var CustomDictionary: OleVariant; 
                                    var IgnoreUppercase: OleVariant; 
                                    var MainDictionary: OleVariant; var SuggestionMode: OleVariant; 
                                    var CustomDictionary2: OleVariant; 
                                    var CustomDictionary3: OleVariant; 
                                    var CustomDictionary4: OleVariant; 
                                    var CustomDictionary5: OleVariant; 
                                    var CustomDictionary6: OleVariant; 
                                    var CustomDictionary7: OleVariant; 
                                    var CustomDictionary8: OleVariant; 
                                    var CustomDictionary9: OleVariant; 
                                    var CustomDictionary10: OleVariant): SpellingSuggestions; safecall;
    function Get_SpellingChecked: WordBool; safecall;
    procedure Set_SpellingChecked(prop: WordBool); safecall;
    function Get_SpellingErrors: ProofreadingErrors; safecall;
    function Get_EnhMetaFileBits: OleVariant; safecall;
    procedure Set_Start(prop: Integer); safecall;
    procedure Set_End_(prop: Integer); safecall;
    function Get_LanguageID: WpsLanguageID; safecall;
    procedure Set_LanguageID(prop: WpsLanguageID); safecall;
    function Get_LanguageIDFarEast: WpsLanguageID; safecall;
    procedure Set_LanguageIDFarEast(prop: WpsLanguageID); safecall;
    function Get_LanguageIDOther: WpsLanguageID; safecall;
    procedure Set_LanguageIDOther(prop: WpsLanguageID); safecall;
    function Get_PlainText: WideString; safecall;
    function Get_IsEndOfRowMark: WordBool; safecall;
    function Get_TwoLinesInOne: WpsTwoLinesInOne; safecall;
    procedure Set_TwoLinesInOne(prop: WpsTwoLinesInOne); safecall;
    function ConvertToTable(var Separator: OleVariant; var NumRows: OleVariant; 
                            var NumColumns: OleVariant; var InitialColumnWidth: OleVariant; 
                            var Format: OleVariant; var ApplyBorders: OleVariant; 
                            var ApplyShading: OleVariant; var ApplyFont: OleVariant; 
                            var ApplyColor: OleVariant; var ApplyHeadingRows: OleVariant; 
                            var ApplyLastRow: OleVariant; var ApplyFirstColumn: OleVariant; 
                            var ApplyLastColumn: OleVariant; var AutoFit: OleVariant; 
                            var AutoFitBehavior: OleVariant; var DefaultTableBehavior: OleVariant): Table; safecall;
    procedure InsertFile(const FileName: WideString; var Range: OleVariant; 
                         var ConfirmConversions: OleVariant; var Link: OleVariant; 
                         var Attachment: OleVariant); safecall;
    function Get_FormFields: FormFields; safecall;
    function Get_Editors: Editors; safecall;
    procedure Relocate(prop: WpsRelocate); safecall;
    function Get_BoldBi: Integer; safecall;
    procedure Set_BoldBi(prop: Integer); safecall;
    function Get_ItalicBi: Integer; safecall;
    procedure Set_ItalicBi(prop: Integer); safecall;
    property Text: WideString read Get_Text write Set_Text;
    property FormattedText: Range read Get_FormattedText write Set_FormattedText;
    property Start: Integer read Get_Start write Set_Start;
    property End_: Integer read Get_End_ write Set_End_;
    property Font: _Font read Get_Font write Set_Font;
    property Duplicate: Range read Get_Duplicate;
    property StoryType: WpsStoryType read Get_StoryType;
    property Tables: Tables read Get_Tables;
    property Footnotes: Footnotes read Get_Footnotes;
    property Endnotes: Endnotes read Get_Endnotes;
    property Comments: Comments read Get_Comments;
    property Cells: Cells read Get_Cells;
    property Sections: Sections read Get_Sections;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Paragraphs: Paragraphs read Get_Paragraphs;
    property Fields: Fields read Get_Fields;
    property Frames: Frames read Get_Frames;
    property ParagraphFormat: _ParagraphFormat read Get_ParagraphFormat write Set_ParagraphFormat;
    property ListFormat: ListFormat read Get_ListFormat;
    property Bookmarks: Bookmarks read Get_Bookmarks;
    property Bold: Integer read Get_Bold write Set_Bold;
    property Italic: Integer read Get_Italic write Set_Italic;
    property Underline: WpsUnderline read Get_Underline write Set_Underline;
    property EmphasisMark: WpsEmphasisMark read Get_EmphasisMark write Set_EmphasisMark;
    property StoryLength: Integer read Get_StoryLength;
    property Hyperlinks: Hyperlinks read Get_Hyperlinks;
    property ListParagraphs: ListParagraphs read Get_ListParagraphs;
    property HighlightColor: WpsColor read Get_HighlightColor write Set_HighlightColor;
    property Columns: Columns read Get_Columns;
    property Rows: Rows read Get_Rows;
    property BookmarkID: Integer read Get_BookmarkID;
    property PreviousBookmarkID: Integer read Get_PreviousBookmarkID;
    property Information[Type_: WpsInformation]: OleVariant read Get_Information;
    property PageSetup: PageSetup read Get_PageSetup write Set_PageSetup;
    property ShapeRange: ShapeRange read Get_ShapeRange;
    property Case_: WpsCharacterCase read Get_Case_ write Set_Case_;
    property Orientation: WpsTextOrientation read Get_Orientation write Set_Orientation;
    property InlineShapes: InlineShapes read Get_InlineShapes;
    property NextStoryRange: Range read Get_NextStoryRange;
    property Document: _Document read Get_Document;
    property FootnoteOptions: FootnoteOptions read Get_FootnoteOptions;
    property EndnoteOptions: EndnoteOptions read Get_EndnoteOptions;
    property CombineCharacters: WordBool read Get_CombineCharacters write Set_CombineCharacters;
    property DropCap: DropCap read Get_DropCap;
    property SymbolFontName: WideString read Get_SymbolFontName;
    property Symbol: Integer read Get_Symbol;
    property WordStat: WordStat read Get_WordStat;
    property Revisions: Revisions read Get_Revisions;
    property SpellingChecked: WordBool read Get_SpellingChecked write Set_SpellingChecked;
    property SpellingErrors: ProofreadingErrors read Get_SpellingErrors;
    property EnhMetaFileBits: OleVariant read Get_EnhMetaFileBits;
    property LanguageID: WpsLanguageID read Get_LanguageID write Set_LanguageID;
    property LanguageIDFarEast: WpsLanguageID read Get_LanguageIDFarEast write Set_LanguageIDFarEast;
    property LanguageIDOther: WpsLanguageID read Get_LanguageIDOther write Set_LanguageIDOther;
    property PlainText: WideString read Get_PlainText;
    property IsEndOfRowMark: WordBool read Get_IsEndOfRowMark;
    property TwoLinesInOne: WpsTwoLinesInOne read Get_TwoLinesInOne write Set_TwoLinesInOne;
    property FormFields: FormFields read Get_FormFields;
    property Editors: Editors read Get_Editors;
    property BoldBi: Integer read Get_BoldBi write Set_BoldBi;
    property ItalicBi: Integer read Get_ItalicBi write Set_ItalicBi;
  end;

// *********************************************************************//
// DispIntf:  RangeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095E-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  RangeDisp = dispinterface
    ['{0002095E-0000-4B30-A977-D214852036FE}']
    property Text: WideString dispid 0;
    property FormattedText: Range dispid 2;
    property Start: Integer dispid 3;
    property End_: Integer dispid 4;
    property Font: _Font dispid 5;
    property Duplicate: Range readonly dispid 6;
    property StoryType: WpsStoryType readonly dispid 7;
    property Tables: Tables readonly dispid 50;
    property Footnotes: Footnotes readonly dispid 54;
    property Endnotes: Endnotes readonly dispid 55;
    property Comments: Comments readonly dispid 56;
    property Cells: Cells readonly dispid 57;
    property Sections: Sections readonly dispid 58;
    property Borders: Borders dispid 1100;
    property Paragraphs: Paragraphs readonly dispid 59;
    property Fields: Fields readonly dispid 64;
    property Frames: Frames readonly dispid 66;
    property ParagraphFormat: _ParagraphFormat dispid 1102;
    property ListFormat: ListFormat readonly dispid 68;
    property Bookmarks: Bookmarks readonly dispid 75;
    property Bold: Integer dispid 130;
    property Italic: Integer dispid 131;
    property Underline: WpsUnderline dispid 139;
    property EmphasisMark: WpsEmphasisMark dispid 140;
    function Style: OleVariant; dispid 151;
    property StoryLength: Integer readonly dispid 152;
    property Hyperlinks: Hyperlinks readonly dispid 156;
    property ListParagraphs: ListParagraphs readonly dispid 157;
    property HighlightColor: WpsColor dispid 301;
    property Columns: Columns readonly dispid 302;
    property Rows: Rows readonly dispid 303;
    property BookmarkID: Integer readonly dispid 308;
    property PreviousBookmarkID: Integer readonly dispid 309;
    property Information[Type_: WpsInformation]: OleVariant readonly dispid 313;
    property PageSetup: PageSetup dispid 1101;
    property ShapeRange: ShapeRange readonly dispid 311;
    property Case_: WpsCharacterCase dispid 312;
    property Orientation: WpsTextOrientation dispid 317;
    property InlineShapes: InlineShapes readonly dispid 319;
    property NextStoryRange: Range readonly dispid 320;
    procedure Select; dispid 65535;
    procedure SetRange(Start: Integer; End_: Integer); dispid 100;
    procedure Collapse(Direction: WpsCollapseDirection); dispid 101;
    procedure InsertBefore(const Text: WideString); dispid 102;
    procedure InsertAfter(const Text: WideString); dispid 104;
    function Next(Unit_: WpsUnits; Count: Integer): Range; dispid 105;
    function Previous(Unit_: WpsUnits; Count: Integer): Range; dispid 106;
    function StartOf(Unit_: WpsUnits; Extend: WpsMovementType): Integer; dispid 107;
    function EndOf(Unit_: WpsUnits; Extend: WpsMovementType): Integer; dispid 108;
    function Move(Unit_: WpsUnits; Count: Integer): Integer; dispid 109;
    function MoveStart(Unit_: WpsUnits; Count: Integer): Integer; dispid 110;
    function MoveEnd(Unit_: WpsUnits; Count: Integer): Integer; dispid 111;
    procedure Cut; dispid 119;
    procedure Copy; dispid 120;
    procedure Paste; dispid 121;
    procedure InsertBreak(Type_: WpsBreakType); dispid 122;
    function InStory(const Range: Range): WordBool; dispid 125;
    function InRange(const Range: Range): WordBool; dispid 126;
    function Delete(Unit_: WpsUnits; Count: Integer): Integer; dispid 127;
    procedure WholeStory; dispid 128;
    function Expand(Unit_: WpsUnits): Integer; dispid 129;
    procedure InsertParagraph; dispid 160;
    procedure InsertParagraphAfter; dispid 161;
    procedure InsertSymbol(CharacterNumber: Integer; const Font: WideString; Unicode: WordBool; 
                           var Bias: OleVariant); dispid 164;
    function IsEqual(const Range: Range): WordBool; dispid 171;
    function GoTo_(What: WpsGoToItem; Which: WpsGoToDirection; Count: Integer; 
                   const Name: WideString): Range; dispid 173;
    function GoToNext(What: WpsGoToItem): Range; dispid 174;
    function GoToPrevious(What: WpsGoToItem): Range; dispid 175;
    procedure InsertParagraphBefore; dispid 212;
    procedure InsertDateTime(const DateTimeFormat: WideString; InsertAsField: WordBool; 
                             InsertAsFullWidth: WordBool; DateLanguage: WpsDateLanguage; 
                             CalendarType: WpsCalendarTypeBi); dispid 444;
    property Document: _Document readonly dispid 409;
    property FootnoteOptions: FootnoteOptions readonly dispid 410;
    property EndnoteOptions: EndnoteOptions readonly dispid 411;
    procedure InsertCaption(var Label_: OleVariant; const Title: WideString; 
                            const TitleAutoText: WideString; Position: WpsCaptionPosition; 
                            ExcludeLabel: WordBool); dispid 417;
    procedure ModifyEnclosure(Style: WpsEncloseStyle; Symbol: WpsEnclosureType; 
                              const EnclosedText: WideString); dispid 223;
    procedure PhoneticGuide(const Text: WideString; Alignment: WpsPhoneticGuideAlignmentType; 
                            Raise_: Integer; FontSize: Integer; const FontName: WideString); dispid 224;
    property CombineCharacters: WordBool dispid 267;
    property DropCap: DropCap readonly dispid 4405;
    procedure EnableDropCap(Position: WpsDropPosition; const FontName: WideString; 
                            LineToDrop: Integer; DistanceFromText: Single); dispid 4404;
    property SymbolFontName: WideString readonly dispid 4407;
    property Symbol: Integer readonly dispid 4406;
    property WordStat: WordStat readonly dispid 4408;
    property Revisions: Revisions readonly dispid 150;
    procedure TCSCConverter(WpsTCSCCDirection: WpsTCSCConverterDirection; CommonTerms: WordBool; 
                            UseVariants: WordBool); dispid 499;
    procedure ReplaceText(const ReplaceText: WideString); dispid 4097;
    procedure CheckSpelling(var CustomDictionary: OleVariant; var IgnoreUppercase: OleVariant; 
                            var AlwaysSuggest: OleVariant; var CustomDictionary2: OleVariant; 
                            var CustomDictionary3: OleVariant; var CustomDictionary4: OleVariant; 
                            var CustomDictionary5: OleVariant; var CustomDictionary6: OleVariant; 
                            var CustomDictionary7: OleVariant; var CustomDictionary8: OleVariant; 
                            var CustomDictionary9: OleVariant; var CustomDictionary10: OleVariant); dispid 205;
    function GetSpellingSuggestions(var CustomDictionary: OleVariant; 
                                    var IgnoreUppercase: OleVariant; 
                                    var MainDictionary: OleVariant; var SuggestionMode: OleVariant; 
                                    var CustomDictionary2: OleVariant; 
                                    var CustomDictionary3: OleVariant; 
                                    var CustomDictionary4: OleVariant; 
                                    var CustomDictionary5: OleVariant; 
                                    var CustomDictionary6: OleVariant; 
                                    var CustomDictionary7: OleVariant; 
                                    var CustomDictionary8: OleVariant; 
                                    var CustomDictionary9: OleVariant; 
                                    var CustomDictionary10: OleVariant): SpellingSuggestions; dispid 209;
    property SpellingChecked: WordBool dispid 261;
    property SpellingErrors: ProofreadingErrors readonly dispid 316;
    property EnhMetaFileBits: OleVariant readonly dispid 345;
    property LanguageID: WpsLanguageID dispid 153;
    property LanguageIDFarEast: WpsLanguageID dispid 321;
    property LanguageIDOther: WpsLanguageID dispid 322;
    property PlainText: WideString readonly dispid 3906;
    property IsEndOfRowMark: WordBool readonly dispid 307;
    property TwoLinesInOne: WpsTwoLinesInOne dispid 266;
    function ConvertToTable(var Separator: OleVariant; var NumRows: OleVariant; 
                            var NumColumns: OleVariant; var InitialColumnWidth: OleVariant; 
                            var Format: OleVariant; var ApplyBorders: OleVariant; 
                            var ApplyShading: OleVariant; var ApplyFont: OleVariant; 
                            var ApplyColor: OleVariant; var ApplyHeadingRows: OleVariant; 
                            var ApplyLastRow: OleVariant; var ApplyFirstColumn: OleVariant; 
                            var ApplyLastColumn: OleVariant; var AutoFit: OleVariant; 
                            var AutoFitBehavior: OleVariant; var DefaultTableBehavior: OleVariant): Table; dispid 498;
    procedure InsertFile(const FileName: WideString; var Range: OleVariant; 
                         var ConfirmConversions: OleVariant; var Link: OleVariant; 
                         var Attachment: OleVariant); dispid 123;
    property FormFields: FormFields readonly dispid 65;
    property Editors: Editors readonly dispid 343;
    procedure Relocate(prop: WpsRelocate); dispid 344;
    property BoldBi: Integer dispid 132;
    property ItalicBi: Integer dispid 133;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: _Font
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020952-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  _Font = interface(_IKsoDispObj)
    ['{00020952-0000-4B30-A977-D214852036FE}']
    function Get_Duplicate: _Font; safecall;
    function Get_Bold: Integer; safecall;
    procedure Set_Bold(prop: Integer); safecall;
    function Get_Italic: Integer; safecall;
    procedure Set_Italic(prop: Integer); safecall;
    function Get_Hidden: Integer; safecall;
    procedure Set_Hidden(prop: Integer); safecall;
    function Get_SmallCaps: Integer; safecall;
    procedure Set_SmallCaps(prop: Integer); safecall;
    function Get_AllCaps: Integer; safecall;
    procedure Set_AllCaps(prop: Integer); safecall;
    function Get_StrikeThrough: Integer; safecall;
    procedure Set_StrikeThrough(prop: Integer); safecall;
    function Get_DoubleStrikeThrough: Integer; safecall;
    procedure Set_DoubleStrikeThrough(prop: Integer); safecall;
    function Get_Color: WpsColor; safecall;
    procedure Set_Color(prop: WpsColor); safecall;
    function Get_Subscript: Integer; safecall;
    procedure Set_Subscript(prop: Integer); safecall;
    function Get_Superscript: Integer; safecall;
    procedure Set_Superscript(prop: Integer); safecall;
    function Get_Underline: WpsUnderline; safecall;
    procedure Set_Underline(prop: WpsUnderline); safecall;
    function Get_Size: Single; safecall;
    procedure Set_Size(prop: Single); safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    function Get_Position: Single; safecall;
    procedure Set_Position(prop: Single); safecall;
    function Get_Spacing: Single; safecall;
    procedure Set_Spacing(prop: Single); safecall;
    function Get_Scaling: Integer; safecall;
    procedure Set_Scaling(prop: Integer); safecall;
    function Get_NameFarEast: WideString; safecall;
    procedure Set_NameFarEast(const prop: WideString); safecall;
    function Get_NameAscii: WideString; safecall;
    procedure Set_NameAscii(const prop: WideString); safecall;
    function Get_NameOther: WideString; safecall;
    procedure Set_NameOther(const prop: WideString); safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Shading: Shading; safecall;
    function Get_EmphasisMark: WpsEmphasisMark; safecall;
    procedure Set_EmphasisMark(prop: WpsEmphasisMark); safecall;
    function Get_DisableCharacterSpaceGrid: WordBool; safecall;
    procedure Set_DisableCharacterSpaceGrid(prop: WordBool); safecall;
    procedure Grow; safecall;
    procedure Shrink; safecall;
    procedure Reset; safecall;
    function Get_HighlightColor: WpsColor; safecall;
    procedure Set_HighlightColor(prop: WpsColor); safecall;
    function Get_UnderlineColor: WpsColor; safecall;
    procedure Set_UnderlineColor(prop: WpsColor); safecall;
    function Get_Kerning: Single; safecall;
    procedure Set_Kerning(prop: Single); safecall;
    procedure SetAsTemplateDefault; safecall;
    function Get_Shadow: Integer; safecall;
    procedure Set_Shadow(prop: Integer); safecall;
    function Get_Outline: Integer; safecall;
    procedure Set_Outline(prop: Integer); safecall;
    function Get_Emboss: Integer; safecall;
    procedure Set_Emboss(prop: Integer); safecall;
    function Get_Engrave: Integer; safecall;
    procedure Set_Engrave(prop: Integer); safecall;
    function Get_SizeBi: Single; safecall;
    procedure Set_SizeBi(prop: Single); safecall;
    function Get_NameBi: WideString; safecall;
    procedure Set_NameBi(const prop: WideString); safecall;
    function Get_BoldBi: Integer; safecall;
    procedure Set_BoldBi(prop: Integer); safecall;
    function Get_ItalicBi: Integer; safecall;
    procedure Set_ItalicBi(prop: Integer); safecall;
    property Duplicate: _Font read Get_Duplicate;
    property Bold: Integer read Get_Bold write Set_Bold;
    property Italic: Integer read Get_Italic write Set_Italic;
    property Hidden: Integer read Get_Hidden write Set_Hidden;
    property SmallCaps: Integer read Get_SmallCaps write Set_SmallCaps;
    property AllCaps: Integer read Get_AllCaps write Set_AllCaps;
    property StrikeThrough: Integer read Get_StrikeThrough write Set_StrikeThrough;
    property DoubleStrikeThrough: Integer read Get_DoubleStrikeThrough write Set_DoubleStrikeThrough;
    property Color: WpsColor read Get_Color write Set_Color;
    property Subscript: Integer read Get_Subscript write Set_Subscript;
    property Superscript: Integer read Get_Superscript write Set_Superscript;
    property Underline: WpsUnderline read Get_Underline write Set_Underline;
    property Size: Single read Get_Size write Set_Size;
    property Name: WideString read Get_Name write Set_Name;
    property Position: Single read Get_Position write Set_Position;
    property Spacing: Single read Get_Spacing write Set_Spacing;
    property Scaling: Integer read Get_Scaling write Set_Scaling;
    property NameFarEast: WideString read Get_NameFarEast write Set_NameFarEast;
    property NameAscii: WideString read Get_NameAscii write Set_NameAscii;
    property NameOther: WideString read Get_NameOther write Set_NameOther;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Shading: Shading read Get_Shading;
    property EmphasisMark: WpsEmphasisMark read Get_EmphasisMark write Set_EmphasisMark;
    property DisableCharacterSpaceGrid: WordBool read Get_DisableCharacterSpaceGrid write Set_DisableCharacterSpaceGrid;
    property HighlightColor: WpsColor read Get_HighlightColor write Set_HighlightColor;
    property UnderlineColor: WpsColor read Get_UnderlineColor write Set_UnderlineColor;
    property Kerning: Single read Get_Kerning write Set_Kerning;
    property Shadow: Integer read Get_Shadow write Set_Shadow;
    property Outline: Integer read Get_Outline write Set_Outline;
    property Emboss: Integer read Get_Emboss write Set_Emboss;
    property Engrave: Integer read Get_Engrave write Set_Engrave;
    property SizeBi: Single read Get_SizeBi write Set_SizeBi;
    property NameBi: WideString read Get_NameBi write Set_NameBi;
    property BoldBi: Integer read Get_BoldBi write Set_BoldBi;
    property ItalicBi: Integer read Get_ItalicBi write Set_ItalicBi;
  end;

// *********************************************************************//
// DispIntf:  _FontDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020952-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  _FontDisp = dispinterface
    ['{00020952-0000-4B30-A977-D214852036FE}']
    property Duplicate: _Font readonly dispid 10;
    property Bold: Integer dispid 130;
    property Italic: Integer dispid 131;
    property Hidden: Integer dispid 132;
    property SmallCaps: Integer dispid 133;
    property AllCaps: Integer dispid 134;
    property StrikeThrough: Integer dispid 135;
    property DoubleStrikeThrough: Integer dispid 136;
    property Color: WpsColor dispid 137;
    property Subscript: Integer dispid 138;
    property Superscript: Integer dispid 139;
    property Underline: WpsUnderline dispid 140;
    property Size: Single dispid 141;
    property Name: WideString dispid 142;
    property Position: Single dispid 143;
    property Spacing: Single dispid 144;
    property Scaling: Integer dispid 145;
    property NameFarEast: WideString dispid 156;
    property NameAscii: WideString dispid 157;
    property NameOther: WideString dispid 158;
    property Borders: Borders dispid 1100;
    property Shading: Shading readonly dispid 153;
    property EmphasisMark: WpsEmphasisMark dispid 154;
    property DisableCharacterSpaceGrid: WordBool dispid 155;
    procedure Grow; dispid 100;
    procedure Shrink; dispid 101;
    procedure Reset; dispid 102;
    property HighlightColor: WpsColor dispid 301;
    property UnderlineColor: WpsColor dispid 166;
    property Kerning: Single dispid 149;
    procedure SetAsTemplateDefault; dispid 103;
    property Shadow: Integer dispid 111;
    property Outline: Integer dispid 104;
    property Emboss: Integer dispid 105;
    property Engrave: Integer dispid 106;
    property SizeBi: Single dispid 107;
    property NameBi: WideString dispid 108;
    property BoldBi: Integer dispid 109;
    property ItalicBi: Integer dispid 110;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Borders
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093C-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Borders = interface(_IKsoDispObj)
    ['{0002093C-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Enable: Integer; safecall;
    procedure Set_Enable(prop: Integer); safecall;
    function Get_DistanceFromTop: Integer; safecall;
    procedure Set_DistanceFromTop(prop: Integer); safecall;
    function Get_Shadow: WordBool; safecall;
    procedure Set_Shadow(prop: WordBool); safecall;
    function Get_InsideLineStyle: WpsLineStyle; safecall;
    procedure Set_InsideLineStyle(prop: WpsLineStyle); safecall;
    function Get_OutsideLineStyle: WpsLineStyle; safecall;
    procedure Set_OutsideLineStyle(prop: WpsLineStyle); safecall;
    function Get_InsideLineWidth: WpsLineWidth; safecall;
    procedure Set_InsideLineWidth(prop: WpsLineWidth); safecall;
    function Get_OutsideLineWidth: WpsLineWidth; safecall;
    procedure Set_OutsideLineWidth(prop: WpsLineWidth); safecall;
    function Get_InsideColor: WpsColor; safecall;
    procedure Set_InsideColor(prop: WpsColor); safecall;
    function Get_OutsideColor: WpsColor; safecall;
    procedure Set_OutsideColor(prop: WpsColor); safecall;
    function Get_DistanceFromLeft: Integer; safecall;
    procedure Set_DistanceFromLeft(prop: Integer); safecall;
    function Get_DistanceFromBottom: Integer; safecall;
    procedure Set_DistanceFromBottom(prop: Integer); safecall;
    function Get_DistanceFromRight: Integer; safecall;
    procedure Set_DistanceFromRight(prop: Integer); safecall;
    function Get_AlwaysInFront: WordBool; safecall;
    procedure Set_AlwaysInFront(prop: WordBool); safecall;
    function Get_SurroundHeader: WordBool; safecall;
    procedure Set_SurroundHeader(prop: WordBool); safecall;
    function Get_SurroundFooter: WordBool; safecall;
    procedure Set_SurroundFooter(prop: WordBool); safecall;
    function Get_JoinBorders: WordBool; safecall;
    procedure Set_JoinBorders(prop: WordBool); safecall;
    function Get_HasHorizontal: WordBool; safecall;
    function Get_HasVertical: WordBool; safecall;
    function Get_DistanceFrom: WpsBorderDistanceFrom; safecall;
    procedure Set_DistanceFrom(prop: WpsBorderDistanceFrom); safecall;
    function Item(Index: WpsBorderType): Border; safecall;
    function Get_EnableFirstPageInSection: WordBool; safecall;
    procedure Set_EnableFirstPageInSection(prop: WordBool); safecall;
    function Get_EnableOtherPagesInSection: WordBool; safecall;
    procedure Set_EnableOtherPagesInSection(prop: WordBool); safecall;
    procedure ApplyPageBordersToAllSections; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Enable: Integer read Get_Enable write Set_Enable;
    property DistanceFromTop: Integer read Get_DistanceFromTop write Set_DistanceFromTop;
    property Shadow: WordBool read Get_Shadow write Set_Shadow;
    property InsideLineStyle: WpsLineStyle read Get_InsideLineStyle write Set_InsideLineStyle;
    property OutsideLineStyle: WpsLineStyle read Get_OutsideLineStyle write Set_OutsideLineStyle;
    property InsideLineWidth: WpsLineWidth read Get_InsideLineWidth write Set_InsideLineWidth;
    property OutsideLineWidth: WpsLineWidth read Get_OutsideLineWidth write Set_OutsideLineWidth;
    property InsideColor: WpsColor read Get_InsideColor write Set_InsideColor;
    property OutsideColor: WpsColor read Get_OutsideColor write Set_OutsideColor;
    property DistanceFromLeft: Integer read Get_DistanceFromLeft write Set_DistanceFromLeft;
    property DistanceFromBottom: Integer read Get_DistanceFromBottom write Set_DistanceFromBottom;
    property DistanceFromRight: Integer read Get_DistanceFromRight write Set_DistanceFromRight;
    property AlwaysInFront: WordBool read Get_AlwaysInFront write Set_AlwaysInFront;
    property SurroundHeader: WordBool read Get_SurroundHeader write Set_SurroundHeader;
    property SurroundFooter: WordBool read Get_SurroundFooter write Set_SurroundFooter;
    property JoinBorders: WordBool read Get_JoinBorders write Set_JoinBorders;
    property HasHorizontal: WordBool read Get_HasHorizontal;
    property HasVertical: WordBool read Get_HasVertical;
    property DistanceFrom: WpsBorderDistanceFrom read Get_DistanceFrom write Set_DistanceFrom;
    property EnableFirstPageInSection: WordBool read Get_EnableFirstPageInSection write Set_EnableFirstPageInSection;
    property EnableOtherPagesInSection: WordBool read Get_EnableOtherPagesInSection write Set_EnableOtherPagesInSection;
  end;

// *********************************************************************//
// DispIntf:  BordersDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093C-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  BordersDisp = dispinterface
    ['{0002093C-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    property Enable: Integer dispid 2;
    property DistanceFromTop: Integer dispid 4;
    property Shadow: WordBool dispid 5;
    property InsideLineStyle: WpsLineStyle dispid 6;
    property OutsideLineStyle: WpsLineStyle dispid 7;
    property InsideLineWidth: WpsLineWidth dispid 8;
    property OutsideLineWidth: WpsLineWidth dispid 9;
    property InsideColor: WpsColor dispid 10;
    property OutsideColor: WpsColor dispid 11;
    property DistanceFromLeft: Integer dispid 20;
    property DistanceFromBottom: Integer dispid 21;
    property DistanceFromRight: Integer dispid 22;
    property AlwaysInFront: WordBool dispid 23;
    property SurroundHeader: WordBool dispid 24;
    property SurroundFooter: WordBool dispid 25;
    property JoinBorders: WordBool dispid 26;
    property HasHorizontal: WordBool readonly dispid 27;
    property HasVertical: WordBool readonly dispid 28;
    property DistanceFrom: WpsBorderDistanceFrom dispid 29;
    function Item(Index: WpsBorderType): Border; dispid 0;
    property EnableFirstPageInSection: WordBool dispid 34;
    property EnableOtherPagesInSection: WordBool dispid 35;
    procedure ApplyPageBordersToAllSections; dispid 36;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Border
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093B-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Border = interface(_IKsoDispObj)
    ['{0002093B-0000-4B30-A977-D214852036FE}']
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(prop: WordBool); safecall;
    function Get_Inside: WordBool; safecall;
    function Get_LineStyle: WpsLineStyle; safecall;
    procedure Set_LineStyle(prop: WpsLineStyle); safecall;
    function Get_LineWidth: WpsLineWidth; safecall;
    procedure Set_LineWidth(prop: WpsLineWidth); safecall;
    function Get_Color: WpsColor; safecall;
    procedure Set_Color(prop: WpsColor); safecall;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property Inside: WordBool read Get_Inside;
    property LineStyle: WpsLineStyle read Get_LineStyle write Set_LineStyle;
    property LineWidth: WpsLineWidth read Get_LineWidth write Set_LineWidth;
    property Color: WpsColor read Get_Color write Set_Color;
  end;

// *********************************************************************//
// DispIntf:  BorderDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093B-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  BorderDisp = dispinterface
    ['{0002093B-0000-4B30-A977-D214852036FE}']
    property Visible: WordBool dispid 0;
    property Inside: WordBool readonly dispid 2;
    property LineStyle: WpsLineStyle dispid 3;
    property LineWidth: WpsLineWidth dispid 4;
    property Color: WpsColor dispid 7;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Shading
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Shading = interface(_IKsoDispObj)
    ['{0002093A-0000-4B30-A977-D214852036FE}']
    function Get_Texture: WpsTextureIndex; safecall;
    procedure Set_Texture(prop: WpsTextureIndex); safecall;
    function Get_ForegroundPatternColor: WpsColor; safecall;
    procedure Set_ForegroundPatternColor(prop: WpsColor); safecall;
    function Get_BackgroundPatternColor: WpsColor; safecall;
    procedure Set_BackgroundPatternColor(prop: WpsColor); safecall;
    property Texture: WpsTextureIndex read Get_Texture write Set_Texture;
    property ForegroundPatternColor: WpsColor read Get_ForegroundPatternColor write Set_ForegroundPatternColor;
    property BackgroundPatternColor: WpsColor read Get_BackgroundPatternColor write Set_BackgroundPatternColor;
  end;

// *********************************************************************//
// DispIntf:  ShadingDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ShadingDisp = dispinterface
    ['{0002093A-0000-4B30-A977-D214852036FE}']
    property Texture: WpsTextureIndex dispid 3;
    property ForegroundPatternColor: WpsColor dispid 4;
    property BackgroundPatternColor: WpsColor dispid 5;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Tables
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094D-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Tables = interface(_IKsoDispObj)
    ['{0002094D-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): Table; safecall;
    function AddOld(const Range: Range; NumRows: Integer; NumColumns: Integer): Table; safecall;
    function Add(const Range: Range; NumRows: Integer; NumColumns: Integer; 
                 var DefaultTableBehavior: OleVariant; var AutoFitBehavior: OleVariant): Table; safecall;
    function Get_NestingLevel: Integer; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property NestingLevel: Integer read Get_NestingLevel;
  end;

// *********************************************************************//
// DispIntf:  TablesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094D-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TablesDisp = dispinterface
    ['{0002094D-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    function Item(Index: Integer): Table; dispid 0;
    function AddOld(const Range: Range; NumRows: Integer; NumColumns: Integer): Table; dispid 4;
    function Add(const Range: Range; NumRows: Integer; NumColumns: Integer; 
                 var DefaultTableBehavior: OleVariant; var AutoFitBehavior: OleVariant): Table; dispid 200;
    property NestingLevel: Integer readonly dispid 100;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Table
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020951-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Table = interface(_IKsoDispObj)
    ['{00020951-0000-4B30-A977-D214852036FE}']
    function Get_Range: Range; safecall;
    function Get_Columns: Columns; safecall;
    function Get_Rows: Rows; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Shading: Shading; safecall;
    function Get_Uniform: WordBool; safecall;
    procedure Select; safecall;
    procedure Delete; safecall;
    function Cell(Row: Integer; Column: Integer): Cell; safecall;
    function Split(var BeforeRow: OleVariant): Table; safecall;
    function Get_Tables: Tables; safecall;
    function Get_NestingLevel: Integer; safecall;
    function Get_AllowPageBreaks: WordBool; safecall;
    procedure Set_AllowPageBreaks(prop: WordBool); safecall;
    function Get_AllowAutoFit: WordBool; safecall;
    procedure Set_AllowAutoFit(prop: WordBool); safecall;
    function Get_PreferredWidth: Single; safecall;
    procedure Set_PreferredWidth(prop: Single); safecall;
    function Get_PreferredWidthType: WpsPreferredWidthType; safecall;
    procedure Set_PreferredWidthType(prop: WpsPreferredWidthType); safecall;
    function Get_TopPadding: Single; safecall;
    procedure Set_TopPadding(prop: Single); safecall;
    function Get_BottomPadding: Single; safecall;
    procedure Set_BottomPadding(prop: Single); safecall;
    function Get_LeftPadding: Single; safecall;
    procedure Set_LeftPadding(prop: Single); safecall;
    function Get_RightPadding: Single; safecall;
    procedure Set_RightPadding(prop: Single); safecall;
    function Get_Spacing: Single; safecall;
    procedure Set_Spacing(prop: Single); safecall;
    function Get_TableDirection: WpsTableDirection; safecall;
    procedure Set_TableDirection(prop: WpsTableDirection); safecall;
    function Get_ID: WideString; safecall;
    procedure Set_ID(const prop: WideString); safecall;
    procedure AutoFitBehavior(Behavior: WpsAutoFitBehavior); safecall;
    function ConvertToText(var Separator: OleVariant; var NestedTables: OleVariant): Range; safecall;
    property Range: Range read Get_Range;
    property Columns: Columns read Get_Columns;
    property Rows: Rows read Get_Rows;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Shading: Shading read Get_Shading;
    property Uniform: WordBool read Get_Uniform;
    property Tables: Tables read Get_Tables;
    property NestingLevel: Integer read Get_NestingLevel;
    property AllowPageBreaks: WordBool read Get_AllowPageBreaks write Set_AllowPageBreaks;
    property AllowAutoFit: WordBool read Get_AllowAutoFit write Set_AllowAutoFit;
    property PreferredWidth: Single read Get_PreferredWidth write Set_PreferredWidth;
    property PreferredWidthType: WpsPreferredWidthType read Get_PreferredWidthType write Set_PreferredWidthType;
    property TopPadding: Single read Get_TopPadding write Set_TopPadding;
    property BottomPadding: Single read Get_BottomPadding write Set_BottomPadding;
    property LeftPadding: Single read Get_LeftPadding write Set_LeftPadding;
    property RightPadding: Single read Get_RightPadding write Set_RightPadding;
    property Spacing: Single read Get_Spacing write Set_Spacing;
    property TableDirection: WpsTableDirection read Get_TableDirection write Set_TableDirection;
    property ID: WideString read Get_ID write Set_ID;
  end;

// *********************************************************************//
// DispIntf:  TableDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020951-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TableDisp = dispinterface
    ['{00020951-0000-4B30-A977-D214852036FE}']
    property Range: Range readonly dispid 0;
    property Columns: Columns readonly dispid 100;
    property Rows: Rows readonly dispid 101;
    property Borders: Borders dispid 1100;
    property Shading: Shading readonly dispid 104;
    property Uniform: WordBool readonly dispid 105;
    procedure Select; dispid 200;
    procedure Delete; dispid 9;
    function Cell(Row: Integer; Column: Integer): Cell; dispid 17;
    function Split(var BeforeRow: OleVariant): Table; dispid 18;
    property Tables: Tables readonly dispid 107;
    property NestingLevel: Integer readonly dispid 108;
    property AllowPageBreaks: WordBool dispid 109;
    property AllowAutoFit: WordBool dispid 110;
    property PreferredWidth: Single dispid 111;
    property PreferredWidthType: WpsPreferredWidthType dispid 112;
    property TopPadding: Single dispid 113;
    property BottomPadding: Single dispid 114;
    property LeftPadding: Single dispid 115;
    property RightPadding: Single dispid 116;
    property Spacing: Single dispid 117;
    property TableDirection: WpsTableDirection dispid 118;
    property ID: WideString dispid 119;
    procedure AutoFitBehavior(Behavior: WpsAutoFitBehavior); dispid 120;
    function ConvertToText(var Separator: OleVariant; var NestedTables: OleVariant): Range; dispid 19;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Columns
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094B-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Columns = interface(_IKsoDispObj)
    ['{0002094B-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_First: Column; safecall;
    function Get_Last: Column; safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Item(Index: Integer): Column; safecall;
    function Add(var BeforeColumn: OleVariant): Column; safecall;
    procedure Select; safecall;
    procedure Delete; safecall;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WpsRulerStyle); safecall;
    procedure AutoFit; safecall;
    procedure DistributeWidth; safecall;
    function Get_NestingLevel: Integer; safecall;
    function Get_PreferredWidth: Single; safecall;
    procedure Set_PreferredWidth(prop: Single); safecall;
    function Get_PreferredWidthType: WpsPreferredWidthType; safecall;
    procedure Set_PreferredWidthType(prop: WpsPreferredWidthType); safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property First: Column read Get_First;
    property Last: Column read Get_Last;
    property Width: Single read Get_Width write Set_Width;
    property Borders: Borders read Get_Borders write Set_Borders;
    property NestingLevel: Integer read Get_NestingLevel;
    property PreferredWidth: Single read Get_PreferredWidth write Set_PreferredWidth;
    property PreferredWidthType: WpsPreferredWidthType read Get_PreferredWidthType write Set_PreferredWidthType;
  end;

// *********************************************************************//
// DispIntf:  ColumnsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094B-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ColumnsDisp = dispinterface
    ['{0002094B-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property First: Column readonly dispid 100;
    property Last: Column readonly dispid 101;
    property Width: Single dispid 3;
    property Borders: Borders dispid 1100;
    function Item(Index: Integer): Column; dispid 0;
    function Add(var BeforeColumn: OleVariant): Column; dispid 5;
    procedure Select; dispid 199;
    procedure Delete; dispid 200;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WpsRulerStyle); dispid 201;
    procedure AutoFit; dispid 202;
    procedure DistributeWidth; dispid 203;
    property NestingLevel: Integer readonly dispid 104;
    property PreferredWidth: Single dispid 105;
    property PreferredWidthType: WpsPreferredWidthType dispid 106;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Column
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Column = interface(_IKsoDispObj)
    ['{0002094F-0000-4B30-A977-D214852036FE}']
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_IsFirst: WordBool; safecall;
    function Get_IsLast: WordBool; safecall;
    function Get_Index: Integer; safecall;
    function Get_Cells: Cells; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Next: Column; safecall;
    function Get_Previous: Column; safecall;
    procedure Select; safecall;
    procedure Delete; safecall;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WpsRulerStyle); safecall;
    procedure AutoFit; safecall;
    function Get_NestingLevel: Integer; safecall;
    function Get_PreferredWidth: Single; safecall;
    procedure Set_PreferredWidth(prop: Single); safecall;
    function Get_PreferredWidthType: WpsPreferredWidthType; safecall;
    procedure Set_PreferredWidthType(prop: WpsPreferredWidthType); safecall;
    property Width: Single read Get_Width write Set_Width;
    property IsFirst: WordBool read Get_IsFirst;
    property IsLast: WordBool read Get_IsLast;
    property Index: Integer read Get_Index;
    property Cells: Cells read Get_Cells;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Next: Column read Get_Next;
    property Previous: Column read Get_Previous;
    property NestingLevel: Integer read Get_NestingLevel;
    property PreferredWidth: Single read Get_PreferredWidth write Set_PreferredWidth;
    property PreferredWidthType: WpsPreferredWidthType read Get_PreferredWidthType write Set_PreferredWidthType;
  end;

// *********************************************************************//
// DispIntf:  ColumnDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ColumnDisp = dispinterface
    ['{0002094F-0000-4B30-A977-D214852036FE}']
    property Width: Single dispid 3;
    property IsFirst: WordBool readonly dispid 4;
    property IsLast: WordBool readonly dispid 5;
    property Index: Integer readonly dispid 6;
    property Cells: Cells readonly dispid 100;
    property Borders: Borders dispid 1100;
    property Next: Column readonly dispid 103;
    property Previous: Column readonly dispid 104;
    procedure Select; dispid 65535;
    procedure Delete; dispid 200;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WpsRulerStyle); dispid 201;
    procedure AutoFit; dispid 202;
    property NestingLevel: Integer readonly dispid 105;
    property PreferredWidth: Single dispid 106;
    property PreferredWidthType: WpsPreferredWidthType dispid 107;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Cells
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Cells = interface(_IKsoDispObj)
    ['{0002094A-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_HeightRule: WpsRowHeightRule; safecall;
    procedure Set_HeightRule(prop: WpsRowHeightRule); safecall;
    function Get_VerticalAlignment: WpsCellVerticalAlignment; safecall;
    procedure Set_VerticalAlignment(prop: WpsCellVerticalAlignment); safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Shading: Shading; safecall;
    function Item(Index: Integer): Cell; safecall;
    function Add(var BeforeCell: OleVariant): Cell; safecall;
    procedure Delete(var ShiftCells: OleVariant); safecall;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WpsRulerStyle); safecall;
    procedure SetHeight(var RowHeight: OleVariant; HeightRule: WpsRowHeightRule); safecall;
    procedure Merge; safecall;
    procedure Split(var NumRows: OleVariant; var NumColumns: OleVariant; 
                    var MergeBeforeSplit: OleVariant); safecall;
    procedure DistributeHeight; safecall;
    procedure DistributeWidth; safecall;
    procedure AutoFit; safecall;
    function Get_NestingLevel: Integer; safecall;
    function Get_PreferredWidth: Single; safecall;
    procedure Set_PreferredWidth(prop: Single); safecall;
    function Get_PreferredWidthType: WpsPreferredWidthType; safecall;
    procedure Set_PreferredWidthType(prop: WpsPreferredWidthType); safecall;
    function Get_Orientation: WpsTextOrientation; safecall;
    procedure Set_Orientation(prop: WpsTextOrientation); safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Width: Single read Get_Width write Set_Width;
    property Height: Single read Get_Height write Set_Height;
    property HeightRule: WpsRowHeightRule read Get_HeightRule write Set_HeightRule;
    property VerticalAlignment: WpsCellVerticalAlignment read Get_VerticalAlignment write Set_VerticalAlignment;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Shading: Shading read Get_Shading;
    property NestingLevel: Integer read Get_NestingLevel;
    property PreferredWidth: Single read Get_PreferredWidth write Set_PreferredWidth;
    property PreferredWidthType: WpsPreferredWidthType read Get_PreferredWidthType write Set_PreferredWidthType;
    property Orientation: WpsTextOrientation read Get_Orientation write Set_Orientation;
  end;

// *********************************************************************//
// DispIntf:  CellsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  CellsDisp = dispinterface
    ['{0002094A-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Width: Single dispid 6;
    property Height: Single dispid 7;
    property HeightRule: WpsRowHeightRule dispid 8;
    property VerticalAlignment: WpsCellVerticalAlignment dispid 1104;
    property Borders: Borders dispid 1100;
    property Shading: Shading readonly dispid 101;
    function Item(Index: Integer): Cell; dispid 0;
    function Add(var BeforeCell: OleVariant): Cell; dispid 4;
    procedure Delete(var ShiftCells: OleVariant); dispid 200;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WpsRulerStyle); dispid 202;
    procedure SetHeight(var RowHeight: OleVariant; HeightRule: WpsRowHeightRule); dispid 203;
    procedure Merge; dispid 204;
    procedure Split(var NumRows: OleVariant; var NumColumns: OleVariant; 
                    var MergeBeforeSplit: OleVariant); dispid 205;
    procedure DistributeHeight; dispid 206;
    procedure DistributeWidth; dispid 207;
    procedure AutoFit; dispid 208;
    property NestingLevel: Integer readonly dispid 102;
    property PreferredWidth: Single dispid 103;
    property PreferredWidthType: WpsPreferredWidthType dispid 104;
    property Orientation: WpsTextOrientation dispid 105;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Cell
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094E-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Cell = interface(_IKsoDispObj)
    ['{0002094E-0000-4B30-A977-D214852036FE}']
    function Get_Range: Range; safecall;
    function Get_RowIndex: Integer; safecall;
    function Get_ColumnIndex: Integer; safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_HeightRule: WpsRowHeightRule; safecall;
    procedure Set_HeightRule(prop: WpsRowHeightRule); safecall;
    function Get_VerticalAlignment: WpsCellVerticalAlignment; safecall;
    procedure Set_VerticalAlignment(prop: WpsCellVerticalAlignment); safecall;
    function Get_Column: Column; safecall;
    function Get_Row: Row; safecall;
    function Get_Next: Cell; safecall;
    function Get_Previous: Cell; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Shading: Shading; safecall;
    procedure Select; safecall;
    procedure Delete(var ShiftCells: OleVariant); safecall;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WpsRulerStyle); safecall;
    procedure SetHeight(var RowHeight: OleVariant; HeightRule: WpsRowHeightRule); safecall;
    procedure Merge(const MergeTo: Cell); safecall;
    procedure Split(var NumRows: OleVariant; var NumColumns: OleVariant); safecall;
    procedure AutoSum; safecall;
    function Get_Tables: Tables; safecall;
    function Get_NestingLevel: Integer; safecall;
    function Get_WordWrap: WordBool; safecall;
    procedure Set_WordWrap(prop: WordBool); safecall;
    function Get_PreferredWidth: Single; safecall;
    procedure Set_PreferredWidth(prop: Single); safecall;
    function Get_FitText: WordBool; safecall;
    procedure Set_FitText(prop: WordBool); safecall;
    function Get_TopPadding: Single; safecall;
    procedure Set_TopPadding(prop: Single); safecall;
    function Get_BottomPadding: Single; safecall;
    procedure Set_BottomPadding(prop: Single); safecall;
    function Get_LeftPadding: Single; safecall;
    procedure Set_LeftPadding(prop: Single); safecall;
    function Get_RightPadding: Single; safecall;
    procedure Set_RightPadding(prop: Single); safecall;
    function Get_ID: WideString; safecall;
    procedure Set_ID(const prop: WideString); safecall;
    function Get_PreferredWidthType: WpsPreferredWidthType; safecall;
    procedure Set_PreferredWidthType(prop: WpsPreferredWidthType); safecall;
    function Get_DiagonalCell: WpsDiagonalCellType; safecall;
    procedure Set_DiagonalCell(prop: WpsDiagonalCellType); safecall;
    function Get_Orientation: WpsTextOrientation; safecall;
    procedure Set_Orientation(prop: WpsTextOrientation); safecall;
    function Get_VertMerge: SYSUINT; safecall;
    procedure Set_VertMerge(prop: SYSUINT); safecall;
    property Range: Range read Get_Range;
    property RowIndex: Integer read Get_RowIndex;
    property ColumnIndex: Integer read Get_ColumnIndex;
    property Width: Single read Get_Width write Set_Width;
    property Height: Single read Get_Height write Set_Height;
    property HeightRule: WpsRowHeightRule read Get_HeightRule write Set_HeightRule;
    property VerticalAlignment: WpsCellVerticalAlignment read Get_VerticalAlignment write Set_VerticalAlignment;
    property Column: Column read Get_Column;
    property Row: Row read Get_Row;
    property Next: Cell read Get_Next;
    property Previous: Cell read Get_Previous;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Shading: Shading read Get_Shading;
    property Tables: Tables read Get_Tables;
    property NestingLevel: Integer read Get_NestingLevel;
    property WordWrap: WordBool read Get_WordWrap write Set_WordWrap;
    property PreferredWidth: Single read Get_PreferredWidth write Set_PreferredWidth;
    property FitText: WordBool read Get_FitText write Set_FitText;
    property TopPadding: Single read Get_TopPadding write Set_TopPadding;
    property BottomPadding: Single read Get_BottomPadding write Set_BottomPadding;
    property LeftPadding: Single read Get_LeftPadding write Set_LeftPadding;
    property RightPadding: Single read Get_RightPadding write Set_RightPadding;
    property ID: WideString read Get_ID write Set_ID;
    property PreferredWidthType: WpsPreferredWidthType read Get_PreferredWidthType write Set_PreferredWidthType;
    property DiagonalCell: WpsDiagonalCellType read Get_DiagonalCell write Set_DiagonalCell;
    property Orientation: WpsTextOrientation read Get_Orientation write Set_Orientation;
    property VertMerge: SYSUINT read Get_VertMerge write Set_VertMerge;
  end;

// *********************************************************************//
// DispIntf:  CellDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094E-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  CellDisp = dispinterface
    ['{0002094E-0000-4B30-A977-D214852036FE}']
    property Range: Range readonly dispid 0;
    property RowIndex: Integer readonly dispid 4;
    property ColumnIndex: Integer readonly dispid 5;
    property Width: Single dispid 6;
    property Height: Single dispid 7;
    property HeightRule: WpsRowHeightRule dispid 8;
    property VerticalAlignment: WpsCellVerticalAlignment dispid 1104;
    property Column: Column readonly dispid 101;
    property Row: Row readonly dispid 102;
    property Next: Cell readonly dispid 103;
    property Previous: Cell readonly dispid 104;
    property Borders: Borders dispid 1100;
    property Shading: Shading readonly dispid 100;
    procedure Select; dispid 65535;
    procedure Delete(var ShiftCells: OleVariant); dispid 200;
    procedure SetWidth(ColumnWidth: Single; RulerStyle: WpsRulerStyle); dispid 202;
    procedure SetHeight(var RowHeight: OleVariant; HeightRule: WpsRowHeightRule); dispid 203;
    procedure Merge(const MergeTo: Cell); dispid 204;
    procedure Split(var NumRows: OleVariant; var NumColumns: OleVariant); dispid 205;
    procedure AutoSum; dispid 206;
    property Tables: Tables readonly dispid 106;
    property NestingLevel: Integer readonly dispid 107;
    property WordWrap: WordBool dispid 108;
    property PreferredWidth: Single dispid 109;
    property FitText: WordBool dispid 110;
    property TopPadding: Single dispid 111;
    property BottomPadding: Single dispid 112;
    property LeftPadding: Single dispid 113;
    property RightPadding: Single dispid 114;
    property ID: WideString dispid 115;
    property PreferredWidthType: WpsPreferredWidthType dispid 116;
    property DiagonalCell: WpsDiagonalCellType dispid 117;
    property Orientation: WpsTextOrientation dispid 105;
    property VertMerge: SYSUINT dispid 118;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Row
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020950-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Row = interface(_IKsoDispObj)
    ['{00020950-0000-4B30-A977-D214852036FE}']
    function Get_Range: Range; safecall;
    function Get_AllowBreakAcrossPages: Integer; safecall;
    procedure Set_AllowBreakAcrossPages(prop: Integer); safecall;
    function Get_Alignment: WpsRowAlignment; safecall;
    procedure Set_Alignment(prop: WpsRowAlignment); safecall;
    function Get_HeadingFormat: Integer; safecall;
    procedure Set_HeadingFormat(prop: Integer); safecall;
    function Get_SpaceBetweenColumns: Single; safecall;
    procedure Set_SpaceBetweenColumns(prop: Single); safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_HeightRule: WpsRowHeightRule; safecall;
    procedure Set_HeightRule(prop: WpsRowHeightRule); safecall;
    function Get_LeftIndent: Single; safecall;
    procedure Set_LeftIndent(prop: Single); safecall;
    function Get_IsLast: WordBool; safecall;
    function Get_IsFirst: WordBool; safecall;
    function Get_Index: Integer; safecall;
    function Get_Cells: Cells; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Next: Row; safecall;
    function Get_Previous: Row; safecall;
    procedure Select; safecall;
    procedure Delete; safecall;
    procedure SetLeftIndent(LeftIndent: Single; RulerStyle: WpsRulerStyle); safecall;
    procedure SetHeight(RowHeight: Single; HeightRule: WpsRowHeightRule); safecall;
    function ConvertToText(var Separator: OleVariant; var NestedTables: OleVariant): Range; safecall;
    function Get_NestingLevel: Integer; safecall;
    function Get_ID: WideString; safecall;
    procedure Set_ID(const prop: WideString); safecall;
    property Range: Range read Get_Range;
    property AllowBreakAcrossPages: Integer read Get_AllowBreakAcrossPages write Set_AllowBreakAcrossPages;
    property Alignment: WpsRowAlignment read Get_Alignment write Set_Alignment;
    property HeadingFormat: Integer read Get_HeadingFormat write Set_HeadingFormat;
    property SpaceBetweenColumns: Single read Get_SpaceBetweenColumns write Set_SpaceBetweenColumns;
    property Height: Single read Get_Height write Set_Height;
    property HeightRule: WpsRowHeightRule read Get_HeightRule write Set_HeightRule;
    property LeftIndent: Single read Get_LeftIndent write Set_LeftIndent;
    property IsLast: WordBool read Get_IsLast;
    property IsFirst: WordBool read Get_IsFirst;
    property Index: Integer read Get_Index;
    property Cells: Cells read Get_Cells;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Next: Row read Get_Next;
    property Previous: Row read Get_Previous;
    property NestingLevel: Integer read Get_NestingLevel;
    property ID: WideString read Get_ID write Set_ID;
  end;

// *********************************************************************//
// DispIntf:  RowDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020950-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  RowDisp = dispinterface
    ['{00020950-0000-4B30-A977-D214852036FE}']
    property Range: Range readonly dispid 0;
    property AllowBreakAcrossPages: Integer dispid 3;
    property Alignment: WpsRowAlignment dispid 4;
    property HeadingFormat: Integer dispid 5;
    property SpaceBetweenColumns: Single dispid 6;
    property Height: Single dispid 7;
    property HeightRule: WpsRowHeightRule dispid 8;
    property LeftIndent: Single dispid 9;
    property IsLast: WordBool readonly dispid 10;
    property IsFirst: WordBool readonly dispid 11;
    property Index: Integer readonly dispid 12;
    property Cells: Cells readonly dispid 100;
    property Borders: Borders dispid 1100;
    property Next: Row readonly dispid 104;
    property Previous: Row readonly dispid 105;
    procedure Select; dispid 65535;
    procedure Delete; dispid 200;
    procedure SetLeftIndent(LeftIndent: Single; RulerStyle: WpsRulerStyle); dispid 202;
    procedure SetHeight(RowHeight: Single; HeightRule: WpsRowHeightRule); dispid 203;
    function ConvertToText(var Separator: OleVariant; var NestedTables: OleVariant): Range; dispid 18;
    property NestingLevel: Integer readonly dispid 106;
    property ID: WideString dispid 107;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Rows
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094C-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Rows = interface(_IKsoDispObj)
    ['{0002094C-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_AllowBreakAcrossPages: Integer; safecall;
    procedure Set_AllowBreakAcrossPages(prop: Integer); safecall;
    function Get_Alignment: WpsRowAlignment; safecall;
    procedure Set_Alignment(prop: WpsRowAlignment); safecall;
    function Get_HeadingFormat: Integer; safecall;
    procedure Set_HeadingFormat(prop: Integer); safecall;
    function Get_SpaceBetweenColumns: Single; safecall;
    procedure Set_SpaceBetweenColumns(prop: Single); safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_HeightRule: WpsRowHeightRule; safecall;
    procedure Set_HeightRule(prop: WpsRowHeightRule); safecall;
    function Get_LeftIndent: Single; safecall;
    procedure Set_LeftIndent(prop: Single); safecall;
    function Get_First: Row; safecall;
    function Get_Last: Row; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Item(Index: Integer): Row; safecall;
    function Add(var BeforeRow: OleVariant): Row; safecall;
    procedure Select; safecall;
    procedure Delete; safecall;
    procedure SetLeftIndent(LeftIndent: Single; RulerStyle: WpsRulerStyle); safecall;
    procedure SetHeight(RowHeight: Single; HeightRule: WpsRowHeightRule); safecall;
    procedure DistributeHeight; safecall;
    function ConvertToText(var Separator: OleVariant; var NestedTables: OleVariant): Range; safecall;
    function Get_WrapAroundText: Integer; safecall;
    procedure Set_WrapAroundText(prop: Integer); safecall;
    function Get_DistanceTop: Single; safecall;
    procedure Set_DistanceTop(prop: Single); safecall;
    function Get_DistanceBottom: Single; safecall;
    procedure Set_DistanceBottom(prop: Single); safecall;
    function Get_DistanceLeft: Single; safecall;
    procedure Set_DistanceLeft(prop: Single); safecall;
    function Get_DistanceRight: Single; safecall;
    procedure Set_DistanceRight(prop: Single); safecall;
    function Get_HorizontalPosition: Single; safecall;
    procedure Set_HorizontalPosition(prop: Single); safecall;
    function Get_VerticalPosition: Single; safecall;
    procedure Set_VerticalPosition(prop: Single); safecall;
    function Get_RelativeHorizontalPosition: WpsRelativeHorizontalPosition; safecall;
    procedure Set_RelativeHorizontalPosition(prop: WpsRelativeHorizontalPosition); safecall;
    function Get_RelativeVerticalPosition: WpsRelativeVerticalPosition; safecall;
    procedure Set_RelativeVerticalPosition(prop: WpsRelativeVerticalPosition); safecall;
    function Get_AllowOverlap: Integer; safecall;
    procedure Set_AllowOverlap(prop: Integer); safecall;
    function Get_NestingLevel: Integer; safecall;
    function Get_TableDirection: WpsTableDirection; safecall;
    procedure Set_TableDirection(prop: WpsTableDirection); safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property AllowBreakAcrossPages: Integer read Get_AllowBreakAcrossPages write Set_AllowBreakAcrossPages;
    property Alignment: WpsRowAlignment read Get_Alignment write Set_Alignment;
    property HeadingFormat: Integer read Get_HeadingFormat write Set_HeadingFormat;
    property SpaceBetweenColumns: Single read Get_SpaceBetweenColumns write Set_SpaceBetweenColumns;
    property Height: Single read Get_Height write Set_Height;
    property HeightRule: WpsRowHeightRule read Get_HeightRule write Set_HeightRule;
    property LeftIndent: Single read Get_LeftIndent write Set_LeftIndent;
    property First: Row read Get_First;
    property Last: Row read Get_Last;
    property Borders: Borders read Get_Borders write Set_Borders;
    property WrapAroundText: Integer read Get_WrapAroundText write Set_WrapAroundText;
    property DistanceTop: Single read Get_DistanceTop write Set_DistanceTop;
    property DistanceBottom: Single read Get_DistanceBottom write Set_DistanceBottom;
    property DistanceLeft: Single read Get_DistanceLeft write Set_DistanceLeft;
    property DistanceRight: Single read Get_DistanceRight write Set_DistanceRight;
    property HorizontalPosition: Single read Get_HorizontalPosition write Set_HorizontalPosition;
    property VerticalPosition: Single read Get_VerticalPosition write Set_VerticalPosition;
    property RelativeHorizontalPosition: WpsRelativeHorizontalPosition read Get_RelativeHorizontalPosition write Set_RelativeHorizontalPosition;
    property RelativeVerticalPosition: WpsRelativeVerticalPosition read Get_RelativeVerticalPosition write Set_RelativeVerticalPosition;
    property AllowOverlap: Integer read Get_AllowOverlap write Set_AllowOverlap;
    property NestingLevel: Integer read Get_NestingLevel;
    property TableDirection: WpsTableDirection read Get_TableDirection write Set_TableDirection;
  end;

// *********************************************************************//
// DispIntf:  RowsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002094C-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  RowsDisp = dispinterface
    ['{0002094C-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property AllowBreakAcrossPages: Integer dispid 3;
    property Alignment: WpsRowAlignment dispid 4;
    property HeadingFormat: Integer dispid 5;
    property SpaceBetweenColumns: Single dispid 6;
    property Height: Single dispid 7;
    property HeightRule: WpsRowHeightRule dispid 8;
    property LeftIndent: Single dispid 9;
    property First: Row readonly dispid 10;
    property Last: Row readonly dispid 11;
    property Borders: Borders dispid 1100;
    function Item(Index: Integer): Row; dispid 0;
    function Add(var BeforeRow: OleVariant): Row; dispid 100;
    procedure Select; dispid 199;
    procedure Delete; dispid 200;
    procedure SetLeftIndent(LeftIndent: Single; RulerStyle: WpsRulerStyle); dispid 202;
    procedure SetHeight(RowHeight: Single; HeightRule: WpsRowHeightRule); dispid 203;
    procedure DistributeHeight; dispid 206;
    function ConvertToText(var Separator: OleVariant; var NestedTables: OleVariant): Range; dispid 210;
    property WrapAroundText: Integer dispid 12;
    property DistanceTop: Single dispid 13;
    property DistanceBottom: Single dispid 14;
    property DistanceLeft: Single dispid 20;
    property DistanceRight: Single dispid 21;
    property HorizontalPosition: Single dispid 15;
    property VerticalPosition: Single dispid 17;
    property RelativeHorizontalPosition: WpsRelativeHorizontalPosition dispid 18;
    property RelativeVerticalPosition: WpsRelativeVerticalPosition dispid 19;
    property AllowOverlap: Integer dispid 22;
    property NestingLevel: Integer readonly dispid 103;
    property TableDirection: WpsTableDirection dispid 104;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Footnotes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020942-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Footnotes = interface(_IKsoDispObj)
    ['{00020942-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): Footnote; safecall;
    function Add(const Range: Range; const Reference: WideString; const Text: WideString): Footnote; safecall;
    procedure Convert; safecall;
    procedure SwapWithEndnotes; safecall;
    function Get_Location: WpsFootnoteLocation; safecall;
    procedure Set_Location(prop: WpsFootnoteLocation); safecall;
    function Get_NumberStyle: WpsNoteNumberStyle; safecall;
    procedure Set_NumberStyle(prop: WpsNoteNumberStyle); safecall;
    function Get_StartingNumber: Integer; safecall;
    procedure Set_StartingNumber(prop: Integer); safecall;
    function Get_NumberingRule: WpsNumberingRule; safecall;
    procedure Set_NumberingRule(prop: WpsNumberingRule); safecall;
    function Get_FootnoteOptions: HResult; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Location: WpsFootnoteLocation read Get_Location write Set_Location;
    property NumberStyle: WpsNoteNumberStyle read Get_NumberStyle write Set_NumberStyle;
    property StartingNumber: Integer read Get_StartingNumber write Set_StartingNumber;
    property NumberingRule: WpsNumberingRule read Get_NumberingRule write Set_NumberingRule;
    property FootnoteOptions: HResult read Get_FootnoteOptions;
  end;

// *********************************************************************//
// DispIntf:  FootnotesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020942-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FootnotesDisp = dispinterface
    ['{00020942-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    function Item(Index: Integer): Footnote; dispid 0;
    function Add(const Range: Range; const Reference: WideString; const Text: WideString): Footnote; dispid 4;
    procedure Convert; dispid 5;
    procedure SwapWithEndnotes; dispid 6;
    property Location: WpsFootnoteLocation dispid 100;
    property NumberStyle: WpsNoteNumberStyle dispid 101;
    property StartingNumber: Integer dispid 102;
    property NumberingRule: WpsNumberingRule dispid 103;
    property FootnoteOptions: HResult readonly dispid 104;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Footnote
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Footnote = interface(_IKsoDispObj)
    ['{0002093F-0000-4B30-A977-D214852036FE}']
    function Get_Range: Range; safecall;
    function Get_Reference: Range; safecall;
    function Get_Index: Integer; safecall;
    procedure Delete; safecall;
    property Range: Range read Get_Range;
    property Reference: Range read Get_Reference;
    property Index: Integer read Get_Index;
  end;

// *********************************************************************//
// DispIntf:  FootnoteDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FootnoteDisp = dispinterface
    ['{0002093F-0000-4B30-A977-D214852036FE}']
    property Range: Range readonly dispid 4;
    property Reference: Range readonly dispid 5;
    property Index: Integer readonly dispid 6;
    procedure Delete; dispid 10;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: FootnoteOptions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00025A24-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FootnoteOptions = interface(_IKsoDispObj)
    ['{00025A24-0000-4B30-A977-D214852036FE}']
    function Get_Location: WpsFootnoteLocation; safecall;
    procedure Set_Location(prop: WpsFootnoteLocation); safecall;
    function Get_NumberStyle: WpsNoteNumberStyle; safecall;
    procedure Set_NumberStyle(prop: WpsNoteNumberStyle); safecall;
    function Get_StartingNumber: Integer; safecall;
    procedure Set_StartingNumber(prop: Integer); safecall;
    function Get_NumberingRule: WpsNumberingRule; safecall;
    procedure Set_NumberingRule(prop: WpsNumberingRule); safecall;
    property Location: WpsFootnoteLocation read Get_Location write Set_Location;
    property NumberStyle: WpsNoteNumberStyle read Get_NumberStyle write Set_NumberStyle;
    property StartingNumber: Integer read Get_StartingNumber write Set_StartingNumber;
    property NumberingRule: WpsNumberingRule read Get_NumberingRule write Set_NumberingRule;
  end;

// *********************************************************************//
// DispIntf:  FootnoteOptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00025A24-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FootnoteOptionsDisp = dispinterface
    ['{00025A24-0000-4B30-A977-D214852036FE}']
    property Location: WpsFootnoteLocation dispid 100;
    property NumberStyle: WpsNoteNumberStyle dispid 101;
    property StartingNumber: Integer dispid 102;
    property NumberingRule: WpsNumberingRule dispid 103;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Endnotes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020941-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Endnotes = interface(_IKsoDispObj)
    ['{00020941-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Location: WpsEndnoteLocation; safecall;
    procedure Set_Location(prop: WpsEndnoteLocation); safecall;
    function Get_NumberStyle: WpsNoteNumberStyle; safecall;
    procedure Set_NumberStyle(prop: WpsNoteNumberStyle); safecall;
    function Get_StartingNumber: Integer; safecall;
    procedure Set_StartingNumber(prop: Integer); safecall;
    function Get_NumberingRule: WpsNumberingRule; safecall;
    procedure Set_NumberingRule(prop: WpsNumberingRule); safecall;
    function Item(Index: Integer): Endnote; safecall;
    function Add(const Range: Range; const Reference: WideString; const Text: WideString): Endnote; safecall;
    procedure Convert; safecall;
    procedure SwapWithFootnotes; safecall;
    function Get_EndnoteOptions: EndnoteOptions; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Location: WpsEndnoteLocation read Get_Location write Set_Location;
    property NumberStyle: WpsNoteNumberStyle read Get_NumberStyle write Set_NumberStyle;
    property StartingNumber: Integer read Get_StartingNumber write Set_StartingNumber;
    property NumberingRule: WpsNumberingRule read Get_NumberingRule write Set_NumberingRule;
    property EndnoteOptions: EndnoteOptions read Get_EndnoteOptions;
  end;

// *********************************************************************//
// DispIntf:  EndnotesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020941-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  EndnotesDisp = dispinterface
    ['{00020941-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property Location: WpsEndnoteLocation dispid 100;
    property NumberStyle: WpsNoteNumberStyle dispid 101;
    property StartingNumber: Integer dispid 102;
    property NumberingRule: WpsNumberingRule dispid 103;
    function Item(Index: Integer): Endnote; dispid 0;
    function Add(const Range: Range; const Reference: WideString; const Text: WideString): Endnote; dispid 4;
    procedure Convert; dispid 5;
    procedure SwapWithFootnotes; dispid 6;
    property EndnoteOptions: EndnoteOptions readonly dispid 104;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Endnote
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093E-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Endnote = interface(_IKsoDispObj)
    ['{0002093E-0000-4B30-A977-D214852036FE}']
    function Get_Range: Range; safecall;
    function Get_Reference: Range; safecall;
    function Get_Index: Integer; safecall;
    procedure Delete; safecall;
    property Range: Range read Get_Range;
    property Reference: Range read Get_Reference;
    property Index: Integer read Get_Index;
  end;

// *********************************************************************//
// DispIntf:  EndnoteDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093E-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  EndnoteDisp = dispinterface
    ['{0002093E-0000-4B30-A977-D214852036FE}']
    property Range: Range readonly dispid 4;
    property Reference: Range readonly dispid 5;
    property Index: Integer readonly dispid 6;
    procedure Delete; dispid 10;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: EndnoteOptions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00023168-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  EndnoteOptions = interface(_IKsoDispObj)
    ['{00023168-0000-4B30-A977-D214852036FE}']
    function Get_Location: WpsEndnoteLocation; safecall;
    procedure Set_Location(prop: WpsEndnoteLocation); safecall;
    function Get_NumberStyle: WpsNoteNumberStyle; safecall;
    procedure Set_NumberStyle(prop: WpsNoteNumberStyle); safecall;
    function Get_StartingNumber: Integer; safecall;
    procedure Set_StartingNumber(prop: Integer); safecall;
    function Get_NumberingRule: WpsNumberingRule; safecall;
    procedure Set_NumberingRule(prop: WpsNumberingRule); safecall;
    property Location: WpsEndnoteLocation read Get_Location write Set_Location;
    property NumberStyle: WpsNoteNumberStyle read Get_NumberStyle write Set_NumberStyle;
    property StartingNumber: Integer read Get_StartingNumber write Set_StartingNumber;
    property NumberingRule: WpsNumberingRule read Get_NumberingRule write Set_NumberingRule;
  end;

// *********************************************************************//
// DispIntf:  EndnoteOptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00023168-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  EndnoteOptionsDisp = dispinterface
    ['{00023168-0000-4B30-A977-D214852036FE}']
    property Location: WpsEndnoteLocation dispid 100;
    property NumberStyle: WpsNoteNumberStyle dispid 101;
    property StartingNumber: Integer dispid 102;
    property NumberingRule: WpsNumberingRule dispid 103;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Comments
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020940-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Comments = interface(_IKsoDispObj)
    ['{00020940-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_ShowBy: WideString; safecall;
    procedure Set_ShowBy(const prop: WideString); safecall;
    function Item(Index: Integer): Comment; safecall;
    function Add(const Range: Range; var Text: OleVariant): Comment; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property ShowBy: WideString read Get_ShowBy write Set_ShowBy;
  end;

// *********************************************************************//
// DispIntf:  CommentsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020940-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  CommentsDisp = dispinterface
    ['{00020940-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property ShowBy: WideString dispid 1003;
    function Item(Index: Integer): Comment; dispid 0;
    function Add(const Range: Range; var Text: OleVariant): Comment; dispid 4;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Comment
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093D-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Comment = interface(_IKsoDispObj)
    ['{0002093D-0000-4B30-A977-D214852036FE}']
    function Get_Range: Range; safecall;
    function Get_Reference: Range; safecall;
    function Get_Scope: Range; safecall;
    function Get_Index: Integer; safecall;
    function Get_Author: WideString; safecall;
    procedure Set_Author(const prop: WideString); safecall;
    function Get_Initial: WideString; safecall;
    procedure Set_Initial(const prop: WideString); safecall;
    function Get_ShowTip: WordBool; safecall;
    procedure Set_ShowTip(prop: WordBool); safecall;
    procedure Delete; safecall;
    procedure Edit; safecall;
    function Get_Date: TDateTime; safecall;
    function Next: Comment; safecall;
    function Get_Color: WpsColor; safecall;
    property Range: Range read Get_Range;
    property Reference: Range read Get_Reference;
    property Scope: Range read Get_Scope;
    property Index: Integer read Get_Index;
    property Author: WideString read Get_Author write Set_Author;
    property Initial: WideString read Get_Initial write Set_Initial;
    property ShowTip: WordBool read Get_ShowTip write Set_ShowTip;
    property Date: TDateTime read Get_Date;
    property Color: WpsColor read Get_Color;
  end;

// *********************************************************************//
// DispIntf:  CommentDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002093D-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  CommentDisp = dispinterface
    ['{0002093D-0000-4B30-A977-D214852036FE}']
    property Range: Range readonly dispid 1003;
    property Reference: Range readonly dispid 1004;
    property Scope: Range readonly dispid 1005;
    property Index: Integer readonly dispid 1006;
    property Author: WideString dispid 1007;
    property Initial: WideString dispid 1008;
    property ShowTip: WordBool dispid 1009;
    procedure Delete; dispid 10;
    procedure Edit; dispid 1011;
    property Date: TDateTime readonly dispid 1010;
    function Next: Comment; dispid 4097;
    property Color: WpsColor readonly dispid 4098;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Sections
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Sections = interface(_IKsoDispObj)
    ['{0002095A-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_First: Section; safecall;
    function Get_Last: Section; safecall;
    function Get_PageSetup: PageSetup; safecall;
    procedure Set_PageSetup(const prop: PageSetup); safecall;
    function Item(Index: Integer): Section; safecall;
    function Add(const Range: Range; var Start: WpsSectionStart): Section; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property First: Section read Get_First;
    property Last: Section read Get_Last;
    property PageSetup: PageSetup read Get_PageSetup write Set_PageSetup;
  end;

// *********************************************************************//
// DispIntf:  SectionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  SectionsDisp = dispinterface
    ['{0002095A-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property First: Section readonly dispid 3;
    property Last: Section readonly dispid 4;
    property PageSetup: PageSetup dispid 1101;
    function Item(Index: Integer): Section; dispid 0;
    function Add(const Range: Range; var Start: WpsSectionStart): Section; dispid 5;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Section
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020959-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Section = interface(_IKsoDispObj)
    ['{00020959-0000-4B30-A977-D214852036FE}']
    function Get_Range: Range; safecall;
    function Get_PageSetup: PageSetup; safecall;
    procedure Set_PageSetup(const prop: PageSetup); safecall;
    function Get_Headers: HeadersFooters; safecall;
    function Get_Footers: HeadersFooters; safecall;
    function Get_Index: Integer; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_ProtectedForForms: WordBool; safecall;
    procedure Set_ProtectedForForms(prop: WordBool); safecall;
    property Range: Range read Get_Range;
    property PageSetup: PageSetup read Get_PageSetup write Set_PageSetup;
    property Headers: HeadersFooters read Get_Headers;
    property Footers: HeadersFooters read Get_Footers;
    property Index: Integer read Get_Index;
    property Borders: Borders read Get_Borders write Set_Borders;
    property ProtectedForForms: WordBool read Get_ProtectedForForms write Set_ProtectedForForms;
  end;

// *********************************************************************//
// DispIntf:  SectionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020959-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  SectionDisp = dispinterface
    ['{00020959-0000-4B30-A977-D214852036FE}']
    property Range: Range readonly dispid 0;
    property PageSetup: PageSetup dispid 1101;
    property Headers: HeadersFooters readonly dispid 121;
    property Footers: HeadersFooters readonly dispid 122;
    property Index: Integer readonly dispid 124;
    property Borders: Borders dispid 1100;
    property ProtectedForForms: WordBool dispid 123;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: PageSetup
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020971-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  PageSetup = interface(_IKsoDispObj)
    ['{00020971-0000-4B30-A977-D214852036FE}']
    function Get_TopMargin: Single; safecall;
    procedure Set_TopMargin(prop: Single); safecall;
    function Get_BottomMargin: Single; safecall;
    procedure Set_BottomMargin(prop: Single); safecall;
    function Get_LeftMargin: Single; safecall;
    procedure Set_LeftMargin(prop: Single); safecall;
    function Get_RightMargin: Single; safecall;
    procedure Set_RightMargin(prop: Single); safecall;
    function Get_Gutter: Single; safecall;
    procedure Set_Gutter(prop: Single); safecall;
    function Get_PageWidth: Single; safecall;
    procedure Set_PageWidth(prop: Single); safecall;
    function Get_PageHeight: Single; safecall;
    procedure Set_PageHeight(prop: Single); safecall;
    function Get_Orientation: WpsOrientation; safecall;
    procedure Set_Orientation(prop: WpsOrientation); safecall;
    function Get_FirstPageTray: WpsPaperTray; safecall;
    procedure Set_FirstPageTray(prop: WpsPaperTray); safecall;
    function Get_OtherPagesTray: WpsPaperTray; safecall;
    procedure Set_OtherPagesTray(prop: WpsPaperTray); safecall;
    function Get_HeaderDistance: Single; safecall;
    procedure Set_HeaderDistance(prop: Single); safecall;
    function Get_FooterDistance: Single; safecall;
    procedure Set_FooterDistance(prop: Single); safecall;
    function Get_SectionStart: WpsSectionStart; safecall;
    procedure Set_SectionStart(prop: WpsSectionStart); safecall;
    function Get_OddAndEvenPagesHeaderFooter: Integer; safecall;
    procedure Set_OddAndEvenPagesHeaderFooter(prop: Integer); safecall;
    function Get_DifferentFirstPageHeaderFooter: Integer; safecall;
    procedure Set_DifferentFirstPageHeaderFooter(prop: Integer); safecall;
    function Get_TextColumns: TextColumns; safecall;
    procedure Set_TextColumns(const prop: TextColumns); safecall;
    function Get_PaperSize: WpsPaperSize; safecall;
    procedure Set_PaperSize(prop: WpsPaperSize); safecall;
    function Get_GutterOnTop: WordBool; safecall;
    procedure Set_GutterOnTop(prop: WordBool); safecall;
    function Get_CharsLine: Single; safecall;
    procedure Set_CharsLine(prop: Single); safecall;
    function Get_LinesPage: Single; safecall;
    procedure Set_LinesPage(prop: Single); safecall;
    function Get_ShowGrid: WordBool; safecall;
    procedure Set_ShowGrid(prop: WordBool); safecall;
    function Get_LayoutMode: WpsLayoutMode; safecall;
    procedure Set_LayoutMode(prop: WpsLayoutMode); safecall;
    procedure TogglePortrait; safecall;
    function Get_GutterPos: WpsGutterStyle; safecall;
    procedure Set_GutterPos(prop: WpsGutterStyle); safecall;
    function Get_Duplicate: PageSetup; safecall;
    function Get_TextOrientation: WpsTextOrientation; safecall;
    procedure Set_TextOrientation(prop: WpsTextOrientation); safecall;
    function Get_MaxLinesPage: Integer; safecall;
    function Get_MaxCharsLine: Integer; safecall;
    function Get_Genko: WordBool; safecall;
    procedure Set_Genko(prop: WordBool); safecall;
    function Get_GenkoGridStyle: WpsGenkoGridStyle; safecall;
    procedure Set_GenkoGridStyle(prop: WpsGenkoGridStyle); safecall;
    function Get_GenkoGridColor: WpsColor; safecall;
    procedure Set_GenkoGridColor(prop: WpsColor); safecall;
    function Get_GenkoPageHeight: Single; safecall;
    procedure Set_GenkoPageHeight(prop: Single); safecall;
    function Get_GenkoPageWidth: Single; safecall;
    procedure Set_GenkoPageWidth(prop: Single); safecall;
    function Get_GenkoPageSize: WpsPaperSize; safecall;
    procedure Set_GenkoPageSize(prop: WpsPaperSize); safecall;
    function Get_GenkoOrientation: WpsOrientation; safecall;
    procedure Set_GenkoOrientation(prop: WpsOrientation); safecall;
    function Get_GenkoFarEastLineBreakControl: WordBool; safecall;
    procedure Set_GenkoFarEastLineBreakControl(prop: WordBool); safecall;
    function Get_GenkoHangingPunctuation: WordBool; safecall;
    procedure Set_GenkoHangingPunctuation(prop: WordBool); safecall;
    function Get_GenkoGrid: WpsGenkoGrid; safecall;
    procedure Set_GenkoGrid(prop: WpsGenkoGrid); safecall;
    function Get_HeaderFooterOrientation: WpsOrientation; safecall;
    procedure Set_HeaderFooterOrientation(prop: WpsOrientation); safecall;
    procedure SetAsTemplateDefault; safecall;
    function Get_BookFoldPrinting: WordBool; safecall;
    procedure Set_BookFoldPrinting(prop: WordBool); safecall;
    function Get_BookFoldRevPrinting: WordBool; safecall;
    procedure Set_BookFoldRevPrinting(prop: WordBool); safecall;
    function Get_BookFoldPrintingSheets: Integer; safecall;
    procedure Set_BookFoldPrintingSheets(prop: Integer); safecall;
    property TopMargin: Single read Get_TopMargin write Set_TopMargin;
    property BottomMargin: Single read Get_BottomMargin write Set_BottomMargin;
    property LeftMargin: Single read Get_LeftMargin write Set_LeftMargin;
    property RightMargin: Single read Get_RightMargin write Set_RightMargin;
    property Gutter: Single read Get_Gutter write Set_Gutter;
    property PageWidth: Single read Get_PageWidth write Set_PageWidth;
    property PageHeight: Single read Get_PageHeight write Set_PageHeight;
    property Orientation: WpsOrientation read Get_Orientation write Set_Orientation;
    property FirstPageTray: WpsPaperTray read Get_FirstPageTray write Set_FirstPageTray;
    property OtherPagesTray: WpsPaperTray read Get_OtherPagesTray write Set_OtherPagesTray;
    property HeaderDistance: Single read Get_HeaderDistance write Set_HeaderDistance;
    property FooterDistance: Single read Get_FooterDistance write Set_FooterDistance;
    property SectionStart: WpsSectionStart read Get_SectionStart write Set_SectionStart;
    property OddAndEvenPagesHeaderFooter: Integer read Get_OddAndEvenPagesHeaderFooter write Set_OddAndEvenPagesHeaderFooter;
    property DifferentFirstPageHeaderFooter: Integer read Get_DifferentFirstPageHeaderFooter write Set_DifferentFirstPageHeaderFooter;
    property TextColumns: TextColumns read Get_TextColumns write Set_TextColumns;
    property PaperSize: WpsPaperSize read Get_PaperSize write Set_PaperSize;
    property GutterOnTop: WordBool read Get_GutterOnTop write Set_GutterOnTop;
    property CharsLine: Single read Get_CharsLine write Set_CharsLine;
    property LinesPage: Single read Get_LinesPage write Set_LinesPage;
    property ShowGrid: WordBool read Get_ShowGrid write Set_ShowGrid;
    property LayoutMode: WpsLayoutMode read Get_LayoutMode write Set_LayoutMode;
    property GutterPos: WpsGutterStyle read Get_GutterPos write Set_GutterPos;
    property Duplicate: PageSetup read Get_Duplicate;
    property TextOrientation: WpsTextOrientation read Get_TextOrientation write Set_TextOrientation;
    property MaxLinesPage: Integer read Get_MaxLinesPage;
    property MaxCharsLine: Integer read Get_MaxCharsLine;
    property Genko: WordBool read Get_Genko write Set_Genko;
    property GenkoGridStyle: WpsGenkoGridStyle read Get_GenkoGridStyle write Set_GenkoGridStyle;
    property GenkoGridColor: WpsColor read Get_GenkoGridColor write Set_GenkoGridColor;
    property GenkoPageHeight: Single read Get_GenkoPageHeight write Set_GenkoPageHeight;
    property GenkoPageWidth: Single read Get_GenkoPageWidth write Set_GenkoPageWidth;
    property GenkoPageSize: WpsPaperSize read Get_GenkoPageSize write Set_GenkoPageSize;
    property GenkoOrientation: WpsOrientation read Get_GenkoOrientation write Set_GenkoOrientation;
    property GenkoFarEastLineBreakControl: WordBool read Get_GenkoFarEastLineBreakControl write Set_GenkoFarEastLineBreakControl;
    property GenkoHangingPunctuation: WordBool read Get_GenkoHangingPunctuation write Set_GenkoHangingPunctuation;
    property GenkoGrid: WpsGenkoGrid read Get_GenkoGrid write Set_GenkoGrid;
    property HeaderFooterOrientation: WpsOrientation read Get_HeaderFooterOrientation write Set_HeaderFooterOrientation;
    property BookFoldPrinting: WordBool read Get_BookFoldPrinting write Set_BookFoldPrinting;
    property BookFoldRevPrinting: WordBool read Get_BookFoldRevPrinting write Set_BookFoldRevPrinting;
    property BookFoldPrintingSheets: Integer read Get_BookFoldPrintingSheets write Set_BookFoldPrintingSheets;
  end;

// *********************************************************************//
// DispIntf:  PageSetupDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020971-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  PageSetupDisp = dispinterface
    ['{00020971-0000-4B30-A977-D214852036FE}']
    property TopMargin: Single dispid 100;
    property BottomMargin: Single dispid 101;
    property LeftMargin: Single dispid 102;
    property RightMargin: Single dispid 103;
    property Gutter: Single dispid 104;
    property PageWidth: Single dispid 105;
    property PageHeight: Single dispid 106;
    property Orientation: WpsOrientation dispid 107;
    property FirstPageTray: WpsPaperTray dispid 108;
    property OtherPagesTray: WpsPaperTray dispid 109;
    property HeaderDistance: Single dispid 112;
    property FooterDistance: Single dispid 113;
    property SectionStart: WpsSectionStart dispid 114;
    property OddAndEvenPagesHeaderFooter: Integer dispid 115;
    property DifferentFirstPageHeaderFooter: Integer dispid 116;
    property TextColumns: TextColumns dispid 119;
    property PaperSize: WpsPaperSize dispid 120;
    property GutterOnTop: WordBool dispid 122;
    property CharsLine: Single dispid 123;
    property LinesPage: Single dispid 124;
    property ShowGrid: WordBool dispid 128;
    property LayoutMode: WpsLayoutMode dispid 131;
    procedure TogglePortrait; dispid 201;
    property GutterPos: WpsGutterStyle dispid 1222;
    property Duplicate: PageSetup readonly dispid 4097;
    property TextOrientation: WpsTextOrientation dispid 4098;
    property MaxLinesPage: Integer readonly dispid 4099;
    property MaxCharsLine: Integer readonly dispid 4100;
    property Genko: WordBool dispid 4101;
    property GenkoGridStyle: WpsGenkoGridStyle dispid 4102;
    property GenkoGridColor: WpsColor dispid 4103;
    property GenkoPageHeight: Single dispid 4104;
    property GenkoPageWidth: Single dispid 4105;
    property GenkoPageSize: WpsPaperSize dispid 4106;
    property GenkoOrientation: WpsOrientation dispid 4107;
    property GenkoFarEastLineBreakControl: WordBool dispid 4108;
    property GenkoHangingPunctuation: WordBool dispid 4109;
    property GenkoGrid: WpsGenkoGrid dispid 4110;
    property HeaderFooterOrientation: WpsOrientation dispid 4111;
    procedure SetAsTemplateDefault; dispid 202;
    property BookFoldPrinting: WordBool dispid 1223;
    property BookFoldRevPrinting: WordBool dispid 1224;
    property BookFoldPrintingSheets: Integer dispid 1225;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: TextColumns
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020973-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TextColumns = interface(_IKsoDispObj)
    ['{00020973-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_EvenlySpaced: Integer; safecall;
    procedure Set_EvenlySpaced(prop: Integer); safecall;
    function Get_LineBetween: Integer; safecall;
    procedure Set_LineBetween(prop: Integer); safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_Spacing: Single; safecall;
    procedure Set_Spacing(prop: Single); safecall;
    function Item(Index: Integer): TextColumn; safecall;
    function Add(var Width: OleVariant; var Spacing: OleVariant; EvenlySpaced: WordBool): TextColumn; safecall;
    procedure SetCount(NumColumns: Integer); safecall;
    function Get_FlowDirection: WpsFlowDirection; safecall;
    procedure Set_FlowDirection(prop: WpsFlowDirection); safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property EvenlySpaced: Integer read Get_EvenlySpaced write Set_EvenlySpaced;
    property LineBetween: Integer read Get_LineBetween write Set_LineBetween;
    property Width: Single read Get_Width write Set_Width;
    property Spacing: Single read Get_Spacing write Set_Spacing;
    property FlowDirection: WpsFlowDirection read Get_FlowDirection write Set_FlowDirection;
  end;

// *********************************************************************//
// DispIntf:  TextColumnsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020973-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TextColumnsDisp = dispinterface
    ['{00020973-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property EvenlySpaced: Integer dispid 100;
    property LineBetween: Integer dispid 101;
    property Width: Single dispid 102;
    property Spacing: Single dispid 103;
    function Item(Index: Integer): TextColumn; dispid 0;
    function Add(var Width: OleVariant; var Spacing: OleVariant; EvenlySpaced: WordBool): TextColumn; dispid 201;
    procedure SetCount(NumColumns: Integer); dispid 202;
    property FlowDirection: WpsFlowDirection dispid 104;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: TextColumn
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020974-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TextColumn = interface(_IKsoDispObj)
    ['{00020974-0000-4B30-A977-D214852036FE}']
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_SpaceAfter: Single; safecall;
    procedure Set_SpaceAfter(prop: Single); safecall;
    property Width: Single read Get_Width write Set_Width;
    property SpaceAfter: Single read Get_SpaceAfter write Set_SpaceAfter;
  end;

// *********************************************************************//
// DispIntf:  TextColumnDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020974-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TextColumnDisp = dispinterface
    ['{00020974-0000-4B30-A977-D214852036FE}']
    property Width: Single dispid 100;
    property SpaceAfter: Single dispid 101;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: HeadersFooters
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020984-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  HeadersFooters = interface(_IKsoDispObj)
    ['{00020984-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: WpsHeaderFooterIndex): HeaderFooter; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  HeadersFootersDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020984-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  HeadersFootersDisp = dispinterface
    ['{00020984-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(Index: WpsHeaderFooterIndex): HeaderFooter; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: HeaderFooter
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020985-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  HeaderFooter = interface(_IKsoDispObj)
    ['{00020985-0000-4B30-A977-D214852036FE}']
    function Get_Range: Range; safecall;
    function Get_Index: WpsHeaderFooterIndex; safecall;
    function Get_IsHeader: WordBool; safecall;
    function Get_Exists: WordBool; safecall;
    procedure Set_Exists(prop: WordBool); safecall;
    function Get_PageNumbers: PageNumbers; safecall;
    function Get_LinkToPrevious: WordBool; safecall;
    procedure Set_LinkToPrevious(prop: WordBool); safecall;
    function Get_Shapes: Shapes; safecall;
    function Get_TextOrientation: WpsTextOrientation; safecall;
    procedure Set_TextOrientation(prop: WpsTextOrientation); safecall;
    property Range: Range read Get_Range;
    property Index: WpsHeaderFooterIndex read Get_Index;
    property IsHeader: WordBool read Get_IsHeader;
    property Exists: WordBool read Get_Exists write Set_Exists;
    property PageNumbers: PageNumbers read Get_PageNumbers;
    property LinkToPrevious: WordBool read Get_LinkToPrevious write Set_LinkToPrevious;
    property Shapes: Shapes read Get_Shapes;
    property TextOrientation: WpsTextOrientation read Get_TextOrientation write Set_TextOrientation;
  end;

// *********************************************************************//
// DispIntf:  HeaderFooterDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020985-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  HeaderFooterDisp = dispinterface
    ['{00020985-0000-4B30-A977-D214852036FE}']
    property Range: Range readonly dispid 0;
    property Index: WpsHeaderFooterIndex readonly dispid 2;
    property IsHeader: WordBool readonly dispid 3;
    property Exists: WordBool dispid 4;
    property PageNumbers: PageNumbers readonly dispid 5;
    property LinkToPrevious: WordBool dispid 6;
    property Shapes: Shapes readonly dispid 7;
    property TextOrientation: WpsTextOrientation dispid 8;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: PageNumbers
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020986-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  PageNumbers = interface(_IKsoDispObj)
    ['{00020986-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_NumberStyle: WpsPageNumberStyle; safecall;
    procedure Set_NumberStyle(prop: WpsPageNumberStyle); safecall;
    function Get_IncludeChapterNumber: WordBool; safecall;
    procedure Set_IncludeChapterNumber(prop: WordBool); safecall;
    function Get_HeadingLevelForChapter: Integer; safecall;
    procedure Set_HeadingLevelForChapter(prop: Integer); safecall;
    function Get_ChapterPageSeparator: WpsSeparatorType; safecall;
    procedure Set_ChapterPageSeparator(prop: WpsSeparatorType); safecall;
    function Get_RestartNumberingAtSection: WordBool; safecall;
    procedure Set_RestartNumberingAtSection(prop: WordBool); safecall;
    function Get_StartingNumber: Integer; safecall;
    procedure Set_StartingNumber(prop: Integer); safecall;
    function Get_ShowFirstPageNumber: WordBool; safecall;
    procedure Set_ShowFirstPageNumber(prop: WordBool); safecall;
    function Item(Index: Integer): PageNumber; safecall;
    function Add(var PageNumberAlignment: OleVariant; var FirstPage: OleVariant): PageNumber; safecall;
    function Get_DoubleQuote: WordBool; safecall;
    procedure Set_DoubleQuote(prop: WordBool); safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property NumberStyle: WpsPageNumberStyle read Get_NumberStyle write Set_NumberStyle;
    property IncludeChapterNumber: WordBool read Get_IncludeChapterNumber write Set_IncludeChapterNumber;
    property HeadingLevelForChapter: Integer read Get_HeadingLevelForChapter write Set_HeadingLevelForChapter;
    property ChapterPageSeparator: WpsSeparatorType read Get_ChapterPageSeparator write Set_ChapterPageSeparator;
    property RestartNumberingAtSection: WordBool read Get_RestartNumberingAtSection write Set_RestartNumberingAtSection;
    property StartingNumber: Integer read Get_StartingNumber write Set_StartingNumber;
    property ShowFirstPageNumber: WordBool read Get_ShowFirstPageNumber write Set_ShowFirstPageNumber;
    property DoubleQuote: WordBool read Get_DoubleQuote write Set_DoubleQuote;
  end;

// *********************************************************************//
// DispIntf:  PageNumbersDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020986-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  PageNumbersDisp = dispinterface
    ['{00020986-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    property NumberStyle: WpsPageNumberStyle dispid 2;
    property IncludeChapterNumber: WordBool dispid 3;
    property HeadingLevelForChapter: Integer dispid 4;
    property ChapterPageSeparator: WpsSeparatorType dispid 5;
    property RestartNumberingAtSection: WordBool dispid 6;
    property StartingNumber: Integer dispid 7;
    property ShowFirstPageNumber: WordBool dispid 8;
    function Item(Index: Integer): PageNumber; dispid 0;
    function Add(var PageNumberAlignment: OleVariant; var FirstPage: OleVariant): PageNumber; dispid 101;
    property DoubleQuote: WordBool dispid 10;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: PageNumber
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020987-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  PageNumber = interface(_IKsoDispObj)
    ['{00020987-0000-4B30-A977-D214852036FE}']
    function Get_Index: Integer; safecall;
    function Get_Alignment: WpsPageNumberAlignment; safecall;
    procedure Set_Alignment(prop: WpsPageNumberAlignment); safecall;
    procedure Select; safecall;
    procedure Copy; safecall;
    procedure Cut; safecall;
    procedure Delete; safecall;
    property Index: Integer read Get_Index;
    property Alignment: WpsPageNumberAlignment read Get_Alignment write Set_Alignment;
  end;

// *********************************************************************//
// DispIntf:  PageNumberDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020987-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  PageNumberDisp = dispinterface
    ['{00020987-0000-4B30-A977-D214852036FE}']
    property Index: Integer readonly dispid 1;
    property Alignment: WpsPageNumberAlignment dispid 3;
    procedure Select; dispid 65535;
    procedure Copy; dispid 101;
    procedure Cut; dispid 102;
    procedure Delete; dispid 103;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Shapes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Shapes = interface(KsoShapes)
    ['{0002099F-0000-4B30-A977-D214852036FE}']
    function Item(var Index: OleVariant): Shape; safecall;
    function AddCallout(Type_: KsoCalloutType; Left: Single; Top: Single; Width: Single; 
                        Height: Single; var Anchor: OleVariant): Shape; safecall;
    function AddConnector(Type_: KsoConnectorType; BeginX: Single; BeginY: Single; EndX: Single; 
                          EndY: Single): Shape; safecall;
    function AddCurve(var SafeArrayOfPoints: OleVariant; var Anchor: OleVariant): Shape; safecall;
    function AddLabel(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                      Height: Single; var Anchor: OleVariant): Shape; safecall;
    function AddLine(BeginX: Single; BeginY: Single; EndX: Single; EndY: Single; 
                     var Anchor: OleVariant): Shape; safecall;
    function AddPicture(const FileName: WideString; var LinkToFile: OleVariant; 
                        var SaveWithDocument: OleVariant; var Left: OleVariant; 
                        var Top: OleVariant; var Width: OleVariant; var Height: OleVariant; 
                        var Anchor: OleVariant): Shape; safecall;
    function AddPolyline(var SafeArrayOfPoints: OleVariant; var Anchor: OleVariant): Shape; safecall;
    function AddShape(Type_: Integer; Left: Single; Top: Single; Width: Single; Height: Single; 
                      var Anchor: OleVariant): Shape; safecall;
    function AddTextEffect(PresetTextEffect: KsoPresetTextEffect; const Text: WideString; 
                           const FontName: WideString; FontSize: Single; FontBold: KsoTriState; 
                           FontItalic: KsoTriState; Left: Single; Top: Single; 
                           var Anchor: OleVariant): Shape; safecall;
    function AddTextbox(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                        Height: Single; var Anchor: OleVariant): Shape; safecall;
    function Range(var Index: OleVariant): ShapeRange; safecall;
    function AddOLEObject(var ClassType: OleVariant; var FileName: OleVariant; 
                          var LinkToFile: OleVariant; var DisplayAsIcon: OleVariant; 
                          var IconFileName: OleVariant; var IconIndex: OleVariant; 
                          var IconLabel: OleVariant; var Left: OleVariant; var Top: OleVariant; 
                          var Width: OleVariant; var Height: OleVariant; var Anchor: OleVariant): Shape; safecall;
    function AddOLEControl(var ClassType: OleVariant; var Left: OleVariant; var Top: OleVariant; 
                           var Width: OleVariant; var Height: OleVariant; var Anchor: OleVariant): Shape; safecall;
    function AddDiagram(Type_: KsoDiagramType; Left: Single; Top: Single; Width: Single; 
                        Height: Single; var Anchor: OleVariant): Shape; safecall;
    function AddCanvas(Left: Single; Top: Single; Width: Single; Height: Single; 
                       var Anchor: OleVariant): Shape; safecall;
    function InsertImagerScan(var Left: OleVariant; var Top: OleVariant; var Width: OleVariant; 
                              var Height: OleVariant; var Anchor: OleVariant): Shape; safecall;
    function BuildFreeform(EditingType: KsoEditingType; X1: Single; Y1: Single): FreeformBuilder; safecall;
  end;

// *********************************************************************//
// DispIntf:  ShapesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ShapesDisp = dispinterface
    ['{0002099F-0000-4B30-A977-D214852036FE}']
    function Item(var Index: OleVariant): Shape; dispid 0;
    function AddCallout(Type_: KsoCalloutType; Left: Single; Top: Single; Width: Single; 
                        Height: Single; var Anchor: OleVariant): Shape; dispid 266;
    function AddConnector(Type_: KsoConnectorType; BeginX: Single; BeginY: Single; EndX: Single; 
                          EndY: Single): Shape; dispid 267;
    function AddCurve(var SafeArrayOfPoints: OleVariant; var Anchor: OleVariant): Shape; dispid 268;
    function AddLabel(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                      Height: Single; var Anchor: OleVariant): Shape; dispid 269;
    function AddLine(BeginX: Single; BeginY: Single; EndX: Single; EndY: Single; 
                     var Anchor: OleVariant): Shape; dispid 270;
    function AddPicture(const FileName: WideString; var LinkToFile: OleVariant; 
                        var SaveWithDocument: OleVariant; var Left: OleVariant; 
                        var Top: OleVariant; var Width: OleVariant; var Height: OleVariant; 
                        var Anchor: OleVariant): Shape; dispid 271;
    function AddPolyline(var SafeArrayOfPoints: OleVariant; var Anchor: OleVariant): Shape; dispid 272;
    function AddShape(Type_: Integer; Left: Single; Top: Single; Width: Single; Height: Single; 
                      var Anchor: OleVariant): Shape; dispid 273;
    function AddTextEffect(PresetTextEffect: KsoPresetTextEffect; const Text: WideString; 
                           const FontName: WideString; FontSize: Single; FontBold: KsoTriState; 
                           FontItalic: KsoTriState; Left: Single; Top: Single; 
                           var Anchor: OleVariant): Shape; dispid 274;
    function AddTextbox(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                        Height: Single; var Anchor: OleVariant): Shape; dispid 275;
    function Range(var Index: OleVariant): ShapeRange; dispid 277;
    function AddOLEObject(var ClassType: OleVariant; var FileName: OleVariant; 
                          var LinkToFile: OleVariant; var DisplayAsIcon: OleVariant; 
                          var IconFileName: OleVariant; var IconIndex: OleVariant; 
                          var IconLabel: OleVariant; var Left: OleVariant; var Top: OleVariant; 
                          var Width: OleVariant; var Height: OleVariant; var Anchor: OleVariant): Shape; dispid 280;
    function AddOLEControl(var ClassType: OleVariant; var Left: OleVariant; var Top: OleVariant; 
                           var Width: OleVariant; var Height: OleVariant; var Anchor: OleVariant): Shape; dispid 358;
    function AddDiagram(Type_: KsoDiagramType; Left: Single; Top: Single; Width: Single; 
                        Height: Single; var Anchor: OleVariant): Shape; dispid 279;
    function AddCanvas(Left: Single; Top: Single; Width: Single; Height: Single; 
                       var Anchor: OleVariant): Shape; dispid 281;
    function InsertImagerScan(var Left: OleVariant; var Top: OleVariant; var Width: OleVariant; 
                              var Height: OleVariant; var Anchor: OleVariant): Shape; dispid 282;
    function BuildFreeform(EditingType: KsoEditingType; X1: Single; Y1: Single): FreeformBuilder; dispid 276;
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
// Interface: Shape
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A0-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Shape = interface(KsoShape)
    ['{000209A0-0000-4B30-A977-D214852036FE}']
    function Duplicate: Shape; safecall;
    function Ungroup: ShapeRange; safecall;
    function Get_ParentGroup: Shape; safecall;
    function Get_RelativeHorizontalPosition: WpsRelativeHorizontalPosition; safecall;
    procedure Set_RelativeHorizontalPosition(prop: WpsRelativeHorizontalPosition); safecall;
    function Get_RelativeVerticalPosition: WpsRelativeVerticalPosition; safecall;
    procedure Set_RelativeVerticalPosition(prop: WpsRelativeVerticalPosition); safecall;
    function Get_LockAnchor: Integer; safecall;
    procedure Set_LockAnchor(prop: Integer); safecall;
    function Get_WrapFormat: WrapFormat; safecall;
    function Get_Anchor: Range; safecall;
    function ConvertToInlineShape: InlineShape; safecall;
    function ConvertToFrame: Frame; safecall;
    procedure Activate; safecall;
    function Get_LayoutInCell: Integer; safecall;
    procedure Set_LayoutInCell(prop: Integer); safecall;
    function Get_TextFrame: TextFrame; safecall;
    function Get_type_: KsoShapeType; safecall;
    function Get_Adjustments: Adjustments; safecall;
    function Get_Callout: CalloutFormat; safecall;
    function Get_ConnectorFormat: ConnectorFormat; safecall;
    function Get_Fill: FillFormat; safecall;
    function Get_GroupItems: GroupShapes; safecall;
    function Get_Line: LineFormat; safecall;
    function Get_Nodes: ShapeNodes; safecall;
    function Get_PictureFormat: PictureFormat; safecall;
    function Get_Shadow: ShadowFormat; safecall;
    function Get_TextEffect: TextEffectFormat; safecall;
    function Get_ThreeD: ThreeDFormat; safecall;
    function Get_Diagram: Diagram; safecall;
    function Get_DiagramNode: DiagramNode; safecall;
    function Get_CanvasItems: CanvasShapes; safecall;
    function Get_OLEFormat: OLEFormat; safecall;
    property ParentGroup: Shape read Get_ParentGroup;
    property RelativeHorizontalPosition: WpsRelativeHorizontalPosition read Get_RelativeHorizontalPosition write Set_RelativeHorizontalPosition;
    property RelativeVerticalPosition: WpsRelativeVerticalPosition read Get_RelativeVerticalPosition write Set_RelativeVerticalPosition;
    property LockAnchor: Integer read Get_LockAnchor write Set_LockAnchor;
    property WrapFormat: WrapFormat read Get_WrapFormat;
    property Anchor: Range read Get_Anchor;
    property LayoutInCell: Integer read Get_LayoutInCell write Set_LayoutInCell;
    property TextFrame: TextFrame read Get_TextFrame;
    property type_: KsoShapeType read Get_type_;
    property Adjustments: Adjustments read Get_Adjustments;
    property Callout: CalloutFormat read Get_Callout;
    property ConnectorFormat: ConnectorFormat read Get_ConnectorFormat;
    property Fill: FillFormat read Get_Fill;
    property GroupItems: GroupShapes read Get_GroupItems;
    property Line: LineFormat read Get_Line;
    property Nodes: ShapeNodes read Get_Nodes;
    property PictureFormat: PictureFormat read Get_PictureFormat;
    property Shadow: ShadowFormat read Get_Shadow;
    property TextEffect: TextEffectFormat read Get_TextEffect;
    property ThreeD: ThreeDFormat read Get_ThreeD;
    property Diagram: Diagram read Get_Diagram;
    property DiagramNode: DiagramNode read Get_DiagramNode;
    property CanvasItems: CanvasShapes read Get_CanvasItems;
    property OLEFormat: OLEFormat read Get_OLEFormat;
  end;

// *********************************************************************//
// DispIntf:  ShapeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A0-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ShapeDisp = dispinterface
    ['{000209A0-0000-4B30-A977-D214852036FE}']
    function Duplicate: Shape; dispid 12;
    function Ungroup: ShapeRange; dispid 23;
    property ParentGroup: Shape readonly dispid 137;
    property RelativeHorizontalPosition: WpsRelativeHorizontalPosition dispid 300;
    property RelativeVerticalPosition: WpsRelativeVerticalPosition dispid 301;
    property LockAnchor: Integer dispid 302;
    property WrapFormat: WrapFormat readonly dispid 303;
    property Anchor: Range readonly dispid 501;
    function ConvertToInlineShape: InlineShape; dispid 25;
    function ConvertToFrame: Frame; dispid 29;
    procedure Activate; dispid 50;
    property LayoutInCell: Integer dispid 145;
    property TextFrame: TextFrame readonly dispid 377;
    property type_: KsoShapeType readonly dispid 380;
    property Adjustments: Adjustments readonly dispid 100;
    property Callout: CalloutFormat readonly dispid 103;
    property ConnectorFormat: ConnectorFormat readonly dispid 106;
    property Fill: FillFormat readonly dispid 107;
    property GroupItems: GroupShapes readonly dispid 108;
    property Line: LineFormat readonly dispid 112;
    property Nodes: ShapeNodes readonly dispid 116;
    property PictureFormat: PictureFormat readonly dispid 118;
    property Shadow: ShadowFormat readonly dispid 119;
    property TextEffect: TextEffectFormat readonly dispid 120;
    property ThreeD: ThreeDFormat readonly dispid 122;
    property Diagram: Diagram readonly dispid 133;
    property DiagramNode: DiagramNode readonly dispid 135;
    property CanvasItems: CanvasShapes readonly dispid 138;
    property OLEFormat: OLEFormat readonly dispid 144;
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
// Interface: ShapeRange
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B5-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ShapeRange = interface(KsoShapeRange)
    ['{000209B5-0000-4B30-A977-D214852036FE}']
    function Item(var Index: OleVariant): Shape; safecall;
    function Duplicate: ShapeRange; safecall;
    function Group: Shape; safecall;
    function Regroup: Shape; safecall;
    function Ungroup: ShapeRange; safecall;
    function Get_ParentGroup: Shape; safecall;
    function Get_RelativeHorizontalPosition: WpsRelativeHorizontalPosition; safecall;
    procedure Set_RelativeHorizontalPosition(prop: WpsRelativeHorizontalPosition); safecall;
    function Get_RelativeVerticalPosition: WpsRelativeVerticalPosition; safecall;
    procedure Set_RelativeVerticalPosition(prop: WpsRelativeVerticalPosition); safecall;
    function Get_LockAnchor: Integer; safecall;
    procedure Set_LockAnchor(prop: Integer); safecall;
    function Get_WrapFormat: WrapFormat; safecall;
    function Get_Anchor: Range; safecall;
    function ConvertToFrame: Frame; safecall;
    function ConvertToInlineShape: InlineShape; safecall;
    procedure Activate; safecall;
    function Get_TextFrame: TextFrame; safecall;
    function Get_type_: KsoShapeType; safecall;
    function Get_Adjustments: Adjustments; safecall;
    function Get_Callout: CalloutFormat; safecall;
    function Get_ConnectorFormat: ConnectorFormat; safecall;
    function Get_Fill: FillFormat; safecall;
    function Get_GroupItems: GroupShapes; safecall;
    function Get_Line: LineFormat; safecall;
    function Get_Nodes: ShapeNodes; safecall;
    function Get_PictureFormat: PictureFormat; safecall;
    function Get_Shadow: ShadowFormat; safecall;
    function Get_TextEffect: TextEffectFormat; safecall;
    function Get_ThreeD: ThreeDFormat; safecall;
    function Get_Diagram: Diagram; safecall;
    function Get_DiagramNode: DiagramNode; safecall;
    function Get_CanvasItems: CanvasShapes; safecall;
    function Get_LayoutInCell: Integer; safecall;
    procedure Set_LayoutInCell(prop: Integer); safecall;
    property ParentGroup: Shape read Get_ParentGroup;
    property RelativeHorizontalPosition: WpsRelativeHorizontalPosition read Get_RelativeHorizontalPosition write Set_RelativeHorizontalPosition;
    property RelativeVerticalPosition: WpsRelativeVerticalPosition read Get_RelativeVerticalPosition write Set_RelativeVerticalPosition;
    property LockAnchor: Integer read Get_LockAnchor write Set_LockAnchor;
    property WrapFormat: WrapFormat read Get_WrapFormat;
    property Anchor: Range read Get_Anchor;
    property TextFrame: TextFrame read Get_TextFrame;
    property type_: KsoShapeType read Get_type_;
    property Adjustments: Adjustments read Get_Adjustments;
    property Callout: CalloutFormat read Get_Callout;
    property ConnectorFormat: ConnectorFormat read Get_ConnectorFormat;
    property Fill: FillFormat read Get_Fill;
    property GroupItems: GroupShapes read Get_GroupItems;
    property Line: LineFormat read Get_Line;
    property Nodes: ShapeNodes read Get_Nodes;
    property PictureFormat: PictureFormat read Get_PictureFormat;
    property Shadow: ShadowFormat read Get_Shadow;
    property TextEffect: TextEffectFormat read Get_TextEffect;
    property ThreeD: ThreeDFormat read Get_ThreeD;
    property Diagram: Diagram read Get_Diagram;
    property DiagramNode: DiagramNode read Get_DiagramNode;
    property CanvasItems: CanvasShapes read Get_CanvasItems;
    property LayoutInCell: Integer read Get_LayoutInCell write Set_LayoutInCell;
  end;

// *********************************************************************//
// DispIntf:  ShapeRangeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B5-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ShapeRangeDisp = dispinterface
    ['{000209B5-0000-4B30-A977-D214852036FE}']
    function Item(var Index: OleVariant): Shape; dispid 0;
    function Duplicate: ShapeRange; dispid 14;
    function Group: Shape; dispid 19;
    function Regroup: Shape; dispid 21;
    function Ungroup: ShapeRange; dispid 27;
    property ParentGroup: Shape readonly dispid 137;
    property RelativeHorizontalPosition: WpsRelativeHorizontalPosition dispid 300;
    property RelativeVerticalPosition: WpsRelativeVerticalPosition dispid 301;
    property LockAnchor: Integer dispid 302;
    property WrapFormat: WrapFormat readonly dispid 303;
    property Anchor: Range readonly dispid 304;
    function ConvertToFrame: Frame; dispid 29;
    function ConvertToInlineShape: InlineShape; dispid 30;
    procedure Activate; dispid 50;
    property TextFrame: TextFrame readonly dispid 121;
    property type_: KsoShapeType readonly dispid 380;
    property Adjustments: Adjustments readonly dispid 100;
    property Callout: CalloutFormat readonly dispid 103;
    property ConnectorFormat: ConnectorFormat readonly dispid 106;
    property Fill: FillFormat readonly dispid 107;
    property GroupItems: GroupShapes readonly dispid 108;
    property Line: LineFormat readonly dispid 112;
    property Nodes: ShapeNodes readonly dispid 116;
    property PictureFormat: PictureFormat readonly dispid 118;
    property Shadow: ShadowFormat readonly dispid 119;
    property TextEffect: TextEffectFormat readonly dispid 120;
    property ThreeD: ThreeDFormat readonly dispid 122;
    property Diagram: Diagram readonly dispid 133;
    property DiagramNode: DiagramNode readonly dispid 135;
    property CanvasItems: CanvasShapes readonly dispid 138;
    property LayoutInCell: Integer dispid 139;
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
// Interface: WrapFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C3-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  WrapFormat = interface(_IKsoDispObj)
    ['{000209C3-0000-4B30-A977-D214852036FE}']
    function Get_type_: WpsWrapType; safecall;
    procedure Set_type_(prop: WpsWrapType); safecall;
    function Get_Side: WpsWrapSideType; safecall;
    procedure Set_Side(prop: WpsWrapSideType); safecall;
    function Get_DistanceTop: Single; safecall;
    procedure Set_DistanceTop(prop: Single); safecall;
    function Get_DistanceBottom: Single; safecall;
    procedure Set_DistanceBottom(prop: Single); safecall;
    function Get_DistanceLeft: Single; safecall;
    procedure Set_DistanceLeft(prop: Single); safecall;
    function Get_DistanceRight: Single; safecall;
    procedure Set_DistanceRight(prop: Single); safecall;
    function Get_AllowOverlap: Integer; safecall;
    procedure Set_AllowOverlap(prop: Integer); safecall;
    property type_: WpsWrapType read Get_type_ write Set_type_;
    property Side: WpsWrapSideType read Get_Side write Set_Side;
    property DistanceTop: Single read Get_DistanceTop write Set_DistanceTop;
    property DistanceBottom: Single read Get_DistanceBottom write Set_DistanceBottom;
    property DistanceLeft: Single read Get_DistanceLeft write Set_DistanceLeft;
    property DistanceRight: Single read Get_DistanceRight write Set_DistanceRight;
    property AllowOverlap: Integer read Get_AllowOverlap write Set_AllowOverlap;
  end;

// *********************************************************************//
// DispIntf:  WrapFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C3-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  WrapFormatDisp = dispinterface
    ['{000209C3-0000-4B30-A977-D214852036FE}']
    property type_: WpsWrapType dispid 100;
    property Side: WpsWrapSideType dispid 101;
    property DistanceTop: Single dispid 102;
    property DistanceBottom: Single dispid 103;
    property DistanceLeft: Single dispid 104;
    property DistanceRight: Single dispid 105;
    property AllowOverlap: Integer dispid 106;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Frame
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Frame = interface(_IKsoDispObj)
    ['{0002092A-0000-4B30-A977-D214852036FE}']
    function Get_HeightRule: WpsFrameSizeRule; safecall;
    procedure Set_HeightRule(prop: WpsFrameSizeRule); safecall;
    function Get_WidthRule: WpsFrameSizeRule; safecall;
    procedure Set_WidthRule(prop: WpsFrameSizeRule); safecall;
    function Get_HorizontalDistanceFromText: Single; safecall;
    procedure Set_HorizontalDistanceFromText(prop: Single); safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_HorizontalPosition: Single; safecall;
    procedure Set_HorizontalPosition(prop: Single); safecall;
    function Get_LockAnchor: WordBool; safecall;
    procedure Set_LockAnchor(prop: WordBool); safecall;
    function Get_RelativeHorizontalPosition: WpsRelativeHorizontalPosition; safecall;
    procedure Set_RelativeHorizontalPosition(prop: WpsRelativeHorizontalPosition); safecall;
    function Get_RelativeVerticalPosition: WpsRelativeVerticalPosition; safecall;
    procedure Set_RelativeVerticalPosition(prop: WpsRelativeVerticalPosition); safecall;
    function Get_VerticalDistanceFromText: Single; safecall;
    procedure Set_VerticalDistanceFromText(prop: Single); safecall;
    function Get_VerticalPosition: Single; safecall;
    procedure Set_VerticalPosition(prop: Single); safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_TextWrap: WordBool; safecall;
    procedure Set_TextWrap(prop: WordBool); safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Range: Range; safecall;
    procedure Delete; safecall;
    procedure Select; safecall;
    procedure Copy; safecall;
    procedure Cut; safecall;
    property HeightRule: WpsFrameSizeRule read Get_HeightRule write Set_HeightRule;
    property WidthRule: WpsFrameSizeRule read Get_WidthRule write Set_WidthRule;
    property HorizontalDistanceFromText: Single read Get_HorizontalDistanceFromText write Set_HorizontalDistanceFromText;
    property Height: Single read Get_Height write Set_Height;
    property HorizontalPosition: Single read Get_HorizontalPosition write Set_HorizontalPosition;
    property LockAnchor: WordBool read Get_LockAnchor write Set_LockAnchor;
    property RelativeHorizontalPosition: WpsRelativeHorizontalPosition read Get_RelativeHorizontalPosition write Set_RelativeHorizontalPosition;
    property RelativeVerticalPosition: WpsRelativeVerticalPosition read Get_RelativeVerticalPosition write Set_RelativeVerticalPosition;
    property VerticalDistanceFromText: Single read Get_VerticalDistanceFromText write Set_VerticalDistanceFromText;
    property VerticalPosition: Single read Get_VerticalPosition write Set_VerticalPosition;
    property Width: Single read Get_Width write Set_Width;
    property TextWrap: WordBool read Get_TextWrap write Set_TextWrap;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Range: Range read Get_Range;
  end;

// *********************************************************************//
// DispIntf:  FrameDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FrameDisp = dispinterface
    ['{0002092A-0000-4B30-A977-D214852036FE}']
    property HeightRule: WpsFrameSizeRule dispid 1;
    property WidthRule: WpsFrameSizeRule dispid 2;
    property HorizontalDistanceFromText: Single dispid 3;
    property Height: Single dispid 4;
    property HorizontalPosition: Single dispid 5;
    property LockAnchor: WordBool dispid 6;
    property RelativeHorizontalPosition: WpsRelativeHorizontalPosition dispid 7;
    property RelativeVerticalPosition: WpsRelativeVerticalPosition dispid 8;
    property VerticalDistanceFromText: Single dispid 9;
    property VerticalPosition: Single dispid 10;
    property Width: Single dispid 11;
    property TextWrap: WordBool dispid 12;
    property Borders: Borders dispid 1100;
    property Range: Range readonly dispid 15;
    procedure Delete; dispid 100;
    procedure Select; dispid 65535;
    procedure Copy; dispid 101;
    procedure Cut; dispid 102;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: InlineShape
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A8-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  InlineShape = interface(_IKsoDispObj)
    ['{000209A8-0000-4B30-A977-D214852036FE}']
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Range: Range; safecall;
    function Get_Field: Field; safecall;
    function Get_type_: WpsInlineShapeType; safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_ScaleHeight: Single; safecall;
    procedure Set_ScaleHeight(prop: Single); safecall;
    function Get_ScaleWidth: Single; safecall;
    procedure Set_ScaleWidth(prop: Single); safecall;
    function Get_LockAspectRatio: KsoTriState; safecall;
    procedure Set_LockAspectRatio(prop: KsoTriState); safecall;
    function Get_Fill: FillFormat; safecall;
    function Get_PictureFormat: PictureFormat; safecall;
    procedure Set_PictureFormat(const prop: PictureFormat); safecall;
    procedure Activate; safecall;
    procedure Reset; safecall;
    procedure Delete; safecall;
    procedure Select; safecall;
    function ConvertToShape: Shape; safecall;
    function Get_TextEffect: TextEffectFormat; safecall;
    procedure Set_TextEffect(const prop: TextEffectFormat); safecall;
    function Get_AlternativeText: WideString; safecall;
    procedure Set_AlternativeText(const prop: WideString); safecall;
    function Get_IsPictureBullet: WordBool; safecall;
    function Get_OLEFormat: OLEFormat; safecall;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Range: Range read Get_Range;
    property Field: Field read Get_Field;
    property type_: WpsInlineShapeType read Get_type_;
    property Height: Single read Get_Height write Set_Height;
    property Width: Single read Get_Width write Set_Width;
    property ScaleHeight: Single read Get_ScaleHeight write Set_ScaleHeight;
    property ScaleWidth: Single read Get_ScaleWidth write Set_ScaleWidth;
    property LockAspectRatio: KsoTriState read Get_LockAspectRatio write Set_LockAspectRatio;
    property Fill: FillFormat read Get_Fill;
    property PictureFormat: PictureFormat read Get_PictureFormat write Set_PictureFormat;
    property TextEffect: TextEffectFormat read Get_TextEffect write Set_TextEffect;
    property AlternativeText: WideString read Get_AlternativeText write Set_AlternativeText;
    property IsPictureBullet: WordBool read Get_IsPictureBullet;
    property OLEFormat: OLEFormat read Get_OLEFormat;
  end;

// *********************************************************************//
// DispIntf:  InlineShapeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A8-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  InlineShapeDisp = dispinterface
    ['{000209A8-0000-4B30-A977-D214852036FE}']
    property Borders: Borders dispid 1100;
    property Range: Range readonly dispid 2;
    property Field: Field readonly dispid 3;
    property type_: WpsInlineShapeType readonly dispid 6;
    property Height: Single dispid 8;
    property Width: Single dispid 9;
    property ScaleHeight: Single dispid 10;
    property ScaleWidth: Single dispid 11;
    property LockAspectRatio: KsoTriState dispid 12;
    property Fill: FillFormat readonly dispid 107;
    property PictureFormat: PictureFormat dispid 118;
    procedure Activate; dispid 100;
    procedure Reset; dispid 101;
    procedure Delete; dispid 102;
    procedure Select; dispid 65535;
    function ConvertToShape: Shape; dispid 104;
    property TextEffect: TextEffectFormat dispid 120;
    property AlternativeText: WideString dispid 131;
    property IsPictureBullet: WordBool readonly dispid 132;
    property OLEFormat: OLEFormat readonly dispid 133;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Field
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Field = interface(_IKsoDispObj)
    ['{0002092F-0000-4B30-A977-D214852036FE}']
    function Get_Code: Range; safecall;
    procedure Set_Code(const prop: Range); safecall;
    function Get_type_: WpsFieldType; safecall;
    function Get_Locked: WordBool; safecall;
    procedure Set_Locked(prop: WordBool); safecall;
    function Get_Kind: WpsFieldKind; safecall;
    function Get_Result: Range; safecall;
    procedure Set_Result(const prop: Range); safecall;
    function Get_Data: WideString; safecall;
    procedure Set_Data(const prop: WideString); safecall;
    function Get_Next: Field; safecall;
    function Get_Previous: Field; safecall;
    function Get_Index: Integer; safecall;
    function Get_ShowCodes: WordBool; safecall;
    procedure Set_ShowCodes(prop: WordBool); safecall;
    procedure Select; safecall;
    function Update: WordBool; safecall;
    procedure Unlink; safecall;
    procedure UpdateSource; safecall;
    procedure DoClick; safecall;
    procedure Copy; safecall;
    procedure Cut; safecall;
    procedure Delete; safecall;
    function Get_IsMergeFormat: WordBool; safecall;
    procedure Set_IsMergeFormat(prop: WordBool); safecall;
    property Code: Range read Get_Code write Set_Code;
    property type_: WpsFieldType read Get_type_;
    property Locked: WordBool read Get_Locked write Set_Locked;
    property Kind: WpsFieldKind read Get_Kind;
    property Result: Range read Get_Result write Set_Result;
    property Data: WideString read Get_Data write Set_Data;
    property Next: Field read Get_Next;
    property Previous: Field read Get_Previous;
    property Index: Integer read Get_Index;
    property ShowCodes: WordBool read Get_ShowCodes write Set_ShowCodes;
    property IsMergeFormat: WordBool read Get_IsMergeFormat write Set_IsMergeFormat;
  end;

// *********************************************************************//
// DispIntf:  FieldDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FieldDisp = dispinterface
    ['{0002092F-0000-4B30-A977-D214852036FE}']
    property Code: Range dispid 0;
    property type_: WpsFieldType readonly dispid 1;
    property Locked: WordBool dispid 2;
    property Kind: WpsFieldKind readonly dispid 3;
    property Result: Range dispid 4;
    property Data: WideString dispid 5;
    property Next: Field readonly dispid 6;
    property Previous: Field readonly dispid 7;
    property Index: Integer readonly dispid 8;
    property ShowCodes: WordBool dispid 9;
    procedure Select; dispid 65535;
    function Update: WordBool; dispid 101;
    procedure Unlink; dispid 102;
    procedure UpdateSource; dispid 103;
    procedure DoClick; dispid 104;
    procedure Copy; dispid 105;
    procedure Cut; dispid 106;
    procedure Delete; dispid 107;
    property IsMergeFormat: WordBool dispid 111;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: FillFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290C8-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FillFormat = interface(KsoFillFormat)
    ['{000290C8-0000-4B30-A977-D214852036FE}']
    function Get_BackColor: ColorFormat; safecall;
    procedure Set_BackColor(const BackColor: ColorFormat); safecall;
    function Get_ForeColor: ColorFormat; safecall;
    procedure Set_ForeColor(const ForeColor: ColorFormat); safecall;
    property BackColor: ColorFormat read Get_BackColor write Set_BackColor;
    property ForeColor: ColorFormat read Get_ForeColor write Set_ForeColor;
  end;

// *********************************************************************//
// DispIntf:  FillFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290C8-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FillFormatDisp = dispinterface
    ['{000290C8-0000-4B30-A977-D214852036FE}']
    property BackColor: ColorFormat dispid 100;
    property ForeColor: ColorFormat dispid 101;
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
// Interface: ColorFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290C6-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ColorFormat = interface(KsoColorFormat)
    ['{000290C6-0000-4B30-A977-D214852036FE}']
  end;

// *********************************************************************//
// DispIntf:  ColorFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290C6-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ColorFormatDisp = dispinterface
    ['{000290C6-0000-4B30-A977-D214852036FE}']
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
// Interface: PictureFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000250C8-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  PictureFormat = interface(KsoPictureFormat)
    ['{000250C8-0000-4B30-A977-D214852036FE}']
    function Get_PictureID: OleVariant; safecall;
    property PictureID: OleVariant read Get_PictureID;
  end;

// *********************************************************************//
// DispIntf:  PictureFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000250C8-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  PictureFormatDisp = dispinterface
    ['{000250C8-0000-4B30-A977-D214852036FE}']
    property PictureID: OleVariant readonly dispid 8192;
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
// Interface: TextEffectFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290CF-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TextEffectFormat = interface(KsoTextEffectFormat)
    ['{000290CF-0000-4B30-A977-D214852036FE}']
  end;

// *********************************************************************//
// DispIntf:  TextEffectFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290CF-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TextEffectFormatDisp = dispinterface
    ['{000290CF-0000-4B30-A977-D214852036FE}']
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
// Interface: OLEFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020322-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  OLEFormat = interface(KsoOLEFormat)
    ['{00020322-0000-4B30-A977-D214852036FE}']
    function Get_ProgID: WideString; safecall;
    function Get_Object_: IDispatch; safecall;
    procedure Activate; safecall;
    procedure DoVerb(VerbIndex: OleVariant); safecall;
    function Get_ClassType: WideString; safecall;
    procedure Set_ClassType(const prop: WideString); safecall;
    property ProgID: WideString read Get_ProgID;
    property Object_: IDispatch read Get_Object_;
    property ClassType: WideString read Get_ClassType write Set_ClassType;
  end;

// *********************************************************************//
// DispIntf:  OLEFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020322-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  OLEFormatDisp = dispinterface
    ['{00020322-0000-4B30-A977-D214852036FE}']
    property ProgID: WideString readonly dispid 1;
    property Object_: IDispatch readonly dispid 2;
    procedure Activate; dispid 3;
    procedure DoVerb(VerbIndex: OleVariant); dispid 4;
    property ClassType: WideString dispid 5;
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
// Interface: TextFrame
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B2-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TextFrame = interface(_IKsoDispObj)
    ['{000209B2-0000-4B30-A977-D214852036FE}']
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
    function Get_TextRange: Range; safecall;
    function Get_ContainingRange: Range; safecall;
    function Get_Next: TextFrame; safecall;
    procedure Set_Next(const prop: TextFrame); safecall;
    function Get_Previous: TextFrame; safecall;
    procedure Set_Previous(const prop: TextFrame); safecall;
    function Get_Overflowing: WordBool; safecall;
    function Get_HasText: WordBool; safecall;
    procedure BreakForwardLink; safecall;
    function ValidLinkTarget(const TargetTextFrame: TextFrame): WordBool; safecall;
    function Get_AutoSize: Integer; safecall;
    procedure Set_AutoSize(prop: Integer); safecall;
    function Get_WordWrap: Integer; safecall;
    procedure Set_WordWrap(prop: Integer); safecall;
    function Get_WrapText: Integer; safecall;
    procedure Set_WrapText(prop: Integer); safecall;
    function Get_RotateText: Integer; safecall;
    procedure Set_RotateText(prop: Integer); safecall;
    function Get_ResizeText: Integer; safecall;
    procedure Set_ResizeText(prop: Integer); safecall;
    property MarginBottom: Single read Get_MarginBottom write Set_MarginBottom;
    property MarginLeft: Single read Get_MarginLeft write Set_MarginLeft;
    property MarginRight: Single read Get_MarginRight write Set_MarginRight;
    property MarginTop: Single read Get_MarginTop write Set_MarginTop;
    property Orientation: KsoTextOrientation read Get_Orientation write Set_Orientation;
    property TextRange: Range read Get_TextRange;
    property ContainingRange: Range read Get_ContainingRange;
    property Next: TextFrame read Get_Next write Set_Next;
    property Previous: TextFrame read Get_Previous write Set_Previous;
    property Overflowing: WordBool read Get_Overflowing;
    property HasText: WordBool read Get_HasText;
    property AutoSize: Integer read Get_AutoSize write Set_AutoSize;
    property WordWrap: Integer read Get_WordWrap write Set_WordWrap;
    property WrapText: Integer read Get_WrapText write Set_WrapText;
    property RotateText: Integer read Get_RotateText write Set_RotateText;
    property ResizeText: Integer read Get_ResizeText write Set_ResizeText;
  end;

// *********************************************************************//
// DispIntf:  TextFrameDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B2-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TextFrameDisp = dispinterface
    ['{000209B2-0000-4B30-A977-D214852036FE}']
    property MarginBottom: Single dispid 100;
    property MarginLeft: Single dispid 101;
    property MarginRight: Single dispid 102;
    property MarginTop: Single dispid 103;
    property Orientation: KsoTextOrientation dispid 104;
    property TextRange: Range readonly dispid 4999;
    property ContainingRange: Range readonly dispid 5000;
    property Next: TextFrame dispid 5001;
    property Previous: TextFrame dispid 5002;
    property Overflowing: WordBool readonly dispid 5003;
    property HasText: WordBool readonly dispid 5008;
    procedure BreakForwardLink; dispid 5004;
    function ValidLinkTarget(const TargetTextFrame: TextFrame): WordBool; dispid 5006;
    property AutoSize: Integer dispid 5009;
    property WordWrap: Integer dispid 5010;
    property WrapText: Integer dispid 5011;
    property RotateText: Integer dispid 5012;
    property ResizeText: Integer dispid 5013;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Adjustments
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C4-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Adjustments = interface(KsoAdjustments)
    ['{000209C4-0000-4B30-A977-D214852036FE}']
  end;

// *********************************************************************//
// DispIntf:  AdjustmentsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C4-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  AdjustmentsDisp = dispinterface
    ['{000209C4-0000-4B30-A977-D214852036FE}']
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
// Interface: CalloutFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C5-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  CalloutFormat = interface(KsoCalloutFormat)
    ['{000209C5-0000-4B30-A977-D214852036FE}']
  end;

// *********************************************************************//
// DispIntf:  CalloutFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C5-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  CalloutFormatDisp = dispinterface
    ['{000209C5-0000-4B30-A977-D214852036FE}']
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
// Interface: ConnectorFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290C7-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ConnectorFormat = interface(KsoConnectorFormat)
    ['{000290C7-0000-4B30-A977-D214852036FE}']
    procedure BeginConnect(const ConnectedShape: Shape; ConnectionSite: SYSINT); safecall;
    procedure EndConnect(const ConnectedShape: Shape; ConnectionSite: SYSINT); safecall;
    function Get_BeginConnectedShape: Shape; safecall;
    function Get_EndConnectedShape: Shape; safecall;
    property BeginConnectedShape: Shape read Get_BeginConnectedShape;
    property EndConnectedShape: Shape read Get_EndConnectedShape;
  end;

// *********************************************************************//
// DispIntf:  ConnectorFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290C7-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ConnectorFormatDisp = dispinterface
    ['{000290C7-0000-4B30-A977-D214852036FE}']
    procedure BeginConnect(const ConnectedShape: Shape; ConnectionSite: SYSINT); dispid 10;
    procedure EndConnect(const ConnectedShape: Shape; ConnectionSite: SYSINT); dispid 12;
    property BeginConnectedShape: Shape readonly dispid 101;
    property EndConnectedShape: Shape readonly dispid 104;
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
// Interface: GroupShapes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290B6-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  GroupShapes = interface(KsoGroupShapes)
    ['{000290B6-0000-4B30-A977-D214852036FE}']
    function Item(Index: SYSINT): Shape; safecall;
    function Range(Index: OleVariant): ShapeRange; safecall;
  end;

// *********************************************************************//
// DispIntf:  GroupShapesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290B6-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  GroupShapesDisp = dispinterface
    ['{000290B6-0000-4B30-A977-D214852036FE}']
    function Item(Index: SYSINT): Shape; dispid 0;
    function Range(Index: OleVariant): ShapeRange; dispid 10;
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
// Interface: LineFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290CA-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  LineFormat = interface(KsoLineFormat)
    ['{000290CA-0000-4B30-A977-D214852036FE}']
    function Get_BackColor: ColorFormat; safecall;
    procedure Set_BackColor(const BackColor: ColorFormat); safecall;
    function Get_ForeColor: ColorFormat; safecall;
    procedure Set_ForeColor(const ForeColor: ColorFormat); safecall;
    property BackColor: ColorFormat read Get_BackColor write Set_BackColor;
    property ForeColor: ColorFormat read Get_ForeColor write Set_ForeColor;
  end;

// *********************************************************************//
// DispIntf:  LineFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290CA-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  LineFormatDisp = dispinterface
    ['{000290CA-0000-4B30-A977-D214852036FE}']
    property BackColor: ColorFormat dispid 100;
    property ForeColor: ColorFormat dispid 108;
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
// Interface: ShapeNodes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290CE-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ShapeNodes = interface(KsoShapeNodes)
    ['{000290CE-0000-4B30-A977-D214852036FE}']
    function Item(Index: SYSINT): ShapeNode; safecall;
  end;

// *********************************************************************//
// DispIntf:  ShapeNodesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290CE-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ShapeNodesDisp = dispinterface
    ['{000290CE-0000-4B30-A977-D214852036FE}']
    function Item(Index: SYSINT): ShapeNode; dispid 0;
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
// Interface: ShapeNode
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290CD-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ShapeNode = interface(KsoShapeNode)
    ['{000290CD-0000-4B30-A977-D214852036FE}']
  end;

// *********************************************************************//
// DispIntf:  ShapeNodeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290CD-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ShapeNodeDisp = dispinterface
    ['{000290CD-0000-4B30-A977-D214852036FE}']
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
// Interface: ShadowFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290CC-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ShadowFormat = interface(KsoShadowFormat)
    ['{000290CC-0000-4B30-A977-D214852036FE}']
    function Get_ForeColor: ColorFormat; safecall;
    procedure Set_ForeColor(const ForeColor: ColorFormat); safecall;
    property ForeColor: ColorFormat read Get_ForeColor write Set_ForeColor;
  end;

// *********************************************************************//
// DispIntf:  ShadowFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290CC-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ShadowFormatDisp = dispinterface
    ['{000290CC-0000-4B30-A977-D214852036FE}']
    property ForeColor: ColorFormat dispid 100;
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
// Interface: ThreeDFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290D0-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ThreeDFormat = interface(KsoThreeDFormat)
    ['{000290D0-0000-4B30-A977-D214852036FE}']
    function Get_ExtrusionColor: ColorFormat; safecall;
    property ExtrusionColor: ColorFormat read Get_ExtrusionColor;
  end;

// *********************************************************************//
// DispIntf:  ThreeDFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290D0-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ThreeDFormatDisp = dispinterface
    ['{000290D0-0000-4B30-A977-D214852036FE}']
    property ExtrusionColor: ColorFormat readonly dispid 101;
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
// Interface: Diagram
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290EC-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Diagram = interface(_KsoDiagram)
    ['{000290EC-0000-4B30-A977-D214852036FE}']
    function Get_Nodes: DiagramNodes; safecall;
    property Nodes: DiagramNodes read Get_Nodes;
  end;

// *********************************************************************//
// DispIntf:  DiagramDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290EC-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DiagramDisp = dispinterface
    ['{000290EC-0000-4B30-A977-D214852036FE}']
    property Nodes: DiagramNodes readonly dispid 101;
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
// Interface: DiagramNodes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290EB-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DiagramNodes = interface(_KsoDiagramNodes)
    ['{000290EB-0000-4B30-A977-D214852036FE}']
    function Item(Index: SYSINT): DiagramNode; safecall;
  end;

// *********************************************************************//
// DispIntf:  DiagramNodesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290EB-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DiagramNodesDisp = dispinterface
    ['{000290EB-0000-4B30-A977-D214852036FE}']
    function Item(Index: SYSINT): DiagramNode; dispid 0;
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
// Interface: DiagramNode
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290E9-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DiagramNode = interface(_KsoDiagramNode)
    ['{000290E9-0000-4B30-A977-D214852036FE}']
    function AddNode(Pos: KsoRelativeNodePosition; NodeType: KsoDiagramNodeType): DiagramNode; safecall;
    procedure MoveNode(const TargetNode: DiagramNode; Pos: KsoRelativeNodePosition); safecall;
    procedure ReplaceNode(const TargetNode: DiagramNode); safecall;
    procedure SwapNode(const TargetNode: DiagramNode; SwapChildren: WordBool); safecall;
    function CloneNode(CopyChildren: WordBool; const TargetNode: DiagramNode; 
                       Pos: KsoRelativeNodePosition): DiagramNode; safecall;
    procedure TransferChildren(const ReceivingNode: DiagramNode); safecall;
    function NextNode: DiagramNode; safecall;
    function PrevNode: DiagramNode; safecall;
    function Get_Children: DiagramNodeChildren; safecall;
    function Get_Shape: Shape; safecall;
    function Get_Root: DiagramNode; safecall;
    function Get_Diagram: Diagram; safecall;
    function Get_TextShape: Shape; safecall;
    property Children: DiagramNodeChildren read Get_Children;
    property Shape: Shape read Get_Shape;
    property Root: DiagramNode read Get_Root;
    property Diagram: Diagram read Get_Diagram;
    property TextShape: Shape read Get_TextShape;
  end;

// *********************************************************************//
// DispIntf:  DiagramNodeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290E9-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DiagramNodeDisp = dispinterface
    ['{000290E9-0000-4B30-A977-D214852036FE}']
    function AddNode(Pos: KsoRelativeNodePosition; NodeType: KsoDiagramNodeType): DiagramNode; dispid 10;
    procedure MoveNode(const TargetNode: DiagramNode; Pos: KsoRelativeNodePosition); dispid 12;
    procedure ReplaceNode(const TargetNode: DiagramNode); dispid 13;
    procedure SwapNode(const TargetNode: DiagramNode; SwapChildren: WordBool); dispid 14;
    function CloneNode(CopyChildren: WordBool; const TargetNode: DiagramNode; 
                       Pos: KsoRelativeNodePosition): DiagramNode; dispid 15;
    procedure TransferChildren(const ReceivingNode: DiagramNode); dispid 16;
    function NextNode: DiagramNode; dispid 17;
    function PrevNode: DiagramNode; dispid 18;
    property Children: DiagramNodeChildren readonly dispid 101;
    property Shape: Shape readonly dispid 102;
    property Root: DiagramNode readonly dispid 103;
    property Diagram: Diagram readonly dispid 104;
    property TextShape: Shape readonly dispid 106;
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
// Interface: DiagramNodeChildren
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290EA-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DiagramNodeChildren = interface(_KsoDiagramNodeChildren)
    ['{000290EA-0000-4B30-A977-D214852036FE}']
    function Item(Index: SYSINT): DiagramNode; safecall;
    function AddNode(Index: SYSINT; NodeType: KsoDiagramNodeType): DiagramNode; safecall;
    function Get_FirstChild: DiagramNode; safecall;
    function Get_LastChild: DiagramNode; safecall;
    property FirstChild: DiagramNode read Get_FirstChild;
    property LastChild: DiagramNode read Get_LastChild;
  end;

// *********************************************************************//
// DispIntf:  DiagramNodeChildrenDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290EA-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DiagramNodeChildrenDisp = dispinterface
    ['{000290EA-0000-4B30-A977-D214852036FE}']
    function Item(Index: SYSINT): DiagramNode; dispid 0;
    function AddNode(Index: SYSINT; NodeType: KsoDiagramNodeType): DiagramNode; dispid 10;
    property FirstChild: DiagramNode readonly dispid 103;
    property LastChild: DiagramNode readonly dispid 104;
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
// Interface: CanvasShapes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00029073-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  CanvasShapes = interface(KsoCanvasShapes)
    ['{00029073-0000-4B30-A977-D214852036FE}']
    function Item(Index: SYSINT): Shape; safecall;
    function AddCallout(Type_: KsoCalloutType; Left: Single; Top: Single; Width: Single; 
                        Height: Single): Shape; safecall;
    function AddConnector(Type_: KsoConnectorType; BeginX: Single; BeginY: Single; EndX: Single; 
                          EndY: Single): Shape; safecall;
    function AddCurve(SafeArrayOfPoints: OleVariant): Shape; safecall;
    function AddLabel(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                      Height: Single): Shape; safecall;
    function AddLine(BeginX: Single; BeginY: Single; EndX: Single; EndY: Single): Shape; safecall;
    function AddPicture(const FileName: WideString; LinkToFile: KsoTriState; 
                        SaveWithDocument: KsoTriState; Left: Single; Top: Single; Width: Single; 
                        Height: Single): Shape; safecall;
    function AddPolyline(SafeArrayOfPoints: OleVariant): Shape; safecall;
    function AddShape(Type_: KsoAutoShapeType; Left: Single; Top: Single; Width: Single; 
                      Height: Single): Shape; safecall;
    function AddTextEffect(PresetTextEffect: KsoPresetTextEffect; const Text: WideString; 
                           const FontName: WideString; FontSize: Single; FontBold: KsoTriState; 
                           FontItalic: KsoTriState; Left: Single; Top: Single): Shape; safecall;
    function AddTextbox(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                        Height: Single): Shape; safecall;
    function BuildFreeform(EditingType: KsoEditingType; X1: Single; Y1: Single): FreeformBuilder; safecall;
    function Range(Index: OleVariant): ShapeRange; safecall;
    function Get_Background: Shape; safecall;
    property Background: Shape read Get_Background;
  end;

// *********************************************************************//
// DispIntf:  CanvasShapesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00029073-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  CanvasShapesDisp = dispinterface
    ['{00029073-0000-4B30-A977-D214852036FE}']
    function Item(Index: SYSINT): Shape; dispid 0;
    function AddCallout(Type_: KsoCalloutType; Left: Single; Top: Single; Width: Single; 
                        Height: Single): Shape; dispid 10;
    function AddConnector(Type_: KsoConnectorType; BeginX: Single; BeginY: Single; EndX: Single; 
                          EndY: Single): Shape; dispid 11;
    function AddCurve(SafeArrayOfPoints: OleVariant): Shape; dispid 12;
    function AddLabel(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                      Height: Single): Shape; dispid 13;
    function AddLine(BeginX: Single; BeginY: Single; EndX: Single; EndY: Single): Shape; dispid 14;
    function AddPicture(const FileName: WideString; LinkToFile: KsoTriState; 
                        SaveWithDocument: KsoTriState; Left: Single; Top: Single; Width: Single; 
                        Height: Single): Shape; dispid 15;
    function AddPolyline(SafeArrayOfPoints: OleVariant): Shape; dispid 16;
    function AddShape(Type_: KsoAutoShapeType; Left: Single; Top: Single; Width: Single; 
                      Height: Single): Shape; dispid 17;
    function AddTextEffect(PresetTextEffect: KsoPresetTextEffect; const Text: WideString; 
                           const FontName: WideString; FontSize: Single; FontBold: KsoTriState; 
                           FontItalic: KsoTriState; Left: Single; Top: Single): Shape; dispid 18;
    function AddTextbox(Orientation: KsoTextOrientation; Left: Single; Top: Single; Width: Single; 
                        Height: Single): Shape; dispid 19;
    function BuildFreeform(EditingType: KsoEditingType; X1: Single; Y1: Single): FreeformBuilder; dispid 20;
    function Range(Index: OleVariant): ShapeRange; dispid 21;
    property Background: Shape readonly dispid 100;
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
// Interface: FreeformBuilder
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290C9-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FreeformBuilder = interface(KsoFreeformBuilder)
    ['{000290C9-0000-4B30-A977-D214852036FE}']
    function ConvertToShape: Shape; safecall;
  end;

// *********************************************************************//
// DispIntf:  FreeformBuilderDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000290C9-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FreeformBuilderDisp = dispinterface
    ['{000290C9-0000-4B30-A977-D214852036FE}']
    function ConvertToShape: Shape; dispid 1610874880;
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
// Interface: Paragraphs
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020958-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Paragraphs = interface(_IKsoDispObj)
    ['{00020958-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_First: Paragraph; safecall;
    function Get_Last: Paragraph; safecall;
    function Get_Format: _ParagraphFormat; safecall;
    procedure Set_Format(const prop: _ParagraphFormat); safecall;
    function Get_Alignment: WpsParagraphAlignment; safecall;
    procedure Set_Alignment(prop: WpsParagraphAlignment); safecall;
    function Get_KeepTogether: Integer; safecall;
    procedure Set_KeepTogether(prop: Integer); safecall;
    function Get_KeepWithNext: Integer; safecall;
    procedure Set_KeepWithNext(prop: Integer); safecall;
    function Get_PageBreakBefore: Integer; safecall;
    procedure Set_PageBreakBefore(prop: Integer); safecall;
    function Get_RightIndent: Single; safecall;
    procedure Set_RightIndent(prop: Single); safecall;
    function Get_LeftIndent: Single; safecall;
    procedure Set_LeftIndent(prop: Single); safecall;
    function Get_FirstLineIndent: Single; safecall;
    procedure Set_FirstLineIndent(prop: Single); safecall;
    function Get_LineSpacing: Single; safecall;
    procedure Set_LineSpacing(prop: Single); safecall;
    function Get_LineSpacingRule: WpsLineSpacing; safecall;
    procedure Set_LineSpacingRule(prop: WpsLineSpacing); safecall;
    function Get_SpaceBefore: Single; safecall;
    procedure Set_SpaceBefore(prop: Single); safecall;
    function Get_SpaceAfter: Single; safecall;
    procedure Set_SpaceAfter(prop: Single); safecall;
    function Get_WidowControl: Integer; safecall;
    procedure Set_WidowControl(prop: Integer); safecall;
    function Get_HangingPunctuation: Integer; safecall;
    procedure Set_HangingPunctuation(prop: Integer); safecall;
    function Get_HalfWidthPunctuationOnTopOfLine: Integer; safecall;
    procedure Set_HalfWidthPunctuationOnTopOfLine(prop: Integer); safecall;
    function Get_AddSpaceBetweenFarEastAndAlpha: Integer; safecall;
    procedure Set_AddSpaceBetweenFarEastAndAlpha(prop: Integer); safecall;
    function Get_AddSpaceBetweenFarEastAndDigit: Integer; safecall;
    procedure Set_AddSpaceBetweenFarEastAndDigit(prop: Integer); safecall;
    function Get_BaseLineAlignment: WpsBaselineAlignment; safecall;
    procedure Set_BaseLineAlignment(prop: WpsBaselineAlignment); safecall;
    function Get_AutoAdjustRightIndent: Integer; safecall;
    procedure Set_AutoAdjustRightIndent(prop: Integer); safecall;
    function Get_DisableLineHeightGrid: Integer; safecall;
    procedure Set_DisableLineHeightGrid(prop: Integer); safecall;
    function Item(Index: Integer): Paragraph; safecall;
    function Add(const Range: Range): Paragraph; safecall;
    procedure CloseUp; safecall;
    procedure OpenUp; safecall;
    procedure OpenOrCloseUp; safecall;
    procedure TabHangingIndent(Count: Smallint); safecall;
    procedure TabIndent(Count: Smallint); safecall;
    procedure Reset; safecall;
    procedure Space1; safecall;
    procedure Space15; safecall;
    procedure Space2; safecall;
    procedure IndentCharWidth(Count: Smallint); safecall;
    procedure IndentFirstLineCharWidth(Count: Smallint); safecall;
    procedure Indent; safecall;
    procedure Outdent; safecall;
    function Get_CharacterUnitRightIndent: Single; safecall;
    procedure Set_CharacterUnitRightIndent(prop: Single); safecall;
    function Get_CharacterUnitLeftIndent: Single; safecall;
    procedure Set_CharacterUnitLeftIndent(prop: Single); safecall;
    function Get_CharacterUnitFirstLineIndent: Single; safecall;
    procedure Set_CharacterUnitFirstLineIndent(prop: Single); safecall;
    function Get_LineUnitBefore: Single; safecall;
    procedure Set_LineUnitBefore(prop: Single); safecall;
    function Get_LineUnitAfter: Single; safecall;
    procedure Set_LineUnitAfter(prop: Single); safecall;
    function Get_SpaceBeforeAuto: Integer; safecall;
    procedure Set_SpaceBeforeAuto(prop: Integer); safecall;
    function Get_SpaceAfterAuto: Integer; safecall;
    procedure Set_SpaceAfterAuto(prop: Integer); safecall;
    procedure IncreaseSpacing; safecall;
    procedure DecreaseSpacing; safecall;
    function Get__Range: Range; safecall;
    procedure OutlineDemote; safecall;
    procedure OutlinePromote; safecall;
    procedure OutlineDemoteToBody; safecall;
    function Get_OutlineLevel: WpsOutlineLevel; safecall;
    procedure Set_OutlineLevel(prop: WpsOutlineLevel); safecall;
    function Get_Style: Style; safecall;
    procedure Set_Style(const prop: Style); safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property First: Paragraph read Get_First;
    property Last: Paragraph read Get_Last;
    property Format: _ParagraphFormat read Get_Format write Set_Format;
    property Alignment: WpsParagraphAlignment read Get_Alignment write Set_Alignment;
    property KeepTogether: Integer read Get_KeepTogether write Set_KeepTogether;
    property KeepWithNext: Integer read Get_KeepWithNext write Set_KeepWithNext;
    property PageBreakBefore: Integer read Get_PageBreakBefore write Set_PageBreakBefore;
    property RightIndent: Single read Get_RightIndent write Set_RightIndent;
    property LeftIndent: Single read Get_LeftIndent write Set_LeftIndent;
    property FirstLineIndent: Single read Get_FirstLineIndent write Set_FirstLineIndent;
    property LineSpacing: Single read Get_LineSpacing write Set_LineSpacing;
    property LineSpacingRule: WpsLineSpacing read Get_LineSpacingRule write Set_LineSpacingRule;
    property SpaceBefore: Single read Get_SpaceBefore write Set_SpaceBefore;
    property SpaceAfter: Single read Get_SpaceAfter write Set_SpaceAfter;
    property WidowControl: Integer read Get_WidowControl write Set_WidowControl;
    property HangingPunctuation: Integer read Get_HangingPunctuation write Set_HangingPunctuation;
    property HalfWidthPunctuationOnTopOfLine: Integer read Get_HalfWidthPunctuationOnTopOfLine write Set_HalfWidthPunctuationOnTopOfLine;
    property AddSpaceBetweenFarEastAndAlpha: Integer read Get_AddSpaceBetweenFarEastAndAlpha write Set_AddSpaceBetweenFarEastAndAlpha;
    property AddSpaceBetweenFarEastAndDigit: Integer read Get_AddSpaceBetweenFarEastAndDigit write Set_AddSpaceBetweenFarEastAndDigit;
    property BaseLineAlignment: WpsBaselineAlignment read Get_BaseLineAlignment write Set_BaseLineAlignment;
    property AutoAdjustRightIndent: Integer read Get_AutoAdjustRightIndent write Set_AutoAdjustRightIndent;
    property DisableLineHeightGrid: Integer read Get_DisableLineHeightGrid write Set_DisableLineHeightGrid;
    property CharacterUnitRightIndent: Single read Get_CharacterUnitRightIndent write Set_CharacterUnitRightIndent;
    property CharacterUnitLeftIndent: Single read Get_CharacterUnitLeftIndent write Set_CharacterUnitLeftIndent;
    property CharacterUnitFirstLineIndent: Single read Get_CharacterUnitFirstLineIndent write Set_CharacterUnitFirstLineIndent;
    property LineUnitBefore: Single read Get_LineUnitBefore write Set_LineUnitBefore;
    property LineUnitAfter: Single read Get_LineUnitAfter write Set_LineUnitAfter;
    property SpaceBeforeAuto: Integer read Get_SpaceBeforeAuto write Set_SpaceBeforeAuto;
    property SpaceAfterAuto: Integer read Get_SpaceAfterAuto write Set_SpaceAfterAuto;
    property _Range: Range read Get__Range;
    property OutlineLevel: WpsOutlineLevel read Get_OutlineLevel write Set_OutlineLevel;
    property Style: Style read Get_Style write Set_Style;
  end;

// *********************************************************************//
// DispIntf:  ParagraphsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020958-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ParagraphsDisp = dispinterface
    ['{00020958-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property First: Paragraph readonly dispid 3;
    property Last: Paragraph readonly dispid 4;
    property Format: _ParagraphFormat dispid 1102;
    property Alignment: WpsParagraphAlignment dispid 101;
    property KeepTogether: Integer dispid 102;
    property KeepWithNext: Integer dispid 103;
    property PageBreakBefore: Integer dispid 104;
    property RightIndent: Single dispid 106;
    property LeftIndent: Single dispid 107;
    property FirstLineIndent: Single dispid 108;
    property LineSpacing: Single dispid 109;
    property LineSpacingRule: WpsLineSpacing dispid 110;
    property SpaceBefore: Single dispid 111;
    property SpaceAfter: Single dispid 112;
    property WidowControl: Integer dispid 114;
    property HangingPunctuation: Integer dispid 119;
    property HalfWidthPunctuationOnTopOfLine: Integer dispid 120;
    property AddSpaceBetweenFarEastAndAlpha: Integer dispid 121;
    property AddSpaceBetweenFarEastAndDigit: Integer dispid 122;
    property BaseLineAlignment: WpsBaselineAlignment dispid 123;
    property AutoAdjustRightIndent: Integer dispid 124;
    property DisableLineHeightGrid: Integer dispid 125;
    function Item(Index: Integer): Paragraph; dispid 0;
    function Add(const Range: Range): Paragraph; dispid 5;
    procedure CloseUp; dispid 301;
    procedure OpenUp; dispid 302;
    procedure OpenOrCloseUp; dispid 303;
    procedure TabHangingIndent(Count: Smallint); dispid 304;
    procedure TabIndent(Count: Smallint); dispid 306;
    procedure Reset; dispid 312;
    procedure Space1; dispid 313;
    procedure Space15; dispid 314;
    procedure Space2; dispid 315;
    procedure IndentCharWidth(Count: Smallint); dispid 320;
    procedure IndentFirstLineCharWidth(Count: Smallint); dispid 322;
    procedure Indent; dispid 333;
    procedure Outdent; dispid 334;
    property CharacterUnitRightIndent: Single dispid 126;
    property CharacterUnitLeftIndent: Single dispid 127;
    property CharacterUnitFirstLineIndent: Single dispid 128;
    property LineUnitBefore: Single dispid 129;
    property LineUnitAfter: Single dispid 130;
    property SpaceBeforeAuto: Integer dispid 132;
    property SpaceAfterAuto: Integer dispid 133;
    procedure IncreaseSpacing; dispid 335;
    procedure DecreaseSpacing; dispid 336;
    property _Range: Range readonly dispid -5;
    procedure OutlineDemote; dispid 352;
    procedure OutlinePromote; dispid 353;
    procedure OutlineDemoteToBody; dispid 354;
    property OutlineLevel: WpsOutlineLevel dispid 355;
    property Style: Style dispid 100;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Paragraph
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020957-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Paragraph = interface(_IKsoDispObj)
    ['{00020957-0000-4B30-A977-D214852036FE}']
    function Get_Range: Range; safecall;
    function Get_Format: _ParagraphFormat; safecall;
    procedure Set_Format(const prop: _ParagraphFormat); safecall;
    function Get_DropCap: DropCap; safecall;
    function Get_Style: Style; safecall;
    procedure Set_Style(const prop: Style); safecall;
    function Get_Alignment: WpsParagraphAlignment; safecall;
    procedure Set_Alignment(prop: WpsParagraphAlignment); safecall;
    function Get_KeepTogether: Integer; safecall;
    procedure Set_KeepTogether(prop: Integer); safecall;
    function Get_KeepWithNext: Integer; safecall;
    procedure Set_KeepWithNext(prop: Integer); safecall;
    function Get_PageBreakBefore: Integer; safecall;
    procedure Set_PageBreakBefore(prop: Integer); safecall;
    function Get_RightIndent: Single; safecall;
    procedure Set_RightIndent(prop: Single); safecall;
    function Get_LeftIndent: Single; safecall;
    procedure Set_LeftIndent(prop: Single); safecall;
    function Get_FirstLineIndent: Single; safecall;
    procedure Set_FirstLineIndent(prop: Single); safecall;
    function Get_LineSpacing: Single; safecall;
    procedure Set_LineSpacing(prop: Single); safecall;
    function Get_LineSpacingRule: WpsLineSpacing; safecall;
    procedure Set_LineSpacingRule(prop: WpsLineSpacing); safecall;
    function Get_SpaceBefore: Single; safecall;
    procedure Set_SpaceBefore(prop: Single); safecall;
    function Get_SpaceAfter: Single; safecall;
    procedure Set_SpaceAfter(prop: Single); safecall;
    function Get_WidowControl: Integer; safecall;
    procedure Set_WidowControl(prop: Integer); safecall;
    function Get_HangingPunctuation: Integer; safecall;
    procedure Set_HangingPunctuation(prop: Integer); safecall;
    function Get_HalfWidthPunctuationOnTopOfLine: Integer; safecall;
    procedure Set_HalfWidthPunctuationOnTopOfLine(prop: Integer); safecall;
    function Get_AddSpaceBetweenFarEastAndAlpha: Integer; safecall;
    procedure Set_AddSpaceBetweenFarEastAndAlpha(prop: Integer); safecall;
    function Get_AddSpaceBetweenFarEastAndDigit: Integer; safecall;
    procedure Set_AddSpaceBetweenFarEastAndDigit(prop: Integer); safecall;
    function Get_BaseLineAlignment: WpsBaselineAlignment; safecall;
    procedure Set_BaseLineAlignment(prop: WpsBaselineAlignment); safecall;
    function Get_AutoAdjustRightIndent: Integer; safecall;
    procedure Set_AutoAdjustRightIndent(prop: Integer); safecall;
    function Get_DisableLineHeightGrid: Integer; safecall;
    procedure Set_DisableLineHeightGrid(prop: Integer); safecall;
    procedure CloseUp; safecall;
    procedure OpenUp; safecall;
    procedure OpenOrCloseUp; safecall;
    procedure TabHangingIndent(Count: Smallint); safecall;
    procedure TabIndent(Count: Smallint); safecall;
    procedure Reset; safecall;
    procedure Space1; safecall;
    procedure Space15; safecall;
    procedure Space2; safecall;
    procedure IndentCharWidth(Count: Smallint); safecall;
    procedure IndentFirstLineCharWidth(Count: Smallint); safecall;
    function Next(Count: Integer): Paragraph; safecall;
    function Previous(Count: Integer): Paragraph; safecall;
    procedure Indent; safecall;
    procedure Outdent; safecall;
    function Get_CharacterUnitRightIndent: Single; safecall;
    procedure Set_CharacterUnitRightIndent(prop: Single); safecall;
    function Get_CharacterUnitLeftIndent: Single; safecall;
    procedure Set_CharacterUnitLeftIndent(prop: Single); safecall;
    function Get_CharacterUnitFirstLineIndent: Single; safecall;
    procedure Set_CharacterUnitFirstLineIndent(prop: Single); safecall;
    function Get_LineUnitBefore: Single; safecall;
    procedure Set_LineUnitBefore(prop: Single); safecall;
    function Get_LineUnitAfter: Single; safecall;
    procedure Set_LineUnitAfter(prop: Single); safecall;
    function Get_SpaceBeforeAuto: Integer; safecall;
    procedure Set_SpaceBeforeAuto(prop: Integer); safecall;
    function Get_SpaceAfterAuto: Integer; safecall;
    procedure Set_SpaceAfterAuto(prop: Integer); safecall;
    function Get_FarEastLineBreakControl: Integer; safecall;
    procedure Set_FarEastLineBreakControl(prop: Integer); safecall;
    function Get_WordWrap: Integer; safecall;
    procedure Set_WordWrap(prop: Integer); safecall;
    function Get_OutlineLevel: WpsOutlineLevel; safecall;
    procedure Set_OutlineLevel(prop: WpsOutlineLevel); safecall;
    procedure OutlineDemote; safecall;
    procedure OutlinePromote; safecall;
    procedure OutlineDemoteToBody; safecall;
    property Range: Range read Get_Range;
    property Format: _ParagraphFormat read Get_Format write Set_Format;
    property DropCap: DropCap read Get_DropCap;
    property Style: Style read Get_Style write Set_Style;
    property Alignment: WpsParagraphAlignment read Get_Alignment write Set_Alignment;
    property KeepTogether: Integer read Get_KeepTogether write Set_KeepTogether;
    property KeepWithNext: Integer read Get_KeepWithNext write Set_KeepWithNext;
    property PageBreakBefore: Integer read Get_PageBreakBefore write Set_PageBreakBefore;
    property RightIndent: Single read Get_RightIndent write Set_RightIndent;
    property LeftIndent: Single read Get_LeftIndent write Set_LeftIndent;
    property FirstLineIndent: Single read Get_FirstLineIndent write Set_FirstLineIndent;
    property LineSpacing: Single read Get_LineSpacing write Set_LineSpacing;
    property LineSpacingRule: WpsLineSpacing read Get_LineSpacingRule write Set_LineSpacingRule;
    property SpaceBefore: Single read Get_SpaceBefore write Set_SpaceBefore;
    property SpaceAfter: Single read Get_SpaceAfter write Set_SpaceAfter;
    property WidowControl: Integer read Get_WidowControl write Set_WidowControl;
    property HangingPunctuation: Integer read Get_HangingPunctuation write Set_HangingPunctuation;
    property HalfWidthPunctuationOnTopOfLine: Integer read Get_HalfWidthPunctuationOnTopOfLine write Set_HalfWidthPunctuationOnTopOfLine;
    property AddSpaceBetweenFarEastAndAlpha: Integer read Get_AddSpaceBetweenFarEastAndAlpha write Set_AddSpaceBetweenFarEastAndAlpha;
    property AddSpaceBetweenFarEastAndDigit: Integer read Get_AddSpaceBetweenFarEastAndDigit write Set_AddSpaceBetweenFarEastAndDigit;
    property BaseLineAlignment: WpsBaselineAlignment read Get_BaseLineAlignment write Set_BaseLineAlignment;
    property AutoAdjustRightIndent: Integer read Get_AutoAdjustRightIndent write Set_AutoAdjustRightIndent;
    property DisableLineHeightGrid: Integer read Get_DisableLineHeightGrid write Set_DisableLineHeightGrid;
    property CharacterUnitRightIndent: Single read Get_CharacterUnitRightIndent write Set_CharacterUnitRightIndent;
    property CharacterUnitLeftIndent: Single read Get_CharacterUnitLeftIndent write Set_CharacterUnitLeftIndent;
    property CharacterUnitFirstLineIndent: Single read Get_CharacterUnitFirstLineIndent write Set_CharacterUnitFirstLineIndent;
    property LineUnitBefore: Single read Get_LineUnitBefore write Set_LineUnitBefore;
    property LineUnitAfter: Single read Get_LineUnitAfter write Set_LineUnitAfter;
    property SpaceBeforeAuto: Integer read Get_SpaceBeforeAuto write Set_SpaceBeforeAuto;
    property SpaceAfterAuto: Integer read Get_SpaceAfterAuto write Set_SpaceAfterAuto;
    property FarEastLineBreakControl: Integer read Get_FarEastLineBreakControl write Set_FarEastLineBreakControl;
    property WordWrap: Integer read Get_WordWrap write Set_WordWrap;
    property OutlineLevel: WpsOutlineLevel read Get_OutlineLevel write Set_OutlineLevel;
  end;

// *********************************************************************//
// DispIntf:  ParagraphDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020957-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ParagraphDisp = dispinterface
    ['{00020957-0000-4B30-A977-D214852036FE}']
    property Range: Range readonly dispid 0;
    property Format: _ParagraphFormat dispid 1102;
    property DropCap: DropCap readonly dispid 13;
    property Style: Style dispid 100;
    property Alignment: WpsParagraphAlignment dispid 101;
    property KeepTogether: Integer dispid 102;
    property KeepWithNext: Integer dispid 103;
    property PageBreakBefore: Integer dispid 104;
    property RightIndent: Single dispid 106;
    property LeftIndent: Single dispid 107;
    property FirstLineIndent: Single dispid 108;
    property LineSpacing: Single dispid 109;
    property LineSpacingRule: WpsLineSpacing dispid 110;
    property SpaceBefore: Single dispid 111;
    property SpaceAfter: Single dispid 112;
    property WidowControl: Integer dispid 114;
    property HangingPunctuation: Integer dispid 119;
    property HalfWidthPunctuationOnTopOfLine: Integer dispid 120;
    property AddSpaceBetweenFarEastAndAlpha: Integer dispid 121;
    property AddSpaceBetweenFarEastAndDigit: Integer dispid 122;
    property BaseLineAlignment: WpsBaselineAlignment dispid 123;
    property AutoAdjustRightIndent: Integer dispid 124;
    property DisableLineHeightGrid: Integer dispid 125;
    procedure CloseUp; dispid 301;
    procedure OpenUp; dispid 302;
    procedure OpenOrCloseUp; dispid 303;
    procedure TabHangingIndent(Count: Smallint); dispid 304;
    procedure TabIndent(Count: Smallint); dispid 306;
    procedure Reset; dispid 312;
    procedure Space1; dispid 313;
    procedure Space15; dispid 314;
    procedure Space2; dispid 315;
    procedure IndentCharWidth(Count: Smallint); dispid 320;
    procedure IndentFirstLineCharWidth(Count: Smallint); dispid 322;
    function Next(Count: Integer): Paragraph; dispid 324;
    function Previous(Count: Integer): Paragraph; dispid 325;
    procedure Indent; dispid 333;
    procedure Outdent; dispid 334;
    property CharacterUnitRightIndent: Single dispid 126;
    property CharacterUnitLeftIndent: Single dispid 127;
    property CharacterUnitFirstLineIndent: Single dispid 128;
    property LineUnitBefore: Single dispid 129;
    property LineUnitAfter: Single dispid 130;
    property SpaceBeforeAuto: Integer dispid 132;
    property SpaceAfterAuto: Integer dispid 133;
    property FarEastLineBreakControl: Integer dispid 117;
    property WordWrap: Integer dispid 118;
    property OutlineLevel: WpsOutlineLevel dispid 202;
    procedure OutlineDemote; dispid 336;
    procedure OutlinePromote; dispid 337;
    procedure OutlineDemoteToBody; dispid 338;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: _ParagraphFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020953-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  _ParagraphFormat = interface(_IKsoDispObj)
    ['{00020953-0000-4B30-A977-D214852036FE}']
    function Get_Duplicate: _ParagraphFormat; safecall;
    function Get_Alignment: WpsParagraphAlignment; safecall;
    procedure Set_Alignment(prop: WpsParagraphAlignment); safecall;
    function Get_KeepTogether: Integer; safecall;
    procedure Set_KeepTogether(prop: Integer); safecall;
    function Get_KeepWithNext: Integer; safecall;
    procedure Set_KeepWithNext(prop: Integer); safecall;
    function Get_PageBreakBefore: Integer; safecall;
    procedure Set_PageBreakBefore(prop: Integer); safecall;
    function Get_RightIndent: Single; safecall;
    procedure Set_RightIndent(prop: Single); safecall;
    function Get_LeftIndent: Single; safecall;
    procedure Set_LeftIndent(prop: Single); safecall;
    function Get_FirstLineIndent: Single; safecall;
    procedure Set_FirstLineIndent(prop: Single); safecall;
    function Get_LineSpacing: Single; safecall;
    procedure Set_LineSpacing(prop: Single); safecall;
    function Get_LineSpacingRule: WpsLineSpacing; safecall;
    procedure Set_LineSpacingRule(prop: WpsLineSpacing); safecall;
    function Get_SpaceBefore: Single; safecall;
    procedure Set_SpaceBefore(prop: Single); safecall;
    function Get_SpaceAfter: Single; safecall;
    procedure Set_SpaceAfter(prop: Single); safecall;
    function Get_WidowControl: Integer; safecall;
    procedure Set_WidowControl(prop: Integer); safecall;
    function Get_HangingPunctuation: Integer; safecall;
    procedure Set_HangingPunctuation(prop: Integer); safecall;
    function Get_HalfWidthPunctuationOnTopOfLine: Integer; safecall;
    procedure Set_HalfWidthPunctuationOnTopOfLine(prop: Integer); safecall;
    function Get_AddSpaceBetweenFarEastAndAlpha: Integer; safecall;
    procedure Set_AddSpaceBetweenFarEastAndAlpha(prop: Integer); safecall;
    function Get_AddSpaceBetweenFarEastAndDigit: Integer; safecall;
    procedure Set_AddSpaceBetweenFarEastAndDigit(prop: Integer); safecall;
    function Get_BaseLineAlignment: WpsBaselineAlignment; safecall;
    procedure Set_BaseLineAlignment(prop: WpsBaselineAlignment); safecall;
    function Get_AutoAdjustRightIndent: Integer; safecall;
    procedure Set_AutoAdjustRightIndent(prop: Integer); safecall;
    function Get_DisableLineHeightGrid: Integer; safecall;
    procedure Set_DisableLineHeightGrid(prop: Integer); safecall;
    function Get_TabStops: TabStops; safecall;
    procedure Set_TabStops(const prop: TabStops); safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Shading: Shading; safecall;
    function Get_OutlineLevel: WpsOutlineLevel; safecall;
    procedure Set_OutlineLevel(prop: WpsOutlineLevel); safecall;
    procedure CloseUp; safecall;
    procedure OpenUp; safecall;
    procedure OpenOrCloseUp; safecall;
    procedure TabHangingIndent(Count: Smallint); safecall;
    procedure TabIndent(Count: Smallint); safecall;
    procedure Reset; safecall;
    procedure Space1; safecall;
    procedure Space15; safecall;
    procedure Space2; safecall;
    procedure IndentCharWidth(Count: Smallint); safecall;
    procedure IndentFirstLineCharWidth(Count: Smallint); safecall;
    function Get_CharacterUnitRightIndent: Single; safecall;
    procedure Set_CharacterUnitRightIndent(prop: Single); safecall;
    function Get_CharacterUnitLeftIndent: Single; safecall;
    procedure Set_CharacterUnitLeftIndent(prop: Single); safecall;
    function Get_CharacterUnitFirstLineIndent: Single; safecall;
    procedure Set_CharacterUnitFirstLineIndent(prop: Single); safecall;
    function Get_LineUnitBefore: Single; safecall;
    procedure Set_LineUnitBefore(prop: Single); safecall;
    function Get_LineUnitAfter: Single; safecall;
    procedure Set_LineUnitAfter(prop: Single); safecall;
    function Get_SpaceBeforeAuto: Integer; safecall;
    procedure Set_SpaceBeforeAuto(prop: Integer); safecall;
    function Get_SpaceAfterAuto: Integer; safecall;
    procedure Set_SpaceAfterAuto(prop: Integer); safecall;
    function Get_LineSpacingRelative: Single; safecall;
    procedure Set_LineSpacingRelative(prop: Single); safecall;
    function Get_FarEastLineBreakControl: Integer; safecall;
    procedure Set_FarEastLineBreakControl(prop: Integer); safecall;
    function Get_WordWrap: Integer; safecall;
    procedure Set_WordWrap(prop: Integer); safecall;
    function Get_ReadingOrder: WpsParagraphReadingOrder; safecall;
    procedure Set_ReadingOrder(prop: WpsParagraphReadingOrder); safecall;
    property Duplicate: _ParagraphFormat read Get_Duplicate;
    property Alignment: WpsParagraphAlignment read Get_Alignment write Set_Alignment;
    property KeepTogether: Integer read Get_KeepTogether write Set_KeepTogether;
    property KeepWithNext: Integer read Get_KeepWithNext write Set_KeepWithNext;
    property PageBreakBefore: Integer read Get_PageBreakBefore write Set_PageBreakBefore;
    property RightIndent: Single read Get_RightIndent write Set_RightIndent;
    property LeftIndent: Single read Get_LeftIndent write Set_LeftIndent;
    property FirstLineIndent: Single read Get_FirstLineIndent write Set_FirstLineIndent;
    property LineSpacing: Single read Get_LineSpacing write Set_LineSpacing;
    property LineSpacingRule: WpsLineSpacing read Get_LineSpacingRule write Set_LineSpacingRule;
    property SpaceBefore: Single read Get_SpaceBefore write Set_SpaceBefore;
    property SpaceAfter: Single read Get_SpaceAfter write Set_SpaceAfter;
    property WidowControl: Integer read Get_WidowControl write Set_WidowControl;
    property HangingPunctuation: Integer read Get_HangingPunctuation write Set_HangingPunctuation;
    property HalfWidthPunctuationOnTopOfLine: Integer read Get_HalfWidthPunctuationOnTopOfLine write Set_HalfWidthPunctuationOnTopOfLine;
    property AddSpaceBetweenFarEastAndAlpha: Integer read Get_AddSpaceBetweenFarEastAndAlpha write Set_AddSpaceBetweenFarEastAndAlpha;
    property AddSpaceBetweenFarEastAndDigit: Integer read Get_AddSpaceBetweenFarEastAndDigit write Set_AddSpaceBetweenFarEastAndDigit;
    property BaseLineAlignment: WpsBaselineAlignment read Get_BaseLineAlignment write Set_BaseLineAlignment;
    property AutoAdjustRightIndent: Integer read Get_AutoAdjustRightIndent write Set_AutoAdjustRightIndent;
    property DisableLineHeightGrid: Integer read Get_DisableLineHeightGrid write Set_DisableLineHeightGrid;
    property TabStops: TabStops read Get_TabStops write Set_TabStops;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Shading: Shading read Get_Shading;
    property OutlineLevel: WpsOutlineLevel read Get_OutlineLevel write Set_OutlineLevel;
    property CharacterUnitRightIndent: Single read Get_CharacterUnitRightIndent write Set_CharacterUnitRightIndent;
    property CharacterUnitLeftIndent: Single read Get_CharacterUnitLeftIndent write Set_CharacterUnitLeftIndent;
    property CharacterUnitFirstLineIndent: Single read Get_CharacterUnitFirstLineIndent write Set_CharacterUnitFirstLineIndent;
    property LineUnitBefore: Single read Get_LineUnitBefore write Set_LineUnitBefore;
    property LineUnitAfter: Single read Get_LineUnitAfter write Set_LineUnitAfter;
    property SpaceBeforeAuto: Integer read Get_SpaceBeforeAuto write Set_SpaceBeforeAuto;
    property SpaceAfterAuto: Integer read Get_SpaceAfterAuto write Set_SpaceAfterAuto;
    property LineSpacingRelative: Single read Get_LineSpacingRelative write Set_LineSpacingRelative;
    property FarEastLineBreakControl: Integer read Get_FarEastLineBreakControl write Set_FarEastLineBreakControl;
    property WordWrap: Integer read Get_WordWrap write Set_WordWrap;
    property ReadingOrder: WpsParagraphReadingOrder read Get_ReadingOrder write Set_ReadingOrder;
  end;

// *********************************************************************//
// DispIntf:  _ParagraphFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020953-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  _ParagraphFormatDisp = dispinterface
    ['{00020953-0000-4B30-A977-D214852036FE}']
    property Duplicate: _ParagraphFormat readonly dispid 10;
    property Alignment: WpsParagraphAlignment dispid 101;
    property KeepTogether: Integer dispid 102;
    property KeepWithNext: Integer dispid 103;
    property PageBreakBefore: Integer dispid 104;
    property RightIndent: Single dispid 106;
    property LeftIndent: Single dispid 107;
    property FirstLineIndent: Single dispid 108;
    property LineSpacing: Single dispid 109;
    property LineSpacingRule: WpsLineSpacing dispid 110;
    property SpaceBefore: Single dispid 111;
    property SpaceAfter: Single dispid 112;
    property WidowControl: Integer dispid 114;
    property HangingPunctuation: Integer dispid 119;
    property HalfWidthPunctuationOnTopOfLine: Integer dispid 120;
    property AddSpaceBetweenFarEastAndAlpha: Integer dispid 121;
    property AddSpaceBetweenFarEastAndDigit: Integer dispid 122;
    property BaseLineAlignment: WpsBaselineAlignment dispid 123;
    property AutoAdjustRightIndent: Integer dispid 124;
    property DisableLineHeightGrid: Integer dispid 125;
    property TabStops: TabStops dispid 1103;
    property Borders: Borders dispid 1100;
    property Shading: Shading readonly dispid 1101;
    property OutlineLevel: WpsOutlineLevel dispid 202;
    procedure CloseUp; dispid 301;
    procedure OpenUp; dispid 302;
    procedure OpenOrCloseUp; dispid 303;
    procedure TabHangingIndent(Count: Smallint); dispid 304;
    procedure TabIndent(Count: Smallint); dispid 306;
    procedure Reset; dispid 312;
    procedure Space1; dispid 313;
    procedure Space15; dispid 314;
    procedure Space2; dispid 315;
    procedure IndentCharWidth(Count: Smallint); dispid 320;
    procedure IndentFirstLineCharWidth(Count: Smallint); dispid 322;
    property CharacterUnitRightIndent: Single dispid 126;
    property CharacterUnitLeftIndent: Single dispid 127;
    property CharacterUnitFirstLineIndent: Single dispid 128;
    property LineUnitBefore: Single dispid 129;
    property LineUnitAfter: Single dispid 130;
    property SpaceBeforeAuto: Integer dispid 132;
    property SpaceAfterAuto: Integer dispid 133;
    property LineSpacingRelative: Single dispid 136;
    property FarEastLineBreakControl: Integer dispid 117;
    property WordWrap: Integer dispid 118;
    property ReadingOrder: WpsParagraphReadingOrder dispid 137;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: TabStops
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020955-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TabStops = interface(_IKsoDispObj)
    ['{00020955-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): TabStop; safecall;
    function Add(Position: Single; Alignment: WpsTabAlignment; Leader: WpsTabLeader): TabStop; safecall;
    procedure ClearAll; safecall;
    function Before(Position: Single): TabStop; safecall;
    function After(Position: Single): TabStop; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  TabStopsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020955-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TabStopsDisp = dispinterface
    ['{00020955-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    function Item(var Index: OleVariant): TabStop; dispid 0;
    function Add(Position: Single; Alignment: WpsTabAlignment; Leader: WpsTabLeader): TabStop; dispid 100;
    procedure ClearAll; dispid 101;
    function Before(Position: Single): TabStop; dispid 102;
    function After(Position: Single): TabStop; dispid 103;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: TabStop
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020954-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TabStop = interface(_IKsoDispObj)
    ['{00020954-0000-4B30-A977-D214852036FE}']
    function Get_Alignment: WpsTabAlignment; safecall;
    procedure Set_Alignment(prop: WpsTabAlignment); safecall;
    function Get_Leader: WpsTabLeader; safecall;
    procedure Set_Leader(prop: WpsTabLeader); safecall;
    function Get_Position: Single; safecall;
    procedure Set_Position(prop: Single); safecall;
    function Get_CustomTab: WordBool; safecall;
    function Get_Next: TabStop; safecall;
    function Get_Previous: TabStop; safecall;
    procedure Clear; safecall;
    property Alignment: WpsTabAlignment read Get_Alignment write Set_Alignment;
    property Leader: WpsTabLeader read Get_Leader write Set_Leader;
    property Position: Single read Get_Position write Set_Position;
    property CustomTab: WordBool read Get_CustomTab;
    property Next: TabStop read Get_Next;
    property Previous: TabStop read Get_Previous;
  end;

// *********************************************************************//
// DispIntf:  TabStopDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020954-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TabStopDisp = dispinterface
    ['{00020954-0000-4B30-A977-D214852036FE}']
    property Alignment: WpsTabAlignment dispid 100;
    property Leader: WpsTabLeader dispid 101;
    property Position: Single dispid 102;
    property CustomTab: WordBool readonly dispid 103;
    property Next: TabStop readonly dispid 104;
    property Previous: TabStop readonly dispid 105;
    procedure Clear; dispid 200;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: DropCap
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020956-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DropCap = interface(_IKsoDispObj)
    ['{00020956-0000-4B30-A977-D214852036FE}']
    function Get_Position: WpsDropPosition; safecall;
    procedure Set_Position(prop: WpsDropPosition); safecall;
    function Get_FontName: WideString; safecall;
    procedure Set_FontName(const prop: WideString); safecall;
    function Get_LinesToDrop: Integer; safecall;
    procedure Set_LinesToDrop(prop: Integer); safecall;
    function Get_DistanceFromText: Single; safecall;
    procedure Set_DistanceFromText(prop: Single); safecall;
    procedure Clear; safecall;
    procedure Enable; safecall;
    function Get_AsiaFont: WordBool; safecall;
    function Get_Range: Range; safecall;
    property Position: WpsDropPosition read Get_Position write Set_Position;
    property FontName: WideString read Get_FontName write Set_FontName;
    property LinesToDrop: Integer read Get_LinesToDrop write Set_LinesToDrop;
    property DistanceFromText: Single read Get_DistanceFromText write Set_DistanceFromText;
    property AsiaFont: WordBool read Get_AsiaFont;
    property Range: Range read Get_Range;
  end;

// *********************************************************************//
// DispIntf:  DropCapDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020956-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DropCapDisp = dispinterface
    ['{00020956-0000-4B30-A977-D214852036FE}']
    property Position: WpsDropPosition dispid 10;
    property FontName: WideString dispid 11;
    property LinesToDrop: Integer dispid 12;
    property DistanceFromText: Single dispid 13;
    procedure Clear; dispid 100;
    procedure Enable; dispid 101;
    property AsiaFont: WordBool readonly dispid 102;
    property Range: Range readonly dispid 103;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Style
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092C-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Style = interface(_IKsoDispObj)
    ['{0002092C-0000-4B30-A977-D214852036FE}']
    function Get_NameLocal: WideString; safecall;
    procedure Set_NameLocal(const prop: WideString); safecall;
    function Get_BaseStyle: OleVariant; safecall;
    procedure Set_BaseStyle(var prop: OleVariant); safecall;
    function Get_type_: WpsStyleType; safecall;
    function Get_BuiltIn: WordBool; safecall;
    function Get_NextParagraphStyle: OleVariant; safecall;
    procedure Set_NextParagraphStyle(var prop: OleVariant); safecall;
    function Get_InUse: WordBool; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_ParagraphFormat: _ParagraphFormat; safecall;
    procedure Set_ParagraphFormat(const prop: _ParagraphFormat); safecall;
    function Get_Font: _Font; safecall;
    procedure Set_Font(const prop: _Font); safecall;
    function Get_Frame: Frame; safecall;
    function Get_AutomaticallyUpdate: WordBool; safecall;
    procedure Set_AutomaticallyUpdate(prop: WordBool); safecall;
    function Get_ListTemplate: ListTemplate; safecall;
    function Get_ListLevelNumber: Integer; safecall;
    function Get_Hidden: WordBool; safecall;
    procedure Set_Hidden(prop: WordBool); safecall;
    procedure Delete; safecall;
    procedure LinkToListTemplate(const ListTemplate: ListTemplate; var ListLevelNumber: Integer); safecall;
    function Get_LinkStyle: OleVariant; safecall;
    procedure Set_LinkStyle(var prop: OleVariant); safecall;
    function Get_NoSpaceBetweenParagraphsOfSameStyle: WordBool; safecall;
    procedure Set_NoSpaceBetweenParagraphsOfSameStyle(prop: WordBool); safecall;
    function Duplicate: Style; safecall;
    procedure Apply; safecall;
    function Get_StyleId: Integer; safecall;
    function Get_ShortcutKey: Integer; safecall;
    procedure Set_ShortcutKey(prop: Integer); safecall;
    procedure SetAsTemplateDefault; safecall;
    property NameLocal: WideString read Get_NameLocal write Set_NameLocal;
    property type_: WpsStyleType read Get_type_;
    property BuiltIn: WordBool read Get_BuiltIn;
    property InUse: WordBool read Get_InUse;
    property Borders: Borders read Get_Borders write Set_Borders;
    property ParagraphFormat: _ParagraphFormat read Get_ParagraphFormat write Set_ParagraphFormat;
    property Font: _Font read Get_Font write Set_Font;
    property Frame: Frame read Get_Frame;
    property AutomaticallyUpdate: WordBool read Get_AutomaticallyUpdate write Set_AutomaticallyUpdate;
    property ListTemplate: ListTemplate read Get_ListTemplate;
    property ListLevelNumber: Integer read Get_ListLevelNumber;
    property Hidden: WordBool read Get_Hidden write Set_Hidden;
    property NoSpaceBetweenParagraphsOfSameStyle: WordBool read Get_NoSpaceBetweenParagraphsOfSameStyle write Set_NoSpaceBetweenParagraphsOfSameStyle;
    property StyleId: Integer read Get_StyleId;
    property ShortcutKey: Integer read Get_ShortcutKey write Set_ShortcutKey;
  end;

// *********************************************************************//
// DispIntf:  StyleDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092C-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  StyleDisp = dispinterface
    ['{0002092C-0000-4B30-A977-D214852036FE}']
    property NameLocal: WideString dispid 0;
    function BaseStyle: OleVariant; dispid 1;
    property type_: WpsStyleType readonly dispid 3;
    property BuiltIn: WordBool readonly dispid 4;
    function NextParagraphStyle: OleVariant; dispid 5;
    property InUse: WordBool readonly dispid 6;
    property Borders: Borders dispid 8;
    property ParagraphFormat: _ParagraphFormat dispid 9;
    property Font: _Font dispid 10;
    property Frame: Frame readonly dispid 11;
    property AutomaticallyUpdate: WordBool dispid 13;
    property ListTemplate: ListTemplate readonly dispid 14;
    property ListLevelNumber: Integer readonly dispid 15;
    property Hidden: WordBool dispid 17;
    procedure Delete; dispid 100;
    procedure LinkToListTemplate(const ListTemplate: ListTemplate; var ListLevelNumber: Integer); dispid 101;
    function LinkStyle: OleVariant; dispid 104;
    property NoSpaceBetweenParagraphsOfSameStyle: WordBool dispid 20;
    function Duplicate: Style; dispid 4097;
    procedure Apply; dispid 26214;
    property StyleId: Integer readonly dispid 65538;
    property ShortcutKey: Integer dispid 65539;
    procedure SetAsTemplateDefault; dispid 65540;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: ListTemplate
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListTemplate = interface(_IKsoDispObj)
    ['{0002098F-0000-4B30-A977-D214852036FE}']
    function Get_OutlineNumbered: WordBool; safecall;
    procedure Set_OutlineNumbered(prop: WordBool); safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    function Get_ListLevels: ListLevels; safecall;
    function Convert(Level: Integer): ListTemplate; safecall;
    property OutlineNumbered: WordBool read Get_OutlineNumbered write Set_OutlineNumbered;
    property Name: WideString read Get_Name write Set_Name;
    property ListLevels: ListLevels read Get_ListLevels;
  end;

// *********************************************************************//
// DispIntf:  ListTemplateDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListTemplateDisp = dispinterface
    ['{0002098F-0000-4B30-A977-D214852036FE}']
    property OutlineNumbered: WordBool dispid 1;
    property Name: WideString dispid 3;
    property ListLevels: ListLevels readonly dispid 2;
    function Convert(Level: Integer): ListTemplate; dispid 101;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: ListLevels
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098E-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListLevels = interface(_IKsoDispObj)
    ['{0002098E-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): ListLevel; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  ListLevelsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098E-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListLevelsDisp = dispinterface
    ['{0002098E-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    function Item(Index: Integer): ListLevel; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: ListLevel
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098D-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListLevel = interface(_IKsoDispObj)
    ['{0002098D-0000-4B30-A977-D214852036FE}']
    function Get_Index: Integer; safecall;
    function Get_NumberFormat: WideString; safecall;
    procedure Set_NumberFormat(const prop: WideString); safecall;
    function Get_TrailingCharacter: WpsTrailingCharacter; safecall;
    procedure Set_TrailingCharacter(prop: WpsTrailingCharacter); safecall;
    function Get_NumberStyle: WpsListNumberStyle; safecall;
    procedure Set_NumberStyle(prop: WpsListNumberStyle); safecall;
    function Get_NumberPosition: Single; safecall;
    procedure Set_NumberPosition(prop: Single); safecall;
    function Get_Alignment: WpsListLevelAlignment; safecall;
    procedure Set_Alignment(prop: WpsListLevelAlignment); safecall;
    function Get_TextPosition: Single; safecall;
    procedure Set_TextPosition(prop: Single); safecall;
    function Get_TabPosition: Single; safecall;
    procedure Set_TabPosition(prop: Single); safecall;
    function Get_ResetOnHigherOld: WordBool; safecall;
    procedure Set_ResetOnHigherOld(prop: WordBool); safecall;
    function Get_StartAt: Integer; safecall;
    procedure Set_StartAt(prop: Integer); safecall;
    function Get_LinkedStyle: WideString; safecall;
    procedure Set_LinkedStyle(const prop: WideString); safecall;
    function Get_Font: _Font; safecall;
    procedure Set_Font(const prop: _Font); safecall;
    function Get_ResetOnHigher: Integer; safecall;
    procedure Set_ResetOnHigher(prop: Integer); safecall;
    function Get_PictureBullet: InlineShape; safecall;
    function ApplyPictureBullet(const FileName: WideString): InlineShape; safecall;
    property Index: Integer read Get_Index;
    property NumberFormat: WideString read Get_NumberFormat write Set_NumberFormat;
    property TrailingCharacter: WpsTrailingCharacter read Get_TrailingCharacter write Set_TrailingCharacter;
    property NumberStyle: WpsListNumberStyle read Get_NumberStyle write Set_NumberStyle;
    property NumberPosition: Single read Get_NumberPosition write Set_NumberPosition;
    property Alignment: WpsListLevelAlignment read Get_Alignment write Set_Alignment;
    property TextPosition: Single read Get_TextPosition write Set_TextPosition;
    property TabPosition: Single read Get_TabPosition write Set_TabPosition;
    property ResetOnHigherOld: WordBool read Get_ResetOnHigherOld write Set_ResetOnHigherOld;
    property StartAt: Integer read Get_StartAt write Set_StartAt;
    property LinkedStyle: WideString read Get_LinkedStyle write Set_LinkedStyle;
    property Font: _Font read Get_Font write Set_Font;
    property ResetOnHigher: Integer read Get_ResetOnHigher write Set_ResetOnHigher;
    property PictureBullet: InlineShape read Get_PictureBullet;
  end;

// *********************************************************************//
// DispIntf:  ListLevelDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002098D-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListLevelDisp = dispinterface
    ['{0002098D-0000-4B30-A977-D214852036FE}']
    property Index: Integer readonly dispid 1;
    property NumberFormat: WideString dispid 2;
    property TrailingCharacter: WpsTrailingCharacter dispid 3;
    property NumberStyle: WpsListNumberStyle dispid 4;
    property NumberPosition: Single dispid 5;
    property Alignment: WpsListLevelAlignment dispid 6;
    property TextPosition: Single dispid 7;
    property TabPosition: Single dispid 8;
    property ResetOnHigherOld: WordBool dispid 9;
    property StartAt: Integer dispid 10;
    property LinkedStyle: WideString dispid 11;
    property Font: _Font dispid 12;
    property ResetOnHigher: Integer dispid 13;
    property PictureBullet: InlineShape readonly dispid 14;
    function ApplyPictureBullet(const FileName: WideString): InlineShape; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Fields
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020930-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Fields = interface(_IKsoDispObj)
    ['{00020930-0000-4B30-A977-D214852036FE}']
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): Field; safecall;
    function Add(const Range: Range; Type_: WpsFieldType; const Text: WideString; 
                 PreserveFormatting: WordBool): Field; safecall;
    procedure ToggleShowCodes; safecall;
    function Update: Integer; safecall;
    procedure Unlink; safecall;
    procedure UpdateSource; safecall;
    function Get_Locked: Integer; safecall;
    procedure Set_Locked(prop: Integer); safecall;
    function Get__NewEnum: IUnknown; safecall;
    property Count: Integer read Get_Count;
    property Locked: Integer read Get_Locked write Set_Locked;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  FieldsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020930-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FieldsDisp = dispinterface
    ['{00020930-0000-4B30-A977-D214852036FE}']
    property Count: Integer readonly dispid 1;
    function Item(Index: Integer): Field; dispid 0;
    function Add(const Range: Range; Type_: WpsFieldType; const Text: WideString; 
                 PreserveFormatting: WordBool): Field; dispid 105;
    procedure ToggleShowCodes; dispid 100;
    function Update: Integer; dispid 101;
    procedure Unlink; dispid 102;
    procedure UpdateSource; dispid 104;
    property Locked: Integer dispid 2;
    property _NewEnum: IUnknown readonly dispid -4;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Frames
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092B-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Frames = interface(_IKsoDispObj)
    ['{0002092B-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): Frame; safecall;
    function Add(const Range: Range): Frame; safecall;
    procedure Delete; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  FramesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092B-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FramesDisp = dispinterface
    ['{0002092B-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(Index: Integer): Frame; dispid 0;
    function Add(const Range: Range): Frame; dispid 100;
    procedure Delete; dispid 101;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: ListFormat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C0-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListFormat = interface(_IKsoDispObj)
    ['{000209C0-0000-4B30-A977-D214852036FE}']
    function Get_ListLevelNumber: Integer; safecall;
    procedure Set_ListLevelNumber(prop: Integer); safecall;
    function Get_List: List; safecall;
    function Get_ListTemplate: ListTemplate; safecall;
    function Get_ListValue: Integer; safecall;
    function Get_SingleList: WordBool; safecall;
    function Get_SingleListTemplate: WordBool; safecall;
    function Get_ListType: WpsListType; safecall;
    function Get_ListString: WideString; safecall;
    function CanContinuePreviousList(const ListTemplate: ListTemplate): WpsContinue; safecall;
    procedure RemoveNumbers(NumberType: WpsNumberType); safecall;
    procedure ConvertNumbersToText(NumberType: WpsNumberType); safecall;
    function CountNumberedItems(NumberType: WpsNumberType; Level: Integer): Integer; safecall;
    procedure ApplyBulletDefaultOld; safecall;
    procedure ApplyNumberDefaultOld; safecall;
    procedure ApplyOutlineNumberDefaultOld; safecall;
    procedure ListOutdent; safecall;
    procedure ListIndent; safecall;
    procedure ApplyListTemplate(const ListTemplate: ListTemplate; 
                                ContinuePreviousList: WpsContinue; ApplyTo: WpsListApplyTo; 
                                DefaultListBehavior: Integer); safecall;
    procedure RestartNumber(fRestart: WordBool); safecall;
    property ListLevelNumber: Integer read Get_ListLevelNumber write Set_ListLevelNumber;
    property List: List read Get_List;
    property ListTemplate: ListTemplate read Get_ListTemplate;
    property ListValue: Integer read Get_ListValue;
    property SingleList: WordBool read Get_SingleList;
    property SingleListTemplate: WordBool read Get_SingleListTemplate;
    property ListType: WpsListType read Get_ListType;
    property ListString: WideString read Get_ListString;
  end;

// *********************************************************************//
// DispIntf:  ListFormatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209C0-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListFormatDisp = dispinterface
    ['{000209C0-0000-4B30-A977-D214852036FE}']
    property ListLevelNumber: Integer dispid 68;
    property List: List readonly dispid 69;
    property ListTemplate: ListTemplate readonly dispid 70;
    property ListValue: Integer readonly dispid 71;
    property SingleList: WordBool readonly dispid 72;
    property SingleListTemplate: WordBool readonly dispid 73;
    property ListType: WpsListType readonly dispid 74;
    property ListString: WideString readonly dispid 75;
    function CanContinuePreviousList(const ListTemplate: ListTemplate): WpsContinue; dispid 184;
    procedure RemoveNumbers(NumberType: WpsNumberType); dispid 185;
    procedure ConvertNumbersToText(NumberType: WpsNumberType); dispid 186;
    function CountNumberedItems(NumberType: WpsNumberType; Level: Integer): Integer; dispid 187;
    procedure ApplyBulletDefaultOld; dispid 188;
    procedure ApplyNumberDefaultOld; dispid 189;
    procedure ApplyOutlineNumberDefaultOld; dispid 190;
    procedure ListOutdent; dispid 210;
    procedure ListIndent; dispid 211;
    procedure ApplyListTemplate(const ListTemplate: ListTemplate; 
                                ContinuePreviousList: WpsContinue; ApplyTo: WpsListApplyTo; 
                                DefaultListBehavior: Integer); dispid 215;
    procedure RestartNumber(fRestart: WordBool); dispid 77;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: List
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020992-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  List = interface(_IKsoDispObj)
    ['{00020992-0000-4B30-A977-D214852036FE}']
    function Get_Range: Range; safecall;
    function Get_ListParagraphs: ListParagraphs; safecall;
    function Get_SingleListTemplate: WordBool; safecall;
    procedure ConvertNumbersToText(var NumberType: WpsNumberType); safecall;
    procedure RemoveNumbers(NumberType: WpsNumberType); safecall;
    function CountNumberedItems(NumberType: WpsNumberType; Level: Integer): Integer; safecall;
    function CanContinuePreviousList(const ListTemplate: ListTemplate): WpsContinue; safecall;
    procedure ApplyListTemplate(const ListTemplate: ListTemplate; 
                                ContinuePreviousList: WpsContinue; DefaultListBehavior: Integer); safecall;
    function Get_StyleName: WideString; safecall;
    property Range: Range read Get_Range;
    property ListParagraphs: ListParagraphs read Get_ListParagraphs;
    property SingleListTemplate: WordBool read Get_SingleListTemplate;
    property StyleName: WideString read Get_StyleName;
  end;

// *********************************************************************//
// DispIntf:  ListDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020992-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListDisp = dispinterface
    ['{00020992-0000-4B30-A977-D214852036FE}']
    property Range: Range readonly dispid 1;
    property ListParagraphs: ListParagraphs readonly dispid 2;
    property SingleListTemplate: WordBool readonly dispid 3;
    procedure ConvertNumbersToText(var NumberType: WpsNumberType); dispid 101;
    procedure RemoveNumbers(NumberType: WpsNumberType); dispid 102;
    function CountNumberedItems(NumberType: WpsNumberType; Level: Integer): Integer; dispid 103;
    function CanContinuePreviousList(const ListTemplate: ListTemplate): WpsContinue; dispid 105;
    procedure ApplyListTemplate(const ListTemplate: ListTemplate; 
                                ContinuePreviousList: WpsContinue; DefaultListBehavior: Integer); dispid 106;
    property StyleName: WideString readonly dispid 4;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: ListParagraphs
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020991-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListParagraphs = interface(_IKsoDispObj)
    ['{00020991-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): Paragraph; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  ListParagraphsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020991-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListParagraphsDisp = dispinterface
    ['{00020991-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    function Item(Index: Integer): Paragraph; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Bookmarks
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020967-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Bookmarks = interface(_IKsoDispObj)
    ['{00020967-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_DefaultSorting: WpsBookmarkSortBy; safecall;
    procedure Set_DefaultSorting(prop: WpsBookmarkSortBy); safecall;
    function Get_ShowHidden: WordBool; safecall;
    procedure Set_ShowHidden(prop: WordBool); safecall;
    function Item(var Index: OleVariant): Bookmark; safecall;
    function Add(const Name: WideString; var Range: OleVariant): Bookmark; safecall;
    function Exists(const Name: WideString): WordBool; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property DefaultSorting: WpsBookmarkSortBy read Get_DefaultSorting write Set_DefaultSorting;
    property ShowHidden: WordBool read Get_ShowHidden write Set_ShowHidden;
  end;

// *********************************************************************//
// DispIntf:  BookmarksDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020967-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  BookmarksDisp = dispinterface
    ['{00020967-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    property DefaultSorting: WpsBookmarkSortBy dispid 3;
    property ShowHidden: WordBool dispid 4;
    function Item(var Index: OleVariant): Bookmark; dispid 0;
    function Add(const Name: WideString; var Range: OleVariant): Bookmark; dispid 5;
    function Exists(const Name: WideString): WordBool; dispid 6;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Bookmark
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020968-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Bookmark = interface(_IKsoDispObj)
    ['{00020968-0000-4B30-A977-D214852036FE}']
    function Get_Name: WideString; safecall;
    function Get_Range: Range; safecall;
    function Get_Empty: WordBool; safecall;
    function Get_Start: Integer; safecall;
    procedure Set_Start(prop: Integer); safecall;
    function Get_End_: Integer; safecall;
    procedure Set_End_(prop: Integer); safecall;
    function Get_Column: WordBool; safecall;
    function Get_StoryType: WpsStoryType; safecall;
    procedure Select; safecall;
    procedure Delete; safecall;
    function Copy(const Name: WideString): Bookmark; safecall;
    property Name: WideString read Get_Name;
    property Range: Range read Get_Range;
    property Empty: WordBool read Get_Empty;
    property Start: Integer read Get_Start write Set_Start;
    property End_: Integer read Get_End_ write Set_End_;
    property Column: WordBool read Get_Column;
    property StoryType: WpsStoryType read Get_StoryType;
  end;

// *********************************************************************//
// DispIntf:  BookmarkDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020968-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  BookmarkDisp = dispinterface
    ['{00020968-0000-4B30-A977-D214852036FE}']
    property Name: WideString readonly dispid 0;
    property Range: Range readonly dispid 1;
    property Empty: WordBool readonly dispid 2;
    property Start: Integer dispid 3;
    property End_: Integer dispid 4;
    property Column: WordBool readonly dispid 5;
    property StoryType: WpsStoryType readonly dispid 6;
    procedure Select; dispid 65535;
    procedure Delete; dispid 11;
    function Copy(const Name: WideString): Bookmark; dispid 12;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Hyperlinks
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099C-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Hyperlinks = interface(_IKsoDispObj)
    ['{0002099C-0000-4B30-A977-D214852036FE}']
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Item(var Index: OleVariant): Hyperlink; safecall;
    function Add(const Anchor: IDispatch; const Address: WideString; const SubAddress: WideString; 
                 const ScreenTip: WideString; const TextToDisplay: WideString; 
                 const Target: WideString): Hyperlink; safecall;
    function _Add(const Anchor: IDispatch; const Address: WideString; const SubAddress: WideString): Hyperlink; safecall;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  HyperlinksDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099C-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  HyperlinksDisp = dispinterface
    ['{0002099C-0000-4B30-A977-D214852036FE}']
    property Count: Integer readonly dispid 1;
    property _NewEnum: IUnknown readonly dispid -4;
    function Item(var Index: OleVariant): Hyperlink; dispid 0;
    function Add(const Anchor: IDispatch; const Address: WideString; const SubAddress: WideString; 
                 const ScreenTip: WideString; const TextToDisplay: WideString; 
                 const Target: WideString): Hyperlink; dispid 101;
    function _Add(const Anchor: IDispatch; const Address: WideString; const SubAddress: WideString): Hyperlink; dispid 100;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Hyperlink
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099D-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Hyperlink = interface(_IKsoDispObj)
    ['{0002099D-0000-4B30-A977-D214852036FE}']
    function Get_Name: WideString; safecall;
    function Get_Range: Range; safecall;
    procedure Delete; safecall;
    procedure AddToFavorites; safecall;
    procedure CreateNewDocument(const FileName: WideString; EditNow: WordBool; Overwrite: WordBool); safecall;
    function Get_Address: WideString; safecall;
    procedure Set_Address(const prop: WideString); safecall;
    function Get_SubAddress: WideString; safecall;
    procedure Set_SubAddress(const prop: WideString); safecall;
    function Get_EmailSubject: WideString; safecall;
    procedure Set_EmailSubject(const prop: WideString); safecall;
    function Get_ScreenTip: WideString; safecall;
    procedure Set_ScreenTip(const prop: WideString); safecall;
    function Get_TextToDisplay: WideString; safecall;
    procedure Set_TextToDisplay(const prop: WideString); safecall;
    function Get_Target: WideString; safecall;
    procedure Set_Target(const prop: WideString); safecall;
    function Get_Field: Field; safecall;
    property Name: WideString read Get_Name;
    property Range: Range read Get_Range;
    property Address: WideString read Get_Address write Set_Address;
    property SubAddress: WideString read Get_SubAddress write Set_SubAddress;
    property EmailSubject: WideString read Get_EmailSubject write Set_EmailSubject;
    property ScreenTip: WideString read Get_ScreenTip write Set_ScreenTip;
    property TextToDisplay: WideString read Get_TextToDisplay write Set_TextToDisplay;
    property Target: WideString read Get_Target write Set_Target;
    property Field: Field read Get_Field;
  end;

// *********************************************************************//
// DispIntf:  HyperlinkDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099D-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  HyperlinkDisp = dispinterface
    ['{0002099D-0000-4B30-A977-D214852036FE}']
    property Name: WideString readonly dispid 1003;
    property Range: Range readonly dispid 1006;
    procedure Delete; dispid 103;
    procedure AddToFavorites; dispid 105;
    procedure CreateNewDocument(const FileName: WideString; EditNow: WordBool; Overwrite: WordBool); dispid 106;
    property Address: WideString dispid 1100;
    property SubAddress: WideString dispid 1101;
    property EmailSubject: WideString dispid 1010;
    property ScreenTip: WideString dispid 1011;
    property TextToDisplay: WideString dispid 1012;
    property Target: WideString dispid 1013;
    property Field: Field readonly dispid 1014;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: InlineShapes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A9-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  InlineShapes = interface(_IKsoDispObj)
    ['{000209A9-0000-4B30-A977-D214852036FE}']
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Item(Index: OleVariant): InlineShape; safecall;
    function AddPicture(const FileName: WideString; var LinkToFile: OleVariant; 
                        var SaveWithDocument: OleVariant; var Range: OleVariant): InlineShape; safecall;
    function AddOLEObject(var ClassType: OleVariant; var FileName: OleVariant; 
                          var LinkToFile: OleVariant; var DisplayAsIcon: OleVariant; 
                          var IconFileName: OleVariant; var IconIndex: OleVariant; 
                          var IconLabel: OleVariant; var Range: OleVariant): InlineShape; safecall;
    function AddOLEControl(var ClassType: OleVariant; var Range: OleVariant): InlineShape; safecall;
    function New(const Range: Range): InlineShape; safecall;
    function AddHorizontalLine(const FileName: WideString; var Range: OleVariant): InlineShape; safecall;
    function AddHorizontalLineStandard(var Range: OleVariant): InlineShape; safecall;
    function AddPictureBullet(const FileName: WideString; var Range: OleVariant): InlineShape; safecall;
    function AddTextEffect(PresetTextEffect: KsoPresetTextEffect; const Text: WideString; 
                           const FontName: WideString; FontSize: Single; FontBold: KsoTriState; 
                           FontItalic: KsoTriState): InlineShape; safecall;
    function InsertImagerScan(var Range: OleVariant): InlineShape; safecall;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  InlineShapesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A9-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  InlineShapesDisp = dispinterface
    ['{000209A9-0000-4B30-A977-D214852036FE}']
    property Count: Integer readonly dispid 1;
    property _NewEnum: IUnknown readonly dispid -4;
    function Item(Index: OleVariant): InlineShape; dispid 0;
    function AddPicture(const FileName: WideString; var LinkToFile: OleVariant; 
                        var SaveWithDocument: OleVariant; var Range: OleVariant): InlineShape; dispid 100;
    function AddOLEObject(var ClassType: OleVariant; var FileName: OleVariant; 
                          var LinkToFile: OleVariant; var DisplayAsIcon: OleVariant; 
                          var IconFileName: OleVariant; var IconIndex: OleVariant; 
                          var IconLabel: OleVariant; var Range: OleVariant): InlineShape; dispid 24;
    function AddOLEControl(var ClassType: OleVariant; var Range: OleVariant): InlineShape; dispid 102;
    function New(const Range: Range): InlineShape; dispid 200;
    function AddHorizontalLine(const FileName: WideString; var Range: OleVariant): InlineShape; dispid 104;
    function AddHorizontalLineStandard(var Range: OleVariant): InlineShape; dispid 105;
    function AddPictureBullet(const FileName: WideString; var Range: OleVariant): InlineShape; dispid 106;
    function AddTextEffect(PresetTextEffect: KsoPresetTextEffect; const Text: WideString; 
                           const FontName: WideString; FontSize: Single; FontBold: KsoTriState; 
                           FontItalic: KsoTriState): InlineShape; dispid 4114;
    function InsertImagerScan(var Range: OleVariant): InlineShape; dispid 4115;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: WordStat
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000219A6-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  WordStat = interface(_IKsoDispObj)
    ['{000219A6-0000-4B30-A977-D214852036FE}']
    function Get_WordCount: Integer; safecall;
    function Get_CharacterCount(IncludeSpace: WordBool): Integer; safecall;
    function Get_ParagraphCount: Integer; safecall;
    function Get_LineCount: Integer; safecall;
    function Get_Phrase(AsiaPhrase: WordBool): Integer; safecall;
    procedure Set_IncludeNotes(CountNotes: WordBool); safecall;
    function Get_IncludeNotes: WordBool; safecall;
    function Get_PageCount: Integer; safecall;
    function Get_FarEastCharacterCount: Integer; safecall;
    property WordCount: Integer read Get_WordCount;
    property CharacterCount[IncludeSpace: WordBool]: Integer read Get_CharacterCount;
    property ParagraphCount: Integer read Get_ParagraphCount;
    property LineCount: Integer read Get_LineCount;
    property Phrase[AsiaPhrase: WordBool]: Integer read Get_Phrase;
    property IncludeNotes: WordBool read Get_IncludeNotes write Set_IncludeNotes;
    property PageCount: Integer read Get_PageCount;
    property FarEastCharacterCount: Integer read Get_FarEastCharacterCount;
  end;

// *********************************************************************//
// DispIntf:  WordStatDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000219A6-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  WordStatDisp = dispinterface
    ['{000219A6-0000-4B30-A977-D214852036FE}']
    property WordCount: Integer readonly dispid 0;
    property CharacterCount[IncludeSpace: WordBool]: Integer readonly dispid 1;
    property ParagraphCount: Integer readonly dispid 2;
    property LineCount: Integer readonly dispid 3;
    property Phrase[AsiaPhrase: WordBool]: Integer readonly dispid 4;
    property IncludeNotes: WordBool dispid 5;
    property PageCount: Integer readonly dispid 6;
    property FarEastCharacterCount: Integer readonly dispid 7;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Revisions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020980-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Revisions = interface(_IKsoDispObj)
    ['{00020980-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): Revision; safecall;
    procedure AcceptAll; safecall;
    procedure RejectAll; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  RevisionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020980-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  RevisionsDisp = dispinterface
    ['{00020980-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 5;
    function Item(Index: Integer): Revision; dispid 0;
    procedure AcceptAll; dispid 101;
    procedure RejectAll; dispid 102;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Revision
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020981-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Revision = interface(_IKsoDispObj)
    ['{00020981-0000-4B30-A977-D214852036FE}']
    function Get_Author: WideString; safecall;
    function Get_Date: TDateTime; safecall;
    function Get_Range: Range; safecall;
    function Get_type_: WpsRevisionType; safecall;
    function Get_Index: Integer; safecall;
    procedure Accept; safecall;
    procedure Reject; safecall;
    function Get_Style: Style; safecall;
    function Get_FormatDescription: WideString; safecall;
    property Author: WideString read Get_Author;
    property Date: TDateTime read Get_Date;
    property Range: Range read Get_Range;
    property type_: WpsRevisionType read Get_type_;
    property Index: Integer read Get_Index;
    property Style: Style read Get_Style;
    property FormatDescription: WideString read Get_FormatDescription;
  end;

// *********************************************************************//
// DispIntf:  RevisionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020981-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  RevisionDisp = dispinterface
    ['{00020981-0000-4B30-A977-D214852036FE}']
    property Author: WideString readonly dispid 1;
    property Date: TDateTime readonly dispid 2;
    property Range: Range readonly dispid 3;
    property type_: WpsRevisionType readonly dispid 4;
    property Index: Integer readonly dispid 5;
    procedure Accept; dispid 101;
    procedure Reject; dispid 102;
    property Style: Style readonly dispid 8;
    property FormatDescription: WideString readonly dispid 9;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: SpellingSuggestions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AA-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  SpellingSuggestions = interface(_IKsoDispObj)
    ['{000209AA-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): SpellingSuggestion; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  SpellingSuggestionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AA-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  SpellingSuggestionsDisp = dispinterface
    ['{000209AA-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(Index: Integer): SpellingSuggestion; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: SpellingSuggestion
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AB-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  SpellingSuggestion = interface(_IKsoDispObj)
    ['{000209AB-0000-4B30-A977-D214852036FE}']
    function Get_Name: WideString; safecall;
    property Name: WideString read Get_Name;
  end;

// *********************************************************************//
// DispIntf:  SpellingSuggestionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AB-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  SpellingSuggestionDisp = dispinterface
    ['{000209AB-0000-4B30-A977-D214852036FE}']
    property Name: WideString readonly dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: ProofreadingErrors
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209BB-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ProofreadingErrors = interface(_IKsoDispObj)
    ['{000209BB-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): Range; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  ProofreadingErrorsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209BB-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ProofreadingErrorsDisp = dispinterface
    ['{000209BB-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(Index: Integer): Range; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: FormFields
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020929-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FormFields = interface(_IKsoDispObj)
    ['{00020929-0000-4B30-A977-D214852036FE}']
    function Get_Count: Integer; safecall;
    function Get_Shaded: WordBool; safecall;
    procedure Set_Shaded(prop: WordBool); safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Item(var Index: OleVariant): FormField; safecall;
    function Add(const Range: Range; Type_: WpsFieldType): FormField; safecall;
    property Count: Integer read Get_Count;
    property Shaded: WordBool read Get_Shaded write Set_Shaded;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  FormFieldsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020929-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FormFieldsDisp = dispinterface
    ['{00020929-0000-4B30-A977-D214852036FE}']
    property Count: Integer readonly dispid 1;
    property Shaded: WordBool dispid 2;
    property _NewEnum: IUnknown readonly dispid -4;
    function Item(var Index: OleVariant): FormField; dispid 0;
    function Add(const Range: Range; Type_: WpsFieldType): FormField; dispid 101;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: FormField
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020928-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FormField = interface(_IKsoDispObj)
    ['{00020928-0000-4B30-A977-D214852036FE}']
    function Get_type_: WpsFieldType; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    function Get_EntryMacro: WideString; safecall;
    procedure Set_EntryMacro(const prop: WideString); safecall;
    function Get_ExitMacro: WideString; safecall;
    procedure Set_ExitMacro(const prop: WideString); safecall;
    function Get_OwnHelp: WordBool; safecall;
    procedure Set_OwnHelp(prop: WordBool); safecall;
    function Get_OwnStatus: WordBool; safecall;
    procedure Set_OwnStatus(prop: WordBool); safecall;
    function Get_HelpText: WideString; safecall;
    procedure Set_HelpText(const prop: WideString); safecall;
    function Get_StatusText: WideString; safecall;
    procedure Set_StatusText(const prop: WideString); safecall;
    function Get_Enabled: WordBool; safecall;
    procedure Set_Enabled(prop: WordBool); safecall;
    function Get_Result: WideString; safecall;
    procedure Set_Result(const prop: WideString); safecall;
    function Get_TextInput: TextInput; safecall;
    function Get_CheckBox: CheckBox; safecall;
    function Get_DropDown: DropDown; safecall;
    function Get_Next: FormField; safecall;
    function Get_Previous: FormField; safecall;
    function Get_CalculateOnExit: WordBool; safecall;
    procedure Set_CalculateOnExit(prop: WordBool); safecall;
    function Get_Range: Range; safecall;
    procedure Select; safecall;
    procedure Copy; safecall;
    procedure Cut; safecall;
    procedure Delete; safecall;
    property type_: WpsFieldType read Get_type_;
    property Name: WideString read Get_Name write Set_Name;
    property EntryMacro: WideString read Get_EntryMacro write Set_EntryMacro;
    property ExitMacro: WideString read Get_ExitMacro write Set_ExitMacro;
    property OwnHelp: WordBool read Get_OwnHelp write Set_OwnHelp;
    property OwnStatus: WordBool read Get_OwnStatus write Set_OwnStatus;
    property HelpText: WideString read Get_HelpText write Set_HelpText;
    property StatusText: WideString read Get_StatusText write Set_StatusText;
    property Enabled: WordBool read Get_Enabled write Set_Enabled;
    property Result: WideString read Get_Result write Set_Result;
    property TextInput: TextInput read Get_TextInput;
    property CheckBox: CheckBox read Get_CheckBox;
    property DropDown: DropDown read Get_DropDown;
    property Next: FormField read Get_Next;
    property Previous: FormField read Get_Previous;
    property CalculateOnExit: WordBool read Get_CalculateOnExit write Set_CalculateOnExit;
    property Range: Range read Get_Range;
  end;

// *********************************************************************//
// DispIntf:  FormFieldDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020928-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FormFieldDisp = dispinterface
    ['{00020928-0000-4B30-A977-D214852036FE}']
    property type_: WpsFieldType readonly dispid 0;
    property Name: WideString dispid 2;
    property EntryMacro: WideString dispid 3;
    property ExitMacro: WideString dispid 4;
    property OwnHelp: WordBool dispid 5;
    property OwnStatus: WordBool dispid 6;
    property HelpText: WideString dispid 7;
    property StatusText: WideString dispid 8;
    property Enabled: WordBool dispid 9;
    property Result: WideString dispid 10;
    property TextInput: TextInput readonly dispid 11;
    property CheckBox: CheckBox readonly dispid 12;
    property DropDown: DropDown readonly dispid 13;
    property Next: FormField readonly dispid 14;
    property Previous: FormField readonly dispid 15;
    property CalculateOnExit: WordBool dispid 16;
    property Range: Range readonly dispid 17;
    procedure Select; dispid 65535;
    procedure Copy; dispid 101;
    procedure Cut; dispid 102;
    procedure Delete; dispid 103;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: TextInput
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020927-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TextInput = interface(_IKsoDispObj)
    ['{00020927-0000-4B30-A977-D214852036FE}']
    function Get_Valid: WordBool; safecall;
    function Get_Default: WideString; safecall;
    procedure Set_Default(const prop: WideString); safecall;
    function Get_type_: WpsTextFormFieldType; safecall;
    function Get_Format: WideString; safecall;
    function Get_Width: Integer; safecall;
    procedure Set_Width(prop: Integer); safecall;
    procedure Clear; safecall;
    procedure EditType(Type_: WpsTextFormFieldType; var Default: OleVariant; 
                       var Format: OleVariant; var Enabled: OleVariant); safecall;
    function Get_Tip: WideString; safecall;
    procedure Set_Tip(const prop: WideString); safecall;
    property Valid: WordBool read Get_Valid;
    property Default: WideString read Get_Default write Set_Default;
    property type_: WpsTextFormFieldType read Get_type_;
    property Format: WideString read Get_Format;
    property Width: Integer read Get_Width write Set_Width;
    property Tip: WideString read Get_Tip write Set_Tip;
  end;

// *********************************************************************//
// DispIntf:  TextInputDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020927-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TextInputDisp = dispinterface
    ['{00020927-0000-4B30-A977-D214852036FE}']
    property Valid: WordBool readonly dispid 0;
    property Default: WideString dispid 1;
    property type_: WpsTextFormFieldType readonly dispid 2;
    property Format: WideString readonly dispid 3;
    property Width: Integer dispid 4;
    procedure Clear; dispid 101;
    procedure EditType(Type_: WpsTextFormFieldType; var Default: OleVariant; 
                       var Format: OleVariant; var Enabled: OleVariant); dispid 102;
    property Tip: WideString dispid 104;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: CheckBox
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020926-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  CheckBox = interface(_IKsoDispObj)
    ['{00020926-0000-4B30-A977-D214852036FE}']
    function Get_Valid: WordBool; safecall;
    function Get_AutoSize: WordBool; safecall;
    procedure Set_AutoSize(prop: WordBool); safecall;
    function Get_Size: Single; safecall;
    procedure Set_Size(prop: Single); safecall;
    function Get_Default: WordBool; safecall;
    procedure Set_Default(prop: WordBool); safecall;
    function Get_Value: WordBool; safecall;
    procedure Set_Value(prop: WordBool); safecall;
    property Valid: WordBool read Get_Valid;
    property AutoSize: WordBool read Get_AutoSize write Set_AutoSize;
    property Size: Single read Get_Size write Set_Size;
    property Default: WordBool read Get_Default write Set_Default;
    property Value: WordBool read Get_Value write Set_Value;
  end;

// *********************************************************************//
// DispIntf:  CheckBoxDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020926-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  CheckBoxDisp = dispinterface
    ['{00020926-0000-4B30-A977-D214852036FE}']
    property Valid: WordBool readonly dispid 0;
    property AutoSize: WordBool dispid 1;
    property Size: Single dispid 2;
    property Default: WordBool dispid 3;
    property Value: WordBool dispid 4;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: DropDown
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020925-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DropDown = interface(_IKsoDispObj)
    ['{00020925-0000-4B30-A977-D214852036FE}']
    function Get_Valid: WordBool; safecall;
    function Get_Default: Integer; safecall;
    procedure Set_Default(prop: Integer); safecall;
    function Get_Value: Integer; safecall;
    procedure Set_Value(prop: Integer); safecall;
    function Get_ListEntries: ListEntries; safecall;
    property Valid: WordBool read Get_Valid;
    property Default: Integer read Get_Default write Set_Default;
    property Value: Integer read Get_Value write Set_Value;
    property ListEntries: ListEntries read Get_ListEntries;
  end;

// *********************************************************************//
// DispIntf:  DropDownDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020925-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DropDownDisp = dispinterface
    ['{00020925-0000-4B30-A977-D214852036FE}']
    property Valid: WordBool readonly dispid 0;
    property Default: Integer dispid 1;
    property Value: Integer dispid 2;
    property ListEntries: ListEntries readonly dispid 3;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: ListEntries
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020924-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListEntries = interface(_IKsoDispObj)
    ['{00020924-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): ListEntry; safecall;
    function Add(const Name: WideString; var Index: OleVariant): ListEntry; safecall;
    procedure Clear; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  ListEntriesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020924-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListEntriesDisp = dispinterface
    ['{00020924-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): ListEntry; dispid 0;
    function Add(const Name: WideString; var Index: OleVariant): ListEntry; dispid 101;
    procedure Clear; dispid 102;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: ListEntry
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020923-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListEntry = interface(_IKsoDispObj)
    ['{00020923-0000-4B30-A977-D214852036FE}']
    function Get_Index: Integer; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    procedure Delete; safecall;
    property Index: Integer read Get_Index;
    property Name: WideString read Get_Name write Set_Name;
  end;

// *********************************************************************//
// DispIntf:  ListEntryDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020923-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListEntryDisp = dispinterface
    ['{00020923-0000-4B30-A977-D214852036FE}']
    property Index: Integer readonly dispid 1;
    property Name: WideString dispid 2;
    procedure Delete; dispid 11;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Editors
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002E08C-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Editors = interface(_IKsoDispObj)
    ['{0002E08C-0000-4B30-A977-D214852036FE}']
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): Editor; safecall;
    function Add(var EditorID: OleVariant): Editor; safecall;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  EditorsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002E08C-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  EditorsDisp = dispinterface
    ['{0002E08C-0000-4B30-A977-D214852036FE}']
    property Count: Integer readonly dispid 2;
    function Item(var Index: OleVariant): Editor; dispid 0;
    function Add(var EditorID: OleVariant): Editor; dispid 501;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Editor
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00027D72-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Editor = interface(_IKsoDispObj)
    ['{00027D72-0000-4B30-A977-D214852036FE}']
    function Get_ID: WideString; safecall;
    function Get_Name: WideString; safecall;
    function Get_Range: Range; safecall;
    function Get_NextRange: Range; safecall;
    procedure Delete; safecall;
    procedure DeleteAll; safecall;
    procedure SelectAll; safecall;
    property ID: WideString read Get_ID;
    property Name: WideString read Get_Name;
    property Range: Range read Get_Range;
    property NextRange: Range read Get_NextRange;
  end;

// *********************************************************************//
// DispIntf:  EditorDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00027D72-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  EditorDisp = dispinterface
    ['{00027D72-0000-4B30-A977-D214852036FE}']
    property ID: WideString readonly dispid 100;
    property Name: WideString readonly dispid 101;
    property Range: Range readonly dispid 102;
    property NextRange: Range readonly dispid 103;
    procedure Delete; dispid 500;
    procedure DeleteAll; dispid 501;
    procedure SelectAll; dispid 502;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Styles
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092D-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Styles = interface(_IKsoDispObj)
    ['{0002092D-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): Style; safecall;
    function Add(const Name: WideString; Type_: WpsStyleType): Style; safecall;
    function Get_InUseCount: Integer; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property InUseCount: Integer read Get_InUseCount;
  end;

// *********************************************************************//
// DispIntf:  StylesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092D-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  StylesDisp = dispinterface
    ['{0002092D-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): Style; dispid 0;
    function Add(const Name: WideString; Type_: WpsStyleType): Style; dispid 100;
    property InUseCount: Integer readonly dispid 4097;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: TablesOfFigures
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020922-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TablesOfFigures = interface(_IKsoDispObj)
    ['{00020922-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): TableOfFigures; safecall;
    function Add(const Range: Range; var Caption: OleVariant; var IncludeLabel: OleVariant; 
                 var UseHeadingStyles: OleVariant; var UpperHeadingLevel: OleVariant; 
                 var LowerHeadingLevel: OleVariant; var UseFields: OleVariant; 
                 var TableID: OleVariant; var RightAlignPageNumbers: OleVariant; 
                 var IncludePageNumbers: OleVariant; var AddedStyles: OleVariant; 
                 var UseHyperlinks: OleVariant; var HidePageNumbersInWeb: OleVariant): TableOfFigures; safecall;
    function Get_Format: WpsTofFormat; safecall;
    procedure Set_Format(prop: WpsTofFormat); safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Format: WpsTofFormat read Get_Format write Set_Format;
  end;

// *********************************************************************//
// DispIntf:  TablesOfFiguresDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020922-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TablesOfFiguresDisp = dispinterface
    ['{00020922-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(Index: Integer): TableOfFigures; dispid 0;
    function Add(const Range: Range; var Caption: OleVariant; var IncludeLabel: OleVariant; 
                 var UseHeadingStyles: OleVariant; var UpperHeadingLevel: OleVariant; 
                 var LowerHeadingLevel: OleVariant; var UseFields: OleVariant; 
                 var TableID: OleVariant; var RightAlignPageNumbers: OleVariant; 
                 var IncludePageNumbers: OleVariant; var AddedStyles: OleVariant; 
                 var UseHyperlinks: OleVariant; var HidePageNumbersInWeb: OleVariant): TableOfFigures; dispid 444;
    property Format: WpsTofFormat dispid 2;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: TableOfFigures
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020921-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TableOfFigures = interface(_IKsoDispObj)
    ['{00020921-0000-4B30-A977-D214852036FE}']
    function Get_Caption: WideString; safecall;
    procedure Set_Caption(const prop: WideString); safecall;
    function Get_IncludeLabel: WordBool; safecall;
    procedure Set_IncludeLabel(prop: WordBool); safecall;
    function Get_RightAlignPageNumbers: WordBool; safecall;
    procedure Set_RightAlignPageNumbers(prop: WordBool); safecall;
    function Get_IncludePageNumbers: WordBool; safecall;
    procedure Set_IncludePageNumbers(prop: WordBool); safecall;
    function Get_Range: Range; safecall;
    function Get_UseFields: WordBool; safecall;
    procedure Set_UseFields(prop: WordBool); safecall;
    function Get_TableID: WideString; safecall;
    procedure Set_TableID(const prop: WideString); safecall;
    function Get_TabLeader: WpsTabLeader; safecall;
    procedure Set_TabLeader(prop: WpsTabLeader); safecall;
    procedure Delete; safecall;
    procedure UpdatePageNumbers; safecall;
    procedure Update; safecall;
    function Get_UseHyperlinks: WordBool; safecall;
    procedure Set_UseHyperlinks(prop: WordBool); safecall;
    property Caption: WideString read Get_Caption write Set_Caption;
    property IncludeLabel: WordBool read Get_IncludeLabel write Set_IncludeLabel;
    property RightAlignPageNumbers: WordBool read Get_RightAlignPageNumbers write Set_RightAlignPageNumbers;
    property IncludePageNumbers: WordBool read Get_IncludePageNumbers write Set_IncludePageNumbers;
    property Range: Range read Get_Range;
    property UseFields: WordBool read Get_UseFields write Set_UseFields;
    property TableID: WideString read Get_TableID write Set_TableID;
    property TabLeader: WpsTabLeader read Get_TabLeader write Set_TabLeader;
    property UseHyperlinks: WordBool read Get_UseHyperlinks write Set_UseHyperlinks;
  end;

// *********************************************************************//
// DispIntf:  TableOfFiguresDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020921-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TableOfFiguresDisp = dispinterface
    ['{00020921-0000-4B30-A977-D214852036FE}']
    property Caption: WideString dispid 1;
    property IncludeLabel: WordBool dispid 2;
    property RightAlignPageNumbers: WordBool dispid 3;
    property IncludePageNumbers: WordBool dispid 7;
    property Range: Range readonly dispid 8;
    property UseFields: WordBool dispid 9;
    property TableID: WideString dispid 10;
    property TabLeader: WpsTabLeader dispid 12;
    procedure Delete; dispid 100;
    procedure UpdatePageNumbers; dispid 101;
    procedure Update; dispid 102;
    property UseHyperlinks: WordBool dispid 13;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: MailMerge
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020920-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMerge = interface(_IKsoDispObj)
    ['{00020920-0000-4B30-A977-D214852036FE}']
    function Get_MainDocumentType: WpsMailMergeMainDocType; safecall;
    procedure Set_MainDocumentType(prop: WpsMailMergeMainDocType); safecall;
    function Get_State: WpsMailMergeState; safecall;
    function Get_Destination: WpsMailMergeDestination; safecall;
    procedure Set_Destination(prop: WpsMailMergeDestination); safecall;
    function Get_DataSource: MailMergeDataSource; safecall;
    function Get_Fields: MailMergeFields; safecall;
    function Get_ViewMailMergeFieldCodes: Integer; safecall;
    procedure Set_ViewMailMergeFieldCodes(prop: Integer); safecall;
    function Get_SuppressBlankLines: WordBool; safecall;
    procedure Set_SuppressBlankLines(prop: WordBool); safecall;
    function Get_MailAsAttachment: WordBool; safecall;
    procedure Set_MailAsAttachment(prop: WordBool); safecall;
    function Get_MailAddressFieldName: WideString; safecall;
    procedure Set_MailAddressFieldName(const prop: WideString); safecall;
    function Get_MailSubject: WideString; safecall;
    procedure Set_MailSubject(const prop: WideString); safecall;
    procedure CreateDataSource(var Name: OleVariant; var PasswordDocument: OleVariant; 
                               var WritePasswordDocument: OleVariant; var HeaderRecord: OleVariant; 
                               var MSQuery: OleVariant; var SQLStatement: OleVariant; 
                               var SQLStatement1: OleVariant; var Connection: OleVariant; 
                               var LinkToSource: OleVariant); safecall;
    procedure CreateHeaderSource(const Name: WideString; var PasswordDocument: OleVariant; 
                                 var WritePasswordDocument: OleVariant; var HeaderRecord: OleVariant); safecall;
    procedure OpenDataSource2000(const Name: WideString; var Format: OleVariant; 
                                 var ConfirmConversions: OleVariant; var ReadOnly: OleVariant; 
                                 var LinkToSource: OleVariant; var AddToRecentFiles: OleVariant; 
                                 var PasswordDocument: OleVariant; 
                                 var PasswordTemplate: OleVariant; var Revert: OleVariant; 
                                 var WritePasswordDocument: OleVariant; 
                                 var WritePasswordTemplate: OleVariant; var Connection: OleVariant; 
                                 var SQLStatement: OleVariant; var SQLStatement1: OleVariant); safecall;
    procedure OpenHeaderSource2000(const Name: WideString; var Format: OleVariant; 
                                   var ConfirmConversions: OleVariant; var ReadOnly: OleVariant; 
                                   var AddToRecentFiles: OleVariant; 
                                   var PasswordDocument: OleVariant; 
                                   var PasswordTemplate: OleVariant; var Revert: OleVariant; 
                                   var WritePasswordDocument: OleVariant; 
                                   var WritePasswordTemplate: OleVariant); safecall;
    procedure Execute(var Pause: OleVariant); safecall;
    procedure Check; safecall;
    procedure EditDataSource; safecall;
    procedure EditHeaderSource; safecall;
    procedure EditMainDocument; safecall;
    procedure UseAddressBook(const Type_: WideString); safecall;
    function Get_HighlightMergeFields: WordBool; safecall;
    procedure Set_HighlightMergeFields(prop: WordBool); safecall;
    function Get_MailFormat: WpsMailMergeMailFormat; safecall;
    procedure Set_MailFormat(prop: WpsMailMergeMailFormat); safecall;
    function Get_ShowSendToCustom: WideString; safecall;
    procedure Set_ShowSendToCustom(const prop: WideString); safecall;
    function Get_WizardState: Integer; safecall;
    procedure Set_WizardState(prop: Integer); safecall;
    procedure OpenDataSource(const Name: WideString; var Format: OleVariant; 
                             var ConfirmConversions: OleVariant; var ReadOnly: OleVariant; 
                             var LinkToSource: OleVariant; var AddToRecentFiles: OleVariant; 
                             var PasswordDocument: OleVariant; var PasswordTemplate: OleVariant; 
                             var Revert: OleVariant; var WritePasswordDocument: OleVariant; 
                             var WritePasswordTemplate: OleVariant; var Connection: OleVariant; 
                             var SQLStatement: OleVariant; var SQLStatement1: OleVariant; 
                             var OpenExclusive: OleVariant; var SubType: OleVariant); safecall;
    procedure OpenHeaderSource(const Name: WideString; var Format: OleVariant; 
                               var ConfirmConversions: OleVariant; var ReadOnly: OleVariant; 
                               var AddToRecentFiles: OleVariant; var PasswordDocument: OleVariant; 
                               var PasswordTemplate: OleVariant; var Revert: OleVariant; 
                               var WritePasswordDocument: OleVariant; 
                               var WritePasswordTemplate: OleVariant; var OpenExclusive: OleVariant); safecall;
    procedure ShowWizard(var InitialState: OleVariant; var ShoWpsocumentStep: OleVariant; 
                         var ShowTemplateStep: OleVariant; var ShoWpsataStep: OleVariant; 
                         var ShowWriteStep: OleVariant; var ShowPreviewStep: OleVariant; 
                         var ShowMergeStep: OleVariant); safecall;
    function Get_MergeFileNameDataField: WideString; safecall;
    procedure Set_MergeFileNameDataField(const prop: WideString); safecall;
    function Get_MergeFileDirectory: WideString; safecall;
    procedure Set_MergeFileDirectory(const prop: WideString); safecall;
    function Get_MergeFileFormat: WideString; safecall;
    procedure Set_MergeFileFormat(const prop: WideString); safecall;
    property MainDocumentType: WpsMailMergeMainDocType read Get_MainDocumentType write Set_MainDocumentType;
    property State: WpsMailMergeState read Get_State;
    property Destination: WpsMailMergeDestination read Get_Destination write Set_Destination;
    property DataSource: MailMergeDataSource read Get_DataSource;
    property Fields: MailMergeFields read Get_Fields;
    property ViewMailMergeFieldCodes: Integer read Get_ViewMailMergeFieldCodes write Set_ViewMailMergeFieldCodes;
    property SuppressBlankLines: WordBool read Get_SuppressBlankLines write Set_SuppressBlankLines;
    property MailAsAttachment: WordBool read Get_MailAsAttachment write Set_MailAsAttachment;
    property MailAddressFieldName: WideString read Get_MailAddressFieldName write Set_MailAddressFieldName;
    property MailSubject: WideString read Get_MailSubject write Set_MailSubject;
    property HighlightMergeFields: WordBool read Get_HighlightMergeFields write Set_HighlightMergeFields;
    property MailFormat: WpsMailMergeMailFormat read Get_MailFormat write Set_MailFormat;
    property ShowSendToCustom: WideString read Get_ShowSendToCustom write Set_ShowSendToCustom;
    property WizardState: Integer read Get_WizardState write Set_WizardState;
    property MergeFileNameDataField: WideString read Get_MergeFileNameDataField write Set_MergeFileNameDataField;
    property MergeFileDirectory: WideString read Get_MergeFileDirectory write Set_MergeFileDirectory;
    property MergeFileFormat: WideString read Get_MergeFileFormat write Set_MergeFileFormat;
  end;

// *********************************************************************//
// DispIntf:  MailMergeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020920-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMergeDisp = dispinterface
    ['{00020920-0000-4B30-A977-D214852036FE}']
    property MainDocumentType: WpsMailMergeMainDocType dispid 1;
    property State: WpsMailMergeState readonly dispid 2;
    property Destination: WpsMailMergeDestination dispid 3;
    property DataSource: MailMergeDataSource readonly dispid 4;
    property Fields: MailMergeFields readonly dispid 5;
    property ViewMailMergeFieldCodes: Integer dispid 6;
    property SuppressBlankLines: WordBool dispid 7;
    property MailAsAttachment: WordBool dispid 8;
    property MailAddressFieldName: WideString dispid 9;
    property MailSubject: WideString dispid 10;
    procedure CreateDataSource(var Name: OleVariant; var PasswordDocument: OleVariant; 
                               var WritePasswordDocument: OleVariant; var HeaderRecord: OleVariant; 
                               var MSQuery: OleVariant; var SQLStatement: OleVariant; 
                               var SQLStatement1: OleVariant; var Connection: OleVariant; 
                               var LinkToSource: OleVariant); dispid 101;
    procedure CreateHeaderSource(const Name: WideString; var PasswordDocument: OleVariant; 
                                 var WritePasswordDocument: OleVariant; var HeaderRecord: OleVariant); dispid 102;
    procedure OpenDataSource2000(const Name: WideString; var Format: OleVariant; 
                                 var ConfirmConversions: OleVariant; var ReadOnly: OleVariant; 
                                 var LinkToSource: OleVariant; var AddToRecentFiles: OleVariant; 
                                 var PasswordDocument: OleVariant; 
                                 var PasswordTemplate: OleVariant; var Revert: OleVariant; 
                                 var WritePasswordDocument: OleVariant; 
                                 var WritePasswordTemplate: OleVariant; var Connection: OleVariant; 
                                 var SQLStatement: OleVariant; var SQLStatement1: OleVariant); dispid 103;
    procedure OpenHeaderSource2000(const Name: WideString; var Format: OleVariant; 
                                   var ConfirmConversions: OleVariant; var ReadOnly: OleVariant; 
                                   var AddToRecentFiles: OleVariant; 
                                   var PasswordDocument: OleVariant; 
                                   var PasswordTemplate: OleVariant; var Revert: OleVariant; 
                                   var WritePasswordDocument: OleVariant; 
                                   var WritePasswordTemplate: OleVariant); dispid 104;
    procedure Execute(var Pause: OleVariant); dispid 105;
    procedure Check; dispid 106;
    procedure EditDataSource; dispid 107;
    procedure EditHeaderSource; dispid 108;
    procedure EditMainDocument; dispid 109;
    procedure UseAddressBook(const Type_: WideString); dispid 111;
    property HighlightMergeFields: WordBool dispid 11;
    property MailFormat: WpsMailMergeMailFormat dispid 12;
    property ShowSendToCustom: WideString dispid 13;
    property WizardState: Integer dispid 14;
    procedure OpenDataSource(const Name: WideString; var Format: OleVariant; 
                             var ConfirmConversions: OleVariant; var ReadOnly: OleVariant; 
                             var LinkToSource: OleVariant; var AddToRecentFiles: OleVariant; 
                             var PasswordDocument: OleVariant; var PasswordTemplate: OleVariant; 
                             var Revert: OleVariant; var WritePasswordDocument: OleVariant; 
                             var WritePasswordTemplate: OleVariant; var Connection: OleVariant; 
                             var SQLStatement: OleVariant; var SQLStatement1: OleVariant; 
                             var OpenExclusive: OleVariant; var SubType: OleVariant); dispid 112;
    procedure OpenHeaderSource(const Name: WideString; var Format: OleVariant; 
                               var ConfirmConversions: OleVariant; var ReadOnly: OleVariant; 
                               var AddToRecentFiles: OleVariant; var PasswordDocument: OleVariant; 
                               var PasswordTemplate: OleVariant; var Revert: OleVariant; 
                               var WritePasswordDocument: OleVariant; 
                               var WritePasswordTemplate: OleVariant; var OpenExclusive: OleVariant); dispid 113;
    procedure ShowWizard(var InitialState: OleVariant; var ShoWpsocumentStep: OleVariant; 
                         var ShowTemplateStep: OleVariant; var ShoWpsataStep: OleVariant; 
                         var ShowWriteStep: OleVariant; var ShowPreviewStep: OleVariant; 
                         var ShowMergeStep: OleVariant); dispid 114;
    property MergeFileNameDataField: WideString dispid 115;
    property MergeFileDirectory: WideString dispid 117;
    property MergeFileFormat: WideString dispid 119;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: MailMergeDataSource
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091D-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMergeDataSource = interface(_IKsoDispObj)
    ['{0002091D-0000-4B30-A977-D214852036FE}']
    function Get_Name: WideString; safecall;
    function Get_HeaderSourceName: WideString; safecall;
    function Get_type_: WpsMailMergeDataSource; safecall;
    function Get_HeaderSourceType: WpsMailMergeDataSource; safecall;
    function Get_ConnectString: WideString; safecall;
    function Get_QueryString: WideString; safecall;
    procedure Set_QueryString(const prop: WideString); safecall;
    function Get_ActiveRecord: WpsMailMergeActiveRecord; safecall;
    procedure Set_ActiveRecord(prop: WpsMailMergeActiveRecord); safecall;
    function Get_FirstRecord: Integer; safecall;
    procedure Set_FirstRecord(prop: Integer); safecall;
    function Get_LastRecord: Integer; safecall;
    procedure Set_LastRecord(prop: Integer); safecall;
    function Get_FieldNames: MailMergeFieldNames; safecall;
    function Get_DataFields: MailMergeDataFields; safecall;
    function FindRecord2000(const FindText: WideString; const Field: WideString): WordBool; safecall;
    function Get_RecordCount: Integer; safecall;
    function Get_Included: WordBool; safecall;
    procedure Set_Included(prop: WordBool); safecall;
    function Get_InvalidAddress: WordBool; safecall;
    procedure Set_InvalidAddress(prop: WordBool); safecall;
    function Get_InvalidComments: WideString; safecall;
    procedure Set_InvalidComments(const prop: WideString); safecall;
    function Get_MappedDataFields: MappedDataFields; safecall;
    function Get_TableName: WideString; safecall;
    function FindRecord(const FindText: WideString; var Field: OleVariant): WordBool; safecall;
    procedure SetAllIncludedFlags(Included: WordBool); safecall;
    procedure SetAllErrorFlags(Invalid: WordBool; const InvalidComment: WideString); safecall;
    procedure Close; safecall;
    property Name: WideString read Get_Name;
    property HeaderSourceName: WideString read Get_HeaderSourceName;
    property type_: WpsMailMergeDataSource read Get_type_;
    property HeaderSourceType: WpsMailMergeDataSource read Get_HeaderSourceType;
    property ConnectString: WideString read Get_ConnectString;
    property QueryString: WideString read Get_QueryString write Set_QueryString;
    property ActiveRecord: WpsMailMergeActiveRecord read Get_ActiveRecord write Set_ActiveRecord;
    property FirstRecord: Integer read Get_FirstRecord write Set_FirstRecord;
    property LastRecord: Integer read Get_LastRecord write Set_LastRecord;
    property FieldNames: MailMergeFieldNames read Get_FieldNames;
    property DataFields: MailMergeDataFields read Get_DataFields;
    property RecordCount: Integer read Get_RecordCount;
    property Included: WordBool read Get_Included write Set_Included;
    property InvalidAddress: WordBool read Get_InvalidAddress write Set_InvalidAddress;
    property InvalidComments: WideString read Get_InvalidComments write Set_InvalidComments;
    property MappedDataFields: MappedDataFields read Get_MappedDataFields;
    property TableName: WideString read Get_TableName;
  end;

// *********************************************************************//
// DispIntf:  MailMergeDataSourceDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091D-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMergeDataSourceDisp = dispinterface
    ['{0002091D-0000-4B30-A977-D214852036FE}']
    property Name: WideString readonly dispid 1;
    property HeaderSourceName: WideString readonly dispid 2;
    property type_: WpsMailMergeDataSource readonly dispid 3;
    property HeaderSourceType: WpsMailMergeDataSource readonly dispid 4;
    property ConnectString: WideString readonly dispid 5;
    property QueryString: WideString dispid 6;
    property ActiveRecord: WpsMailMergeActiveRecord dispid 7;
    property FirstRecord: Integer dispid 8;
    property LastRecord: Integer dispid 9;
    property FieldNames: MailMergeFieldNames readonly dispid 10;
    property DataFields: MailMergeDataFields readonly dispid 11;
    function FindRecord2000(const FindText: WideString; const Field: WideString): WordBool; dispid 101;
    property RecordCount: Integer readonly dispid 12;
    property Included: WordBool dispid 13;
    property InvalidAddress: WordBool dispid 14;
    property InvalidComments: WideString dispid 15;
    property MappedDataFields: MappedDataFields readonly dispid 16;
    property TableName: WideString readonly dispid 17;
    function FindRecord(const FindText: WideString; var Field: OleVariant): WordBool; dispid 102;
    procedure SetAllIncludedFlags(Included: WordBool); dispid 103;
    procedure SetAllErrorFlags(Invalid: WordBool; const InvalidComment: WideString); dispid 104;
    procedure Close; dispid 105;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: MailMergeFieldNames
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091C-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMergeFieldNames = interface(_IKsoDispObj)
    ['{0002091C-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): MailMergeFieldName; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  MailMergeFieldNamesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091C-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMergeFieldNamesDisp = dispinterface
    ['{0002091C-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): MailMergeFieldName; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: MailMergeFieldName
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091B-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMergeFieldName = interface(_IKsoDispObj)
    ['{0002091B-0000-4B30-A977-D214852036FE}']
    function Get_Name: WideString; safecall;
    function Get_Index: Integer; safecall;
    property Name: WideString read Get_Name;
    property Index: Integer read Get_Index;
  end;

// *********************************************************************//
// DispIntf:  MailMergeFieldNameDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091B-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMergeFieldNameDisp = dispinterface
    ['{0002091B-0000-4B30-A977-D214852036FE}']
    property Name: WideString readonly dispid 0;
    property Index: Integer readonly dispid 1;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: MailMergeDataFields
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMergeDataFields = interface(_IKsoDispObj)
    ['{0002091A-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): MailMergeDataField; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  MailMergeDataFieldsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMergeDataFieldsDisp = dispinterface
    ['{0002091A-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): MailMergeDataField; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: MailMergeDataField
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020919-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMergeDataField = interface(_IKsoDispObj)
    ['{00020919-0000-4B30-A977-D214852036FE}']
    function Get_Value: WideString; safecall;
    function Get_Name: WideString; safecall;
    function Get_Index: Integer; safecall;
    property Value: WideString read Get_Value;
    property Name: WideString read Get_Name;
    property Index: Integer read Get_Index;
  end;

// *********************************************************************//
// DispIntf:  MailMergeDataFieldDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020919-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMergeDataFieldDisp = dispinterface
    ['{00020919-0000-4B30-A977-D214852036FE}']
    property Value: WideString readonly dispid 0;
    property Name: WideString readonly dispid 1;
    property Index: Integer readonly dispid 2;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: MappedDataFields
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00026814-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MappedDataFields = interface(_IKsoDispObj)
    ['{00026814-0000-4B30-A977-D214852036FE}']
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Item(Index: WpsMappedDataFields): MappedDataField; safecall;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  MappedDataFieldsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00026814-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MappedDataFieldsDisp = dispinterface
    ['{00026814-0000-4B30-A977-D214852036FE}']
    property Count: Integer readonly dispid 1;
    property _NewEnum: IUnknown readonly dispid -4;
    function Item(Index: WpsMappedDataFields): MappedDataField; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: MappedDataField
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00021669-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MappedDataField = interface(_IKsoDispObj)
    ['{00021669-0000-4B30-A977-D214852036FE}']
    function Get_Index: Integer; safecall;
    function Get_DataFieldName: WideString; safecall;
    function Get_Name: WideString; safecall;
    function Get_Value: WideString; safecall;
    function Get_DataFieldIndex: Integer; safecall;
    procedure Set_DataFieldIndex(prop: Integer); safecall;
    property Index: Integer read Get_Index;
    property DataFieldName: WideString read Get_DataFieldName;
    property Name: WideString read Get_Name;
    property Value: WideString read Get_Value;
    property DataFieldIndex: Integer read Get_DataFieldIndex write Set_DataFieldIndex;
  end;

// *********************************************************************//
// DispIntf:  MappedDataFieldDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00021669-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MappedDataFieldDisp = dispinterface
    ['{00021669-0000-4B30-A977-D214852036FE}']
    property Index: Integer readonly dispid 1;
    property DataFieldName: WideString readonly dispid 2;
    property Name: WideString readonly dispid 0;
    property Value: WideString readonly dispid 4;
    property DataFieldIndex: Integer dispid 5;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: MailMergeFields
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMergeFields = interface(_IKsoDispObj)
    ['{0002091F-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): MailMergeField; safecall;
    function Add(const Range: Range; const Name: WideString): MailMergeField; safecall;
    function AddAsk(const Range: Range; const Name: WideString; var Prompt: OleVariant; 
                    var DefaultAskText: OleVariant; var AskOnce: OleVariant): MailMergeField; safecall;
    function AddFillIn(const Range: Range; var Prompt: OleVariant; 
                       var DefaultFillInText: OleVariant; var AskOnce: OleVariant): MailMergeField; safecall;
    function AddIf(const Range: Range; const MergeField: WideString; 
                   Comparison: WpsMailMergeComparison; var CompareTo: OleVariant; 
                   var TrueAutoText: OleVariant; var TrueText: OleVariant; 
                   var FalseAutoText: OleVariant; var FalseText: OleVariant): MailMergeField; safecall;
    function AddMergeRec(const Range: Range): MailMergeField; safecall;
    function AddMergeSeq(const Range: Range): MailMergeField; safecall;
    function AddNext(const Range: Range): MailMergeField; safecall;
    function AddNextIf(const Range: Range; const MergeField: WideString; 
                       Comparison: WpsMailMergeComparison; var CompareTo: OleVariant): MailMergeField; safecall;
    function AddSet(const Range: Range; const Name: WideString; var ValueText: OleVariant; 
                    var ValueAutoText: OleVariant): MailMergeField; safecall;
    function AddSkipIf(const Range: Range; const MergeField: WideString; 
                       Comparison: WpsMailMergeComparison; var CompareTo: OleVariant): MailMergeField; safecall;
    function Update: Integer; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  MailMergeFieldsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMergeFieldsDisp = dispinterface
    ['{0002091F-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(Index: Integer): MailMergeField; dispid 0;
    function Add(const Range: Range; const Name: WideString): MailMergeField; dispid 101;
    function AddAsk(const Range: Range; const Name: WideString; var Prompt: OleVariant; 
                    var DefaultAskText: OleVariant; var AskOnce: OleVariant): MailMergeField; dispid 102;
    function AddFillIn(const Range: Range; var Prompt: OleVariant; 
                       var DefaultFillInText: OleVariant; var AskOnce: OleVariant): MailMergeField; dispid 103;
    function AddIf(const Range: Range; const MergeField: WideString; 
                   Comparison: WpsMailMergeComparison; var CompareTo: OleVariant; 
                   var TrueAutoText: OleVariant; var TrueText: OleVariant; 
                   var FalseAutoText: OleVariant; var FalseText: OleVariant): MailMergeField; dispid 104;
    function AddMergeRec(const Range: Range): MailMergeField; dispid 105;
    function AddMergeSeq(const Range: Range): MailMergeField; dispid 106;
    function AddNext(const Range: Range): MailMergeField; dispid 107;
    function AddNextIf(const Range: Range; const MergeField: WideString; 
                       Comparison: WpsMailMergeComparison; var CompareTo: OleVariant): MailMergeField; dispid 108;
    function AddSet(const Range: Range; const Name: WideString; var ValueText: OleVariant; 
                    var ValueAutoText: OleVariant): MailMergeField; dispid 109;
    function AddSkipIf(const Range: Range; const MergeField: WideString; 
                       Comparison: WpsMailMergeComparison; var CompareTo: OleVariant): MailMergeField; dispid 110;
    function Update: Integer; dispid 111;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: MailMergeField
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091E-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMergeField = interface(_IKsoDispObj)
    ['{0002091E-0000-4B30-A977-D214852036FE}']
    function Get_type_: WpsFieldType; safecall;
    function Get_Locked: WordBool; safecall;
    procedure Set_Locked(prop: WordBool); safecall;
    function Get_Code: Range; safecall;
    procedure Set_Code(const prop: Range); safecall;
    function Get_Next: MailMergeField; safecall;
    function Get_Previous: MailMergeField; safecall;
    procedure Select; safecall;
    procedure Copy; safecall;
    procedure Cut; safecall;
    procedure Delete; safecall;
    function Update: WordBool; safecall;
    procedure Unlink; safecall;
    property type_: WpsFieldType read Get_type_;
    property Locked: WordBool read Get_Locked write Set_Locked;
    property Code: Range read Get_Code write Set_Code;
    property Next: MailMergeField read Get_Next;
    property Previous: MailMergeField read Get_Previous;
  end;

// *********************************************************************//
// DispIntf:  MailMergeFieldDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002091E-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  MailMergeFieldDisp = dispinterface
    ['{0002091E-0000-4B30-A977-D214852036FE}']
    property type_: WpsFieldType readonly dispid 0;
    property Locked: WordBool dispid 3;
    property Code: Range dispid 5;
    property Next: MailMergeField readonly dispid 8;
    property Previous: MailMergeField readonly dispid 9;
    procedure Select; dispid 65535;
    procedure Copy; dispid 105;
    procedure Cut; dispid 106;
    procedure Delete; dispid 107;
    function Update: WordBool; dispid 108;
    procedure Unlink; dispid 109;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: TablesOfContents
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020914-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TablesOfContents = interface(_IKsoDispObj)
    ['{00020914-0000-4B30-A977-D214852036FE}']
    function Get_Format: WpsTocFormat; safecall;
    procedure Set_Format(prop: WpsTocFormat); safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): TableOfContents; safecall;
    function Add(const Range: Range; UseHeadingStyles: WordBool; UpperHeadingLevel: Integer; 
                 LowerHeadingLevel: Integer; UseFields: WordBool; var TableID: OleVariant; 
                 RightAlignPageNumbers: WordBool; IncludePageNumbers: WordBool; 
                 const AddedStyles: WideString; UseHyperlinks: WordBool; 
                 HidePageNumbersInWeb: WordBool; UseOutlineLevels: WordBool): TableOfContents; safecall;
    function Get__NewEnum: IUnknown; safecall;
    property Format: WpsTocFormat read Get_Format write Set_Format;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  TablesOfContentsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020914-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TablesOfContentsDisp = dispinterface
    ['{00020914-0000-4B30-A977-D214852036FE}']
    property Format: WpsTocFormat dispid 2;
    property Count: Integer readonly dispid 1;
    function Item(Index: Integer): TableOfContents; dispid 0;
    function Add(const Range: Range; UseHeadingStyles: WordBool; UpperHeadingLevel: Integer; 
                 LowerHeadingLevel: Integer; UseFields: WordBool; var TableID: OleVariant; 
                 RightAlignPageNumbers: WordBool; IncludePageNumbers: WordBool; 
                 const AddedStyles: WideString; UseHyperlinks: WordBool; 
                 HidePageNumbersInWeb: WordBool; UseOutlineLevels: WordBool): TableOfContents; dispid 103;
    property _NewEnum: IUnknown readonly dispid -4;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: TableOfContents
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020913-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TableOfContents = interface(_IKsoDispObj)
    ['{00020913-0000-4B30-A977-D214852036FE}']
    function Get_UseHeadingStyles: WordBool; safecall;
    procedure Set_UseHeadingStyles(prop: WordBool); safecall;
    function Get_UseFields: WordBool; safecall;
    procedure Set_UseFields(prop: WordBool); safecall;
    function Get_TableID: WideString; safecall;
    procedure Set_TableID(const prop: WideString); safecall;
    function Get_RightAlignPageNumbers: WordBool; safecall;
    procedure Set_RightAlignPageNumbers(prop: WordBool); safecall;
    function Get_IncludePageNumbers: WordBool; safecall;
    procedure Set_IncludePageNumbers(prop: WordBool); safecall;
    function Get_Range: Range; safecall;
    function Get_TabLeader: WpsTabLeader; safecall;
    procedure Set_TabLeader(prop: WpsTabLeader); safecall;
    procedure Delete; safecall;
    function Get_LowerHeadingLevel: Integer; safecall;
    procedure Set_LowerHeadingLevel(prop: Integer); safecall;
    function Get_UpperHeadingLevel: Integer; safecall;
    procedure Set_UpperHeadingLevel(prop: Integer); safecall;
    procedure UpdatePageNumbers; safecall;
    procedure Update; safecall;
    function Get_UseHyperlinks: WordBool; safecall;
    procedure Set_UseHyperlinks(prop: WordBool); safecall;
    property UseHeadingStyles: WordBool read Get_UseHeadingStyles write Set_UseHeadingStyles;
    property UseFields: WordBool read Get_UseFields write Set_UseFields;
    property TableID: WideString read Get_TableID write Set_TableID;
    property RightAlignPageNumbers: WordBool read Get_RightAlignPageNumbers write Set_RightAlignPageNumbers;
    property IncludePageNumbers: WordBool read Get_IncludePageNumbers write Set_IncludePageNumbers;
    property Range: Range read Get_Range;
    property TabLeader: WpsTabLeader read Get_TabLeader write Set_TabLeader;
    property LowerHeadingLevel: Integer read Get_LowerHeadingLevel write Set_LowerHeadingLevel;
    property UpperHeadingLevel: Integer read Get_UpperHeadingLevel write Set_UpperHeadingLevel;
    property UseHyperlinks: WordBool read Get_UseHyperlinks write Set_UseHyperlinks;
  end;

// *********************************************************************//
// DispIntf:  TableOfContentsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020913-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TableOfContentsDisp = dispinterface
    ['{00020913-0000-4B30-A977-D214852036FE}']
    property UseHeadingStyles: WordBool dispid 1;
    property UseFields: WordBool dispid 2;
    property TableID: WideString dispid 5;
    property RightAlignPageNumbers: WordBool dispid 7;
    property IncludePageNumbers: WordBool dispid 8;
    property Range: Range readonly dispid 9;
    property TabLeader: WpsTabLeader dispid 10;
    procedure Delete; dispid 100;
    property LowerHeadingLevel: Integer dispid 4;
    property UpperHeadingLevel: Integer dispid 3;
    procedure UpdatePageNumbers; dispid 101;
    procedure Update; dispid 102;
    property UseHyperlinks: WordBool dispid 11;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Windows
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020961-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Windows = interface(_IKsoDispObj)
    ['{00020961-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): Window; safecall;
    function Add(const Window: Window): Window; safecall;
    procedure Arrange(ArrangeStyle: Integer); safecall;
    function CompareSideBySideWith(var Document: OleVariant): WordBool; safecall;
    function BreakSideBySide: WordBool; safecall;
    procedure ResetPositionsSideBySide; safecall;
    function Get_SyncScrollingSideBySide: WordBool; safecall;
    procedure Set_SyncScrollingSideBySide(prop: WordBool); safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property SyncScrollingSideBySide: WordBool read Get_SyncScrollingSideBySide write Set_SyncScrollingSideBySide;
  end;

// *********************************************************************//
// DispIntf:  WindowsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020961-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  WindowsDisp = dispinterface
    ['{00020961-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    function Item(Index: Integer): Window; dispid 0;
    function Add(const Window: Window): Window; dispid 10;
    procedure Arrange(ArrangeStyle: Integer); dispid 11;
    function CompareSideBySideWith(var Document: OleVariant): WordBool; dispid 12;
    function BreakSideBySide: WordBool; dispid 13;
    procedure ResetPositionsSideBySide; dispid 14;
    property SyncScrollingSideBySide: WordBool dispid 1003;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Window
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020962-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Window = interface(_IKsoDispObj)
    ['{00020962-0000-4B30-A977-D214852036FE}']
    function Get_ActivePane: Pane; safecall;
    function Get_Document: _Document; safecall;
    function Get_Panes: Panes; safecall;
    function Get_Selection: Selection; safecall;
    function Get_Left: Integer; safecall;
    procedure Set_Left(prop: Integer); safecall;
    function Get_Top: Integer; safecall;
    procedure Set_Top(prop: Integer); safecall;
    function Get_Width: Integer; safecall;
    procedure Set_Width(prop: Integer); safecall;
    function Get_Height: Integer; safecall;
    procedure Set_Height(prop: Integer); safecall;
    function Get_Split: WordBool; safecall;
    procedure Set_Split(prop: WordBool); safecall;
    function Get_SplitVertical: Integer; safecall;
    procedure Set_SplitVertical(prop: Integer); safecall;
    function Get_Caption: WideString; safecall;
    procedure Set_Caption(const prop: WideString); safecall;
    function Get_WindowState: WpsWindowState; safecall;
    procedure Set_WindowState(prop: WpsWindowState); safecall;
    function Get_DisplayRulers: WordBool; safecall;
    procedure Set_DisplayRulers(prop: WordBool); safecall;
    function Get_DisplayVerticalRuler: WordBool; safecall;
    procedure Set_DisplayVerticalRuler(prop: WordBool); safecall;
    function Get_View: View; safecall;
    function Get_type_: WpsWindowType; safecall;
    function Get_Next: Window; safecall;
    function Get_Previous: Window; safecall;
    function Get_WindowNumber: Integer; safecall;
    function Get_DisplayVerticalScrollBar: WordBool; safecall;
    procedure Set_DisplayVerticalScrollBar(prop: WordBool); safecall;
    function Get_DisplayHorizontalScrollBar: WordBool; safecall;
    procedure Set_DisplayHorizontalScrollBar(prop: WordBool); safecall;
    function Get_DisplayScreenTips: WordBool; safecall;
    procedure Set_DisplayScreenTips(prop: WordBool); safecall;
    function Get_HorizontalPercentScrolled: Integer; safecall;
    procedure Set_HorizontalPercentScrolled(prop: Integer); safecall;
    function Get_VerticalPercentScrolled: Integer; safecall;
    procedure Set_VerticalPercentScrolled(prop: Integer); safecall;
    function Get_Active: WordBool; safecall;
    function Get_Index: Integer; safecall;
    procedure Activate; safecall;
    procedure Close(SaveChanges: WpsSaveOptions; RouteDocument: WordBool); safecall;
    procedure LargeScroll(Down: Integer; Up: Integer; ToRight: Integer; ToLeft: Integer); safecall;
    procedure SmallScroll(Down: Integer; Up: Integer; ToRight: Integer; ToLeft: Integer); safecall;
    function NewWindow: Window; safecall;
    procedure PageScroll(Down: Integer; Up: Integer); safecall;
    procedure SetFocus; safecall;
    function RangeFromPoint(x: Integer; y: Integer): IDispatch; safecall;
    procedure ScrollIntoView(const obj: IDispatch; var Start: OleVariant); safecall;
    procedure GetPoint(out ScreenPixelsLeft: Integer; out ScreenPixelsTop: Integer; 
                       out ScreenPixelsWidth: Integer; out ScreenPixelsHeight: Integer; 
                       const obj: IDispatch); safecall;
    function Get_UsableWidth: Integer; safecall;
    function Get_UsableHeight: Integer; safecall;
    function Get_DisplayRightRuler: WordBool; safecall;
    procedure Set_DisplayRightRuler(prop: WordBool); safecall;
    function Get_DisplayLeftScrollBar: WordBool; safecall;
    procedure Set_DisplayLeftScrollBar(prop: WordBool); safecall;
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(prop: WordBool); safecall;
    procedure PrintOut(Background: WordBool; Append: WordBool; Range: WpsPrintOutRange; 
                       const OutputFileName: WideString; From: SYSINT; To_: SYSINT; 
                       Item: WpsPrintOutItem; Copies: SYSINT; const Pages: WideString; 
                       PageType: WpsPrintOutPages; PrintToFile: WordBool; Collate: WordBool; 
                       var ActivePrinterMacGX: OleVariant; ManualDuplexPrint: WordBool; 
                       PrintZoomColumn: Integer; PrintZoomRow: Integer; 
                       PrintZoomPaperWidth: Integer; PrintZoomPaperHeight: Integer; 
                       FlipPrint: WordBool; PaperTray: WpsPaperTray; CutterLine: WordBool; 
                       PaperOrder: WpsPaperOrder); safecall;
    procedure ToggleShowAllReviewers; safecall;
    function Get_CommandBars: _CommandBars; safecall;
    procedure ScreenCoordToDoc(SrcLeft: Integer; SrcTop: Integer; SrcWidth: Integer; 
                               SrcHeight: Integer; out NewLeft: Single; out NewTop: Single; 
                               out NewWidth: Single; out NewHeight: Single); safecall;
    procedure DocCoordToScreen(SrcLeft: Single; SrcTop: Single; SrcWidth: Single; 
                               SrcHeight: Single; out NewLeft: Integer; out NewTop: Integer; 
                               out NewWidth: Integer; out NewHeight: Integer); safecall;
    function Get_DocumentMap: WordBool; safecall;
    procedure Set_DocumentMap(prop: WordBool); safecall;
    property ActivePane: Pane read Get_ActivePane;
    property Document: _Document read Get_Document;
    property Panes: Panes read Get_Panes;
    property Selection: Selection read Get_Selection;
    property Left: Integer read Get_Left write Set_Left;
    property Top: Integer read Get_Top write Set_Top;
    property Width: Integer read Get_Width write Set_Width;
    property Height: Integer read Get_Height write Set_Height;
    property Split: WordBool read Get_Split write Set_Split;
    property SplitVertical: Integer read Get_SplitVertical write Set_SplitVertical;
    property Caption: WideString read Get_Caption write Set_Caption;
    property WindowState: WpsWindowState read Get_WindowState write Set_WindowState;
    property DisplayRulers: WordBool read Get_DisplayRulers write Set_DisplayRulers;
    property DisplayVerticalRuler: WordBool read Get_DisplayVerticalRuler write Set_DisplayVerticalRuler;
    property View: View read Get_View;
    property type_: WpsWindowType read Get_type_;
    property Next: Window read Get_Next;
    property Previous: Window read Get_Previous;
    property WindowNumber: Integer read Get_WindowNumber;
    property DisplayVerticalScrollBar: WordBool read Get_DisplayVerticalScrollBar write Set_DisplayVerticalScrollBar;
    property DisplayHorizontalScrollBar: WordBool read Get_DisplayHorizontalScrollBar write Set_DisplayHorizontalScrollBar;
    property DisplayScreenTips: WordBool read Get_DisplayScreenTips write Set_DisplayScreenTips;
    property HorizontalPercentScrolled: Integer read Get_HorizontalPercentScrolled write Set_HorizontalPercentScrolled;
    property VerticalPercentScrolled: Integer read Get_VerticalPercentScrolled write Set_VerticalPercentScrolled;
    property Active: WordBool read Get_Active;
    property Index: Integer read Get_Index;
    property UsableWidth: Integer read Get_UsableWidth;
    property UsableHeight: Integer read Get_UsableHeight;
    property DisplayRightRuler: WordBool read Get_DisplayRightRuler write Set_DisplayRightRuler;
    property DisplayLeftScrollBar: WordBool read Get_DisplayLeftScrollBar write Set_DisplayLeftScrollBar;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property CommandBars: _CommandBars read Get_CommandBars;
    property DocumentMap: WordBool read Get_DocumentMap write Set_DocumentMap;
  end;

// *********************************************************************//
// DispIntf:  WindowDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020962-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  WindowDisp = dispinterface
    ['{00020962-0000-4B30-A977-D214852036FE}']
    property ActivePane: Pane readonly dispid 1;
    property Document: _Document readonly dispid 2;
    property Panes: Panes readonly dispid 3;
    property Selection: Selection readonly dispid 4;
    property Left: Integer dispid 5;
    property Top: Integer dispid 6;
    property Width: Integer dispid 7;
    property Height: Integer dispid 8;
    property Split: WordBool dispid 9;
    property SplitVertical: Integer dispid 10;
    property Caption: WideString dispid 0;
    property WindowState: WpsWindowState dispid 11;
    property DisplayRulers: WordBool dispid 12;
    property DisplayVerticalRuler: WordBool dispid 13;
    property View: View readonly dispid 14;
    property type_: WpsWindowType readonly dispid 15;
    property Next: Window readonly dispid 16;
    property Previous: Window readonly dispid 17;
    property WindowNumber: Integer readonly dispid 18;
    property DisplayVerticalScrollBar: WordBool dispid 19;
    property DisplayHorizontalScrollBar: WordBool dispid 20;
    property DisplayScreenTips: WordBool dispid 22;
    property HorizontalPercentScrolled: Integer dispid 23;
    property VerticalPercentScrolled: Integer dispid 24;
    property Active: WordBool readonly dispid 26;
    property Index: Integer readonly dispid 28;
    procedure Activate; dispid 100;
    procedure Close(SaveChanges: WpsSaveOptions; RouteDocument: WordBool); dispid 102;
    procedure LargeScroll(Down: Integer; Up: Integer; ToRight: Integer; ToLeft: Integer); dispid 103;
    procedure SmallScroll(Down: Integer; Up: Integer; ToRight: Integer; ToLeft: Integer); dispid 104;
    function NewWindow: Window; dispid 105;
    procedure PageScroll(Down: Integer; Up: Integer); dispid 108;
    procedure SetFocus; dispid 109;
    function RangeFromPoint(x: Integer; y: Integer): IDispatch; dispid 110;
    procedure ScrollIntoView(const obj: IDispatch; var Start: OleVariant); dispid 111;
    procedure GetPoint(out ScreenPixelsLeft: Integer; out ScreenPixelsTop: Integer; 
                       out ScreenPixelsWidth: Integer; out ScreenPixelsHeight: Integer; 
                       const obj: IDispatch); dispid 112;
    property UsableWidth: Integer readonly dispid 31;
    property UsableHeight: Integer readonly dispid 32;
    property DisplayRightRuler: WordBool dispid 35;
    property DisplayLeftScrollBar: WordBool dispid 34;
    property Visible: WordBool dispid 36;
    procedure PrintOut(Background: WordBool; Append: WordBool; Range: WpsPrintOutRange; 
                       const OutputFileName: WideString; From: SYSINT; To_: SYSINT; 
                       Item: WpsPrintOutItem; Copies: SYSINT; const Pages: WideString; 
                       PageType: WpsPrintOutPages; PrintToFile: WordBool; Collate: WordBool; 
                       var ActivePrinterMacGX: OleVariant; ManualDuplexPrint: WordBool; 
                       PrintZoomColumn: Integer; PrintZoomRow: Integer; 
                       PrintZoomPaperWidth: Integer; PrintZoomPaperHeight: Integer; 
                       FlipPrint: WordBool; PaperTray: WpsPaperTray; CutterLine: WordBool; 
                       PaperOrder: WpsPaperOrder); dispid 445;
    procedure ToggleShowAllReviewers; dispid 446;
    property CommandBars: _CommandBars readonly dispid 447;
    procedure ScreenCoordToDoc(SrcLeft: Integer; SrcTop: Integer; SrcWidth: Integer; 
                               SrcHeight: Integer; out NewLeft: Single; out NewTop: Single; 
                               out NewWidth: Single; out NewHeight: Single); dispid 1610809406;
    procedure DocCoordToScreen(SrcLeft: Single; SrcTop: Single; SrcWidth: Single; 
                               SrcHeight: Single; out NewLeft: Integer; out NewTop: Integer; 
                               out NewWidth: Integer; out NewHeight: Integer); dispid 1610809407;
    property DocumentMap: WordBool dispid 25;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Pane
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020960-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Pane = interface(_IKsoDispObj)
    ['{00020960-0000-4B30-A977-D214852036FE}']
    function Get_Document: _Document; safecall;
    function Get_Selection: Selection; safecall;
    function Get_DisplayRulers: WordBool; safecall;
    procedure Set_DisplayRulers(prop: WordBool); safecall;
    function Get_DisplayVerticalRuler: WordBool; safecall;
    procedure Set_DisplayVerticalRuler(prop: WordBool); safecall;
    function Get_Zooms: Zooms; safecall;
    function Get_Index: Integer; safecall;
    function Get_View: View; safecall;
    function Get_Next: Pane; safecall;
    function Get_Previous: Pane; safecall;
    function Get_HorizontalPercentScrolled: Integer; safecall;
    procedure Set_HorizontalPercentScrolled(prop: Integer); safecall;
    function Get_VerticalPercentScrolled: Integer; safecall;
    procedure Set_VerticalPercentScrolled(prop: Integer); safecall;
    function Get_BrowseWidth: Integer; safecall;
    procedure Activate; safecall;
    procedure Close; safecall;
    procedure LargeScroll(Down: Integer; Up: Integer; ToRight: Integer; ToLeft: Integer); safecall;
    procedure SmallScroll(Down: Integer; Up: Integer; ToRight: Integer; ToLeft: Integer); safecall;
    procedure AutoScroll(Velocity: Integer); safecall;
    procedure PageScroll(Down: Integer; Up: Integer); safecall;
    function Get_Pages: Pages; safecall;
    procedure Seralize(const bstrFileName: WideString); safecall;
    property Document: _Document read Get_Document;
    property Selection: Selection read Get_Selection;
    property DisplayRulers: WordBool read Get_DisplayRulers write Set_DisplayRulers;
    property DisplayVerticalRuler: WordBool read Get_DisplayVerticalRuler write Set_DisplayVerticalRuler;
    property Zooms: Zooms read Get_Zooms;
    property Index: Integer read Get_Index;
    property View: View read Get_View;
    property Next: Pane read Get_Next;
    property Previous: Pane read Get_Previous;
    property HorizontalPercentScrolled: Integer read Get_HorizontalPercentScrolled write Set_HorizontalPercentScrolled;
    property VerticalPercentScrolled: Integer read Get_VerticalPercentScrolled write Set_VerticalPercentScrolled;
    property BrowseWidth: Integer read Get_BrowseWidth;
    property Pages: Pages read Get_Pages;
  end;

// *********************************************************************//
// DispIntf:  PaneDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020960-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  PaneDisp = dispinterface
    ['{00020960-0000-4B30-A977-D214852036FE}']
    property Document: _Document readonly dispid 1;
    property Selection: Selection readonly dispid 3;
    property DisplayRulers: WordBool dispid 4;
    property DisplayVerticalRuler: WordBool dispid 5;
    property Zooms: Zooms readonly dispid 7;
    property Index: Integer readonly dispid 9;
    property View: View readonly dispid 10;
    property Next: Pane readonly dispid 11;
    property Previous: Pane readonly dispid 12;
    property HorizontalPercentScrolled: Integer dispid 13;
    property VerticalPercentScrolled: Integer dispid 14;
    property BrowseWidth: Integer readonly dispid 17;
    procedure Activate; dispid 100;
    procedure Close; dispid 101;
    procedure LargeScroll(Down: Integer; Up: Integer; ToRight: Integer; ToLeft: Integer); dispid 102;
    procedure SmallScroll(Down: Integer; Up: Integer; ToRight: Integer; ToLeft: Integer); dispid 103;
    procedure AutoScroll(Velocity: Integer); dispid 104;
    procedure PageScroll(Down: Integer; Up: Integer); dispid 105;
    property Pages: Pages readonly dispid 19;
    procedure Seralize(const bstrFileName: WideString); dispid 112;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Selection
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020975-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Selection = interface(_IKsoDispObj)
    ['{00020975-0000-4B30-A977-D214852036FE}']
    function Get_Text: WideString; safecall;
    procedure Set_Text(const prop: WideString); safecall;
    function Get_FormattedText: Range; safecall;
    procedure Set_FormattedText(const prop: Range); safecall;
    function Get_Start: Integer; safecall;
    function Get_End_: Integer; safecall;
    function Get_Font: _Font; safecall;
    procedure Set_Font(const prop: _Font); safecall;
    function Get_type_: WpsSelectionType; safecall;
    function Get_StoryType: WpsStoryType; safecall;
    function Get_Style: OleVariant; safecall;
    procedure Set_Style(var prop: OleVariant); safecall;
    function Get_Tables: Tables; safecall;
    function Get_Footnotes: Footnotes; safecall;
    function Get_Endnotes: Endnotes; safecall;
    function Get_FootnoteOptions: FootnoteOptions; safecall;
    function Get_EndnoteOptions: EndnoteOptions; safecall;
    function Get_Comments: Comments; safecall;
    function Get_Cells: Cells; safecall;
    function Get_Sections: Sections; safecall;
    function Get_Paragraphs: Paragraphs; safecall;
    function Get_Borders: Borders; safecall;
    procedure Set_Borders(const prop: Borders); safecall;
    function Get_Fields: Fields; safecall;
    function Get_Frames: Frames; safecall;
    function Get_ParagraphFormat: _ParagraphFormat; safecall;
    procedure Set_ParagraphFormat(const prop: _ParagraphFormat); safecall;
    function Get_PageSetup: PageSetup; safecall;
    procedure Set_PageSetup(const prop: PageSetup); safecall;
    function Get_Bookmarks: Bookmarks; safecall;
    function Get_StoryLength: Integer; safecall;
    function Get_Hyperlinks: Hyperlinks; safecall;
    function Get_Columns: Columns; safecall;
    function Get_Rows: Rows; safecall;
    function Get_BookmarkID: Integer; safecall;
    function Get_PreviousBookmarkID: Integer; safecall;
    function Get_Range: Range; safecall;
    function Get_Information(Type_: WpsInformation): OleVariant; safecall;
    function Get_Flags: WpsSelectionFlags; safecall;
    procedure Set_Flags(prop: WpsSelectionFlags); safecall;
    function Get_Active: WordBool; safecall;
    function Get_StartIsActive: WordBool; safecall;
    procedure Set_StartIsActive(prop: WordBool); safecall;
    function Get_IPAtEndOfLine: WordBool; safecall;
    function Get_ExtendMode: WordBool; safecall;
    procedure Set_ExtendMode(prop: WordBool); safecall;
    function Get_Orientation: WpsTextOrientation; safecall;
    procedure Set_Orientation(prop: WpsTextOrientation); safecall;
    function Get_InlineShapes: InlineShapes; safecall;
    function Get_Document: _Document; safecall;
    function Get_ShapeRange: ShapeRange; safecall;
    procedure Select; safecall;
    procedure SetRange(Start: Integer; End_: Integer); safecall;
    procedure Collapse(Direction: WpsCollapseDirection); safecall;
    procedure InsertBefore(const Text: WideString); safecall;
    procedure InsertAfter(const Text: WideString); safecall;
    function Next(Unit_: WpsUnits; Count: Integer): Range; safecall;
    function Previous(Unit_: WpsUnits; Count: Integer): Range; safecall;
    function StartOf(Unit_: WpsUnits; Extend: WpsMovementType): Integer; safecall;
    function EndOf(Unit_: WpsUnits; Extend: WpsMovementType): Integer; safecall;
    function Move(Unit_: WpsUnits; Count: Integer): Integer; safecall;
    function MoveStart(Unit_: WpsUnits; Count: Integer): Integer; safecall;
    function MoveEnd(Unit_: WpsUnits; Count: Integer): Integer; safecall;
    procedure Cut; safecall;
    procedure Copy; safecall;
    procedure Paste; safecall;
    procedure InsertBreak(Type_: WpsBreakType); safecall;
    function InStory(const Range: Range): WordBool; safecall;
    function InRange(const Range: Range): WordBool; safecall;
    function Delete(Unit_: WpsUnits; Count: Integer): Integer; safecall;
    function Expand(Unit_: WpsUnits): Integer; safecall;
    procedure InsertParagraph; safecall;
    procedure InsertParagraphAfter; safecall;
    procedure InsertSymbol(CharacterNumber: Integer; const Font: WideString; Unicode: WordBool; 
                           var Bias: OleVariant); safecall;
    function IsEqual(const Range: Range): WordBool; safecall;
    function GoTo_(What: WpsGoToItem; Which: WpsGoToDirection; Count: Integer; 
                   const Name: WideString): Range; safecall;
    function GoToNext(What: WpsGoToItem): Range; safecall;
    function GoToPrevious(What: WpsGoToItem): Range; safecall;
    procedure PasteSpecial(IconIndex: Integer; Link: WordBool; Placement: WpsOLEPlacement; 
                           DisplayAsIcon: WordBool; var DataType: OleVariant; 
                           const IconFileName: WideString; const IconLabel: WideString); safecall;
    function NextField: Field; safecall;
    procedure InsertParagraphBefore; safecall;
    procedure InsertCells(var ShiftCells: OleVariant); safecall;
    function MoveLeft(Unit_: WpsUnits; Count: Integer; Extend: WpsMovementType): Integer; safecall;
    function MoveRight(Unit_: WpsUnits; Count: Integer; Extend: WpsMovementType): Integer; safecall;
    function MoveUp(Unit_: WpsUnits; Count: Integer; Extend: WpsMovementType): Integer; safecall;
    function MoveDown(Unit_: WpsUnits; Count: Integer; Extend: WpsMovementType): Integer; safecall;
    function HomeKey(Unit_: WpsUnits; Extend: WpsMovementType): Integer; safecall;
    function EndKey(Unit_: WpsUnits; Extend: WpsMovementType): Integer; safecall;
    procedure EscapeKey; safecall;
    procedure TypeText(const Text: WideString); safecall;
    procedure CopyFormat; safecall;
    procedure PasteFormat; safecall;
    procedure TypeParagraph; safecall;
    procedure TypeBackspace; safecall;
    procedure SelectColumn; safecall;
    procedure CreateTextbox(uOrientation: KsoTextOrientation); safecall;
    procedure WholeStory; safecall;
    procedure SelectRow; safecall;
    procedure SplitTable; safecall;
    procedure InsertRows(var NumRows: OleVariant); safecall;
    procedure InsertColumns; safecall;
    procedure PasteAsNestedTable; safecall;
    procedure InsertRowsBelow(var NumRows: OleVariant); safecall;
    procedure InsertColumnsRight; safecall;
    procedure InsertRowsAbove(var NumRows: OleVariant); safecall;
    procedure BoldRun; safecall;
    procedure ItalicRun; safecall;
    procedure InsertDateTime(const DateTimeFormat: WideString; InsertAsField: WordBool; 
                             InsertAsFullWidth: WordBool; DateLanguage: WpsDateLanguage; 
                             CalendarType: WpsCalendarTypeBi); safecall;
    procedure PasteAppendTable; safecall;
    procedure InsertCaption(var Label_: OleVariant; const Title: WideString; 
                            const TitleAutoText: WideString; Position: WpsCaptionPosition; 
                            ExcludeLabel: WordBool); safecall;
    procedure ClearFormatting; safecall;
    function Get_Find: Find; safecall;
    function Get_WordStat: WordStat; safecall;
    function Get_HeaderFooter: HeaderFooter; safecall;
    function Get_Case_: WpsCharacterCase; safecall;
    procedure Set_Case_(prop: WpsCharacterCase); safecall;
    procedure EnableDropCap(Position: WpsDropPosition; const FontName: WideString; 
                            LineToDrop: Integer; DistanceFromText: Single); safecall;
    function Get_DropCap: DropCap; safecall;
    function Get_HighlightColor: WpsColor; safecall;
    procedure Set_HighlightColor(prop: WpsColor); safecall;
    procedure ApplyListTemplate(const ListTemplate: ListTemplate; 
                                ContinuePreviousList: WpsContinue; ApplyTo: WpsListApplyTo; 
                                DefaultListBehavior: Integer; UseContinueType: WpsListGalleryType; 
                                l1BasedContinueLevel: Integer); safecall;
    function Get_ListType: WpsListType; safecall;
    function Get_ListTemplate: ListTemplate; safecall;
    function CanContinuePreviousList(const ListTemplate: ListTemplate): WpsContinue; safecall;
    procedure RemoveNumbers(NumberType: WpsNumberType; ApplyTo: WpsListApplyTo); safecall;
    procedure PasteAndFormat(Type_: WpsRecoveryType); safecall;
    procedure TCSCConverter(WpsTCSCCDirection: WpsTCSCConverterDirection; CommonTerms: WordBool; 
                            UseVariants: WordBool); safecall;
    procedure InsertFile(const FileName: WideString; var Range: OleVariant; 
                         var ConfirmConversions: OleVariant; var Link: OleVariant; 
                         var Attachment: OleVariant); safecall;
    function Get_EnhMetaFileBits: OleVariant; safecall;
    procedure Set_Start(prop: Integer); safecall;
    procedure Set_End_(prop: Integer); safecall;
    function Get_LanguageID: WpsLanguageID; safecall;
    procedure Set_LanguageID(prop: WpsLanguageID); safecall;
    function Get_LanguageIDFarEast: WpsLanguageID; safecall;
    procedure Set_LanguageIDFarEast(prop: WpsLanguageID); safecall;
    function Get_LanguageIDOther: WpsLanguageID; safecall;
    procedure Set_LanguageIDOther(prop: WpsLanguageID); safecall;
    function ConvertToTable(var Separator: OleVariant; var NumRows: OleVariant; 
                            var NumColumns: OleVariant; var InitialColumnWidth: OleVariant; 
                            var Format: OleVariant; var ApplyBorders: OleVariant; 
                            var ApplyShading: OleVariant; var ApplyFont: OleVariant; 
                            var ApplyColor: OleVariant; var ApplyHeadingRows: OleVariant; 
                            var ApplyLastRow: OleVariant; var ApplyFirstColumn: OleVariant; 
                            var ApplyLastColumn: OleVariant; var AutoFit: OleVariant; 
                            var AutoFitBehavior: OleVariant; var DefaultTableBehavior: OleVariant): Table; safecall;
    function Get_FormFields: FormFields; safecall;
    function Get_Editors: Editors; safecall;
    procedure InsertCrossReference(var ReferenceType: OleVariant; ReferenceKind: WpsReferenceKind; 
                                   var ReferenceItem: OleVariant; 
                                   var InsertAsHyperlink: OleVariant; 
                                   var IncludePosition: OleVariant; 
                                   var SeparateNumbers: OleVariant; var SeparatorString: OleVariant); safecall;
    property Text: WideString read Get_Text write Set_Text;
    property FormattedText: Range read Get_FormattedText write Set_FormattedText;
    property Start: Integer read Get_Start write Set_Start;
    property End_: Integer read Get_End_ write Set_End_;
    property Font: _Font read Get_Font write Set_Font;
    property type_: WpsSelectionType read Get_type_;
    property StoryType: WpsStoryType read Get_StoryType;
    property Tables: Tables read Get_Tables;
    property Footnotes: Footnotes read Get_Footnotes;
    property Endnotes: Endnotes read Get_Endnotes;
    property FootnoteOptions: FootnoteOptions read Get_FootnoteOptions;
    property EndnoteOptions: EndnoteOptions read Get_EndnoteOptions;
    property Comments: Comments read Get_Comments;
    property Cells: Cells read Get_Cells;
    property Sections: Sections read Get_Sections;
    property Paragraphs: Paragraphs read Get_Paragraphs;
    property Borders: Borders read Get_Borders write Set_Borders;
    property Fields: Fields read Get_Fields;
    property Frames: Frames read Get_Frames;
    property ParagraphFormat: _ParagraphFormat read Get_ParagraphFormat write Set_ParagraphFormat;
    property PageSetup: PageSetup read Get_PageSetup write Set_PageSetup;
    property Bookmarks: Bookmarks read Get_Bookmarks;
    property StoryLength: Integer read Get_StoryLength;
    property Hyperlinks: Hyperlinks read Get_Hyperlinks;
    property Columns: Columns read Get_Columns;
    property Rows: Rows read Get_Rows;
    property BookmarkID: Integer read Get_BookmarkID;
    property PreviousBookmarkID: Integer read Get_PreviousBookmarkID;
    property Range: Range read Get_Range;
    property Information[Type_: WpsInformation]: OleVariant read Get_Information;
    property Flags: WpsSelectionFlags read Get_Flags write Set_Flags;
    property Active: WordBool read Get_Active;
    property StartIsActive: WordBool read Get_StartIsActive write Set_StartIsActive;
    property IPAtEndOfLine: WordBool read Get_IPAtEndOfLine;
    property ExtendMode: WordBool read Get_ExtendMode write Set_ExtendMode;
    property Orientation: WpsTextOrientation read Get_Orientation write Set_Orientation;
    property InlineShapes: InlineShapes read Get_InlineShapes;
    property Document: _Document read Get_Document;
    property ShapeRange: ShapeRange read Get_ShapeRange;
    property Find: Find read Get_Find;
    property WordStat: WordStat read Get_WordStat;
    property HeaderFooter: HeaderFooter read Get_HeaderFooter;
    property Case_: WpsCharacterCase read Get_Case_ write Set_Case_;
    property DropCap: DropCap read Get_DropCap;
    property HighlightColor: WpsColor read Get_HighlightColor write Set_HighlightColor;
    property ListType: WpsListType read Get_ListType;
    property ListTemplate: ListTemplate read Get_ListTemplate;
    property EnhMetaFileBits: OleVariant read Get_EnhMetaFileBits;
    property LanguageID: WpsLanguageID read Get_LanguageID write Set_LanguageID;
    property LanguageIDFarEast: WpsLanguageID read Get_LanguageIDFarEast write Set_LanguageIDFarEast;
    property LanguageIDOther: WpsLanguageID read Get_LanguageIDOther write Set_LanguageIDOther;
    property FormFields: FormFields read Get_FormFields;
    property Editors: Editors read Get_Editors;
  end;

// *********************************************************************//
// DispIntf:  SelectionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020975-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  SelectionDisp = dispinterface
    ['{00020975-0000-4B30-A977-D214852036FE}']
    property Text: WideString dispid 0;
    property FormattedText: Range dispid 2;
    property Start: Integer dispid 3;
    property End_: Integer dispid 4;
    property Font: _Font dispid 5;
    property type_: WpsSelectionType readonly dispid 6;
    property StoryType: WpsStoryType readonly dispid 7;
    function Style: OleVariant; dispid 8;
    property Tables: Tables readonly dispid 50;
    property Footnotes: Footnotes readonly dispid 54;
    property Endnotes: Endnotes readonly dispid 55;
    property FootnoteOptions: FootnoteOptions readonly dispid 1024;
    property EndnoteOptions: EndnoteOptions readonly dispid 1025;
    property Comments: Comments readonly dispid 56;
    property Cells: Cells readonly dispid 57;
    property Sections: Sections readonly dispid 58;
    property Paragraphs: Paragraphs readonly dispid 59;
    property Borders: Borders dispid 1100;
    property Fields: Fields readonly dispid 64;
    property Frames: Frames readonly dispid 66;
    property ParagraphFormat: _ParagraphFormat dispid 1102;
    property PageSetup: PageSetup dispid 1101;
    property Bookmarks: Bookmarks readonly dispid 75;
    property StoryLength: Integer readonly dispid 152;
    property Hyperlinks: Hyperlinks readonly dispid 156;
    property Columns: Columns readonly dispid 302;
    property Rows: Rows readonly dispid 303;
    property BookmarkID: Integer readonly dispid 308;
    property PreviousBookmarkID: Integer readonly dispid 309;
    property Range: Range readonly dispid 400;
    property Information[Type_: WpsInformation]: OleVariant readonly dispid 401;
    property Flags: WpsSelectionFlags dispid 402;
    property Active: WordBool readonly dispid 403;
    property StartIsActive: WordBool dispid 404;
    property IPAtEndOfLine: WordBool readonly dispid 405;
    property ExtendMode: WordBool dispid 406;
    property Orientation: WpsTextOrientation dispid 410;
    property InlineShapes: InlineShapes readonly dispid 411;
    property Document: _Document readonly dispid 1003;
    property ShapeRange: ShapeRange readonly dispid 1004;
    procedure Select; dispid 65535;
    procedure SetRange(Start: Integer; End_: Integer); dispid 100;
    procedure Collapse(Direction: WpsCollapseDirection); dispid 101;
    procedure InsertBefore(const Text: WideString); dispid 102;
    procedure InsertAfter(const Text: WideString); dispid 104;
    function Next(Unit_: WpsUnits; Count: Integer): Range; dispid 105;
    function Previous(Unit_: WpsUnits; Count: Integer): Range; dispid 106;
    function StartOf(Unit_: WpsUnits; Extend: WpsMovementType): Integer; dispid 107;
    function EndOf(Unit_: WpsUnits; Extend: WpsMovementType): Integer; dispid 108;
    function Move(Unit_: WpsUnits; Count: Integer): Integer; dispid 109;
    function MoveStart(Unit_: WpsUnits; Count: Integer): Integer; dispid 110;
    function MoveEnd(Unit_: WpsUnits; Count: Integer): Integer; dispid 111;
    procedure Cut; dispid 119;
    procedure Copy; dispid 120;
    procedure Paste; dispid 121;
    procedure InsertBreak(Type_: WpsBreakType); dispid 122;
    function InStory(const Range: Range): WordBool; dispid 125;
    function InRange(const Range: Range): WordBool; dispid 126;
    function Delete(Unit_: WpsUnits; Count: Integer): Integer; dispid 127;
    function Expand(Unit_: WpsUnits): Integer; dispid 129;
    procedure InsertParagraph; dispid 160;
    procedure InsertParagraphAfter; dispid 161;
    procedure InsertSymbol(CharacterNumber: Integer; const Font: WideString; Unicode: WordBool; 
                           var Bias: OleVariant); dispid 164;
    function IsEqual(const Range: Range): WordBool; dispid 171;
    function GoTo_(What: WpsGoToItem; Which: WpsGoToDirection; Count: Integer; 
                   const Name: WideString): Range; dispid 173;
    function GoToNext(What: WpsGoToItem): Range; dispid 174;
    function GoToPrevious(What: WpsGoToItem): Range; dispid 175;
    procedure PasteSpecial(IconIndex: Integer; Link: WordBool; Placement: WpsOLEPlacement; 
                           DisplayAsIcon: WordBool; var DataType: OleVariant; 
                           const IconFileName: WideString; const IconLabel: WideString); dispid 176;
    function NextField: Field; dispid 178;
    procedure InsertParagraphBefore; dispid 212;
    procedure InsertCells(var ShiftCells: OleVariant); dispid 214;
    function MoveLeft(Unit_: WpsUnits; Count: Integer; Extend: WpsMovementType): Integer; dispid 500;
    function MoveRight(Unit_: WpsUnits; Count: Integer; Extend: WpsMovementType): Integer; dispid 501;
    function MoveUp(Unit_: WpsUnits; Count: Integer; Extend: WpsMovementType): Integer; dispid 502;
    function MoveDown(Unit_: WpsUnits; Count: Integer; Extend: WpsMovementType): Integer; dispid 503;
    function HomeKey(Unit_: WpsUnits; Extend: WpsMovementType): Integer; dispid 504;
    function EndKey(Unit_: WpsUnits; Extend: WpsMovementType): Integer; dispid 505;
    procedure EscapeKey; dispid 506;
    procedure TypeText(const Text: WideString); dispid 507;
    procedure CopyFormat; dispid 509;
    procedure PasteFormat; dispid 510;
    procedure TypeParagraph; dispid 512;
    procedure TypeBackspace; dispid 513;
    procedure SelectColumn; dispid 516;
    procedure CreateTextbox(uOrientation: KsoTextOrientation); dispid 523;
    procedure WholeStory; dispid 524;
    procedure SelectRow; dispid 525;
    procedure SplitTable; dispid 526;
    procedure InsertRows(var NumRows: OleVariant); dispid 528;
    procedure InsertColumns; dispid 529;
    procedure PasteAsNestedTable; dispid 533;
    procedure InsertRowsBelow(var NumRows: OleVariant); dispid 537;
    procedure InsertColumnsRight; dispid 538;
    procedure InsertRowsAbove(var NumRows: OleVariant); dispid 539;
    procedure BoldRun; dispid 602;
    procedure ItalicRun; dispid 603;
    procedure InsertDateTime(const DateTimeFormat: WideString; InsertAsField: WordBool; 
                             InsertAsFullWidth: WordBool; DateLanguage: WpsDateLanguage; 
                             CalendarType: WpsCalendarTypeBi); dispid 444;
    procedure PasteAppendTable; dispid 1010;
    procedure InsertCaption(var Label_: OleVariant; const Title: WideString; 
                            const TitleAutoText: WideString; Position: WpsCaptionPosition; 
                            ExcludeLabel: WordBool); dispid 417;
    procedure ClearFormatting; dispid 1009;
    property Find: Find readonly dispid 262;
    property WordStat: WordStat readonly dispid 4096;
    property HeaderFooter: HeaderFooter readonly dispid 306;
    property Case_: WpsCharacterCase dispid 307;
    procedure EnableDropCap(Position: WpsDropPosition; const FontName: WideString; 
                            LineToDrop: Integer; DistanceFromText: Single); dispid 4404;
    property DropCap: DropCap readonly dispid 4405;
    property HighlightColor: WpsColor dispid 4397;
    procedure ApplyListTemplate(const ListTemplate: ListTemplate; 
                                ContinuePreviousList: WpsContinue; ApplyTo: WpsListApplyTo; 
                                DefaultListBehavior: Integer; UseContinueType: WpsListGalleryType; 
                                l1BasedContinueLevel: Integer); dispid 4406;
    property ListType: WpsListType readonly dispid 4407;
    property ListTemplate: ListTemplate readonly dispid 4408;
    function CanContinuePreviousList(const ListTemplate: ListTemplate): WpsContinue; dispid 4409;
    procedure RemoveNumbers(NumberType: WpsNumberType; ApplyTo: WpsListApplyTo); dispid 4416;
    procedure PasteAndFormat(Type_: WpsRecoveryType); dispid 412;
    procedure TCSCConverter(WpsTCSCCDirection: WpsTCSCConverterDirection; CommonTerms: WordBool; 
                            UseVariants: WordBool); dispid 674;
    procedure InsertFile(const FileName: WideString; var Range: OleVariant; 
                         var ConfirmConversions: OleVariant; var Link: OleVariant; 
                         var Attachment: OleVariant); dispid 123;
    property EnhMetaFileBits: OleVariant readonly dispid 315;
    property LanguageID: WpsLanguageID dispid 153;
    property LanguageIDFarEast: WpsLanguageID dispid 154;
    property LanguageIDOther: WpsLanguageID dispid 155;
    function ConvertToTable(var Separator: OleVariant; var NumRows: OleVariant; 
                            var NumColumns: OleVariant; var InitialColumnWidth: OleVariant; 
                            var Format: OleVariant; var ApplyBorders: OleVariant; 
                            var ApplyShading: OleVariant; var ApplyFont: OleVariant; 
                            var ApplyColor: OleVariant; var ApplyHeadingRows: OleVariant; 
                            var ApplyLastRow: OleVariant; var ApplyFirstColumn: OleVariant; 
                            var ApplyLastColumn: OleVariant; var AutoFit: OleVariant; 
                            var AutoFitBehavior: OleVariant; var DefaultTableBehavior: OleVariant): Table; dispid 457;
    property FormFields: FormFields readonly dispid 65;
    property Editors: Editors readonly dispid 313;
    procedure InsertCrossReference(var ReferenceType: OleVariant; ReferenceKind: WpsReferenceKind; 
                                   var ReferenceItem: OleVariant; 
                                   var InsertAsHyperlink: OleVariant; 
                                   var IncludePosition: OleVariant; 
                                   var SeparateNumbers: OleVariant; var SeparatorString: OleVariant); dispid 418;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Find
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B0-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Find = interface(_IKsoDispObj)
    ['{000209B0-0000-4B30-A977-D214852036FE}']
    function Get_Forward: WordBool; safecall;
    procedure Set_Forward(prop: WordBool); safecall;
    function Get_Font: _Font; safecall;
    procedure Set_Font(const prop: _Font); safecall;
    function Get_Found: WordBool; safecall;
    function Get_MatchAllWordForms: WordBool; safecall;
    procedure Set_MatchAllWordForms(prop: WordBool); safecall;
    function Get_MatchCase: WordBool; safecall;
    procedure Set_MatchCase(prop: WordBool); safecall;
    function Get_MatchWildcards: WordBool; safecall;
    procedure Set_MatchWildcards(prop: WordBool); safecall;
    function Get_MatchWholeWord: WordBool; safecall;
    procedure Set_MatchWholeWord(prop: WordBool); safecall;
    function Get_MatchByte: WordBool; safecall;
    procedure Set_MatchByte(prop: WordBool); safecall;
    function Get_ParagraphFormat: _ParagraphFormat; safecall;
    procedure Set_ParagraphFormat(const prop: _ParagraphFormat); safecall;
    function Get_Style: OleVariant; safecall;
    procedure Set_Style(var prop: OleVariant); safecall;
    function Get_Text: WideString; safecall;
    procedure Set_Text(const prop: WideString); safecall;
    function Get_Highlight: Integer; safecall;
    procedure Set_Highlight(prop: Integer); safecall;
    function Get_Replacement: Replacement; safecall;
    function Get_Frame: Frame; safecall;
    function Get_Wrap: WpsFindWrap; safecall;
    procedure Set_Wrap(prop: WpsFindWrap); safecall;
    function Get_Format: WordBool; safecall;
    procedure Set_Format(prop: WordBool); safecall;
    function Get_MatchControl: WordBool; safecall;
    procedure Set_MatchControl(prop: WordBool); safecall;
    procedure ClearFormatting; safecall;
    function Execute(var FindText: OleVariant; var MatchCase: OleVariant; 
                     var MatchWholeWord: OleVariant; var MatchWildcards: OleVariant; 
                     var MatchSoundsLike: OleVariant; var MatchAllWordForms: OleVariant; 
                     var Forward: OleVariant; var Wrap: OleVariant; var Format: OleVariant; 
                     var ReplaceWith: OleVariant; var Replace: OleVariant; 
                     var MatchKashida: OleVariant; var MatchDiacritics: OleVariant; 
                     var MatchAlefHamza: OleVariant; var MatchControl: OleVariant): WordBool; safecall;
    function Get_Scope: WpsFindScope; safecall;
    procedure Set_Scope(prop: WpsFindScope); safecall;
    function Get_SelectionMatch: WordBool; safecall;
    function Get_MatchFuzzy: WordBool; safecall;
    procedure Set_MatchFuzzy(prop: WordBool); safecall;
    function Get_MatchSoundsLike: WordBool; safecall;
    procedure Set_MatchSoundsLike(prop: WordBool); safecall;
    property Forward: WordBool read Get_Forward write Set_Forward;
    property Font: _Font read Get_Font write Set_Font;
    property Found: WordBool read Get_Found;
    property MatchAllWordForms: WordBool read Get_MatchAllWordForms write Set_MatchAllWordForms;
    property MatchCase: WordBool read Get_MatchCase write Set_MatchCase;
    property MatchWildcards: WordBool read Get_MatchWildcards write Set_MatchWildcards;
    property MatchWholeWord: WordBool read Get_MatchWholeWord write Set_MatchWholeWord;
    property MatchByte: WordBool read Get_MatchByte write Set_MatchByte;
    property ParagraphFormat: _ParagraphFormat read Get_ParagraphFormat write Set_ParagraphFormat;
    property Text: WideString read Get_Text write Set_Text;
    property Highlight: Integer read Get_Highlight write Set_Highlight;
    property Replacement: Replacement read Get_Replacement;
    property Frame: Frame read Get_Frame;
    property Wrap: WpsFindWrap read Get_Wrap write Set_Wrap;
    property Format: WordBool read Get_Format write Set_Format;
    property MatchControl: WordBool read Get_MatchControl write Set_MatchControl;
    property Scope: WpsFindScope read Get_Scope write Set_Scope;
    property SelectionMatch: WordBool read Get_SelectionMatch;
    property MatchFuzzy: WordBool read Get_MatchFuzzy write Set_MatchFuzzy;
    property MatchSoundsLike: WordBool read Get_MatchSoundsLike write Set_MatchSoundsLike;
  end;

// *********************************************************************//
// DispIntf:  FindDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B0-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FindDisp = dispinterface
    ['{000209B0-0000-4B30-A977-D214852036FE}']
    property Forward: WordBool dispid 10;
    property Font: _Font dispid 11;
    property Found: WordBool readonly dispid 12;
    property MatchAllWordForms: WordBool dispid 13;
    property MatchCase: WordBool dispid 14;
    property MatchWildcards: WordBool dispid 15;
    property MatchWholeWord: WordBool dispid 17;
    property MatchByte: WordBool dispid 41;
    property ParagraphFormat: _ParagraphFormat dispid 18;
    function Style: OleVariant; dispid 19;
    property Text: WideString dispid 22;
    property Highlight: Integer dispid 24;
    property Replacement: Replacement readonly dispid 25;
    property Frame: Frame readonly dispid 26;
    property Wrap: WpsFindWrap dispid 27;
    property Format: WordBool dispid 28;
    property MatchControl: WordBool dispid 103;
    procedure ClearFormatting; dispid 31;
    function Execute(var FindText: OleVariant; var MatchCase: OleVariant; 
                     var MatchWholeWord: OleVariant; var MatchWildcards: OleVariant; 
                     var MatchSoundsLike: OleVariant; var MatchAllWordForms: OleVariant; 
                     var Forward: OleVariant; var Wrap: OleVariant; var Format: OleVariant; 
                     var ReplaceWith: OleVariant; var Replace: OleVariant; 
                     var MatchKashida: OleVariant; var MatchDiacritics: OleVariant; 
                     var MatchAlefHamza: OleVariant; var MatchControl: OleVariant): WordBool; dispid 444;
    property Scope: WpsFindScope dispid 111;
    property SelectionMatch: WordBool readonly dispid 112;
    property MatchFuzzy: WordBool dispid 29;
    property MatchSoundsLike: WordBool dispid 16;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Replacement
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B1-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Replacement = interface(_IKsoDispObj)
    ['{000209B1-0000-4B30-A977-D214852036FE}']
    function Get_Text: WideString; safecall;
    procedure Set_Text(const prop: WideString); safecall;
    procedure ClearFormatting; safecall;
    property Text: WideString read Get_Text write Set_Text;
  end;

// *********************************************************************//
// DispIntf:  ReplacementDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B1-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ReplacementDisp = dispinterface
    ['{000209B1-0000-4B30-A977-D214852036FE}']
    property Text: WideString dispid 15;
    procedure ClearFormatting; dispid 20;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Zooms
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A7-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Zooms = interface(_IKsoDispObj)
    ['{000209A7-0000-4B30-A977-D214852036FE}']
    function Item(Index: WpsViewType): Zoom; safecall;
  end;

// *********************************************************************//
// DispIntf:  ZoomsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A7-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ZoomsDisp = dispinterface
    ['{000209A7-0000-4B30-A977-D214852036FE}']
    function Item(Index: WpsViewType): Zoom; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Zoom
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A6-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Zoom = interface(_IKsoDispObj)
    ['{000209A6-0000-4B30-A977-D214852036FE}']
    function Get_Percentage: Integer; safecall;
    procedure Set_Percentage(prop: Integer); safecall;
    function Get_PageFit: WpsPageFit; safecall;
    procedure Set_PageFit(prop: WpsPageFit); safecall;
    function Get_PageRows: Integer; safecall;
    procedure Set_PageRows(prop: Integer); safecall;
    function Get_PageColumns: Integer; safecall;
    procedure Set_PageColumns(prop: Integer); safecall;
    property Percentage: Integer read Get_Percentage write Set_Percentage;
    property PageFit: WpsPageFit read Get_PageFit write Set_PageFit;
    property PageRows: Integer read Get_PageRows write Set_PageRows;
    property PageColumns: Integer read Get_PageColumns write Set_PageColumns;
  end;

// *********************************************************************//
// DispIntf:  ZoomDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A6-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ZoomDisp = dispinterface
    ['{000209A6-0000-4B30-A977-D214852036FE}']
    property Percentage: Integer dispid 0;
    property PageFit: WpsPageFit dispid 1;
    property PageRows: Integer dispid 2;
    property PageColumns: Integer dispid 3;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: View
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A5-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  View = interface(_IKsoDispObj)
    ['{000209A5-0000-4B30-A977-D214852036FE}']
    function Get_type_: WpsViewType; safecall;
    procedure Set_type_(prop: WpsViewType); safecall;
    function Get_FullScreen: WordBool; safecall;
    procedure Set_FullScreen(prop: WordBool); safecall;
    function Get_ShowAll: WordBool; safecall;
    procedure Set_ShowAll(prop: WordBool); safecall;
    function Get_ShowFieldCodes: WordBool; safecall;
    procedure Set_ShowFieldCodes(prop: WordBool); safecall;
    function Get_Zoom: Zoom; safecall;
    function Get_ShowObjectAnchors: WordBool; safecall;
    procedure Set_ShowObjectAnchors(prop: WordBool); safecall;
    function Get_ShowTextBoundaries: WordBool; safecall;
    procedure Set_ShowTextBoundaries(prop: WordBool); safecall;
    function Get_ShowHighlight: WordBool; safecall;
    procedure Set_ShowHighlight(prop: WordBool); safecall;
    function Get_ShowDrawings: WordBool; safecall;
    procedure Set_ShowDrawings(prop: WordBool); safecall;
    function Get_ShowTabs: WordBool; safecall;
    procedure Set_ShowTabs(prop: WordBool); safecall;
    function Get_ShowSpaces: WordBool; safecall;
    procedure Set_ShowSpaces(prop: WordBool); safecall;
    function Get_ShowParagraphs: WordBool; safecall;
    procedure Set_ShowParagraphs(prop: WordBool); safecall;
    function Get_ShowHiddenText: WordBool; safecall;
    procedure Set_ShowHiddenText(prop: WordBool); safecall;
    function Get_ShowPicturePlaceHolders: WordBool; safecall;
    procedure Set_ShowPicturePlaceHolders(prop: WordBool); safecall;
    function Get_ShowBookmarks: WordBool; safecall;
    procedure Set_ShowBookmarks(prop: WordBool); safecall;
    function Get_FieldShading: WpsFieldShading; safecall;
    procedure Set_FieldShading(prop: WpsFieldShading); safecall;
    function Get_TableGridlines: WordBool; safecall;
    procedure Set_TableGridlines(prop: WordBool); safecall;
    function Get_SeekView: WpsSeekView; safecall;
    procedure Set_SeekView(prop: WpsSeekView); safecall;
    function Get_ShowOptionalBreaks: WordBool; safecall;
    procedure Set_ShowOptionalBreaks(prop: WordBool); safecall;
    function Get_DisplayPageBoundaries: WordBool; safecall;
    procedure Set_DisplayPageBoundaries(prop: WordBool); safecall;
    function Get_ShowRevisionsAndComments: WordBool; safecall;
    procedure Set_ShowRevisionsAndComments(prop: WordBool); safecall;
    function Get_ShowComments: WordBool; safecall;
    procedure Set_ShowComments(prop: WordBool); safecall;
    function Get_ShowInsertionsAndDeletions: WordBool; safecall;
    procedure Set_ShowInsertionsAndDeletions(prop: WordBool); safecall;
    function Get_ShowFormatChanges: WordBool; safecall;
    procedure Set_ShowFormatChanges(prop: WordBool); safecall;
    function Get_RevisionsView: WpsRevisionsView; safecall;
    procedure Set_RevisionsView(prop: WpsRevisionsView); safecall;
    function Get_RevisionsMode: WpsRevisionsMode; safecall;
    procedure Set_RevisionsMode(prop: WpsRevisionsMode); safecall;
    function Get_RevisionsBalloonWidth: Single; safecall;
    procedure Set_RevisionsBalloonWidth(prop: Single); safecall;
    function Get_RevisionsBalloonWidthType: WpsRevisionsBalloonWidthType; safecall;
    procedure Set_RevisionsBalloonWidthType(prop: WpsRevisionsBalloonWidthType); safecall;
    function Get_RevisionsBalloonSide: WpsRevisionsBalloonMargin; safecall;
    procedure Set_RevisionsBalloonSide(prop: WpsRevisionsBalloonMargin); safecall;
    function Get_Reviewers: Reviewers; safecall;
    function Get_RevisionsBalloonShowConnectingLines: WordBool; safecall;
    procedure Set_RevisionsBalloonShowConnectingLines(prop: WordBool); safecall;
    procedure PreviousHeaderFooter; safecall;
    procedure NextHeaderFooter; safecall;
    function Get_Selection: Selection; safecall;
    procedure ScrollIntoView(const obj: IDispatch; var Start: OleVariant); safecall;
    function Get_PageLayoutOrientation: WpsPagesOrientation; safecall;
    procedure Set_PageLayoutOrientation(prop: WpsPagesOrientation); safecall;
    function Get_ShowBreakHolder: WordBool; safecall;
    procedure Set_ShowBreakHolder(prop: WordBool); safecall;
    procedure CollapseOutline(const prop: Range); safecall;
    procedure ExpandOutline(const prop: Range); safecall;
    procedure ShowHeading(Level: Integer); safecall;
    function Get_ShowFirstLineOnly: WordBool; safecall;
    procedure Set_ShowFirstLineOnly(prop: WordBool); safecall;
    function Get_ShowFormat: WordBool; safecall;
    procedure Set_ShowFormat(prop: WordBool); safecall;
    procedure ShowAllHeadings; safecall;
    property type_: WpsViewType read Get_type_ write Set_type_;
    property FullScreen: WordBool read Get_FullScreen write Set_FullScreen;
    property ShowAll: WordBool read Get_ShowAll write Set_ShowAll;
    property ShowFieldCodes: WordBool read Get_ShowFieldCodes write Set_ShowFieldCodes;
    property Zoom: Zoom read Get_Zoom;
    property ShowObjectAnchors: WordBool read Get_ShowObjectAnchors write Set_ShowObjectAnchors;
    property ShowTextBoundaries: WordBool read Get_ShowTextBoundaries write Set_ShowTextBoundaries;
    property ShowHighlight: WordBool read Get_ShowHighlight write Set_ShowHighlight;
    property ShowDrawings: WordBool read Get_ShowDrawings write Set_ShowDrawings;
    property ShowTabs: WordBool read Get_ShowTabs write Set_ShowTabs;
    property ShowSpaces: WordBool read Get_ShowSpaces write Set_ShowSpaces;
    property ShowParagraphs: WordBool read Get_ShowParagraphs write Set_ShowParagraphs;
    property ShowHiddenText: WordBool read Get_ShowHiddenText write Set_ShowHiddenText;
    property ShowPicturePlaceHolders: WordBool read Get_ShowPicturePlaceHolders write Set_ShowPicturePlaceHolders;
    property ShowBookmarks: WordBool read Get_ShowBookmarks write Set_ShowBookmarks;
    property FieldShading: WpsFieldShading read Get_FieldShading write Set_FieldShading;
    property TableGridlines: WordBool read Get_TableGridlines write Set_TableGridlines;
    property SeekView: WpsSeekView read Get_SeekView write Set_SeekView;
    property ShowOptionalBreaks: WordBool read Get_ShowOptionalBreaks write Set_ShowOptionalBreaks;
    property DisplayPageBoundaries: WordBool read Get_DisplayPageBoundaries write Set_DisplayPageBoundaries;
    property ShowRevisionsAndComments: WordBool read Get_ShowRevisionsAndComments write Set_ShowRevisionsAndComments;
    property ShowComments: WordBool read Get_ShowComments write Set_ShowComments;
    property ShowInsertionsAndDeletions: WordBool read Get_ShowInsertionsAndDeletions write Set_ShowInsertionsAndDeletions;
    property ShowFormatChanges: WordBool read Get_ShowFormatChanges write Set_ShowFormatChanges;
    property RevisionsView: WpsRevisionsView read Get_RevisionsView write Set_RevisionsView;
    property RevisionsMode: WpsRevisionsMode read Get_RevisionsMode write Set_RevisionsMode;
    property RevisionsBalloonWidth: Single read Get_RevisionsBalloonWidth write Set_RevisionsBalloonWidth;
    property RevisionsBalloonWidthType: WpsRevisionsBalloonWidthType read Get_RevisionsBalloonWidthType write Set_RevisionsBalloonWidthType;
    property RevisionsBalloonSide: WpsRevisionsBalloonMargin read Get_RevisionsBalloonSide write Set_RevisionsBalloonSide;
    property Reviewers: Reviewers read Get_Reviewers;
    property RevisionsBalloonShowConnectingLines: WordBool read Get_RevisionsBalloonShowConnectingLines write Set_RevisionsBalloonShowConnectingLines;
    property Selection: Selection read Get_Selection;
    property PageLayoutOrientation: WpsPagesOrientation read Get_PageLayoutOrientation write Set_PageLayoutOrientation;
    property ShowBreakHolder: WordBool read Get_ShowBreakHolder write Set_ShowBreakHolder;
    property ShowFirstLineOnly: WordBool read Get_ShowFirstLineOnly write Set_ShowFirstLineOnly;
    property ShowFormat: WordBool read Get_ShowFormat write Set_ShowFormat;
  end;

// *********************************************************************//
// DispIntf:  ViewDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A5-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ViewDisp = dispinterface
    ['{000209A5-0000-4B30-A977-D214852036FE}']
    property type_: WpsViewType dispid 0;
    property FullScreen: WordBool dispid 1;
    property ShowAll: WordBool dispid 3;
    property ShowFieldCodes: WordBool dispid 4;
    property Zoom: Zoom readonly dispid 10;
    property ShowObjectAnchors: WordBool dispid 11;
    property ShowTextBoundaries: WordBool dispid 12;
    property ShowHighlight: WordBool dispid 13;
    property ShowDrawings: WordBool dispid 14;
    property ShowTabs: WordBool dispid 15;
    property ShowSpaces: WordBool dispid 16;
    property ShowParagraphs: WordBool dispid 17;
    property ShowHiddenText: WordBool dispid 19;
    property ShowPicturePlaceHolders: WordBool dispid 21;
    property ShowBookmarks: WordBool dispid 22;
    property FieldShading: WpsFieldShading dispid 23;
    property TableGridlines: WordBool dispid 25;
    property SeekView: WpsSeekView dispid 28;
    property ShowOptionalBreaks: WordBool dispid 31;
    property DisplayPageBoundaries: WordBool dispid 32;
    property ShowRevisionsAndComments: WordBool dispid 34;
    property ShowComments: WordBool dispid 35;
    property ShowInsertionsAndDeletions: WordBool dispid 36;
    property ShowFormatChanges: WordBool dispid 37;
    property RevisionsView: WpsRevisionsView dispid 38;
    property RevisionsMode: WpsRevisionsMode dispid 39;
    property RevisionsBalloonWidth: Single dispid 40;
    property RevisionsBalloonWidthType: WpsRevisionsBalloonWidthType dispid 41;
    property RevisionsBalloonSide: WpsRevisionsBalloonMargin dispid 42;
    property Reviewers: Reviewers readonly dispid 43;
    property RevisionsBalloonShowConnectingLines: WordBool dispid 44;
    procedure PreviousHeaderFooter; dispid 105;
    procedure NextHeaderFooter; dispid 106;
    property Selection: Selection readonly dispid 4097;
    procedure ScrollIntoView(const obj: IDispatch; var Start: OleVariant); dispid 111;
    property PageLayoutOrientation: WpsPagesOrientation dispid 45;
    property ShowBreakHolder: WordBool dispid 18;
    procedure CollapseOutline(const prop: Range); dispid 256;
    procedure ExpandOutline(const prop: Range); dispid 257;
    procedure ShowHeading(Level: Integer); dispid 258;
    property ShowFirstLineOnly: WordBool dispid 259;
    property ShowFormat: WordBool dispid 260;
    procedure ShowAllHeadings; dispid 261;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Reviewers
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020C9A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Reviewers = interface(_IKsoDispObj)
    ['{00020C9A-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): Reviewer; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  ReviewersDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020C9A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ReviewersDisp = dispinterface
    ['{00020C9A-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    function Item(var Index: OleVariant): Reviewer; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Reviewer
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000204AE-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Reviewer = interface(_IKsoDispObj)
    ['{000204AE-0000-4B30-A977-D214852036FE}']
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(prop: WordBool); safecall;
    function Get_Name: WideString; safecall;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property Name: WideString read Get_Name;
  end;

// *********************************************************************//
// DispIntf:  ReviewerDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000204AE-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ReviewerDisp = dispinterface
    ['{000204AE-0000-4B30-A977-D214852036FE}']
    property Visible: WordBool dispid 0;
    property Name: WideString readonly dispid 1;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Pages
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00027402-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Pages = interface(_IKsoDispObj)
    ['{00027402-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): Page; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  PagesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00027402-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  PagesDisp = dispinterface
    ['{00027402-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    function Item(Index: Integer): Page; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Page
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000240A9-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Page = interface(_IKsoDispObj)
    ['{000240A9-0000-4B30-A977-D214852036FE}']
    function Get_Left: Integer; safecall;
    function Get_Top: Integer; safecall;
    function Get_Width: Integer; safecall;
    function Get_Height: Integer; safecall;
    function Get_Breaks: Breaks; safecall;
    function Get_Index: Integer; safecall;
    function Get_Number: Integer; safecall;
    function Get_NumberName: WideString; safecall;
    procedure ExportToEmf(const bstrFileName: WideString); safecall;
    property Left: Integer read Get_Left;
    property Top: Integer read Get_Top;
    property Width: Integer read Get_Width;
    property Height: Integer read Get_Height;
    property Breaks: Breaks read Get_Breaks;
    property Index: Integer read Get_Index;
    property Number: Integer read Get_Number;
    property NumberName: WideString read Get_NumberName;
  end;

// *********************************************************************//
// DispIntf:  PageDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000240A9-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  PageDisp = dispinterface
    ['{000240A9-0000-4B30-A977-D214852036FE}']
    property Left: Integer readonly dispid 2;
    property Top: Integer readonly dispid 3;
    property Width: Integer readonly dispid 4;
    property Height: Integer readonly dispid 5;
    property Breaks: Breaks readonly dispid 7;
    property Index: Integer readonly dispid 4097;
    property Number: Integer readonly dispid 4098;
    property NumberName: WideString readonly dispid 4099;
    procedure ExportToEmf(const bstrFileName: WideString); dispid 4100;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Breaks
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00029309-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Breaks = interface(_IKsoDispObj)
    ['{00029309-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): Break; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  BreaksDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00029309-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  BreaksDisp = dispinterface
    ['{00029309-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    function Item(Index: Integer): Break; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Break
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00025BF1-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Break = interface(_IKsoDispObj)
    ['{00025BF1-0000-4B30-A977-D214852036FE}']
    function Get_Range: Range; safecall;
    function Get_PageIndex: Integer; safecall;
    property Range: Range read Get_Range;
    property PageIndex: Integer read Get_PageIndex;
  end;

// *********************************************************************//
// DispIntf:  BreakDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00025BF1-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  BreakDisp = dispinterface
    ['{00025BF1-0000-4B30-A977-D214852036FE}']
    property Range: Range readonly dispid 2;
    property PageIndex: Integer readonly dispid 3;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Panes
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Panes = interface(_IKsoDispObj)
    ['{0002095F-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): Pane; safecall;
    function Add(SplitVertical: Integer): Pane; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  PanesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002095F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  PanesDisp = dispinterface
    ['{0002095F-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    function Item(Index: Integer): Pane; dispid 0;
    function Add(SplitVertical: Integer): Pane; dispid 3;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: ListTemplates
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020990-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListTemplates = interface(_IKsoDispObj)
    ['{00020990-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): ListTemplate; safecall;
    function Add(OutlineNumbered: WordBool; const Name: WideString): ListTemplate; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  ListTemplatesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020990-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListTemplatesDisp = dispinterface
    ['{00020990-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    function Item(Index: Integer): ListTemplate; dispid 0;
    function Add(OutlineNumbered: WordBool; const Name: WideString): ListTemplate; dispid 100;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Lists
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020993-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Lists = interface(_IKsoDispObj)
    ['{00020993-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): List; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  ListsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020993-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListsDisp = dispinterface
    ['{00020993-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    function Item(Index: Integer): List; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: ExtraColors
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00027778-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ExtraColors = interface(_IKsoDispObj)
    ['{00027778-0000-4B30-A977-D214852036FE}']
    function Get_Count: Integer; safecall;
    procedure Add(val: Integer); safecall;
    function Item(Index: OleVariant): Integer; safecall;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  ExtraColorsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00027778-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ExtraColorsDisp = dispinterface
    ['{00027778-0000-4B30-A977-D214852036FE}']
    property Count: Integer readonly dispid 509743104;
    procedure Add(val: Integer); dispid 509747200;
    function Item(Index: OleVariant): Integer; dispid 509747201;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Variables
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020965-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Variables = interface(_IKsoDispObj)
    ['{00020965-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): Variable; safecall;
    function Add(const Name: WideString; var Value: OleVariant): Variable; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  VariablesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020965-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  VariablesDisp = dispinterface
    ['{00020965-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    function Item(var Index: OleVariant): Variable; dispid 0;
    function Add(const Name: WideString; var Value: OleVariant): Variable; dispid 7;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Variable
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020966-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Variable = interface(_IKsoDispObj)
    ['{00020966-0000-4B30-A977-D214852036FE}']
    function Get_Name: WideString; safecall;
    function Get_Value: WideString; safecall;
    procedure Set_Value(const prop: WideString); safecall;
    function Get_Index: Integer; safecall;
    procedure Delete; safecall;
    property Name: WideString read Get_Name;
    property Value: WideString read Get_Value write Set_Value;
    property Index: Integer read Get_Index;
  end;

// *********************************************************************//
// DispIntf:  VariableDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020966-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  VariableDisp = dispinterface
    ['{00020966-0000-4B30-A977-D214852036FE}']
    property Name: WideString readonly dispid 1;
    property Value: WideString dispid 0;
    property Index: Integer readonly dispid 2;
    procedure Delete; dispid 11;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: _WebOptions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020AA1-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  _WebOptions = interface(_IKsoDispObj)
    ['{00020AA1-0000-4B30-A977-D214852036FE}']
    procedure Set_Encoding(prop: KsoEncoding); safecall;
    function Get_Encoding: KsoEncoding; safecall;
    property Encoding: KsoEncoding read Get_Encoding write Set_Encoding;
  end;

// *********************************************************************//
// DispIntf:  _WebOptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020AA1-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  _WebOptionsDisp = dispinterface
    ['{00020AA1-0000-4B30-A977-D214852036FE}']
    property Encoding: KsoEncoding dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: ListGalleries
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020995-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListGalleries = interface(_IKsoDispObj)
    ['{00020995-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: WpsListGalleryType): ListGallery; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  ListGalleriesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020995-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ListGalleriesDisp = dispinterface
    ['{00020995-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 2;
    function Item(Index: WpsListGalleryType): ListGallery; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: ListGallery
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020994-0000-0000-C000-000000000046}
// *********************************************************************//
  ListGallery = interface(IDispatch)
    ['{00020994-0000-0000-C000-000000000046}']
    function Get_ListTemplates: ListTemplates; safecall;
    function Get_Application: _Application; safecall;
    function Get_Creator: Integer; safecall;
    function Get_Parent: IDispatch; safecall;
    function Get_Modified(Index: Integer): WordBool; safecall;
    procedure Reset(Index: Integer); safecall;
    property ListTemplates: ListTemplates read Get_ListTemplates;
    property Application: _Application read Get_Application;
    property Creator: Integer read Get_Creator;
    property Parent: IDispatch read Get_Parent;
    property Modified[Index: Integer]: WordBool read Get_Modified;
  end;

// *********************************************************************//
// DispIntf:  ListGalleryDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020994-0000-0000-C000-000000000046}
// *********************************************************************//
  ListGalleryDisp = dispinterface
    ['{00020994-0000-0000-C000-000000000046}']
    property ListTemplates: ListTemplates readonly dispid 1;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    property Modified[Index: Integer]: WordBool readonly dispid 101;
    procedure Reset(Index: Integer); dispid 100;
  end;

// *********************************************************************//
// Interface: FileConverters
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FileConverters = interface(_IKsoDispObj)
    ['{0002099A-0000-4B30-A977-D214852036FE}']
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Item(var Index: OleVariant): FileConverter; safecall;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  FileConvertersDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002099A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FileConvertersDisp = dispinterface
    ['{0002099A-0000-4B30-A977-D214852036FE}']
    property Count: Integer readonly dispid 1;
    property _NewEnum: IUnknown readonly dispid -4;
    function Item(var Index: OleVariant): FileConverter; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: FileConverter
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020999-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FileConverter = interface(_IKsoDispObj)
    ['{00020999-0000-4B30-A977-D214852036FE}']
    function Get_FormatName: WideString; safecall;
    function Get__className: WideString; safecall;
    function Get_SaveFormat: Integer; safecall;
    function Get_OpenFormat: Integer; safecall;
    function Get_CanSave: WordBool; safecall;
    function Get_CanOpen: WordBool; safecall;
    function Get_Extensions: WideString; safecall;
    property FormatName: WideString read Get_FormatName;
    property _className: WideString read Get__className;
    property SaveFormat: Integer read Get_SaveFormat;
    property OpenFormat: Integer read Get_OpenFormat;
    property CanSave: WordBool read Get_CanSave;
    property CanOpen: WordBool read Get_CanOpen;
    property Extensions: WideString read Get_Extensions;
  end;

// *********************************************************************//
// DispIntf:  FileConverterDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020999-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FileConverterDisp = dispinterface
    ['{00020999-0000-4B30-A977-D214852036FE}']
    property FormatName: WideString readonly dispid 0;
    property _className: WideString readonly dispid 1;
    property SaveFormat: Integer readonly dispid 2;
    property OpenFormat: Integer readonly dispid 3;
    property CanSave: WordBool readonly dispid 4;
    property CanOpen: WordBool readonly dispid 5;
    property Extensions: WideString readonly dispid 8;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: CaptionLabels
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020978-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  CaptionLabels = interface(_IKsoDispObj)
    ['{00020978-0000-4B30-A977-D214852036FE}']
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): CaptionLabel; safecall;
    function Add(const Name: WideString): CaptionLabel; safecall;
    function Get__NewEnum: IUnknown; safecall;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  CaptionLabelsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020978-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  CaptionLabelsDisp = dispinterface
    ['{00020978-0000-4B30-A977-D214852036FE}']
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): CaptionLabel; dispid 0;
    function Add(const Name: WideString): CaptionLabel; dispid 100;
    property _NewEnum: IUnknown readonly dispid -4;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: CaptionLabel
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020979-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  CaptionLabel = interface(_IKsoDispObj)
    ['{00020979-0000-4B30-A977-D214852036FE}']
    function Get_Name: WideString; safecall;
    function Get_BuiltIn: Integer; safecall;
    function Get_ID: WpsCaptionLabelID; safecall;
    function Get_IncludeChapterNumber: Integer; safecall;
    procedure Set_IncludeChapterNumber(prop: Integer); safecall;
    function Get_NumberStyle: WpsCaptionNumberStyle; safecall;
    procedure Set_NumberStyle(prop: WpsCaptionNumberStyle); safecall;
    function Get_ChapterStyleLevel: Integer; safecall;
    procedure Set_ChapterStyleLevel(prop: Integer); safecall;
    function Get_Separator: WpsSeparatorType; safecall;
    procedure Set_Separator(prop: WpsSeparatorType); safecall;
    function Get_Position: WpsCaptionPosition; safecall;
    procedure Set_Position(prop: WpsCaptionPosition); safecall;
    procedure Delete; safecall;
    property Name: WideString read Get_Name;
    property BuiltIn: Integer read Get_BuiltIn;
    property ID: WpsCaptionLabelID read Get_ID;
    property IncludeChapterNumber: Integer read Get_IncludeChapterNumber write Set_IncludeChapterNumber;
    property NumberStyle: WpsCaptionNumberStyle read Get_NumberStyle write Set_NumberStyle;
    property ChapterStyleLevel: Integer read Get_ChapterStyleLevel write Set_ChapterStyleLevel;
    property Separator: WpsSeparatorType read Get_Separator write Set_Separator;
    property Position: WpsCaptionPosition read Get_Position write Set_Position;
  end;

// *********************************************************************//
// DispIntf:  CaptionLabelDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020979-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  CaptionLabelDisp = dispinterface
    ['{00020979-0000-4B30-A977-D214852036FE}']
    property Name: WideString readonly dispid 0;
    property BuiltIn: Integer readonly dispid 1;
    property ID: WpsCaptionLabelID readonly dispid 2;
    property IncludeChapterNumber: Integer dispid 3;
    property NumberStyle: WpsCaptionNumberStyle dispid 4;
    property ChapterStyleLevel: Integer dispid 5;
    property Separator: WpsSeparatorType dispid 6;
    property Position: WpsCaptionPosition dispid 7;
    procedure Delete; dispid 100;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Options
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B7-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Options = interface(_IKsoDispObj)
    ['{000209B7-0000-4B30-A977-D214852036FE}']
    function Get_AllowAccentedUppercase: WordBool; safecall;
    procedure Set_AllowAccentedUppercase(prop: WordBool); safecall;
    function Get_BlueScreen: WordBool; safecall;
    procedure Set_BlueScreen(prop: WordBool); safecall;
    function Get_MeasurementUnit: WpsMeasurementUnits; safecall;
    procedure Set_MeasurementUnit(prop: WpsMeasurementUnits); safecall;
    function Get_UpdateFieldsAtPrint: WordBool; safecall;
    procedure Set_UpdateFieldsAtPrint(prop: WordBool); safecall;
    function Get_PrintFieldCodes: WordBool; safecall;
    procedure Set_PrintFieldCodes(prop: WordBool); safecall;
    function Get_PrintComments: WordBool; safecall;
    procedure Set_PrintComments(prop: WordBool); safecall;
    function Get_PrintHiddenText: WordBool; safecall;
    procedure Set_PrintHiddenText(prop: WordBool); safecall;
    function Get_PrintDrawingObjects: WordBool; safecall;
    procedure Set_PrintDrawingObjects(prop: WordBool); safecall;
    function Get_CreateBackup: WordBool; safecall;
    procedure Set_CreateBackup(prop: WordBool); safecall;
    function Get_BackgroundSave: WordBool; safecall;
    procedure Set_BackgroundSave(prop: WordBool); safecall;
    function Get_DefaultFilePath(Path: WpsDefaultFilePath): WideString; safecall;
    procedure Set_DefaultFilePath(Path: WpsDefaultFilePath; const prop: WideString); safecall;
    function Get_ReplaceSelection: WordBool; safecall;
    procedure Set_ReplaceSelection(prop: WordBool); safecall;
    function Get_AllowDragAndDrop: WordBool; safecall;
    procedure Set_AllowDragAndDrop(prop: WordBool); safecall;
    function Get_TabIndentKey: WordBool; safecall;
    procedure Set_TabIndentKey(prop: WordBool); safecall;
    function Get_SnapToGrid: WordBool; safecall;
    procedure Set_SnapToGrid(prop: WordBool); safecall;
    function Get_GridDistanceHorizontal: Single; safecall;
    procedure Set_GridDistanceHorizontal(prop: Single); safecall;
    function Get_GridDistanceVertical: Single; safecall;
    procedure Set_GridDistanceVertical(prop: Single); safecall;
    function Get_GridOriginHorizontal: Single; safecall;
    procedure Set_GridOriginHorizontal(prop: Single); safecall;
    function Get_GridOriginVertical: Single; safecall;
    procedure Set_GridOriginVertical(prop: Single); safecall;
    function Get_DefaultHighlightColor: WpsColor; safecall;
    procedure Set_DefaultHighlightColor(prop: WpsColor); safecall;
    function Get_DefaultBorderLineStyle: WpsLineStyle; safecall;
    procedure Set_DefaultBorderLineStyle(prop: WpsLineStyle); safecall;
    function Get_IgnoreUppercase: WordBool; safecall;
    procedure Set_IgnoreUppercase(prop: WordBool); safecall;
    function Get_IgnoreMixedDigits: WordBool; safecall;
    procedure Set_IgnoreMixedDigits(prop: WordBool); safecall;
    function Get_DefaultBorderLineWidth: WpsLineWidth; safecall;
    procedure Set_DefaultBorderLineWidth(prop: WpsLineWidth); safecall;
    function Get_DefaultOpenFormat: WpsOpenFormat; safecall;
    procedure Set_DefaultOpenFormat(prop: WpsOpenFormat); safecall;
    function Get_MapPaperSize: WordBool; safecall;
    procedure Set_MapPaperSize(prop: WordBool); safecall;
    function Get_DisplayGridLines: WordBool; safecall;
    procedure Set_DisplayGridLines(prop: WordBool); safecall;
    function Get_DefaultBorderColor: WpsColor; safecall;
    procedure Set_DefaultBorderColor(prop: WpsColor); safecall;
    function Get_AllowPixelUnits: WordBool; safecall;
    procedure Set_AllowPixelUnits(prop: WordBool); safecall;
    function Get_UseCharacterUnit: WordBool; safecall;
    procedure Set_UseCharacterUnit(prop: WordBool); safecall;
    function Get_CtrlClickHyperlinkToOpen: WordBool; safecall;
    procedure Set_CtrlClickHyperlinkToOpen(prop: WordBool); safecall;
    function Get_PrintOddPagesInAscendingOrder: WordBool; safecall;
    procedure Set_PrintOddPagesInAscendingOrder(prop: WordBool); safecall;
    function Get_PrintEvenPagesInAscendingOrder: WordBool; safecall;
    procedure Set_PrintEvenPagesInAscendingOrder(prop: WordBool); safecall;
    function Get_PrintReverse: WordBool; safecall;
    procedure Set_PrintReverse(prop: WordBool); safecall;
    function Get_PrintHiddenTextMode: WpsPrintHiddenTextMode; safecall;
    procedure Set_PrintHiddenTextMode(prop: WpsPrintHiddenTextMode); safecall;
    function Get_InsertedTextMark: WpsInsertedTextMark; safecall;
    procedure Set_InsertedTextMark(prop: WpsInsertedTextMark); safecall;
    function Get_DeletedTextMark: WpsDeletedTextMark; safecall;
    procedure Set_DeletedTextMark(prop: WpsDeletedTextMark); safecall;
    function Get_RevisedLinesMark: WpsRevisedLinesMark; safecall;
    procedure Set_RevisedLinesMark(prop: WpsRevisedLinesMark); safecall;
    function Get_InsertedTextColor: WpsColor; safecall;
    procedure Set_InsertedTextColor(prop: WpsColor); safecall;
    function Get_DeletedTextColor: WpsColor; safecall;
    procedure Set_DeletedTextColor(prop: WpsColor); safecall;
    function Get_RevisedLinesColor: WpsColor; safecall;
    procedure Set_RevisedLinesColor(prop: WpsColor); safecall;
    function Get_RevisedPropertiesMark: WpsRevisedPropertiesMark; safecall;
    procedure Set_RevisedPropertiesMark(prop: WpsRevisedPropertiesMark); safecall;
    function Get_RevisedPropertiesColor: WpsColor; safecall;
    procedure Set_RevisedPropertiesColor(prop: WpsColor); safecall;
    function Get_RevisionsBalloonPrintOrientation: WpsRevisionsBalloonPrintOrientation; safecall;
    procedure Set_RevisionsBalloonPrintOrientation(prop: WpsRevisionsBalloonPrintOrientation); safecall;
    function Get_CommentsColor: WpsColor; safecall;
    procedure Set_CommentsColor(prop: WpsColor); safecall;
    function Get_RevisionsBalloonTitle: WpsRevisionsBalloonTitle; safecall;
    procedure Set_RevisionsBalloonTitle(prop: WpsRevisionsBalloonTitle); safecall;
    function Get_CreateBackupOnFirstSave: WordBool; safecall;
    procedure Set_CreateBackupOnFirstSave(prop: WordBool); safecall;
    function Get_SuggestFromMainDictionaryOnly: WordBool; safecall;
    procedure Set_SuggestFromMainDictionaryOnly(prop: WordBool); safecall;
    function Get_SuggestSpellingCorrections: WordBool; safecall;
    procedure Set_SuggestSpellingCorrections(prop: WordBool); safecall;
    function Get_SaveInterval: Integer; safecall;
    procedure Set_SaveInterval(prop: Integer); safecall;
    function Get_AllowClickAndTypeMouse: WordBool; safecall;
    procedure Set_AllowClickAndTypeMouse(prop: WordBool); safecall;
    function Get_DiscernPerson: WordBool; safecall;
    procedure Set_DiscernPerson(prop: WordBool); safecall;
    function Get_DiscernLocation: WordBool; safecall;
    procedure Set_DiscernLocation(prop: WordBool); safecall;
    function Get_MatchFuzzyCase: WordBool; safecall;
    procedure Set_MatchFuzzyCase(prop: WordBool); safecall;
    function Get_MatchFuzzyByte: WordBool; safecall;
    procedure Set_MatchFuzzyByte(prop: WordBool); safecall;
    function Get_MatchFuzzyHiragana: WordBool; safecall;
    procedure Set_MatchFuzzyHiragana(prop: WordBool); safecall;
    function Get_MatchFuzzySmallKana: WordBool; safecall;
    procedure Set_MatchFuzzySmallKana(prop: WordBool); safecall;
    function Get_MatchFuzzyDash: WordBool; safecall;
    procedure Set_MatchFuzzyDash(prop: WordBool); safecall;
    function Get_MatchFuzzyIterationMark: WordBool; safecall;
    procedure Set_MatchFuzzyIterationMark(prop: WordBool); safecall;
    function Get_MatchFuzzyKanji: WordBool; safecall;
    procedure Set_MatchFuzzyKanji(prop: WordBool); safecall;
    function Get_MatchFuzzyOldKana: WordBool; safecall;
    procedure Set_MatchFuzzyOldKana(prop: WordBool); safecall;
    function Get_MatchFuzzyProlongedSoundMark: WordBool; safecall;
    procedure Set_MatchFuzzyProlongedSoundMark(prop: WordBool); safecall;
    function Get_MatchFuzzyDZ: WordBool; safecall;
    procedure Set_MatchFuzzyDZ(prop: WordBool); safecall;
    function Get_MatchFuzzyBV: WordBool; safecall;
    procedure Set_MatchFuzzyBV(prop: WordBool); safecall;
    function Get_MatchFuzzyTC: WordBool; safecall;
    procedure Set_MatchFuzzyTC(prop: WordBool); safecall;
    function Get_MatchFuzzyHF: WordBool; safecall;
    procedure Set_MatchFuzzyHF(prop: WordBool); safecall;
    function Get_MatchFuzzyZJ: WordBool; safecall;
    procedure Set_MatchFuzzyZJ(prop: WordBool); safecall;
    function Get_MatchFuzzyAY: WordBool; safecall;
    procedure Set_MatchFuzzyAY(prop: WordBool); safecall;
    function Get_MatchFuzzyKiKu: WordBool; safecall;
    procedure Set_MatchFuzzyKiKu(prop: WordBool); safecall;
    function Get_MatchFuzzyPunctuation: WordBool; safecall;
    procedure Set_MatchFuzzyPunctuation(prop: WordBool); safecall;
    function Get_MatchFuzzySpace: WordBool; safecall;
    procedure Set_MatchFuzzySpace(prop: WordBool); safecall;
    function Get_Overtype: WordBool; safecall;
    procedure Set_Overtype(prop: WordBool); safecall;
    function Get_ShowDocumentMapOnLeft: WordBool; safecall;
    procedure Set_ShowDocumentMapOnLeft(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeApplyNumberedLists: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeApplyNumberedLists(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeReplaceOrdinals: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeReplaceOrdinals(prop: WordBool); safecall;
    function Get_AutoFormatAsYouTypeReplaceHyperlinks: WordBool; safecall;
    procedure Set_AutoFormatAsYouTypeReplaceHyperlinks(prop: WordBool); safecall;
    function Get_AutoApplyAsYouTypeAdjustWordWrap: WordBool; safecall;
    procedure Set_AutoApplyAsYouTypeAdjustWordWrap(prop: WordBool); safecall;
    function Get_CheckSpellingAsYouType: WordBool; safecall;
    procedure Set_CheckSpellingAsYouType(prop: WordBool); safecall;
    function Get_IgnoreInternetAndFileAddresses: WordBool; safecall;
    procedure Set_IgnoreInternetAndFileAddresses(prop: WordBool); safecall;
    function Get_DisplayPasteOptions: WordBool; safecall;
    procedure Set_DisplayPasteOptions(prop: WordBool); safecall;
    function Get_HideTaskPanelAfterDocumentOpen: WordBool; safecall;
    procedure Set_HideTaskPanelAfterDocumentOpen(prop: WordBool); safecall;
    function Get_EnableObjectBrowser: WordBool; safecall;
    procedure Set_EnableObjectBrowser(prop: WordBool); safecall;
    function Get_PasteFormatMode: WpsRecoveryType; safecall;
    procedure Set_PasteFormatMode(prop: WpsRecoveryType); safecall;
    function Get_SmartParaSelection: WordBool; safecall;
    procedure Set_SmartParaSelection(prop: WordBool); safecall;
    function Get_DocumentTabWidth: Integer; safecall;
    procedure Set_DocumentTabWidth(prop: Integer); safecall;
    function Get_DocumentTabPosition: WpsDocumentTabPosition; safecall;
    procedure Set_DocumentTabPosition(prop: WpsDocumentTabPosition); safecall;
    function Get_DocumentTabShowCloseBtn: WordBool; safecall;
    procedure Set_DocumentTabShowCloseBtn(prop: WordBool); safecall;
    function Get_DocumentTabDbClose: WordBool; safecall;
    procedure Set_DocumentTabDbClose(prop: WordBool); safecall;
    function Get_DocumentTabDbNew: WordBool; safecall;
    procedure Set_DocumentTabDbNew(prop: WordBool); safecall;
    function Get_DocumentTabSwitch: WordBool; safecall;
    procedure Set_DocumentTabSwitch(prop: WordBool); safecall;
    function Get_DocumentTabWarning: WordBool; safecall;
    procedure Set_DocumentTabWarning(prop: WordBool); safecall;
    function Get_DocumentViewDirection: WpsDocumentViewDirection; safecall;
    procedure Set_DocumentViewDirection(prop: WpsDocumentViewDirection); safecall;
    function Get_ShowTableAdjustLabel: WordBool; safecall;
    procedure Set_ShowTableAdjustLabel(prop: WordBool); safecall;
    property AllowAccentedUppercase: WordBool read Get_AllowAccentedUppercase write Set_AllowAccentedUppercase;
    property BlueScreen: WordBool read Get_BlueScreen write Set_BlueScreen;
    property MeasurementUnit: WpsMeasurementUnits read Get_MeasurementUnit write Set_MeasurementUnit;
    property UpdateFieldsAtPrint: WordBool read Get_UpdateFieldsAtPrint write Set_UpdateFieldsAtPrint;
    property PrintFieldCodes: WordBool read Get_PrintFieldCodes write Set_PrintFieldCodes;
    property PrintComments: WordBool read Get_PrintComments write Set_PrintComments;
    property PrintHiddenText: WordBool read Get_PrintHiddenText write Set_PrintHiddenText;
    property PrintDrawingObjects: WordBool read Get_PrintDrawingObjects write Set_PrintDrawingObjects;
    property CreateBackup: WordBool read Get_CreateBackup write Set_CreateBackup;
    property BackgroundSave: WordBool read Get_BackgroundSave write Set_BackgroundSave;
    property DefaultFilePath[Path: WpsDefaultFilePath]: WideString read Get_DefaultFilePath write Set_DefaultFilePath;
    property ReplaceSelection: WordBool read Get_ReplaceSelection write Set_ReplaceSelection;
    property AllowDragAndDrop: WordBool read Get_AllowDragAndDrop write Set_AllowDragAndDrop;
    property TabIndentKey: WordBool read Get_TabIndentKey write Set_TabIndentKey;
    property SnapToGrid: WordBool read Get_SnapToGrid write Set_SnapToGrid;
    property GridDistanceHorizontal: Single read Get_GridDistanceHorizontal write Set_GridDistanceHorizontal;
    property GridDistanceVertical: Single read Get_GridDistanceVertical write Set_GridDistanceVertical;
    property GridOriginHorizontal: Single read Get_GridOriginHorizontal write Set_GridOriginHorizontal;
    property GridOriginVertical: Single read Get_GridOriginVertical write Set_GridOriginVertical;
    property DefaultHighlightColor: WpsColor read Get_DefaultHighlightColor write Set_DefaultHighlightColor;
    property DefaultBorderLineStyle: WpsLineStyle read Get_DefaultBorderLineStyle write Set_DefaultBorderLineStyle;
    property IgnoreUppercase: WordBool read Get_IgnoreUppercase write Set_IgnoreUppercase;
    property IgnoreMixedDigits: WordBool read Get_IgnoreMixedDigits write Set_IgnoreMixedDigits;
    property DefaultBorderLineWidth: WpsLineWidth read Get_DefaultBorderLineWidth write Set_DefaultBorderLineWidth;
    property DefaultOpenFormat: WpsOpenFormat read Get_DefaultOpenFormat write Set_DefaultOpenFormat;
    property MapPaperSize: WordBool read Get_MapPaperSize write Set_MapPaperSize;
    property DisplayGridLines: WordBool read Get_DisplayGridLines write Set_DisplayGridLines;
    property DefaultBorderColor: WpsColor read Get_DefaultBorderColor write Set_DefaultBorderColor;
    property AllowPixelUnits: WordBool read Get_AllowPixelUnits write Set_AllowPixelUnits;
    property UseCharacterUnit: WordBool read Get_UseCharacterUnit write Set_UseCharacterUnit;
    property CtrlClickHyperlinkToOpen: WordBool read Get_CtrlClickHyperlinkToOpen write Set_CtrlClickHyperlinkToOpen;
    property PrintOddPagesInAscendingOrder: WordBool read Get_PrintOddPagesInAscendingOrder write Set_PrintOddPagesInAscendingOrder;
    property PrintEvenPagesInAscendingOrder: WordBool read Get_PrintEvenPagesInAscendingOrder write Set_PrintEvenPagesInAscendingOrder;
    property PrintReverse: WordBool read Get_PrintReverse write Set_PrintReverse;
    property PrintHiddenTextMode: WpsPrintHiddenTextMode read Get_PrintHiddenTextMode write Set_PrintHiddenTextMode;
    property InsertedTextMark: WpsInsertedTextMark read Get_InsertedTextMark write Set_InsertedTextMark;
    property DeletedTextMark: WpsDeletedTextMark read Get_DeletedTextMark write Set_DeletedTextMark;
    property RevisedLinesMark: WpsRevisedLinesMark read Get_RevisedLinesMark write Set_RevisedLinesMark;
    property InsertedTextColor: WpsColor read Get_InsertedTextColor write Set_InsertedTextColor;
    property DeletedTextColor: WpsColor read Get_DeletedTextColor write Set_DeletedTextColor;
    property RevisedLinesColor: WpsColor read Get_RevisedLinesColor write Set_RevisedLinesColor;
    property RevisedPropertiesMark: WpsRevisedPropertiesMark read Get_RevisedPropertiesMark write Set_RevisedPropertiesMark;
    property RevisedPropertiesColor: WpsColor read Get_RevisedPropertiesColor write Set_RevisedPropertiesColor;
    property RevisionsBalloonPrintOrientation: WpsRevisionsBalloonPrintOrientation read Get_RevisionsBalloonPrintOrientation write Set_RevisionsBalloonPrintOrientation;
    property CommentsColor: WpsColor read Get_CommentsColor write Set_CommentsColor;
    property RevisionsBalloonTitle: WpsRevisionsBalloonTitle read Get_RevisionsBalloonTitle write Set_RevisionsBalloonTitle;
    property CreateBackupOnFirstSave: WordBool read Get_CreateBackupOnFirstSave write Set_CreateBackupOnFirstSave;
    property SuggestFromMainDictionaryOnly: WordBool read Get_SuggestFromMainDictionaryOnly write Set_SuggestFromMainDictionaryOnly;
    property SuggestSpellingCorrections: WordBool read Get_SuggestSpellingCorrections write Set_SuggestSpellingCorrections;
    property SaveInterval: Integer read Get_SaveInterval write Set_SaveInterval;
    property AllowClickAndTypeMouse: WordBool read Get_AllowClickAndTypeMouse write Set_AllowClickAndTypeMouse;
    property DiscernPerson: WordBool read Get_DiscernPerson write Set_DiscernPerson;
    property DiscernLocation: WordBool read Get_DiscernLocation write Set_DiscernLocation;
    property MatchFuzzyCase: WordBool read Get_MatchFuzzyCase write Set_MatchFuzzyCase;
    property MatchFuzzyByte: WordBool read Get_MatchFuzzyByte write Set_MatchFuzzyByte;
    property MatchFuzzyHiragana: WordBool read Get_MatchFuzzyHiragana write Set_MatchFuzzyHiragana;
    property MatchFuzzySmallKana: WordBool read Get_MatchFuzzySmallKana write Set_MatchFuzzySmallKana;
    property MatchFuzzyDash: WordBool read Get_MatchFuzzyDash write Set_MatchFuzzyDash;
    property MatchFuzzyIterationMark: WordBool read Get_MatchFuzzyIterationMark write Set_MatchFuzzyIterationMark;
    property MatchFuzzyKanji: WordBool read Get_MatchFuzzyKanji write Set_MatchFuzzyKanji;
    property MatchFuzzyOldKana: WordBool read Get_MatchFuzzyOldKana write Set_MatchFuzzyOldKana;
    property MatchFuzzyProlongedSoundMark: WordBool read Get_MatchFuzzyProlongedSoundMark write Set_MatchFuzzyProlongedSoundMark;
    property MatchFuzzyDZ: WordBool read Get_MatchFuzzyDZ write Set_MatchFuzzyDZ;
    property MatchFuzzyBV: WordBool read Get_MatchFuzzyBV write Set_MatchFuzzyBV;
    property MatchFuzzyTC: WordBool read Get_MatchFuzzyTC write Set_MatchFuzzyTC;
    property MatchFuzzyHF: WordBool read Get_MatchFuzzyHF write Set_MatchFuzzyHF;
    property MatchFuzzyZJ: WordBool read Get_MatchFuzzyZJ write Set_MatchFuzzyZJ;
    property MatchFuzzyAY: WordBool read Get_MatchFuzzyAY write Set_MatchFuzzyAY;
    property MatchFuzzyKiKu: WordBool read Get_MatchFuzzyKiKu write Set_MatchFuzzyKiKu;
    property MatchFuzzyPunctuation: WordBool read Get_MatchFuzzyPunctuation write Set_MatchFuzzyPunctuation;
    property MatchFuzzySpace: WordBool read Get_MatchFuzzySpace write Set_MatchFuzzySpace;
    property Overtype: WordBool read Get_Overtype write Set_Overtype;
    property ShowDocumentMapOnLeft: WordBool read Get_ShowDocumentMapOnLeft write Set_ShowDocumentMapOnLeft;
    property AutoFormatAsYouTypeApplyNumberedLists: WordBool read Get_AutoFormatAsYouTypeApplyNumberedLists write Set_AutoFormatAsYouTypeApplyNumberedLists;
    property AutoFormatAsYouTypeReplaceOrdinals: WordBool read Get_AutoFormatAsYouTypeReplaceOrdinals write Set_AutoFormatAsYouTypeReplaceOrdinals;
    property AutoFormatAsYouTypeReplaceHyperlinks: WordBool read Get_AutoFormatAsYouTypeReplaceHyperlinks write Set_AutoFormatAsYouTypeReplaceHyperlinks;
    property AutoApplyAsYouTypeAdjustWordWrap: WordBool read Get_AutoApplyAsYouTypeAdjustWordWrap write Set_AutoApplyAsYouTypeAdjustWordWrap;
    property CheckSpellingAsYouType: WordBool read Get_CheckSpellingAsYouType write Set_CheckSpellingAsYouType;
    property IgnoreInternetAndFileAddresses: WordBool read Get_IgnoreInternetAndFileAddresses write Set_IgnoreInternetAndFileAddresses;
    property DisplayPasteOptions: WordBool read Get_DisplayPasteOptions write Set_DisplayPasteOptions;
    property HideTaskPanelAfterDocumentOpen: WordBool read Get_HideTaskPanelAfterDocumentOpen write Set_HideTaskPanelAfterDocumentOpen;
    property EnableObjectBrowser: WordBool read Get_EnableObjectBrowser write Set_EnableObjectBrowser;
    property PasteFormatMode: WpsRecoveryType read Get_PasteFormatMode write Set_PasteFormatMode;
    property SmartParaSelection: WordBool read Get_SmartParaSelection write Set_SmartParaSelection;
    property DocumentTabWidth: Integer read Get_DocumentTabWidth write Set_DocumentTabWidth;
    property DocumentTabPosition: WpsDocumentTabPosition read Get_DocumentTabPosition write Set_DocumentTabPosition;
    property DocumentTabShowCloseBtn: WordBool read Get_DocumentTabShowCloseBtn write Set_DocumentTabShowCloseBtn;
    property DocumentTabDbClose: WordBool read Get_DocumentTabDbClose write Set_DocumentTabDbClose;
    property DocumentTabDbNew: WordBool read Get_DocumentTabDbNew write Set_DocumentTabDbNew;
    property DocumentTabSwitch: WordBool read Get_DocumentTabSwitch write Set_DocumentTabSwitch;
    property DocumentTabWarning: WordBool read Get_DocumentTabWarning write Set_DocumentTabWarning;
    property DocumentViewDirection: WpsDocumentViewDirection read Get_DocumentViewDirection write Set_DocumentViewDirection;
    property ShowTableAdjustLabel: WordBool read Get_ShowTableAdjustLabel write Set_ShowTableAdjustLabel;
  end;

// *********************************************************************//
// DispIntf:  OptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B7-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  OptionsDisp = dispinterface
    ['{000209B7-0000-4B30-A977-D214852036FE}']
    property AllowAccentedUppercase: WordBool dispid 1;
    property BlueScreen: WordBool dispid 20;
    property MeasurementUnit: WpsMeasurementUnits dispid 26;
    property UpdateFieldsAtPrint: WordBool dispid 30;
    property PrintFieldCodes: WordBool dispid 32;
    property PrintComments: WordBool dispid 33;
    property PrintHiddenText: WordBool dispid 34;
    property PrintDrawingObjects: WordBool dispid 38;
    property CreateBackup: WordBool dispid 41;
    property BackgroundSave: WordBool dispid 46;
    property DefaultFilePath[Path: WpsDefaultFilePath]: WideString dispid 65;
    property ReplaceSelection: WordBool dispid 67;
    property AllowDragAndDrop: WordBool dispid 68;
    property TabIndentKey: WordBool dispid 72;
    property SnapToGrid: WordBool dispid 79;
    property GridDistanceHorizontal: Single dispid 81;
    property GridDistanceVertical: Single dispid 82;
    property GridOriginHorizontal: Single dispid 83;
    property GridOriginVertical: Single dispid 84;
    property DefaultHighlightColor: WpsColor dispid 274;
    property DefaultBorderLineStyle: WpsLineStyle dispid 275;
    property IgnoreUppercase: WordBool dispid 280;
    property IgnoreMixedDigits: WordBool dispid 281;
    property DefaultBorderLineWidth: WpsLineWidth dispid 284;
    property DefaultOpenFormat: WpsOpenFormat dispid 286;
    property MapPaperSize: WordBool dispid 289;
    property DisplayGridLines: WordBool dispid 306;
    property DefaultBorderColor: WpsColor dispid 337;
    property AllowPixelUnits: WordBool dispid 345;
    property UseCharacterUnit: WordBool dispid 346;
    property CtrlClickHyperlinkToOpen: WordBool dispid 435;
    property PrintOddPagesInAscendingOrder: WordBool dispid 330;
    property PrintEvenPagesInAscendingOrder: WordBool dispid 331;
    property PrintReverse: WordBool dispid 288;
    property PrintHiddenTextMode: WpsPrintHiddenTextMode dispid 4096;
    property InsertedTextMark: WpsInsertedTextMark dispid 57;
    property DeletedTextMark: WpsDeletedTextMark dispid 58;
    property RevisedLinesMark: WpsRevisedLinesMark dispid 59;
    property InsertedTextColor: WpsColor dispid 60;
    property DeletedTextColor: WpsColor dispid 61;
    property RevisedLinesColor: WpsColor dispid 62;
    property RevisedPropertiesMark: WpsRevisedPropertiesMark dispid 76;
    property RevisedPropertiesColor: WpsColor dispid 77;
    property RevisionsBalloonPrintOrientation: WpsRevisionsBalloonPrintOrientation dispid 453;
    property CommentsColor: WpsColor dispid 454;
    property RevisionsBalloonTitle: WpsRevisionsBalloonTitle dispid 4038;
    property CreateBackupOnFirstSave: WordBool dispid 4039;
    property SuggestFromMainDictionaryOnly: WordBool dispid 282;
    property SuggestSpellingCorrections: WordBool dispid 283;
    property SaveInterval: Integer dispid 45;
    property AllowClickAndTypeMouse: WordBool dispid 285;
    property DiscernPerson: WordBool dispid 287;
    property DiscernLocation: WordBool dispid 290;
    property MatchFuzzyCase: WordBool dispid 512;
    property MatchFuzzyByte: WordBool dispid 513;
    property MatchFuzzyHiragana: WordBool dispid 514;
    property MatchFuzzySmallKana: WordBool dispid 515;
    property MatchFuzzyDash: WordBool dispid 516;
    property MatchFuzzyIterationMark: WordBool dispid 517;
    property MatchFuzzyKanji: WordBool dispid 518;
    property MatchFuzzyOldKana: WordBool dispid 519;
    property MatchFuzzyProlongedSoundMark: WordBool dispid 520;
    property MatchFuzzyDZ: WordBool dispid 521;
    property MatchFuzzyBV: WordBool dispid 528;
    property MatchFuzzyTC: WordBool dispid 529;
    property MatchFuzzyHF: WordBool dispid 530;
    property MatchFuzzyZJ: WordBool dispid 531;
    property MatchFuzzyAY: WordBool dispid 532;
    property MatchFuzzyKiKu: WordBool dispid 533;
    property MatchFuzzyPunctuation: WordBool dispid 534;
    property MatchFuzzySpace: WordBool dispid 535;
    property Overtype: WordBool dispid 66;
    property ShowDocumentMapOnLeft: WordBool dispid 536;
    property AutoFormatAsYouTypeApplyNumberedLists: WordBool dispid 263;
    property AutoFormatAsYouTypeReplaceOrdinals: WordBool dispid 266;
    property AutoFormatAsYouTypeReplaceHyperlinks: WordBool dispid 272;
    property AutoApplyAsYouTypeAdjustWordWrap: WordBool dispid 460;
    property CheckSpellingAsYouType: WordBool dispid 276;
    property IgnoreInternetAndFileAddresses: WordBool dispid 278;
    property DisplayPasteOptions: WordBool dispid 439;
    property HideTaskPanelAfterDocumentOpen: WordBool dispid 461;
    property EnableObjectBrowser: WordBool dispid 560;
    property PasteFormatMode: WpsRecoveryType dispid 546;
    property SmartParaSelection: WordBool dispid 452;
    property DocumentTabWidth: Integer dispid 561;
    property DocumentTabPosition: WpsDocumentTabPosition dispid 562;
    property DocumentTabShowCloseBtn: WordBool dispid 563;
    property DocumentTabDbClose: WordBool dispid 564;
    property DocumentTabDbNew: WordBool dispid 565;
    property DocumentTabSwitch: WordBool dispid 566;
    property DocumentTabWarning: WordBool dispid 567;
    property DocumentViewDirection: WpsDocumentViewDirection dispid 273;
    property ShowTableAdjustLabel: WordBool dispid 568;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: RecentFiles
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020963-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  RecentFiles = interface(_IKsoDispObj)
    ['{00020963-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Maximum: Integer; safecall;
    procedure Set_Maximum(prop: Integer); safecall;
    function Item(Index: Integer): RecentFile; safecall;
    function Add(var Document: OleVariant; ReadOnly: WordBool): RecentFile; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Maximum: Integer read Get_Maximum write Set_Maximum;
  end;

// *********************************************************************//
// DispIntf:  RecentFilesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020963-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  RecentFilesDisp = dispinterface
    ['{00020963-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    property Maximum: Integer dispid 2;
    function Item(Index: Integer): RecentFile; dispid 0;
    function Add(var Document: OleVariant; ReadOnly: WordBool): RecentFile; dispid 3;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: RecentFile
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020964-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  RecentFile = interface(_IKsoDispObj)
    ['{00020964-0000-4B30-A977-D214852036FE}']
    function Get_Name: WideString; safecall;
    function Get_Index: Integer; safecall;
    function Get_ReadOnly: WordBool; safecall;
    procedure Set_ReadOnly(prop: WordBool); safecall;
    function Get_Path: WideString; safecall;
    function Open: _Document; safecall;
    procedure Delete; safecall;
    property Name: WideString read Get_Name;
    property Index: Integer read Get_Index;
    property ReadOnly: WordBool read Get_ReadOnly write Set_ReadOnly;
    property Path: WideString read Get_Path;
  end;

// *********************************************************************//
// DispIntf:  RecentFileDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020964-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  RecentFileDisp = dispinterface
    ['{00020964-0000-4B30-A977-D214852036FE}']
    property Name: WideString readonly dispid 0;
    property Index: Integer readonly dispid 1;
    property ReadOnly: WordBool dispid 2;
    property Path: WideString readonly dispid 3;
    function Open: _Document; dispid 4;
    procedure Delete; dispid 5;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Template
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Template = interface(_IKsoDispObj)
    ['{0002096A-0000-4B30-A977-D214852036FE}']
    function Get_Name: WideString; safecall;
    function Get_Path: WideString; safecall;
    function Get_Saved: WordBool; safecall;
    procedure Set_Saved(prop: WordBool); safecall;
    function Get_type_: WpsTemplateType; safecall;
    function Get_FullName: WideString; safecall;
    function OpenAsDocument: _Document; safecall;
    procedure Save; safecall;
    property Name: WideString read Get_Name;
    property Path: WideString read Get_Path;
    property Saved: WordBool read Get_Saved write Set_Saved;
    property type_: WpsTemplateType read Get_type_;
    property FullName: WideString read Get_FullName;
  end;

// *********************************************************************//
// DispIntf:  TemplateDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TemplateDisp = dispinterface
    ['{0002096A-0000-4B30-A977-D214852036FE}']
    property Name: WideString readonly dispid 0;
    property Path: WideString readonly dispid 1;
    property Saved: WordBool dispid 5;
    property type_: WpsTemplateType readonly dispid 6;
    property FullName: WideString readonly dispid 7;
    function OpenAsDocument: _Document; dispid 100;
    procedure Save; dispid 101;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Templates
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A2-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Templates = interface(_IKsoDispObj)
    ['{000209A2-0000-4B30-A977-D214852036FE}']
    function Get_Count: Integer; safecall;
    function Get__NewEnum: IUnknown; safecall;
    function Item(var Index: OleVariant): Template; safecall;
    property Count: Integer read Get_Count;
    property _NewEnum: IUnknown read Get__NewEnum;
  end;

// *********************************************************************//
// DispIntf:  TemplatesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209A2-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  TemplatesDisp = dispinterface
    ['{000209A2-0000-4B30-A977-D214852036FE}']
    property Count: Integer readonly dispid 1;
    property _NewEnum: IUnknown readonly dispid -4;
    function Item(var Index: OleVariant): Template; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Browser
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092E-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Browser = interface(_IKsoDispObj)
    ['{0002092E-0000-4B30-A977-D214852036FE}']
    function Get_Target: WpsBrowseTarget; safecall;
    procedure Set_Target(prop: WpsBrowseTarget); safecall;
    procedure Next; safecall;
    procedure Previous; safecall;
    property Target: WpsBrowseTarget read Get_Target write Set_Target;
  end;

// *********************************************************************//
// DispIntf:  BrowserDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002092E-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  BrowserDisp = dispinterface
    ['{0002092E-0000-4B30-A977-D214852036FE}']
    property Target: WpsBrowseTarget dispid 1;
    procedure Next; dispid 101;
    procedure Previous; dispid 102;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: System
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020935-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  System = interface(_IKsoDispObj)
    ['{00020935-0000-4B30-A977-D214852036FE}']
    function Get_OperatingSystem: WideString; safecall;
    function Get_ProcessorType: WideString; safecall;
    function Get_Version: WideString; safecall;
    function Get_FreeDiskSpace: Integer; safecall;
    function Get_Country: WpsCountry; safecall;
    function Get_LanguageDesignation: WideString; safecall;
    function Get_HorizontalResolution: Integer; safecall;
    function Get_VerticalResolution: Integer; safecall;
    function Get_ProfileString(const Section: WideString; const Key: WideString): WideString; safecall;
    procedure Set_ProfileString(const Section: WideString; const Key: WideString; 
                                const prop: WideString); safecall;
    function Get_PrivateProfileString(const FileName: WideString; const Section: WideString; 
                                      const Key: WideString): WideString; safecall;
    procedure Set_PrivateProfileString(const FileName: WideString; const Section: WideString; 
                                       const Key: WideString; const prop: WideString); safecall;
    function Get_MathCoprocessorInstalled: WordBool; safecall;
    function Get_ComputerType: WideString; safecall;
    function Get_MacintoshName: WideString; safecall;
    function Get_QuickDrawInstalled: WordBool; safecall;
    function Get_Cursor: WpsCursorType; safecall;
    procedure Set_Cursor(prop: WpsCursorType); safecall;
    procedure MSInfo; safecall;
    procedure Connect(const Path: WideString; var Drive: OleVariant; var Password: OleVariant); safecall;
    property OperatingSystem: WideString read Get_OperatingSystem;
    property ProcessorType: WideString read Get_ProcessorType;
    property Version: WideString read Get_Version;
    property FreeDiskSpace: Integer read Get_FreeDiskSpace;
    property Country: WpsCountry read Get_Country;
    property LanguageDesignation: WideString read Get_LanguageDesignation;
    property HorizontalResolution: Integer read Get_HorizontalResolution;
    property VerticalResolution: Integer read Get_VerticalResolution;
    property ProfileString[const Section: WideString; const Key: WideString]: WideString read Get_ProfileString write Set_ProfileString;
    property PrivateProfileString[const FileName: WideString; const Section: WideString; 
                                  const Key: WideString]: WideString read Get_PrivateProfileString write Set_PrivateProfileString;
    property MathCoprocessorInstalled: WordBool read Get_MathCoprocessorInstalled;
    property ComputerType: WideString read Get_ComputerType;
    property MacintoshName: WideString read Get_MacintoshName;
    property QuickDrawInstalled: WordBool read Get_QuickDrawInstalled;
    property Cursor: WpsCursorType read Get_Cursor write Set_Cursor;
  end;

// *********************************************************************//
// DispIntf:  SystemDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020935-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  SystemDisp = dispinterface
    ['{00020935-0000-4B30-A977-D214852036FE}']
    property OperatingSystem: WideString readonly dispid 1;
    property ProcessorType: WideString readonly dispid 2;
    property Version: WideString readonly dispid 3;
    property FreeDiskSpace: Integer readonly dispid 4;
    property Country: WpsCountry readonly dispid 5;
    property LanguageDesignation: WideString readonly dispid 6;
    property HorizontalResolution: Integer readonly dispid 7;
    property VerticalResolution: Integer readonly dispid 8;
    property ProfileString[const Section: WideString; const Key: WideString]: WideString dispid 9;
    property PrivateProfileString[const FileName: WideString; const Section: WideString; 
                                  const Key: WideString]: WideString dispid 10;
    property MathCoprocessorInstalled: WordBool readonly dispid 11;
    property ComputerType: WideString readonly dispid 12;
    property MacintoshName: WideString readonly dispid 14;
    property QuickDrawInstalled: WordBool readonly dispid 15;
    property Cursor: WpsCursorType dispid 16;
    procedure MSInfo; dispid 101;
    procedure Connect(const Path: WideString; var Drive: OleVariant; var Password: OleVariant); dispid 102;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: PdfExportOptions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00021000-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  PdfExportOptions = interface(_IKsoDispObj)
    ['{00021000-0000-4B30-A977-D214852036FE}']
    function Get_ViewPdfAuto: WordBool; safecall;
    procedure Set_ViewPdfAuto(prop: WordBool); safecall;
    function Get_ConvertTablesOfContents: WordBool; safecall;
    procedure Set_ConvertTablesOfContents(prop: WordBool); safecall;
    function Get_ConvertCommentsMode: WpsPdfCommentsMode; safecall;
    procedure Set_ConvertCommentsMode(prop: WpsPdfCommentsMode); safecall;
    function Get_ConvertFootnotes: WordBool; safecall;
    procedure Set_ConvertFootnotes(prop: WordBool); safecall;
    function Get_ConvertEndnotes: WordBool; safecall;
    procedure Set_ConvertEndnotes(prop: WordBool); safecall;
    function Get_ConvertSummaryInfo: WordBool; safecall;
    procedure Set_ConvertSummaryInfo(prop: WordBool); safecall;
    function Get_EditRight(Rights: WpsPdfEditRight): WordBool; safecall;
    procedure Set_EditRight(Rights: WpsPdfEditRight; prop: WordBool); safecall;
    function Get_PrintRight: WpsPdfPrintRight; safecall;
    procedure Set_PrintRight(prop: WpsPdfPrintRight); safecall;
    function Get_CopyRight: WpsPdfCopyRight; safecall;
    procedure Set_CopyRight(prop: WpsPdfCopyRight); safecall;
    function Get_ConvertHyperlinks: WordBool; safecall;
    procedure Set_ConvertHyperlinks(prop: WordBool); safecall;
    function Get_ConvertTitleStyles: WordBool; safecall;
    procedure Set_ConvertTitleStyles(prop: WordBool); safecall;
    function Get_ConvertBuildinStyles: WordBool; safecall;
    procedure Set_ConvertBuildinStyles(prop: WordBool); safecall;
    function Get_ConvertBookmark: WordBool; safecall;
    procedure Set_ConvertBookmark(prop: WordBool); safecall;
    function Get_ConvertCustomStyles: WordBool; safecall;
    procedure Set_ConvertCustomStyles(prop: WordBool); safecall;
    property ViewPdfAuto: WordBool read Get_ViewPdfAuto write Set_ViewPdfAuto;
    property ConvertTablesOfContents: WordBool read Get_ConvertTablesOfContents write Set_ConvertTablesOfContents;
    property ConvertCommentsMode: WpsPdfCommentsMode read Get_ConvertCommentsMode write Set_ConvertCommentsMode;
    property ConvertFootnotes: WordBool read Get_ConvertFootnotes write Set_ConvertFootnotes;
    property ConvertEndnotes: WordBool read Get_ConvertEndnotes write Set_ConvertEndnotes;
    property ConvertSummaryInfo: WordBool read Get_ConvertSummaryInfo write Set_ConvertSummaryInfo;
    property EditRight[Rights: WpsPdfEditRight]: WordBool read Get_EditRight write Set_EditRight;
    property PrintRight: WpsPdfPrintRight read Get_PrintRight write Set_PrintRight;
    property CopyRight: WpsPdfCopyRight read Get_CopyRight write Set_CopyRight;
    property ConvertHyperlinks: WordBool read Get_ConvertHyperlinks write Set_ConvertHyperlinks;
    property ConvertTitleStyles: WordBool read Get_ConvertTitleStyles write Set_ConvertTitleStyles;
    property ConvertBuildinStyles: WordBool read Get_ConvertBuildinStyles write Set_ConvertBuildinStyles;
    property ConvertBookmark: WordBool read Get_ConvertBookmark write Set_ConvertBookmark;
    property ConvertCustomStyles: WordBool read Get_ConvertCustomStyles write Set_ConvertCustomStyles;
  end;

// *********************************************************************//
// DispIntf:  PdfExportOptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00021000-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  PdfExportOptionsDisp = dispinterface
    ['{00021000-0000-4B30-A977-D214852036FE}']
    property ViewPdfAuto: WordBool dispid 1;
    property ConvertTablesOfContents: WordBool dispid 2;
    property ConvertCommentsMode: WpsPdfCommentsMode dispid 3;
    property ConvertFootnotes: WordBool dispid 4;
    property ConvertEndnotes: WordBool dispid 5;
    property ConvertSummaryInfo: WordBool dispid 6;
    property EditRight[Rights: WpsPdfEditRight]: WordBool dispid 7;
    property PrintRight: WpsPdfPrintRight dispid 8;
    property CopyRight: WpsPdfCopyRight dispid 9;
    property ConvertHyperlinks: WordBool dispid 10;
    property ConvertTitleStyles: WordBool dispid 11;
    property ConvertBuildinStyles: WordBool dispid 12;
    property ConvertBookmark: WordBool dispid 13;
    property ConvertCustomStyles: WordBool dispid 14;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Dictionaries
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AC-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Dictionaries = interface(_IKsoDispObj)
    ['{000209AC-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Get_Maximum: Integer; safecall;
    function Get_ActiveCustomDictionary: Dictionary; safecall;
    procedure Set_ActiveCustomDictionary(const prop: Dictionary); safecall;
    function Item(var Index: OleVariant): Dictionary; safecall;
    function Add(const FileName: WideString): Dictionary; safecall;
    procedure ClearAll; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
    property Maximum: Integer read Get_Maximum;
    property ActiveCustomDictionary: Dictionary read Get_ActiveCustomDictionary write Set_ActiveCustomDictionary;
  end;

// *********************************************************************//
// DispIntf:  DictionariesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AC-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DictionariesDisp = dispinterface
    ['{000209AC-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    property Maximum: Integer readonly dispid 2;
    property ActiveCustomDictionary: Dictionary dispid 3;
    function Item(var Index: OleVariant): Dictionary; dispid 0;
    function Add(const FileName: WideString): Dictionary; dispid 101;
    procedure ClearAll; dispid 102;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Dictionary
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AD-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Dictionary = interface(_IKsoDispObj)
    ['{000209AD-0000-4B30-A977-D214852036FE}']
    function Get_Name: WideString; safecall;
    function Get_Path: WideString; safecall;
    function Get_ReadOnly: WordBool; safecall;
    procedure Delete; safecall;
    function Get_Included: WordBool; safecall;
    procedure Set_Included(prop: WordBool); safecall;
    procedure AppendNewWord(const bstrnewword: WideString); safecall;
    procedure DeleteWord(const bstrword: WideString); safecall;
    function ScanDictionary: WideString; safecall;
    property Name: WideString read Get_Name;
    property Path: WideString read Get_Path;
    property ReadOnly: WordBool read Get_ReadOnly;
    property Included: WordBool read Get_Included write Set_Included;
  end;

// *********************************************************************//
// DispIntf:  DictionaryDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209AD-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DictionaryDisp = dispinterface
    ['{000209AD-0000-4B30-A977-D214852036FE}']
    property Name: WideString readonly dispid 0;
    property Path: WideString readonly dispid 1;
    property ReadOnly: WordBool readonly dispid 3;
    procedure Delete; dispid 101;
    property Included: WordBool dispid 102;
    procedure AppendNewWord(const bstrnewword: WideString); dispid 103;
    procedure DeleteWord(const bstrword: WideString); dispid 104;
    function ScanDictionary: WideString; dispid 105;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Dialogs
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020910-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Dialogs = interface(_IKsoDispObj)
    ['{00020910-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(Index: WpsDialog): Dialog; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  DialogsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020910-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DialogsDisp = dispinterface
    ['{00020910-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(Index: WpsDialog): Dialog; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: Dialog
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B8-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  Dialog = interface(_IKsoDispObj)
    ['{000209B8-0000-4B30-A977-D214852036FE}']
    function Get_DefaultTab: WpsDialogTab; safecall;
    procedure Set_DefaultTab(prop: WpsDialogTab); safecall;
    function Get_type_: WpsDialog; safecall;
    function Show(var TimeOut: OleVariant): Integer; safecall;
    function Display(var TimeOut: OleVariant): Integer; safecall;
    procedure Execute; safecall;
    procedure Update; safecall;
    function Get_CommandName: WideString; safecall;
    property DefaultTab: WpsDialogTab read Get_DefaultTab write Set_DefaultTab;
    property type_: WpsDialog read Get_type_;
    property CommandName: WideString read Get_CommandName;
  end;

// *********************************************************************//
// DispIntf:  DialogDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {000209B8-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DialogDisp = dispinterface
    ['{000209B8-0000-4B30-A977-D214852036FE}']
    property DefaultTab: WpsDialogTab dispid 32002;
    property type_: WpsDialog readonly dispid 0;
    function Show(var TimeOut: OleVariant): Integer; dispid 336;
    function Display(var TimeOut: OleVariant): Integer; dispid 338;
    procedure Execute; dispid 32001;
    procedure Update; dispid 302;
    property CommandName: WideString readonly dispid 32006;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: AutoCorrect
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020948-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  AutoCorrect = interface(_IKsoDispObj)
    ['{00020948-0000-4B30-A977-D214852036FE}']
    function Get_CorrectDays: WordBool; safecall;
    procedure Set_CorrectDays(prop: WordBool); safecall;
    function Get_CorrectInitialCaps: WordBool; safecall;
    procedure Set_CorrectInitialCaps(prop: WordBool); safecall;
    function Get_CorrectSentenceCaps: WordBool; safecall;
    procedure Set_CorrectSentenceCaps(prop: WordBool); safecall;
    function Get_ReplaceText: WordBool; safecall;
    procedure Set_ReplaceText(prop: WordBool); safecall;
    function Get_FirstLetterAutoAdd: WordBool; safecall;
    procedure Set_FirstLetterAutoAdd(prop: WordBool); safecall;
    function Get_TwoInitialCapsAutoAdd: WordBool; safecall;
    procedure Set_TwoInitialCapsAutoAdd(prop: WordBool); safecall;
    function Get_CorrectCapsLock: WordBool; safecall;
    procedure Set_CorrectCapsLock(prop: WordBool); safecall;
    function Get_CorrectHangulAndAlphabet: WordBool; safecall;
    procedure Set_CorrectHangulAndAlphabet(prop: WordBool); safecall;
    function Get_HangulAndAlphabetAutoAdd: WordBool; safecall;
    procedure Set_HangulAndAlphabetAutoAdd(prop: WordBool); safecall;
    function Get_ReplaceTextFromSpellingChecker: WordBool; safecall;
    procedure Set_ReplaceTextFromSpellingChecker(prop: WordBool); safecall;
    function Get_OtherCorrectionsAutoAdd: WordBool; safecall;
    procedure Set_OtherCorrectionsAutoAdd(prop: WordBool); safecall;
    function Get_CorrectKeyboardSetting: WordBool; safecall;
    procedure Set_CorrectKeyboardSetting(prop: WordBool); safecall;
    function Get_CorrectTableCells: WordBool; safecall;
    procedure Set_CorrectTableCells(prop: WordBool); safecall;
    function Get_DisplayAutoCorrectOptions: WordBool; safecall;
    procedure Set_DisplayAutoCorrectOptions(prop: WordBool); safecall;
    property CorrectDays: WordBool read Get_CorrectDays write Set_CorrectDays;
    property CorrectInitialCaps: WordBool read Get_CorrectInitialCaps write Set_CorrectInitialCaps;
    property CorrectSentenceCaps: WordBool read Get_CorrectSentenceCaps write Set_CorrectSentenceCaps;
    property ReplaceText: WordBool read Get_ReplaceText write Set_ReplaceText;
    property FirstLetterAutoAdd: WordBool read Get_FirstLetterAutoAdd write Set_FirstLetterAutoAdd;
    property TwoInitialCapsAutoAdd: WordBool read Get_TwoInitialCapsAutoAdd write Set_TwoInitialCapsAutoAdd;
    property CorrectCapsLock: WordBool read Get_CorrectCapsLock write Set_CorrectCapsLock;
    property CorrectHangulAndAlphabet: WordBool read Get_CorrectHangulAndAlphabet write Set_CorrectHangulAndAlphabet;
    property HangulAndAlphabetAutoAdd: WordBool read Get_HangulAndAlphabetAutoAdd write Set_HangulAndAlphabetAutoAdd;
    property ReplaceTextFromSpellingChecker: WordBool read Get_ReplaceTextFromSpellingChecker write Set_ReplaceTextFromSpellingChecker;
    property OtherCorrectionsAutoAdd: WordBool read Get_OtherCorrectionsAutoAdd write Set_OtherCorrectionsAutoAdd;
    property CorrectKeyboardSetting: WordBool read Get_CorrectKeyboardSetting write Set_CorrectKeyboardSetting;
    property CorrectTableCells: WordBool read Get_CorrectTableCells write Set_CorrectTableCells;
    property DisplayAutoCorrectOptions: WordBool read Get_DisplayAutoCorrectOptions write Set_DisplayAutoCorrectOptions;
  end;

// *********************************************************************//
// DispIntf:  AutoCorrectDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {00020948-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  AutoCorrectDisp = dispinterface
    ['{00020948-0000-4B30-A977-D214852036FE}']
    property CorrectDays: WordBool dispid 1;
    property CorrectInitialCaps: WordBool dispid 2;
    property CorrectSentenceCaps: WordBool dispid 3;
    property ReplaceText: WordBool dispid 4;
    property FirstLetterAutoAdd: WordBool dispid 8;
    property TwoInitialCapsAutoAdd: WordBool dispid 10;
    property CorrectCapsLock: WordBool dispid 11;
    property CorrectHangulAndAlphabet: WordBool dispid 12;
    property HangulAndAlphabetAutoAdd: WordBool dispid 14;
    property ReplaceTextFromSpellingChecker: WordBool dispid 15;
    property OtherCorrectionsAutoAdd: WordBool dispid 16;
    property CorrectKeyboardSetting: WordBool dispid 18;
    property CorrectTableCells: WordBool dispid 19;
    property DisplayAutoCorrectOptions: WordBool dispid 20;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: FontNames
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FontNames = interface(_IKsoDispObj)
    ['{0002096F-0000-4B30-A977-D214852036FE}']
    function Get_Count: Integer; safecall;
    function Item(Index: Integer): WideString; safecall;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  FontNamesDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002096F-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  FontNamesDisp = dispinterface
    ['{0002096F-0000-4B30-A977-D214852036FE}']
    property Count: Integer readonly dispid 2;
    function Item(Index: Integer): WideString; dispid 0;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// Interface: AutoCaptions
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  AutoCaptions = interface(IDispatch)
    ['{0002097A-0000-4B30-A977-D214852036FE}']
    function Get__NewEnum: IUnknown; safecall;
    function Get_Count: Integer; safecall;
    function Item(var Index: OleVariant): AutoCaption; safecall;
    procedure CancelAutoInsert; safecall;
    property _NewEnum: IUnknown read Get__NewEnum;
    property Count: Integer read Get_Count;
  end;

// *********************************************************************//
// DispIntf:  AutoCaptionsDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097A-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  AutoCaptionsDisp = dispinterface
    ['{0002097A-0000-4B30-A977-D214852036FE}']
    property _NewEnum: IUnknown readonly dispid -4;
    property Count: Integer readonly dispid 1;
    function Item(var Index: OleVariant): AutoCaption; dispid 0;
    procedure CancelAutoInsert; dispid 100;
  end;

// *********************************************************************//
// Interface: AutoCaption
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097B-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  AutoCaption = interface(_IKsoDispObj)
    ['{0002097B-0000-4B30-A977-D214852036FE}']
    function Get_Name: WideString; safecall;
    function Get_AutoInsert: WordBool; safecall;
    procedure Set_AutoInsert(prop: WordBool); safecall;
    function Get_Index: Integer; safecall;
    function Get_CaptionLabel: OleVariant; safecall;
    procedure Set_CaptionLabel(var prop: OleVariant); safecall;
    property Name: WideString read Get_Name;
    property AutoInsert: WordBool read Get_AutoInsert write Set_AutoInsert;
    property Index: Integer read Get_Index;
  end;

// *********************************************************************//
// DispIntf:  AutoCaptionDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {0002097B-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  AutoCaptionDisp = dispinterface
    ['{0002097B-0000-4B30-A977-D214852036FE}']
    property Name: WideString readonly dispid 0;
    property AutoInsert: WordBool dispid 1;
    property Index: Integer readonly dispid 2;
    function CaptionLabel: OleVariant; dispid 3;
    property Application: _Application readonly dispid 1000;
    property Creator: Integer readonly dispid 1001;
    property Parent: IDispatch readonly dispid 1002;
    procedure zimp_DispObj_Reserved1; dispid 268435201;
    procedure zimp_DispObj_Reserved2; dispid 268435202;
    procedure zimp_DispObj_Reserved3; dispid 268435203;
    procedure zimp_DispObj_Reserved4; dispid 268435204;
    procedure zimp_DispObj_Reserved5; dispid 268435205;
    property Description: WideString readonly dispid 268435206;
  end;

// *********************************************************************//
// DispIntf:  ApplicationEvents
// Flags:     (4112) Hidden Dispatchable
// GUID:      {000209FE-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  ApplicationEvents = dispinterface
    ['{000209FE-0000-4B30-A977-D214852036FE}']
    procedure Startup; dispid 1;
    procedure Quit; dispid 2;
    procedure DocumentChange; dispid 3;
    procedure DocumentOpen(const Doc: _Document); dispid 4;
    procedure DocumentBeforeClose(const Doc: _Document; var Cancel: WordBool); dispid 6;
    procedure DocumentBeforePrint(const Doc: _Document; var Cancel: WordBool); dispid 7;
    procedure DocumentBeforeSave(const Doc: _Document; var SaveAsUI: WordBool; var Cancel: WordBool); dispid 8;
    procedure NewDocument(const Doc: _Document); dispid 9;
    procedure WindowActivate(const Doc: _Document; const Wn: Window); dispid 10;
    procedure WindowDeactivate(const Doc: _Document; const Wn: Window); dispid 11;
    procedure WindowSelectionChange(const Sel: Selection); dispid 12;
    procedure WindowBeforeRightClick(const Sel: Selection; var Cancel: WordBool); dispid 13;
    procedure WindowBeforeDoubleClick(const Sel: Selection; var Cancel: WordBool); dispid 14;
    procedure WindowSize(const Doc: _Document; const Wn: Window); dispid 25;
    procedure DocumentBeforePrintPreview(const Doc: _Document; var Cancel: WordBool); dispid 32;
    procedure DocumentAfterPrintPreviewClose(const Doc: _Document; var Cancel: WordBool); dispid 33;
  end;

// *********************************************************************//
// DispIntf:  DocumentEvents
// Flags:     (4112) Hidden Dispatchable
// GUID:      {000209F6-0000-4B30-A977-D214852036FE}
// *********************************************************************//
  DocumentEvents = dispinterface
    ['{000209F6-0000-4B30-A977-D214852036FE}']
    procedure New; dispid 4;
    procedure Open; dispid 5;
    procedure Close; dispid 6;
  end;

// *********************************************************************//
// Interface: _OLEControl
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {000209A4-0000-0000-C000-000000000046}
// *********************************************************************//
  _OLEControl = interface(IDispatch)
    ['{000209A4-0000-0000-C000-000000000046}']
    function Get_Left: Single; safecall;
    procedure Set_Left(prop: Single); safecall;
    function Get_Top: Single; safecall;
    procedure Set_Top(prop: Single); safecall;
    function Get_Height: Single; safecall;
    procedure Set_Height(prop: Single); safecall;
    function Get_Width: Single; safecall;
    procedure Set_Width(prop: Single); safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const prop: WideString); safecall;
    function Get_Automation: IDispatch; safecall;
    procedure Select; safecall;
    procedure Copy; safecall;
    procedure Cut; safecall;
    procedure Delete; safecall;
    procedure Activate; safecall;
    function Get_AltHTML: WideString; safecall;
    procedure Set_AltHTML(const prop: WideString); safecall;
    property Left: Single read Get_Left write Set_Left;
    property Top: Single read Get_Top write Set_Top;
    property Height: Single read Get_Height write Set_Height;
    property Width: Single read Get_Width write Set_Width;
    property Name: WideString read Get_Name write Set_Name;
    property Automation: IDispatch read Get_Automation;
    property AltHTML: WideString read Get_AltHTML write Set_AltHTML;
  end;

// *********************************************************************//
// DispIntf:  _OLEControlDisp
// Flags:     (4432) Hidden Dual OleAutomation Dispatchable
// GUID:      {000209A4-0000-0000-C000-000000000046}
// *********************************************************************//
  _OLEControlDisp = dispinterface
    ['{000209A4-0000-0000-C000-000000000046}']
    property Left: Single dispid -2147417853;
    property Top: Single dispid -2147417852;
    property Height: Single dispid -2147417851;
    property Width: Single dispid -2147417850;
    property Name: WideString dispid -2147418112;
    property Automation: IDispatch readonly dispid -2147417849;
    procedure Select; dispid -2147417568;
    procedure Copy; dispid -2147417560;
    procedure Cut; dispid -2147417559;
    procedure Delete; dispid -2147417520;
    procedure Activate; dispid -2147417519;
    property AltHTML: WideString dispid -2147415101;
  end;

// *********************************************************************//
// DispIntf:  OCXEvents
// Flags:     (4112) Hidden Dispatchable
// GUID:      {000209F3-0000-0000-C000-000000000046}
// *********************************************************************//
  OCXEvents = dispinterface
    ['{000209F3-0000-0000-C000-000000000046}']
    procedure GotFocus; dispid -2147417888;
    procedure LostFocus; dispid -2147417887;
  end;

// *********************************************************************//
// The Class CoOLECtrl provides a Create and CreateRemote method to          
// create instances of the default interface _OLEControl exposed by              
// the CoClass OLECtrl. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoOLECtrl = class
    class function Create: _OLEControl;
    class function CreateRemote(const MachineName: string): _OLEControl;
  end;

// *********************************************************************//
// The Class CoApplication provides a Create and CreateRemote method to          
// create instances of the default interface _Application exposed by              
// the CoClass Application. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoApplication = class
    class function Create: _Application;
    class function CreateRemote(const MachineName: string): _Application;
  end;

// *********************************************************************//
// The Class CoDocument provides a Create and CreateRemote method to          
// create instances of the default interface _Document exposed by              
// the CoClass Document. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoDocument = class
    class function Create: _Document;
    class function CreateRemote(const MachineName: string): _Document;
  end;

// *********************************************************************//
// The Class CoFont provides a Create and CreateRemote method to          
// create instances of the default interface _Font exposed by              
// the CoClass Font. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoFont = class
    class function Create: _Font;
    class function CreateRemote(const MachineName: string): _Font;
  end;

// *********************************************************************//
// The Class CoParagraphFormat provides a Create and CreateRemote method to          
// create instances of the default interface _ParagraphFormat exposed by              
// the CoClass ParagraphFormat. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoParagraphFormat = class
    class function Create: _ParagraphFormat;
    class function CreateRemote(const MachineName: string): _ParagraphFormat;
  end;

// *********************************************************************//
// The Class CoWebOptions provides a Create and CreateRemote method to          
// create instances of the default interface _WebOptions exposed by              
// the CoClass WebOptions. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoWebOptions = class
    class function Create: _WebOptions;
    class function CreateRemote(const MachineName: string): _WebOptions;
  end;

implementation

uses ComObj;

class function CoOLECtrl.Create: _OLEControl;
begin
  Result := CreateComObject(CLASS_OLECtrl) as _OLEControl;
end;

class function CoOLECtrl.CreateRemote(const MachineName: string): _OLEControl;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_OLECtrl) as _OLEControl;
end;

class function CoApplication.Create: _Application;
begin
  Result := CreateComObject(CLASS_Application) as _Application;
end;

class function CoApplication.CreateRemote(const MachineName: string): _Application;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Application) as _Application;
end;

class function CoDocument.Create: _Document;
begin
  Result := CreateComObject(CLASS_Document) as _Document;
end;

class function CoDocument.CreateRemote(const MachineName: string): _Document;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Document) as _Document;
end;

class function CoFont.Create: _Font;
begin
  Result := CreateComObject(CLASS_Font) as _Font;
end;

class function CoFont.CreateRemote(const MachineName: string): _Font;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Font) as _Font;
end;

class function CoParagraphFormat.Create: _ParagraphFormat;
begin
  Result := CreateComObject(CLASS_ParagraphFormat) as _ParagraphFormat;
end;

class function CoParagraphFormat.CreateRemote(const MachineName: string): _ParagraphFormat;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ParagraphFormat) as _ParagraphFormat;
end;

class function CoWebOptions.Create: _WebOptions;
begin
  Result := CreateComObject(CLASS_WebOptions) as _WebOptions;
end;

class function CoWebOptions.CreateRemote(const MachineName: string): _WebOptions;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_WebOptions) as _WebOptions;
end;

end.
