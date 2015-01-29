{**
 * <p>Title: ChangeSkin  </p>
 * <p>Copyright: Copyright (c) 2006/11/23</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 皮肤选择界面</p>
 *}

unit U_ChangeSkin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, WinSkinData, ComCtrls, ExtCtrls;

type
  TF_ChangeSkin = class(TForm)
    Label1: TLabel;
    btnOk: TBitBtn;
    btnCancel: TBitBtn;
    btnPreview: TBitBtn;
    lvwSkins: TListView;
    lblCurrSkin: TLabel;
    pnl1: TPanel;
    imgPreView: TImage;
    pnl2: TPanel;
    bvl1: TBevel;
    lbl1: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure btnPreviewClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure lvwSkinsSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);

  private
    { Private declarations } 
    FsSkinPath: string;
    FsOldSkin : string;
    FsNewSkin : string;
    FSkinData : TSkinData;
    FbBtnEnabled: Boolean;
    FSkin: string;

    procedure ApplySkin( sSkinFileName: string; bRefresh: Boolean = True );overload;
    procedure SetBtnStatus( bEnanle: Boolean );
  public
    property Skin: string read FSkin;    // 最终选择的皮肤
    
  public
    { Public declarations }
    procedure PassSkinObject( var ASkinData: TSkinData);
    procedure ReadSkinfile( apath : string );
  end;

  function GetValidSkinFileName(ASkinName: string):string;  
  function GetValidSkinName(ASkinFileName: string): string;
  procedure ApplySkin( var ASkinData: TSkinData; sSkinFileName: string );overload;

implementation

{$R *.dfm}

uses
  U_Const, U_CommonFunc, U_Search;

const
   SKIN_NOSKIN = 'No Skin';

function GetValidSkinFileName(ASkinName: string):string;
begin
  Result := ASkinName;
  if Result = '' then
    Exit;
  if ExtractFileExt(Result) <> '.skn' then
    Result := Result + '.skn';
  if Pos(':',Result) = 0 then
    Result := commonFunc.AppRootPath + PATH_SKIN + Result;
end;

function GetValidSkinName(ASkinFileName: string): string;
var
  nPosLastSlash: Integer;
begin
  nPosLastSlash := commonFunc.PosLast('\', ASkinFileName);
  if nPosLastSlash > 0 then
    Result := Copy(ASkinFileName, nPosLastSlash+1, Length(ASkinFileName))
  else
    Result := ASkinFileName;
end;

procedure ApplySkin( var ASkinData: TSkinData; sSkinFileName: string );overload;
var
  sOldSkin: string;
begin
  sOldSkin := ASkinData.SkinFile;
  try
    //ASkinData.Active := False;
    if GetValidSkinName(sSkinFileName) = SKIN_NOSKIN then
      ASkinData.SkinFile := ''
    else if FileExists( GetValidSkinFileName(sSkinFileName) ) then
      ASkinData.SkinFile:= GetValidSkinFileName(sSkinFileName);

    if (sSkinFileName <> '') and not ASkinData.Active then
      ASkinData.Active := True;
  except
    ASkinData.SkinFile := sOldSkin;
    ASkinData.Active := True;
  end;
end;

{ TFChangeMainStyle }

procedure TF_ChangeSkin.ReadSkinfile(apath: string);
var
  sts: Integer ;
  SR: TSearchRec;
  list: Tstringlist;

  procedure AddFile;
  begin
    list.add(RemoveFileExt(SR.Name));
    with lvwSkins.Items.Add do
    begin
      Caption := RemoveFileExt(SR.Name);
//      SubItems.Add(apath + SR.Name);
    end;
  end;

//// ReadSkinfile begin
begin
  lvwSkins.Items.BeginUpdate;
  lvwSkins.Clear;
  lvwSkins.Items.Add.Caption := SKIN_NOSKIN;
  list:=Tstringlist.create;
  try
    sts := FindFirst( apath + '*.skn' , faAnyFile , SR );
    if sts = 0 then begin
        if ( SR.Name <> '.' ) and ( SR.Name <> '..' ) then begin
            if pos('.', SR.Name) <> 0 then
              Addfile;
        end;
        while FindNext( SR ) = 0 do begin
            if ( SR.Name <> '.' ) and ( SR.Name <> '..' ) then begin
                if Pos('.', SR.Name) <> 0 then  Addfile;
            end;
        end;
    end ;
    FindClose( SR ) ;
    list.sort;
    lvwSkins.Items.EndUpdate;
    //cbbSkins.items.assign(list);
  finally
    list.free;
  end;
end;

procedure TF_ChangeSkin.FormCreate(Sender: TObject);
begin
  FSkin := '';
end;

procedure TF_ChangeSkin.btnOkClick(Sender: TObject);
begin
  if FSkin = '' then
  begin
    ApplySkin(FsSkinPath+FsNewSkin, False);  
    FSkin := FsNewSkin;
//    FSkin := RemoveFileExt( ExtractFileName(FsNewSkin) );
  end;
  ModalResult:=mrok;
end;

procedure TF_ChangeSkin.ApplySkin( sSkinFileName: string; bRefresh: Boolean = True );
var
  sOldSkin: string;
begin
  sOldSkin := FSkinData.SkinFile;
  try
    if not bRefresh then       // 先把Active置为False则不刷新
      FSkinData.Active := False;

    if GetValidSkinName(sSkinFileName) = SKIN_NOSKIN then
      FSkinData.SkinFile := ''
    else if FileExists( GetValidSkinFileName(sSkinFileName) ) then
      FSkinData.SkinFile := GetValidSkinFileName( sSkinFileName );

    if not FSkinData.Active then
      FSkinData.Active := True;
  except
    FSkinData.SkinFile := sOldSkin;
    FSkinData.Active := True;
  end;
end;

procedure TF_ChangeSkin.btnPreviewClick(Sender: TObject);
begin
  ApplySkin( FsSkinPath + FsNewSkin );
end;

procedure TF_ChangeSkin.btnCancelClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TF_ChangeSkin.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  if ( FSkin <> '' ) and ( FsNewSkin <> '' ) and
      ( not SameText( FsOldSkin, FsNewSkin ) ) then
    ApplySkin( FsSkinPath + FsOldSkin, False );
end;

procedure TF_ChangeSkin.PassSkinObject( var ASkinData: TSkinData);
begin
  FSkinData := ASkinData;
  FsSkinPath := commonFunc.AppRootPath + PATH_SKIN;
  ReadSkinfile(FsSkinPath);
  FsOldSkin := RemoveFileExt( ExtractFileName( FSkinData.SkinFile ) );
  FsNewSkin := FsOldSkin;
end;

procedure TF_ChangeSkin.FormShow(Sender: TObject);
begin
  lblCurrSkin.Caption := FsOldSkin;
  SetBtnStatus(False);
end;

procedure TF_ChangeSkin.SetBtnStatus(bEnanle: Boolean);
begin
  if FbBtnEnabled <> bEnanle then
  begin
    FbBtnEnabled := bEnanle;
    btnOk.Enabled := bEnanle;
    btnPreview.Enabled := bEnanle;
  end;
end;

procedure TF_ChangeSkin.lvwSkinsSelectItem(Sender: TObject;
  Item: TListItem; Selected: Boolean);
var
  sJPG: string;
begin
  if Assigned( lvwSkins.Selected ) then
  begin
    FsNewSkin := lvwSkins.Selected.Caption;
    SetBtnStatus(True);
    sJPG := FsSkinPath + FsNewSkin + '.jpg';
    if FileExists(sJPG) then
      commonFunc.LoadPhoto(imgPreview, sJPG)
    else
    begin
      commonFunc.ClearPhoto(imgPreView);
      //DrawStrOnImg(imgPreView,'没有预览');
    end;
  end
  else
  begin
    SetBtnStatus(False);
    commonFunc.ClearPhoto(imgPreView);
  end;
end;

end.
