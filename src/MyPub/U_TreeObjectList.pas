{**
 * <p>Title: U_TreeObjectList  </p>
 * <p>Copyright: Copyright (c) 2008/12/19</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 以树方式存储数据的类，用于被继承扩展，方便在TreeView中显示
 *}
unit U_TreeObjectList;

interface

uses
  Classes, Contnrs, SysUtils;

type
  TTreeObjectList = class; 
  TTreeObjectItem = class; 
  TFuncCompareTreeObjectItem = function (item1, item2: TTreeObjectItem):Boolean;
  
  TTreeObjectItem = class
  private                     
    FCheckState: Integer;            // 选中状态： 0未选中 1选中 2灰色
    FDisableCheckCascade: Boolean;   // 屏蔽类自身的check级联功能，在界面显示时可设置为true
    function GetChecked: Boolean;
    procedure SetChecked(value: boolean);
    procedure SetDisableCheckCascade(value: Boolean);
  protected                
    OwnerList: TTreeObjectList;
    procedure CheckGray();
  public
    Name: string;
    AllowGrayed: Boolean;
//    Data: TObject;
    Parent: TTreeObjectItem;     // 仅根节点可为nil
    Childs: TTreeObjectList;     // 子元素列表
  public
    constructor Create;
    destructor Destroy;override;
    function GetNextSibling: TTreeObjectItem;
    function CustomFindChildItem(compFunc: TFuncCompareTreeObjectItem;
      itemToFind: TTreeObjectItem): TTreeObjectItem;
    function AllChecked: Boolean;
    function IsLeaf: Boolean;
  published
    property Checked: Boolean read GetChecked write SetChecked;
    property CheckState: Integer read FCheckState write FCheckState;
    property DisableCheckCascade: Boolean read FDisableCheckCascade
      write SetDisableCheckCascade;
  end;

  TTreeObjectList = class(TObjectList)
  private                      
    FDisableCheckCascade: Boolean;
    function GetItem(index: Integer): TTreeObjectItem;
    procedure SetItem(index: Integer; value: TTreeObjectItem);
    function GetChecked: Boolean;
    procedure SetChecked(value: boolean);    
    procedure SetDisableCheckCascade(value: Boolean);
  protected         
    OwerItem: TTreeObjectItem;
  public
    AllowGrayed: Boolean;
    property Items[index: Integer]: TTreeObjectItem read GetItem write SetItem;
  public
    constructor Create(AOwerItem: TTreeObjectItem);
    destructor Destroy;override; 
    function Add(treeItem: TTreeObjectItem): Integer;
    function CustomFind(compFunc: TFuncCompareTreeObjectItem;
      itemToFind: TTreeObjectItem): TTreeObjectItem;
    function AllChecked: Boolean;
    function HasChecked: Boolean;
    procedure SetCheckedItemByStr(sCheckedItems: string; sSeparator: Char = ',');
    function GetSelectedItemNames(sSeparator: string = ','): string;
  published
    property Checked: Boolean read GetChecked write SetChecked;
    property DisableCheckCascade: Boolean read FDisableCheckCascade
      write SetDisableCheckCascade;
  end;

implementation

const
  C_nChecked_UnCheck = 0;
  C_nChecked_Checked = 1;
  C_nChecked_Grayed  = 2;

{ TTreeObjectList }

function TTreeObjectList.Add(treeItem: TTreeObjectItem): Integer;
begin
  Result := inherited Add(treeItem);
  Items[Result].Parent := OwerItem;
  Items[Result].OwnerList := Self;
end;

constructor TTreeObjectList.Create(AOwerItem: TTreeObjectItem);
begin
  inherited Create();
  OwerItem := AOwerItem;
  if OwerItem <> nil then
    OwerItem.Childs := Self;
  DisableCheckCascade := False;
end;

function TTreeObjectList.CustomFind(
  compFunc: TFuncCompareTreeObjectItem;
  itemToFind: TTreeObjectItem): TTreeObjectItem;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Count - 1 do
  begin
    Result := Items[i].CustomFindChildItem(compFunc, itemToFind);
    if Result <> nil then
    begin
      Break;
    end;  
  end;
end;

destructor TTreeObjectList.Destroy;
//var
//  i: Integer;
begin
//  for i := Count - 1 downto 0 do
//  begin
//    Items[i].Free;
//  end;
  inherited;
end;

function TTreeObjectList.GetItem(index: Integer): TTreeObjectItem;
begin
  Result := inherited GetItem(index) as TTreeObjectItem;
end;   

function TTreeObjectList.GetChecked: Boolean;
var
  i: Integer;
begin
  Result := True;
  for i := 0 to Count - 1 do
  begin
    Result := Result and Items[i].Checked;
  end;  
end;

procedure TTreeObjectList.SetChecked(value: boolean);
var
  i: Integer;
begin
  for i := 0 to Count - 1 do
  begin
    Items[i].Checked := value;
  end;  
end;

procedure TTreeObjectList.SetItem(index: Integer; value: TTreeObjectItem);
begin
  inherited SetItem(index ,value);
end;

function TTreeObjectList.AllChecked: Boolean;
var
  i: Integer;
begin
  Result := True;
  for i := 0 to Count - 1 do
  begin
    if not Items[i].AllChecked then
    begin
      Result := False;
      Break;
    end;
  end;
end;  

function TTreeObjectList.HasChecked: Boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to Count - 1 do
  begin
    if Items[i].Checked then
    begin
      Result := True;
      Break;
    end;
  end;
end;

procedure TTreeObjectList.SetDisableCheckCascade(value: Boolean);
var
  i: Integer;
begin
  Self.FDisableCheckCascade := value;
  for i := 0 to Count - 1 do
  begin
    Items[i].DisableCheckCascade := value;
  end;  
end;

procedure TTreeObjectList.SetCheckedItemByStr(sCheckedItems: string;
  sSeparator: Char);
var
  i: Integer;
  slst: TStrings;
begin
  slst := TStringList.Create;
  try
    ExtractStrings([sSeparator], [' '], PChar(sCheckedItems), slst);
    for i := 0 to Count - 1 do
    begin
      if -1 <> slst.IndexOf(Items[i].Name) then
      begin
        Items[i].Checked := True;
      end;
      if not Items[i].IsLeaf then
      begin
        Items[i].Childs.SetCheckedItemByStr(sCheckedItems, sSeparator);
      end;    
    end;
  finally
    slst.Free;
  end;
end;

function TTreeObjectList.GetSelectedItemNames(sSeparator: string): string;
var
  i: Integer;
  sResult: string;
begin
  sResult := '';
  for i := 0 to Count - 1 do
  begin
    if sResult = '' then
    begin
      if Items[i].IsLeaf then
      begin
        if Items[i].Checked then
          sResult := Items[i].Name
      end
      else
        sResult := Items[i].Childs.GetSelectedItemNames(sSeparator);
    end
    else
    begin                
      if Items[i].IsLeaf then
      begin
        if Items[i].Checked then
          sResult := sResult + sSeparator + Items[i].Name
      end
      else
        sResult := sResult + sSeparator + Items[i].Childs.
          GetSelectedItemNames(sSeparator);
    end;
  end;
  Result := sResult;
end;

{ TTreeObjectItem }

constructor TTreeObjectItem.Create;
begin
  Childs := TTreeObjectList.Create(Self);
  FCheckState := C_nChecked_UnCheck;
  FDisableCheckCascade := False;
end;
 
destructor TTreeObjectItem.Destroy;
begin
  if Assigned(Childs) then
    FreeAndNil(Childs);
  inherited;
end;

function TTreeObjectItem.CustomFindChildItem(
  compFunc: TFuncCompareTreeObjectItem;
  itemToFind: TTreeObjectItem): TTreeObjectItem;
begin
  Result := nil;
  if compFunc(Self, itemToFind) then
  begin
    Result := Self;
    Exit;
  end;  
  if Assigned(Childs) and (Childs.Count > 0) then
  begin
    Result := Childs.CustomFind(compFunc, itemToFind);
  end;
end;

function TTreeObjectItem.GetNextSibling: TTreeObjectItem;
var
  nIndex: Integer;
begin
  Result := nil;
  if Parent = nil then
    Exit;
  nIndex := Parent.Childs.IndexOf(Self);
  if (nIndex = -1) then
  begin
    raise Exception.Create(
      '(TTreeObjectItem.GetNextSibling)对象不在父对象的子列表中');
    Exit;
  end;
  if nIndex = Parent.Childs.Count - 1 then
    Exit;
  Result := Parent.Childs.Items[nIndex+1];
end;

function TTreeObjectItem.GetChecked: Boolean;
begin
  Result := FCheckState in [C_nChecked_Checked, C_nChecked_Grayed];
end;

procedure TTreeObjectItem.SetChecked(value: boolean);
begin
  if value then
    FCheckState := C_nChecked_Checked
  else
    FCheckState := C_nChecked_UnCheck;
    
  if not DisableCheckCascade and Assigned(Childs) then
    Childs.Checked := value;
  if not DisableCheckCascade and Assigned(Parent) and Assigned(OwnerList) then
  begin
    if OwnerList.AllChecked then
      Parent.CheckState := C_nChecked_Checked
    else if OwnerList.HasChecked then
    begin
      if AllowGrayed then
        Parent.CheckState := C_nChecked_Grayed
      else
        Parent.CheckState := C_nChecked_UnCheck;
    end
    else
      Parent.CheckState := C_nChecked_UnCheck;
  end;
end;

procedure TTreeObjectItem.CheckGray;
begin
  if not AllowGrayed then
    Exit;
  if not AllChecked then
    FCheckState := C_nChecked_Grayed;
end;

function TTreeObjectItem.AllChecked: Boolean;
begin
  Result := FCheckState = C_nChecked_Checked; // 实心选中状态，即为都选中的状态
end;

procedure TTreeObjectItem.SetDisableCheckCascade(value: Boolean);
begin
  Self.FDisableCheckCascade := value;
  if Assigned(Childs) then
    Childs.DisableCheckCascade := value;
end;

function TTreeObjectItem.IsLeaf: Boolean;
begin
  Result := (Childs = nil) or (Childs.Count = 0);
end;

end.
