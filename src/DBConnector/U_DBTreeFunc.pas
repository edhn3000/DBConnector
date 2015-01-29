unit U_DBTreeFunc;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ActnList,
  ADODB, ActnMan, ActnMenus, DB,
  U_TableInfo, U_DBConnectInterface, U_OracleDBConnect,
  U_DBEngineInterface;

const
  IMAGEINDEX_ROOT_NORMAL   = 0;
  IMAGEINDEX_ROOT_EXPANDED = 1;
  IMAGEINDEX_TABLE   = 2;
  IMAGEINDEX_PROC    = 3;
  IMAGEINDEX_TRIG    = 5;
  IMAGEINDEX_VIEW    = 6;
  IMAGEINDEX_FIELD   = 4;
  IMAGEINDEX_PRIFIELD = 8;

type
  { 固定节点类型 }
  TFixNodeType = (fntUnknown, fntAllTbl, fntAllView, fntAllProc, fntAllTri,
                  fntAcsMoudle, fntAcsReport, fntAcsMacro,
                  fntUsers, fntSequences, fntOraIndexs, fntOraSynonyms, fntOraDBLinks);
  TNodeTag = (ntgNone, ntgTypeRoot, ntgTable, ntgView, ntgField,
             ntgProc, ntgTrig, ntgIndex,
             ntgOraDBLink, ntgUser, ntgOraSynonyms,
             ntgAcsMoudle, ntgAcsReport, ntgAcsMacro);

  { 固定节点结构 }
  TFixNode = record
    NodeType: TFixNodeType;
    NodeText: string;
    sql_user: string;
    sql_all:  string;
  end;      
  TAryNodes = array of TFixNode;
             
  TDBTreeNodeData = class
  public
    NodeTag: TNodeTag;
    IsPrimaryField: Boolean;
  public
    constructor Create(tag: TNodeTag);
  end;

  TDBTree = class
  private
    FDBConnect: IDBConnect;
    function Connected: Boolean;
  protected    
    procedure initDBTreeForExport(tvw: TTreeView);
  public              
    procedure SetDBConnect(dbc: IDBConnect);

    procedure clearDBTree(tvw: TTreeView);
    // 初始化数据库浏览树的根节点
    procedure initDBTree(tvw: TTreeView);
    // 初始化数据库浏览树的根节点

    // 给数据库的浏览树节点增加子节点
    procedure AddDBTreeNodeChilds(tvw: TTreeView; Atnd: TTreeNode;
      bSystemObjects: Boolean; bSort: Boolean = false); 

    // 设置树节点图标索引 索引对应的具体图标见U_DataMoudle.ilTree控件
    class procedure SetNodeImage(ANode: TTreeNode);
  end;

  function GetAryFixNodes(dbt: TDBType): TAryNodes;  
  function IsFixNode(sNodeText: string; dbt: TDBType): Boolean;

const    
  gC_asFixNodes_Default: array[0..3] of TFixNode = (
    (NodeType: fntAllTbl;    NodeText: 'Tables';     sql_user: ''; sql_all: ''),
    (NodeType: fntAllView;   NodeText: 'Views';      sql_user: ''; sql_all: ''),
    (NodeType: fntAllProc;   NodeText: 'Procedures'; sql_user: ''; sql_all: ''),
    (NodeType: fntAllTri;    NodeText: 'Triggers';   sql_user: ''; sql_all: '')
    );

  gC_asFixNodes_Oracle: array[0..7] of TFixNode = (
    (NodeType: fntAllTbl;      NodeText: 'Tables';
     sql_user: SQL_ORACLE_GETTABLES_USER; sql_all: SQL_ORACLE_GETTEBLES_ALL),
    (NodeType: fntAllView;     NodeText: 'Views';
     sql_user: SQL_ORACLE_GETVIEWS_USER;  sql_all: SQL_ORACLE_GETVIEWS_ALL),
    (NodeType: fntAllProc;     NodeText: 'Procedures';
     sql_user: SQL_ORACLE_GETPROCS_USER;  sql_all: SQL_ORACLE_GETPROCS_ALL),
    (NodeType: fntAllTri;      NodeText: 'Triggers';
     sql_user: SQL_ORACLE_GETTRIGS_USER;  sql_all: SQL_ORACLE_GETTRIGS_ALL),
    (NodeType: fntOraIndexs;   NodeText: 'Indexs';
     sql_user: SQL_ORACLE_GETINDEXES_USER; sql_all: SQL_ORACLE_GETINDEXES_ALL),
    (NodeType: fntUsers;       NodeText: 'Users';
     sql_user: SQL_ORACLE_GETDBLINKS_USER; sql_all: SQL_ORACLE_GETDBLINKS_ALL),
    (NodeType: fntOraSynonyms; NodeText: 'Synonyms';
     sql_user: SQL_ORACLE_GETSYNONYMS_USER; sql_all: SQL_ORACLE_GETSYNONYMS_ALL),
    (NodeType: fntOraDBLinks;  NodeText: 'DBLinks';
     sql_user: SQL_ORACLE_GETDBLINKS_USER; sql_all: SQL_ORACLE_GETDBLINKS_ALL)
    );

  gC_asFixNodes_Access: array[0..4] of TFixNode = (
    (NodeType: fntAllTbl;    NodeText: 'Tables';  sql_user: ''; sql_all: ''),
    (NodeType: fntAllView;   NodeText: 'Querys';  sql_user: ''; sql_all: ''),
    (NodeType: fntAcsMoudle; NodeText: 'Moudles'; sql_user: ''; sql_all: ''),
    (NodeType: fntAcsReport; NodeText: 'Reports'; sql_user: ''; sql_all: ''),
    (NodeType: fntAcsMacro;  NodeText: 'Macros';  sql_user: ''; sql_all: '')
    );  

  gC_asFixNodes_MySQL: array[0..3] of TFixNode = (
    (NodeType: fntAllTbl;    NodeText: 'Tables';     sql_user: ''; sql_all: ''),
    (NodeType: fntAllView;   NodeText: 'Views';      sql_user: ''; sql_all: ''),
    (NodeType: fntAllProc;   NodeText: 'Procedures'; sql_user: ''; sql_all: ''),
    (NodeType: fntAllTri;    NodeText: 'Triggers';   sql_user: ''; sql_all: '')
    );

var
  g_DBTreeFunc: TDBTree;

implementation

function GetAryFixNodes(dbt: TDBType): TAryNodes;
var
  i: Integer;
begin
  case dbt of
  dbtOracle:
  begin
    SetLength(Result, Length(gC_asFixNodes_Oracle));
    for i := 0 to High(Result) do
    begin
      Result[i] := gC_asFixNodes_Oracle[i];
    end;  
  end;
  dbtAccess, dbtAccess2007:
  begin
    SetLength(Result, Length(gC_asFixNodes_Access));   
    for i := 0 to High(Result) do
    begin                
      Result[i] := gC_asFixNodes_Access[i];
    end;
  end;
  else
  begin
    SetLength(Result, Length(gC_asFixNodes_Default)); 
    for i := 0 to High(Result) do
    begin                     
      Result[i] := gC_asFixNodes_Default[i];
    end;  
  end;
  end;
end;  

function IsFixNode(sNodeText: string; dbt: TDBType): Boolean;
var
  i: Integer;
  aryNodes: TAryNodes;
begin
  Result := False;
  aryNodes := GetAryFixNodes(dbt);
  for i := Low(aryNodes) to High(aryNodes) do
  begin
    if SameText(sNodeText, aryNodes[i].NodeText) then
    begin
      Result := True;
      Break;
    end;
  end;
end;
               

{ TDBTreeNodeData }

constructor TDBTreeNodeData.Create(tag: TNodeTag);
begin
  NodeTag := tag;
end;

{ TDBTree }

class procedure TDBTree.SetNodeImage(ANode: TTreeNode);
begin
  if ANode.Data = nil then
    Exit;
  case TDBTreeNodeData(ANode.Data).NodeTag of
    ntgTypeRoot:
    begin
      if not ANode.HasChildren then
      begin
        ANode.ImageIndex := IMAGEINDEX_ROOT_EXPANDED;
        ANode.SelectedIndex := IMAGEINDEX_ROOT_EXPANDED;
      end 
      else
      begin
        if (ANode.getFirstChild = nil) then
        begin
          ANode.ImageIndex := IMAGEINDEX_ROOT_NORMAL;
          ANode.SelectedIndex := IMAGEINDEX_ROOT_NORMAL;
        end
        else if not ANode.getFirstChild.IsVisible then
        begin
          ANode.ImageIndex := IMAGEINDEX_ROOT_NORMAL;
          ANode.SelectedIndex := IMAGEINDEX_ROOT_NORMAL;
        end
        else
        begin
          ANode.ImageIndex := IMAGEINDEX_ROOT_EXPANDED;
          ANode.SelectedIndex := IMAGEINDEX_ROOT_EXPANDED;
        end;
      end;
    end;
    ntgTable:
    begin
      ANode.ImageIndex := IMAGEINDEX_TABLE;
      ANode.SelectedIndex := IMAGEINDEX_TABLE;
    end;
    ntgProc:
    begin
      ANode.ImageIndex := IMAGEINDEX_PROC;
      ANode.SelectedIndex := IMAGEINDEX_PROC;
    end;
    ntgField:
    begin
      if TDBTreeNodeData(ANode.Data).IsPrimaryField then
      begin
        ANode.ImageIndex := IMAGEINDEX_PRIFIELD;
        ANode.SelectedIndex := IMAGEINDEX_PRIFIELD;
      end
      else
      begin          
        ANode.ImageIndex := IMAGEINDEX_FIELD;
        ANode.SelectedIndex := IMAGEINDEX_FIELD;
      end;
    end;
    ntgTrig:
    begin
      ANode.ImageIndex := IMAGEINDEX_TRIG;
      ANode.SelectedIndex := IMAGEINDEX_TRIG
    end;
    ntgView:
    begin
      ANode.ImageIndex := IMAGEINDEX_VIEW;
      ANode.SelectedIndex := IMAGEINDEX_VIEW;
    end;
    else
    begin
      ANode.ImageIndex := IMAGEINDEX_VIEW;
      ANode.SelectedIndex := IMAGEINDEX_VIEW;
    end;
  end;
end;

procedure TDBTree.clearDBTree(tvw: TTreeView);
var
  i: Integer;
begin
  for i := tvw.Items.Count - 1 downto 0 do
  begin
    if (tvw.Items[i].Data <> nil)
       and (TObject(tvw.Items[i].Data) is TObject) then
      TObject(tvw.Items[i].Data).Free;
  end;
  tvw.Items.Clear;
end;

procedure TDBTree.initDBTree(tvw: TTreeView);
var
  tndFixNode: TTreeNode;
  i: Integer;
  aryNodes: TAryNodes;
  db: IDBConnect;
begin
  db := FDBConnect;
  tvw.Items.BeginUpdate;
  clearDBTree(tvw);
  // 根据数据库类型初始化树
  if Self.Connected then
    aryNodes := GetAryFixNodes(db.DBEngine.GetDBType)
  else
    aryNodes := GetAryFixNodes(dbtUnKnown);
  for i:= Low(aryNodes) to High(aryNodes) do
  begin
    tndFixNode := tvw.Items.AddChild(nil, aryNodes[i].NodeText);
    tndFixNode.Data := TDBTreeNodeData.Create(ntgTypeRoot);
    tndFixNode.HasChildren := True;
    SetNodeImage(tndFixNode);
  end;
  tvw.Items.EndUpdate;
end;  

procedure TDBTree.initDBTreeForExport(tvw: TTreeView);
var
  tndFixNode: TTreeNode;
  i: Integer;
  aryNodes: TAryNodes;
begin
  tvw.Items.BeginUpdate;
  for i := tvw.Items.Count - 1 downto 0 do
  begin
    if (tvw.Items[i].Data <> nil)
       and (TObject(tvw.Items[i].Data) is TObject) then
      TObject(tvw.Items[i].Data).Free;
  end;
  tvw.Items.Clear;
  // 始终使用UnKnown类型
  if Self.Connected then
    aryNodes := GetAryFixNodes(dbtUnKnown)
  else
    aryNodes := GetAryFixNodes(dbtUnKnown);
  for i:= Low(aryNodes) to High(aryNodes) do
  begin
    tndFixNode := tvw.Items.AddChild(nil, aryNodes[i].NodeText);
    tndFixNode.Data := TDBTreeNodeData.Create(ntgTypeRoot);    // 表示是一类节点的根节点
    tndFixNode.HasChildren := True;
    SetNodeImage(tndFixNode);
  end;
  tvw.Items.EndUpdate;
end; 

procedure TDBTree.AddDBTreeNodeChilds(tvw: TTreeView; Atnd: TTreeNode;
  bSystemObjects: Boolean; bSort: Boolean = false);
var
  qry: TDataSet;
  slstNodes: TStringList;
  db: IDBConnect;

  procedure DoAddChilds(ANode: TTreeNode; list: TStringList; ntg: TNodeTag);
  var
    i: Integer;
    tndAdded: TTreeNode;
  begin
//    with F_MAIN do
//    begin
      ANode.DeleteChildren;
      if bSort then
        list.Sort;
      for i := 0 to list.Count - 1 do
      begin
        tndAdded := tvw.Items.AddChild(ANode, list.Strings[i]);
        tndAdded.Data := TDBTreeNodeData.Create(ntg);          // 节点标记TNodeTag类型
        if ntg in [ntgTable] then
          tndAdded.HasChildren := True;        // 始终认为有字节点

        TDBTree.SetNodeImage(tndAdded);  
        Application.ProcessMessages;
      end;
//    end;
  end;

  procedure ProcessTypeRoot(db: IDBConnect; node: TTreeNode);
  var
    i: Integer;   
    fnt: TFixNodeType;  
    aryNodes: TAryNodes;
  begin
    // 分析节点具体类型
    fnt := fntUnknown;
    SetLength(aryNodes, 0);
    aryNodes := GetAryFixNodes(db.DBEngine.GetDBType);
    for i := 0 to High(aryNodes) do
    begin
      if node.Text = aryNodes[i].NodeText then
      begin
        fnt := aryNodes[i].NodeType;
        Break;
      end;
    end;

    // 显示用户对象还是系统对象
    db.SystemObject := bSystemObjects;
    case fnt of
      fntAllTbl:
      begin
        db.GetTableNames(slstNodes, qry);
        DoAddChilds(node, slstNodes, ntgTable);
      end;
      fntAllProc:
      begin
        db.GetProcedureNames(slstNodes, qry);
        DoAddChilds(node, slstNodes, ntgProc);
      end;
      fntAllTri:
      begin
        db.GetTriggerNames(slstNodes, qry);
        DoAddChilds(node, slstNodes, ntgTrig);
      end;
      fntAllView:
      begin
        db.GetViewNames(slstNodes, qry);
        DoAddChilds(node, slstNodes, ntgView);
      end;
      fntAcsMoudle:
      begin
        db.GetMoudleNames(slstNodes, qry);
        DoAddChilds(node, slstNodes, ntgAcsMoudle);
      end;
      fntAcsReport:
      begin
        db.GetReportNames(slstNodes, qry);
        DoAddChilds(node, slstNodes, ntgAcsReport);
      end;
      fntAcsMacro:
      begin
        db.GetMacroNames(slstNodes, qry);
        DoAddChilds(node, slstNodes, ntgAcsMacro);
      end;
      fntOraIndexs:
      begin
        db.GetIndexNames(slstNodes);
        DoAddChilds(node, slstNodes, ntgIndex);
      end;
      fntOraSynonyms:
      begin
        db.GetSynonymNames(slstNodes);
        DoAddChilds(node, slstNodes, ntgOraSynonyms);
      end;
      fntOraDBLinks:
      begin
        db.GetDBLinkNames(slstNodes);
        DoAddChilds(node, slstNodes, ntgOraDBLink);
      end;
      fntUsers:
      begin
        db.GetUsers(slstNodes);
        DoAddChilds(node, slstNodes, ntgUser);
      end;
    end;
  end;

  procedure ProcessTable(db: IDBConnect; node: TTreeNode);
  var
    tableInfo: TTableInfo;
    childNode: TTreeNode;
    n: Integer;
  begin
    tableInfo := db.GetTableInfo(Atnd.Text);
    try
      node.DeleteChildren;   // 去除"正在加载..."节点
      if bSort then
        tableInfo.Fields.SortByFieldName;
      for n := 0 to tableInfo.Fields.Count - 1 do
      begin
        childNode := tvw.Items.AddChild(Atnd, Format('%s %s %s', [
          tableInfo.Fields.Items[n].FieldName,
          tableInfo.Fields.Items[n].DataTypeStr,
          tableInfo.Fields.Items[n].GetNullableStr]));
        if tableInfo.Fields.Items[n].IsPrimary then
        begin
          childNode.Data := TDBTreeNodeData.Create(ntgField);  //TObject(ntgPrimaryField)
          TDBTreeNodeData(childNode.Data).IsPrimaryField := True;
        end
        else
          childNode.Data := TDBTreeNodeData.Create(ntgField); // TObject(ntgField);
        TDBTree.SetNodeImage(childNode);
      end;
    finally
      tableInfo.Free;
    end;
  end;   
//////// AddChilds Begin
begin
  if not Assigned(Atnd) then
    Exit;
  if not Self.Connected then
    Exit; 
  db := FDBConnect;
  // 已有子节点
  if (Atnd.getFirstChild <> nil) then
    Exit;
  // 或者已尝试过添加但没有子节点
  if (not Atnd.HasChildren)
     or ((Atnd.ImageIndex = IMAGEINDEX_ROOT_EXPANDED)
          and (Atnd.getFirstChild = nil)) then
    Exit;      
  // 显示用户对象还是系统对象
  db.SystemObject := bSystemObjects;

  tvw.Items.BeginUpdate;
  tvw.Items.AddChild(Atnd, '正在加载...').ImageIndex := -1;
  tvw.Items.EndUpdate;  
  Atnd.Expand(False);
  
  Application.ProcessMessages;
  qry := db.GetNewQuery;
  slstNodes := TStringList.Create;
  tvw.Items.BeginUpdate;
  try
    case TDBTreeNodeData(Atnd.Data).NodeTag of
      ntgTypeRoot:
        ProcessTypeRoot(FDBConnect, Atnd);
      ntgTable:
        ProcessTable(FDBConnect, Atnd);
      else
      begin
        Atnd.DeleteChildren;     // 去除"正在加载..."节点
      end;
    end;

    if Atnd.getFirstChild = nil then
    begin
      Atnd.HasChildren := False;
    end;
    
    qry.Close;
    slstNodes.Clear;
  finally
    tvw.Items.EndUpdate;
    qry.Free;
    slstNodes.Free;
  end;
end;

procedure TDBTree.SetDBConnect(dbc: IDBConnect);
begin
  Self.FDBConnect := dbc;
end;

function TDBTree.Connected: Boolean;
begin
  Result := Assigned(FDBConnect) and FDBConnect.Connected;
end;

initialization
  g_DBTreeFunc := TDBTree.Create;

finalization
  g_DBTreeFunc.Free;

end.
