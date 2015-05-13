{**
 * <p>Title: UF_ExportTables  </p>
 * <p>Copyright: Copyright (c) 20010/2/21</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 导出表界面</p>
 *}
unit UF_ExportTables;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls, CnCheckTreeView,
  U_Pub, DB, U_DBTreeFunc, U_TableInfo, U_DBEngineInterface,
  U_ThreadUtil,
  U_DBConnectInterface;

type
  TF_ExportTables = class(TForm)
    tvwObjects: TCnCheckTreeView;
    pnlBottom: TPanel;
    Label1: TLabel;
    edtExportFile: TEdit;
    btnSelect: TButton;
    btnOk: TButton;
    StatusBar1: TStatusBar;
    Label2: TLabel;
    grpOption: TGroupBox;
    chkDropTables: TCheckBox;
    chkCreateTables: TCheckBox;
    chkDisableTriggers: TCheckBox;
    chkDeleteTables: TCheckBox;
    btnOpenDir: TButton;
    pnlClient: TPanel;
    chkRemoveBreakLine: TCheckBox;
    chkClobInFile: TCheckBox;
    chkExportData: TCheckBox;
    procedure btnSelectClick(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnOpenDirClick(Sender: TObject);
    procedure chkDropTablesClick(Sender: TObject);
  private
    { Private declarations }
    FDBConnect: IDBConnect;
    FThread: TCustomThread;
    FThreadStop: Boolean; 
    FExporting: Boolean;
    FExportOption: TDBExportOption;
    // 导出建表脚本
    procedure DoExport();    
    procedure DoBeforeExport();
    procedure DoAfterExport();
    procedure ExportTable(node: TTreeNode; sPath: string);
    procedure ExportView(node: TTreeNode; sPath: string);
    procedure ExportProc(node: TTreeNode; sPath: string);
//    procedure AddExportTableSql(list: TStrings; sTableName: string);
    procedure AddDBTreeNodeChilds(tvw: TTreeView; Atnd: TTreeNode;
      bSystemObjects: Boolean; bSort: Boolean = false);
    procedure ShowStatus(s: string);
    procedure ExportTrig(node: TTreeNode; sPath: string);
    procedure DisableControls;
    procedure EnableControls;
  public
    procedure setDBConnect(dbConnect: IDBConnect);
    { Public declarations }
  end;

implementation

{$R *.dfm}

uses
  U_FileUtil, ShellAPI;
   
const
  C_nNodeStateIndex_UnChecked = 1;
  C_nNodeStateIndex_Checked = 2;
  
  C_BTNOK_CAPTION_RUN  = '导出';
  C_BTNOK_CAPTION_STOP = '停止'; 
  C_BTNOK_CAPTION_STOPING = '正在停止...';

{ TF_ExportTables }      

procedure TF_ExportTables.setDBConnect(dbConnect: IDBConnect);
begin
  FDBConnect := dbConnect;
end;

//procedure TF_ExportTables.AddExportTableSql(list: TStrings; sTableName: string);
//begin
//  if chkDropTable.Checked then
//  begin
//    list.Add(Format('drop table %s;', [sTableName]));
//  end;
//  if chkCreateTable.Checked then
//  begin
//    FDBConnect.ExportTableCreateSQL(sTableName, list);
//  end;
//  if chkDeleteRecords.Checked then
//  begin
//    list.Add(Format('delete from %s;', [sTableName]));
//  end;
//  FDBConnect.ExportTableToSql();
//end;      
         

procedure TF_ExportTables.ExportProc(node: TTreeNode; sPath: string);
var
  sName: string;
  slst: TStrings;
begin
  sName := node.Text;     
  ShowStatus(Format('正在导出%s',[sName]));
  slst := TStringList.Create;
  try
    FDBConnect.GetProcSource(sName, slst);
    slst.SaveToFile(sPath + C_FilePrefix_CreateProcdure + sName + '.sql');
  finally
    slst.Free;
  end;
end;

procedure TF_ExportTables.ExportTrig(node: TTreeNode; sPath: string);
var
  sName: string;
  slst: TStrings;
begin
  sName := node.Text;     
  ShowStatus(Format('正在导出%s',[sName]));
  slst := TStringList.Create;
  try
    FDBConnect.GetTrigSource(sName, slst);
    slst.SaveToFile(sPath + C_FilePrefix_CreateTrigger + sName + '.sql');
  finally
    slst.Free;
  end;
end;

procedure TF_ExportTables.ExportTable(node: TTreeNode; sPath: string);
var
  sName: string;
  slst: TStrings;
begin
  sName := node.Text;     // tablename
  ShowStatus(Format('正在导出%s',[sName]));
  slst := TStringList.Create;
  try
    if FExportOption.DropTables then
    begin
      slst.Add(Format('drop table %s;',[sName]));
    end;
    if FExportOption.CreateTables then
    begin
      FDBConnect.GenTableCreateSQL(sName, slst);
      // save to CT_xxxx.sql
      slst.SaveToFile(sPath + C_FilePrefix_CreateTable + sName + '.sql');
    end;
    // gen insert table & save to IT_xxxx.sql
    FDBConnect.ExportTableData(sName,
      sPath + C_FilePrefix_InsertTable + sName + '.sql', FExportOption);
  finally
    slst.Free;
  end;
end;

procedure TF_ExportTables.ExportView(node: TTreeNode; sPath: string); 
var
  sName: string;
  slst: TStrings;
begin
  sName := node.Text;     
  ShowStatus(Format('正在导出%s',[sName]));
  slst := TStringList.Create;
  try
    FDBConnect.GetViewSource(sName, slst);
    slst.SaveToFile(sPath + C_FilePrefix_CreateView + sName + '.sql');
  finally
    slst.Free;
  end;
end;

procedure TF_ExportTables.DoExport();
var
  node: TTreeNode;
  i: Integer;
  sPath: string;
begin
  sPath := edtExportFile.Text;
  
  if Copy(sPath, Length(sPath), 1 ) <> '\' then
    sPath := sPath + '\';
  for i := 0 to tvwObjects.Items.Count -1 do
  begin
    if FThreadStop then
      Break;
    node := tvwObjects.Items[i];
    if node.StateIndex <> C_nNodeStateIndex_Checked then
      Continue;
    case TDBTreeNodeData(node.Data).NodeTag of
    ntgTable:
      ExportTable(node, sPath);
    ntgView:
      ExportView(node, sPath);
    ntgProc:
      ExportProc(node, sPath);
    ntgTrig:
      ExportTrig(node, sPath);
    else
      //
    end;
  end;
end;    

procedure TF_ExportTables.DoBeforeExport;
begin
  btnOk.Caption := C_BTNOK_CAPTION_STOP; 
  DisableControls;  
  FExporting := True; 
  Application.ProcessMessages;
end;

procedure TF_ExportTables.DoAfterExport;
begin
  btnOk.Caption := C_BTNOK_CAPTION_RUN;
  EnableControls;    
  FExporting := False;

  if FThreadStop then 
    ShowStatus('用户中止导出！')
  else
    ShowStatus('导出完毕！');
end;

procedure TF_ExportTables.DisableControls;
var
  i: Integer;
begin
  for i := 0 to grpOption.ControlCount - 1 do
  begin
    grpOption.Controls[i].Enabled := False;
    if grpOption.Controls[i] is TCheckBox then
      (grpOption.Controls[i] as TCheckBox).Color := clBtnFace;
  end;
  edtExportFile.Enabled := False;
  edtExportFile.Color := clBtnFace;
end;

procedure TF_ExportTables.EnableControls;
var
  i: Integer;
begin
  for i := 0 to grpOption.ControlCount - 1 do
  begin
    grpOption.Controls[i].Enabled := True;
    if grpOption.Controls[i] is TCheckBox then
      (grpOption.Controls[i] as TCheckBox).Color := clWindow;
  end;        
  edtExportFile.Enabled := True;  
  edtExportFile.Color := clWindow;
end;

procedure TF_ExportTables.btnOkClick(Sender: TObject);
begin
  if FExporting then
  begin
    btnOk.Caption := C_BTNOK_CAPTION_STOPING;
    FThreadStop := True;
    FDBConnect.StopExec;
  end
  else
  begin
    if not DirectoryExists(edtExportFile.Text) then
    begin
      if IDYES = MessageBox(Handle, PChar('导出目录'
        + edtExportFile.Text + '不存在，是否创建？'),
        '提示', MB_YESNO + MB_ICONQUESTION) then
      begin
        if not ForceDirectories(edtExportFile.Text) then
        begin
          Exit;
        end
      end
      else
        Exit;
    end;
    FThreadStop := False;
    FExporting := True;
    FExportOption.DropTables := chkDropTables.Checked;
    FExportOption.CreateTables := chkCreateTables.Checked;
    FExportOption.DeleteTables := chkDeleteTables.Checked;
    FExportOption.DisableTriggers := chkDisableTriggers.Checked;
    FExportOption.RemoveBreakLine := chkRemoveBreakLine.Checked;
    FExportOption.ClobInFile := chkClobInFile.Checked;
    FExportOption.ExportData := chkExportData.Checked;

    FThread := TCustomThread.Create(TCustomThreadProxy.Create(DoExport), True);
    FThread.DoBeforeExecute := TCustomThreadProxy.Create(DoBeforeExport);
    FThread.DoAfterExecute := TCustomThreadProxy.Create(DoAfterExport);
    FThread.Resume;
  end;
end;

procedure TF_ExportTables.btnSelectClick(Sender: TObject);
begin
  edtExportFile.Text := TFileUtil.BrowseFolder(edtExportFile.Text);
end;

procedure TF_ExportTables.AddDBTreeNodeChilds(tvw: TTreeView; Atnd: TTreeNode;
  bSystemObjects: Boolean; bSort: Boolean = false);
var
  qry: TDataSet;
  slstNodes: TStringList;

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
//        if ntg in [ntgTable] then
//          tndAdded.HasChildren := True;        // 始终认为有字节点

        TDBTree.SetNodeImage(tndAdded);
        tndAdded.StateIndex := C_nNodeStateIndex_UnChecked;
      end;
//    end;
  end;
  
  procedure ProcessTypeRoot(node: TTreeNode);
  var
    i: Integer;   
    fnt: TFixNodeType;  
    aryNodes: TAryNodes;
  begin
    // 分析节点具体类型
    fnt := fntUnknown;
    SetLength(aryNodes, 0);
    aryNodes := GetAryFixNodes(FDBConnect.DBEngine.GetDBType);
    for i := 0 to High(aryNodes) do
    begin
      if node.Text = aryNodes[i].NodeText then
      begin
        fnt := aryNodes[i].NodeType;
        Break;
      end;
    end;

    // 显示用户对象还是系统对象
    FDBConnect.SystemObject := bSystemObjects;
    case fnt of
      fntAllTbl:
      begin
        FDBConnect.GetTableNames(slstNodes, qry);
        DoAddChilds(node, slstNodes, ntgTable);
      end;
      fntAllProc:
      begin
        FDBConnect.GetProcedureNames(slstNodes, qry);
        DoAddChilds(node, slstNodes, ntgProc);
      end;
      fntAllTri:
      begin
        FDBConnect.GetTriggerNames(slstNodes, qry);
        DoAddChilds(node, slstNodes, ntgTrig);
      end;
      fntAllView:
      begin
        FDBConnect.GetViewNames(slstNodes, qry);
        DoAddChilds(node, slstNodes, ntgView);
      end;
//      fntAcsMoudle:
//      begin 
//        (FDBConnect as TAccessDBConnect).GetMoudleNames(slstNodes, qry);
//        DoAddChilds(node, slstNodes, ntgView);
//      end;
//      fntAcsReport:
//      begin
//        (FDBConnect as TAccessDBConnect).GetReportNames(slstNodes, qry);
//        DoAddChilds(node, slstNodes, ntgView);
//      end;
//      fntAcsMacro:
//      begin
//        (FDBConnect as TAccessDBConnect).GetMacroNames(slstNodes, qry);
//        DoAddChilds(node, slstNodes, ntgView);
//      end;
      fntOraIndexs:
      begin   
        FDBConnect.GetIndexNames(
          slstNodes);
        DoAddChilds(node, slstNodes, ntgView);
      end;
//      fntOraSynonyms:
//      begin
//        (FDBConnect as TOracleDBConnect).GetSynonymNames(slstNodes);
//        DoAddChilds(node, slstNodes, ntgView);
//      end;
//      fntOraDBLinks:
//      begin
//        (FDBConnect as TOracleDBConnect).GetDBLinkNames(slstNodes);
//        DoAddChilds(node, slstNodes, ntgView);
//      end;
//      fntUsers:
//      begin      
//        (FDBConnect as TOracleDBConnect).GetUsers(slstNodes);
//        DoAddChilds(node, slstNodes, ntgView);
//      end;
    end;
  end;

  procedure ProcessTable(node: TTreeNode);
  var
    tableInfo: TTableInfo;
    childNode: TTreeNode;
    n: Integer;
  begin
    tableInfo := FDBConnect.GetTableInfo(Atnd.Text);
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
        childNode.Data := TDBTreeNodeData.Create(ntgField);

        if tableInfo.Fields.Items[n].IsPrimary then
          TDBTreeNodeData(childNode.Data).IsPrimaryField := True;
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
  if not g_Global.Connected then
    Exit;
  // 已有子节点
  if (Atnd.getFirstChild <> nil) then
    Exit;
  // 或者已尝试过添加但没有子节点
  if (not Atnd.HasChildren)
     or ((Atnd.ImageIndex = IMAGEINDEX_ROOT_EXPANDED)
          and (Atnd.getFirstChild = nil)) then
    Exit;      
  // 显示用户对象还是系统对象
  FDBConnect.SystemObject := bSystemObjects;

  qry := FDBConnect.GetNewQuery;
  slstNodes := TStringList.Create;
//  tvw.Items.BeginUpdate;
  try
    case TDBTreeNodeData(Atnd.Data).NodeTag of
      ntgTypeRoot:
      begin
        ProcessTypeRoot(Atnd);
      end;  
      ntgTable:
      begin
        ProcessTable(Atnd);
      end;
      else
        Atnd.DeleteChildren;     // 去除"正在加载..."节点
    end;

    if Atnd.getFirstChild = nil then
      Atnd.HasChildren := False;
    qry.Close;
    
//    slstNodes.Clear;
  finally
//    tvw.Items.EndUpdate;
    qry.Free;
    slstNodes.Free;
  end;
end;

procedure TF_ExportTables.FormShow(Sender: TObject);
var
  tndFixNode: TTreeNode;
  i: Integer;
  aryNodes: TAryNodes;
begin
  with tvwObjects.Items do
  begin
    tvwObjects.Items.BeginUpdate;
    tvwObjects.Items.Clear;

    aryNodes := GetAryFixNodes(dbtUnKnown);
    for i:= Low(aryNodes) to High(aryNodes) do
    begin
      tndFixNode := tvwObjects.Items.AddChild(nil, aryNodes[i].NodeText);
      tndFixNode.Data := TDBTreeNodeData.Create(ntgTypeRoot);
      tndFixNode.HasChildren := True;
      TDBTree.SetNodeImage(tndFixNode);
      tndFixNode.StateIndex := C_nNodeStateIndex_UnChecked;

      AddDBTreeNodeChilds(tvwObjects, tndFixNode, False, True);
    end;  
    tvwObjects.Items.EndUpdate;
  end;
  btnOk.Caption := C_BTNOK_CAPTION_RUN;
end;

procedure TF_ExportTables.ShowStatus(s: string);
begin
  StatusBar1.Panels[0].Text := s;
  Application.ProcessMessages;
end;

procedure TF_ExportTables.FormCreate(Sender: TObject);
begin
  FThreadStop := False;
  edtExportFile.Text := ExtractFilePath(Application.ExeName) + 'DBExport\';
  FExportOption := TDBExportOption.Create;
end;

procedure TF_ExportTables.FormDestroy(Sender: TObject);
begin
  FExportOption.Free;
end;

procedure TF_ExportTables.btnOpenDirClick(Sender: TObject);
begin
  ShellExecute( Application.Handle, 'open', 'explorer.exe',
    Pchar(edtExportFile.Text), nil, sw_shownormal);
end;

procedure TF_ExportTables.chkDropTablesClick(Sender: TObject);
begin
  if chkDropTables.Checked then
    chkCreateTables.Checked := True;
end;

end.
