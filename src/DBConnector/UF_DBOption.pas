unit UF_DBOption;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls, Menus, ImgList, UF_BaseDialog,
  U_DBConfig, U_DBEngineInterface, U_Base64Codec, jpeg;

type
  TF_DBOption = class(TF_BaseDialog)
    pgcOptions: TPageControl;
    ts1: TTabSheet;
    lbl1: TLabel;
    lbl6: TLabel;
    edtAcsSource: TEdit;
    edtAcsSec: TEdit;
    btnChooseDB: TButton;
    btnChooseSec: TButton;
    chkAcs2007: TCheckBox;
    ts2: TTabSheet;
    lblOraSID: TLabel;
    lblOraIP: TLabel;
    lblOraPort: TLabel;
    lblOraPro: TLabel;
    edtOraIP: TEdit;
    edtOraPort: TEdit;
    cbbOraPro: TComboBox;
    tsSybase: TTabSheet;
    edtSybServerName: TEdit;
    pnl1: TPanel;
    btnReadConfig: TButton;
    pmConfig: TPopupMenu;
    ilTree: TImageList;
    btnApply: TButton;
    grp1: TGroupBox;
    lbl2: TLabel;
    lbl5: TLabel;
    cbbDBType: TComboBox;
    cbbEngineType: TComboBox;
    grp2: TGroupBox;
    lbl7: TLabel;
    lbl8: TLabel;
    cbbUser: TComboBox;
    edtPwd: TEdit;
    chkIsLocalTNS: TCheckBox;
    tsMySql: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    lbl9: TLabel;
    lbl10: TLabel;
    Label4: TLabel;
    GroupBox1: TGroupBox;
    lbl11: TLabel;
    edtRecentDBCount: TEdit;
    btnTest: TButton;
    cbbOraSID: TComboBox;
    chkSingle: TCheckBox;
    edtmslHost: TComboBox;
    edtmslDataBase: TComboBox;
    Label5: TLabel;
    edtSybPort: TEdit;
    Label6: TLabel;
    Label7: TLabel;
    cbbSybDataBase: TComboBox;
    Label8: TLabel;
    edtMysqlCharset: TEdit;
    tsSqlite: TTabSheet;
    Label9: TLabel;
    edtSqliteDBFile: TEdit;
    btnChooseSqliteDB: TButton;
    Label10: TLabel;
    procedure FormShow(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure btnChooseDBClick(Sender: TObject);
    procedure btnChooseSecClick(Sender: TObject);
    procedure cbbDBTypeChange(Sender: TObject);
    procedure pgcOptionsChange(Sender: TObject);
    procedure btnReadConfigClick(Sender: TObject);
    procedure btnApplyClick(Sender: TObject);
    procedure cbbEngineTypeChange(Sender: TObject);
    procedure btnAddToConfigClick(Sender: TObject);
    procedure chkIsLocalTNSClick(Sender: TObject);
    procedure edtRecentDBCountKeyPress(Sender: TObject; var Key: Char);
    procedure btnTestClick(Sender: TObject);
    procedure cbbOraSIDClick(Sender: TObject);
    procedure edtmslDataBaseDropDown(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    FDBConfigList: TDBConfigList;

    function Apply: Boolean;
                            
    procedure LoadOdbc;
    procedure LoadTnsNames;

    procedure PmConfigItemCLick(Sender: TObject);
    procedure ChangeCurrnetUser;
    procedure SetCtrlEnable(ctrl: TWinControl; bEnable: Boolean);
    procedure RefreshUI;
    procedure DBEngineTypeToCbxIndex(dbet: TDBEngineType);
    procedure CbxIndexToDBEngineType(var dbet: TDBEngineType);
    procedure CbxIndexToDBType(var dbt: TDBType);
    procedure DBTypeToCbxIndex(dbt: TDBType);

  protected                         
    procedure InitByDBInfo(Adbi: TDBConfig);      
    function LoadSettings: Boolean;
    function SaveSettings: Boolean;
    
    function EncodeStr(s: string): string;
    function DecodeStr(s: string): string;
  public
    { Public declarations }
    procedure SetDBConfigList(value: TDBConfigList);
  end;

var
  F_DBOption: TF_DBOption;

implementation

{$R *.dfm}    

uses
  U_Ini, U_DBUtil;


const
  C_nDBTypeItemIndex_None   = 0;
  C_nDBTypeItemIndex_Access = 1;
  C_nDBTypeItemIndex_Oracle = 2;
  C_nDBTypeItemIndex_Sybase = 3;
  C_nDBTypeItemIndex_MySQL  = 4;
  C_nDBTypeItemIndex_SQLite = 5;
  CBB_DBTYPE_ITEMS: array[0..5] of String = ('', 'Access', 'Oracle', 'Sybase',
    'MySQL', 'SQLite');

  C_nDBEngineTypeItemIndex_Auto = 0;
  C_nDBEngineTypeItemIndex_Ado  = 1;
  C_nDBEngineTypeItemIndex_Bde  = 2;
  C_nDBEngineTypeItemIndex_DBExpress  = 9;
  C_nDBEngineTypeItemIndex_SQLite  = 3;
  CBB_DBENGINETYPE_ITEMS: array[0..3] of String = ('自动选择', 'ADO驱动', 'BDE驱动', 'SQLite驱动');

  C_PageIndex_Access = 0;
  C_PageIndex_Oracle = 1;
  C_PageIndex_Sybase = 2;
  C_PageIndex_MySQL  = 3;
  C_PageIndex_SQLite = 4;

  C_nSybDBTypeItemIndex_ASE = 0;
  C_nSybDBTypeItemIndex_ASA = 1;

procedure TF_DBOption.SetDBConfigList(value: TDBConfigList);
var
  dbnamesFile: string;
begin
  FDBConfigList := value;
  dbnamesFile := ExtractFilePath(Application.ExeName) + 'DBNames.xml';
  if FileExists(dbnamesFile) then
  begin
    FDBConfigList.Clear;
    FDBConfigList.LoadFromFile(dbnamesFile);
  end;
end;

function TF_DBOption.EncodeStr(s: string): string;
begin
  Result := Base64Encode(s);
end;

function TF_DBOption.DecodeStr(s: string): string;
begin
  Result := Base64Decode(s);
end;

function TF_DBOption.LoadSettings: Boolean;
var
  mni: TMenuItem;
  i: Integer;
  bIsFix: Boolean;
  dbcfg: TDBConfig;
  showName: string;
  function FillStrLength(s: string; nLen: Integer): string;
  begin
    while Length(s) < nLen do
      s := s + ' ';
    Result := s;
  end;
begin
  with GlobalParams do
  begin
    chkSingle.Checked := SingleInstance;
    case DBconfig.DBType of
      dbtAccess, dbtAccess2007:
      begin
        cbbDBType.ItemIndex := C_nDBTypeItemIndex_Access;
        pgcOptions.ActivePageIndex := C_PageIndex_Access;
        cbbUser.Text := DBconfig.AcsUser;
        edtPwd.Text := DBconfig.getEncodePwd(dbtAccess);// Base64Encode(AcsPwd);
        chkAcs2007.Checked := (DBconfig.DBType = dbtAccess2007);
      end;
      dbtOracle:
      begin
        cbbDBType.ItemIndex := C_nDBTypeItemIndex_Oracle;
        pgcOptions.ActivePageIndex := C_PageIndex_Oracle;
        cbbUser.Text := DBconfig.OraUser;
        edtPwd.Text := DBconfig.getEncodePwd(dbtOracle);
      end;
      dbtSybase:
      begin
        cbbDBType.ItemIndex := C_nDBTypeItemIndex_Sybase;
        pgcOptions.ActivePageIndex := C_PageIndex_Sybase;
        cbbUser.Text := DBconfig.SybUser;
        edtPwd.Text := DBconfig.getEncodePwd(dbtSyBase);
//        cbbSybaseType.ItemIndex := SybType;
      end;
      dbtMySQL:
      begin
        cbbDBType.ItemIndex := C_nDBTypeItemIndex_MySQL;
        pgcOptions.ActivePageIndex := C_PageIndex_MySQL;
        cbbUser.Text := DBconfig.mslUser;
        edtPwd.Text := DBconfig.getEncodePwd(dbtMySQL);
      end;
      dbtSqlite:
      begin
        cbbDBType.ItemIndex := C_nDBTypeItemIndex_SQLite;
        pgcOptions.ActivePageIndex := C_PageIndex_SQLite;
      end;
      else
      begin
        cbbDBType.ItemIndex := 0;
        pgcOptions.ActivePageIndex := 0;
      end;
    end;
    DBEngineTypeToCbxIndex(DBconfig.DBEngineType);
    DBTypeToCbxIndex(DBconfig.DBType);
    cbbOraPro.ItemIndex := cbbOraPro.Items.IndexOf(UpperCase(DBconfig.Protocol));

    edtAcsSource.Text := DBconfig.AcsDataSource;
    edtAcsSec.Text := DBconfig.AcsSecuredDB;

    cbbOraSID.Text := DBconfig.SID;
    edtOraIP.Text := DBconfig.Host;
    edtOraPort.Text := DBconfig.Port;
    chkIsLocalTNS.Checked := DBconfig.IsLocalTns;

    edtSybServerName.Text := DBconfig.SybServerName;
    edtSybPort.Text := DBconfig.SybPort;
    cbbSybDataBase.Text := DBconfig.SybDatabaseName;

    edtmslHost.Text := DBconfig.mslHost;
    edtmslDataBase.Text := DBconfig.mslDataBase;
    edtMysqlCharset.Text := DBconfig.mslCharset;

    edtSqliteDBFile.Text := DBconfig.sltDB;

    edtRecentDBCount.Text := IntToStr(RecentDBCount);
  end;

  pmConfig.Items.Clear;
  bIsFix := False;
  if Assigned(FDBConfigList) then
  begin
    for i := 0 to FDBConfigList.Count - 1 do
    begin
      dbcfg := FDBConfigList.Items[i];
      if not bIsFix and dbcfg.IsFixDB then
      begin
        mni := TMenuItem.Create(Self);
        mni.Caption := '-';
        pmConfig.Items.Add(mni);
        bIsFix := True;
        Continue;
      end;
      mni := TMenuItem.Create(Self);
      showName := dbcfg.DBShownName;
      if showName = '' then
        showName := dbcfg.GetSeparatedDataSource;
      mni.Caption := Format( '%s_%s %s@%s', [
        FillStrLength(DBTypeToStr(dbcfg.DBType,False), 6),
        FillStrLength(DBEngineTypeToStr(dbcfg.DBEngineType, False), 4),
        dbcfg.UserName, showName]);
      mni.AutoHotkeys := maManual;
      mni.ShortCut := 0;
      mni.Tag := i;
      mni.OnClick := PmConfigItemCLick;

      pmConfig.Items.Add(mni);
    end;
  end;

  LoadTnsNames;
  LoadOdbc;
  Result := True;
end;

function TF_DBOption.SaveSettings: Boolean;
begin
  with GlobalParams do
  begin
    SingleInstance := chkSingle.Checked;
    case cbbDBType.ItemIndex of
      C_nDBTypeItemIndex_Access:
      begin
        if chkAcs2007.Checked then
          DBconfig.DBType := dbtAccess2007
        else
          DBconfig.DBType := dbtAccess;
        DBconfig.AcsUser := cbbUser.Text;
        DBconfig.setEncodePwd(dbtAccess, edtPwd.Text);
      end;
      C_nDBTypeItemIndex_Oracle:
      begin
        DBconfig.DBType := dbtOracle;
        DBconfig.OraUser := cbbUser.Text;
        DBconfig.setEncodePwd(dbtOracle, edtPwd.Text);
      end;
      C_nDBTypeItemIndex_Sybase:
      begin
        DBconfig.DBType := dbtSyBase;
        DBconfig.SybUser := cbbUser.Text;
        DBconfig.setEncodePwd(dbtSyBase, edtPwd.Text);
        CbxIndexToDBType(DBconfig.DBType);
      end;
      C_nDBTypeItemIndex_MySQL:
      begin
        DBconfig.DBType := dbtMySQL;
        DBconfig.mslUser := cbbUser.Text;
        DBconfig.setEncodePwd(dbtMySQL, edtPwd.Text);
      end;
      C_nDBTypeItemIndex_SQLite:
      begin
        DBconfig.DBType := dbtSqlite;
      end
    else
    end;
    CbxIndexToDBEngineType(DBconfig.DBEngineType);

    if cbbOraPro.ItemIndex <> -1 then
      DBconfig.Protocol := cbbOraPro.Items[cbbOraPro.ItemIndex];

    DBconfig.AcsDataSource := edtAcsSource.Text;
    DBconfig.AcsSecuredDB := edtAcsSec.Text;

    DBconfig.SID := cbbOraSID.Text;
    DBconfig.Host := edtOraIP.Text;
    DBconfig.Port := edtOraPort.Text;

    DBconfig.SybServerName := edtSybServerName.Text;
    DBconfig.SybDatabaseName := cbbSybDataBase.Text;
    DBconfig.SybPort := edtSybPort.Text;

    DBconfig.mslHost := edtmslHost.Text;
    DBconfig.mslDataBase := edtmslDataBase.Text;
    DBconfig.mslCharset := edtMysqlCharset.Text;

    DBconfig.sltDB := edtSqliteDBFile.Text;

    DBconfig.IsLocalTns := chkIsLocalTNS.Checked;
    RecentDBCount := StrToInt(edtRecentDBCount.Text);
    if Assigned(FDBConfigList) then
      FDBConfigList.RecentDBCount := RecentDBCount;

    SaveSettings;
  end;

  Result := True;
end;

procedure TF_DBOption.FormCreate(Sender: TObject);
var
  i: Integer;
begin
  inherited;

  cbbDBType.Items.Clear;
  for I := 0 to High(CBB_DBTYPE_ITEMS) do begin
    cbbDBType.Items.Add(CBB_DBTYPE_ITEMS[i]);
  end;
end;

procedure TF_DBOption.FormShow(Sender: TObject);
begin
  inherited;
  ChangeMainDescription('设置数据库参数');
  ChangeSubDescription('修改数据库参数后保存。点击『读取配置』可从配置文件中读取。');

  LoadSettings;
  RefreshUI;
  // TODO: 注册表HKEY_LOCAL_MACHINE\SOFTWARE\ORACLE\KEY_OraDb10g_home1\  ORACLE_HOME  + NETWORK\ADMIN\tnsnames.ora
end;

procedure TF_DBOption.btnOkClick(Sender: TObject);
begin
  if Apply then
    ModalResult := mrOK;
end;

procedure TF_DBOption.btnChooseDBClick(Sender: TObject);
var
  sInitDir: string;
  sDBExt: string;
begin
  with TOpenDialog.Create(Self) do
  try
    Filter := 'Access DB|*.mdb;*.accdb|*.*|*.*';
    sInitDir := ExtractFileDir(edtAcsSource.Text);
    if sInitDir = '' then
      sInitDir := GlobalParams.LastDir;
    if sInitDir = '' then
      sInitDir := ExtractFilePath(Application.ExeName);

    InitialDir := sInitDir;
    if Execute then
    begin
      edtAcsSource.Text := FileName;
      GlobalParams.LastDir := FileName;

      sDBExt := ExtractFileExt(FileName);
      if sDBExt = '.accdb' then
        chkAcs2007.Checked := True
      else if sDBExt = '.mdb' then
        chkAcs2007.Checked := False;
    end;
  finally
    Free;
  end;
end;

procedure TF_DBOption.btnChooseSecClick(Sender: TObject);
var
  sInitDir: string;
begin
  with TOpenDialog.Create(Self) do
  try
    Filter := '*.mdw|*.mdw|*.*|*.*';
    sInitDir := ExtractFileDir(edtAcsSec.Text);
    if sInitDir = '' then
      sInitDir := GlobalParams.LastDir;
    if sInitDir = '' then
      sInitDir := ExtractFilePath(Application.ExeName);

    InitialDir := sInitDir;
    if Execute then
    begin
      edtAcsSec.Text := FileName;
      GlobalParams.LastDir := FileName;
    end;
  finally
    Free;
  end;
end;

procedure TF_DBOption.cbbDBTypeChange(Sender: TObject);
begin
  if (cbbDBType.ItemIndex > 0) and (pgcOptions.PageCount >cbbDBType.ItemIndex-1) then
    pgcOptions.ActivePageIndex := cbbDBType.ItemIndex-1;
  pgcOptions.Enabled := cbbDBType.ItemIndex <> C_nDBTypeItemIndex_None;
  ChangeCurrnetUser;
end;

procedure TF_DBOption.pgcOptionsChange(Sender: TObject);
begin
  if pgcOptions.ActivePageIndex + 1 < cbbDBType.Items.Count then
  cbbDBType.ItemIndex := pgcOptions.ActivePageIndex + 1;
  ChangeCurrnetUser;
end;

procedure TF_DBOption.PmConfigItemCLick(Sender: TObject);
var
  dbi: TDBConfig;
  i: Integer;
  slst: TStrings;
begin
  with Sender as TMenuItem do
  begin
    dbi := FDBConfigList.Items[Tag];
    InitByDBInfo(dbi);
    cbbUser.Items.Clear;
    ExtractStrings([','], [' '], PChar(dbi.UserName), cbbUser.Items);
//    fStrUtil.Split(dbi.UserName, ',', cbbUser.Items);
    // 多个用户名
    for i := 0 to FDBConfigList.Count -1 do
    begin
      if i = Tag then Continue;
      if dbi.IsSameDB(FDBConfigList.Items[i]) then
      begin
        slst := TStringList.Create;
        ExtractStrings([','], [' '],
          PChar(FDBConfigList.Items[i].UserName), slst);
//        fStrUtil.Split(FDBConfigList.Items[i].UserName, ',', nil);
        cbbUser.Items.AddStrings(slst);
        slst.Free;
      end;
    end;
    if dbi.IsPasswordEncoded then
      edtPwd.Text := DecodeStr(dbi.Password)
    else
      edtPwd.Text := dbi.Password;
  end;
end;

procedure TF_DBOption.btnReadConfigClick(Sender: TObject);
var
  X, Y: Integer;
  control: TControl;
begin
  X := btnReadConfig.Left + btnReadConfig.Width;
  Y := btnReadConfig.Top;
  control := btnReadConfig;
  while not (control.Parent is TForm) do
  begin
    X := X + control.Parent.Left;
    Y := Y + control.Parent.Top;
    control := control.Parent;
  end;
  X := X + Self.Left + Self.Width - Self.ClientWidth;
  Y := Y + Self.Top + Self.Height - Self.ClientHeight;

  pmConfig.Popup(X, Y);
end;

procedure TF_DBOption.InitByDBInfo(Adbi: TDBConfig);
var
  slst: TStrings;
begin
  slst := TStringList.Create;
  try
    with Adbi do
    begin
      case DBType of
         dbtAccess, dbtAccess2007:
          begin
            cbbDBType.ItemIndex := C_nDBTypeItemIndex_Access;
            pgcOptions.ActivePageIndex := C_PageIndex_Access;
            edtAcsSource.Text := AcsDataSource;
            edtAcsSec.Text := AcsSecuredDB;
            chkAcs2007.Checked := (DBType = dbtAccess2007);
          end;
        dbtOracle:
          begin
            cbbDBType.ItemIndex := C_nDBTypeItemIndex_Oracle;
            pgcOptions.ActivePageIndex := C_PageIndex_Oracle;
            cbbOraSID.Text := SID;
            chkIsLocalTNS.Checked := IsLocalTns;
            if not IsLocalTns then
            begin
              edtOraIP.Text := Host;
              edtOraPort.Text := Port;
              if SameText(Protocol, 'tcp') then
                cbbOraPro.ItemIndex := 1
              else if SameText(Protocol, 'tcps') then
                cbbOraPro.ItemIndex := 2
              else if SameText(Protocol, 'ipc') then
                cbbOraPro.ItemIndex := 3
              else
                cbbOraPro.ItemIndex := 0;
            end;
          end;
        dbtSybase:
        begin
          cbbDBType.ItemIndex := C_nDBTypeItemIndex_Sybase;
          pgcOptions.ActivePageIndex := C_PageIndex_Sybase;
          edtSybServerName.Text := SybServerName;
          edtSybPort.Text := SybPort;
          cbbSybDataBase.Text := SybDatabaseName;
        end;
        dbtMySql:
        begin
          cbbDBType.ItemIndex := C_nDBTypeItemIndex_MySQL;
          pgcOptions.ActivePageIndex := C_PageIndex_MySQL;
          edtmslHost.Text := mslHost;
          edtmslDataBase.Text := mslDataBase;
          edtMysqlCharset.Text := mslCharset;
        end;
        dbtSqlite:
        begin
          cbbDBType.ItemIndex := C_nDBTypeItemIndex_SQLite;
          pgcOptions.ActivePageIndex := C_PageIndex_SQLite;
          edtSqliteDBFile.Text := sltDB;
        end;
        else
        begin
          cbbDBType.ItemIndex := 0;
          pgcOptions.ActivePageIndex := 0;
        end;
      end;
      DBEngineTypeToCbxIndex(DBEngineType);
      slst.Clear;
      ExtractStrings([','], [' '], PChar(UserName), slst);
//      fStrUtil.Split(UserName, ',', slst);
      while slst.Count < 1 do
        slst.Add('');
      cbbUser.Text := slst[0];
      edtPwd.Text := Password;
    end;
  finally
    slst.Free;
  end;
end;

procedure TF_DBOption.ChangeCurrnetUser;
begin
  cbbUser.Items.Clear;
  with GlobalParams do
  case pgcOptions.ActivePageIndex of
    C_PageIndex_Access:
    begin
      cbbUser.Text := DBconfig.AcsUser;
      edtPwd.Text := DBconfig.getEncodePwd(dbtAccess);
    end;
    C_PageIndex_Oracle:
    begin
      cbbUser.Text := DBconfig.OraUser;
      edtPwd.Text := DBconfig.getEncodePwd(dbtOracle);
    end;
    C_PageIndex_Sybase:
    begin
      cbbUser.Text := DBconfig.SybUser;
      edtPwd.Text := DBconfig.getEncodePwd(dbtSyBase);
    end;
    C_PageIndex_MySQL:
    begin
      cbbUser.Text := DBconfig.mslUser;
      edtPwd.Text := DBconfig.getEncodePwd(dbtMySQL);
    end;
  end;
end;    

function TF_DBOption.Apply: Boolean;
begin
  Result := False;
  if cbbDBType.ItemIndex = 0 then
  begin
    MessageBox(Handle, '请选择一个数据库类型！', '提示',
      MB_OK + MB_ICONINFORMATION);
    Exit;
  end;

  Result := SaveSettings;
end;

procedure TF_DBOption.btnApplyClick(Sender: TObject);
begin
  Apply;
end;

procedure TF_DBOption.cbbEngineTypeChange(Sender: TObject);
var
  bLocalTnsOnly: Boolean;
begin
  bLocalTnsOnly := (cbbEngineType.ItemIndex in [
    C_nDBEngineTypeItemIndex_Bde, C_nDBEngineTypeItemIndex_DBExpress]);
  chkIsLocalTNS.Checked := bLocalTnsOnly;
  SetCtrlEnable(chkIsLocalTNS, not bLocalTnsOnly);
end;

procedure TF_DBOption.SetCtrlEnable(ctrl: TWinControl; bEnable: Boolean);
begin
  ctrl.Enabled := bEnable;
  if bEnable then
  begin
    if ctrl is TEdit then
      (ctrl as TEdit).Color := clWindow;
    if ctrl is TCheckBox then
      (ctrl as TCheckBox).Color := clWindow;  
    if ctrl is TComboBox then
      (ctrl as TComboBox).Color := clWindow;
  end
  else
  begin
    if ctrl is TEdit then
      (ctrl as TEdit).Color := clBtnFace;
    if ctrl is TCheckBox then
      (ctrl as TCheckBox).Color := clBtnFace;
    if ctrl is TComboBox then
      (ctrl as TComboBox).Color := clBtnFace;
  end;
end;

procedure TF_DBOption.btnAddToConfigClick(Sender: TObject);
//var
//  cfg: TDBConfig;
begin
  inherited;
  if IDYES = MessageBox(Handle, '是否要将配置写入的配置文件中以便以后读取？',
     '确认', MB_YESNO + MB_ICONINFORMATION)  then
  begin
//    cfg := TDBConfig.Create;
//    cfg.DBShownName
  end;
end;

procedure TF_DBOption.RefreshUI;
//var
//  bLocalTnsOnly: Boolean;
begin
//  bLocalTnsOnly := (cbbEngineType.ItemIndex in [
//    C_nDBEngineTypeItemIndex_Bde, C_nDBEngineTypeItemIndex_DBExpress]);
//  chkIsLocalTNS.Checked := bLocalTnsOnly;
//  SetCtrlEnable(chkIsLocalTNS, not bLocalTnsOnly);
end;

procedure TF_DBOption.chkIsLocalTNSClick(Sender: TObject);
begin
  inherited;
  SetCtrlEnable(edtOraIP, not chkIsLocalTNS.Checked);
  SetCtrlEnable(cbbOraPro, not chkIsLocalTNS.Checked);
  SetCtrlEnable(edtOraPort, not chkIsLocalTNS.Checked);
end;

procedure TF_DBOption.DBEngineTypeToCbxIndex(dbet: TDBEngineType);
begin  
  case dbet of
    dbetAuto: cbbEngineType.ItemIndex := C_nDBEngineTypeItemIndex_Auto;
    dbetADO: cbbEngineType.ItemIndex := C_nDBEngineTypeItemIndex_Ado;
    dbetBDE: cbbEngineType.ItemIndex := C_nDBEngineTypeItemIndex_Bde;
    dbetDBExpress: cbbEngineType.ItemIndex := C_nDBEngineTypeItemIndex_DBExpress;
    dbetSqlite: cbbEngineType.ItemIndex := C_nDBEngineTypeItemIndex_SQLite;
    else
  end;
end;

procedure TF_DBOption.CbxIndexToDBEngineType(var dbet: TDBEngineType);
begin
  case cbbEngineType.ItemIndex of
    C_nDBEngineTypeItemIndex_Auto: dbet := dbetAuto;
    C_nDBEngineTypeItemIndex_Ado: dbet := dbetADO;
    C_nDBEngineTypeItemIndex_Bde: dbet := dbetBDE;
    C_nDBEngineTypeItemIndex_DBExpress: dbet := dbetDBExpress;
    C_nDBEngineTypeItemIndex_SQLite : dbet := dbetSqlite;
    else
  end;
end;

procedure TF_DBOption.CbxIndexToDBType(var dbt: TDBType);
begin
//  case cbbSybaseType.ItemIndex of
//    C_nSybDBTypeItemIndex_ASE: dbt := dbtSybase;
//    C_nSybDBTypeItemIndex_ASA: dbt := dbtSybaseASA;
//    else
//  end;
end;

procedure TF_DBOption.DBTypeToCbxIndex(dbt: TDBType);
begin
//  case dbt of
//    dbtSybase: cbbSybaseType.ItemIndex := C_nSybDBTypeItemIndex_ASE;
//    dbtSybaseASA: cbbSybaseType.ItemIndex := C_nSybDBTypeItemIndex_ASA;
//    else
//  end;
end;

procedure TF_DBOption.edtRecentDBCountKeyPress(Sender: TObject;
  var Key: Char);
begin
  inherited;
  if not CharInSet(Key, ['0'..'9']) then
//  if not (Key in ['0'..'9']) then
    Key := #0;
end;

procedure TF_DBOption.btnTestClick(Sender: TObject);
begin
  inherited;
  //
end;   

procedure TF_DBOption.LoadOdbc;
begin
  TDBUtil.LoadODBCNames(edtmslHost.Items);
end;

procedure TF_DBOption.LoadTnsNames;
begin
  cbbOraSID.Items.Clear;
  try
    TDBUtil.LoadOracleTnsNames(cbbOraSID.Items);
  except
    // no error for loading
  end;
end;

procedure TF_DBOption.cbbOraSIDClick(Sender: TObject);
begin
  inherited;
  chkIsLocalTNS.Checked := True;
end;

procedure TF_DBOption.edtmslDataBaseDropDown(Sender: TObject);
begin
  edtmslDataBase.Items.Clear;
  if Trim(edtmslHost.Text) <> '' then
  begin  
    try
      TDBUtil.LoadMysqlDatabaseNames(edtmslHost.Text, edtmslDataBase.Items);
    except
      // init database names  no error report
    end;  
  end;
end;

end.
