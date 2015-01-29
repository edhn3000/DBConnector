{**
 * <p>Title: U_DBConfig  </p>
 * <p>Copyright: Copyright (c) </p>
 * @author fengyq
 * @version 1.0           
 * <p>Description: 数据库配置单元 </p>
 * <p>Description: 该单元包含配置方面的管理，可将配置保存到文件；
      另外TDBConfig类旨在将DBConnector简化连接数据库参数格式。
     </p>
 *}
unit U_DBConfig;

interface

uses
  ADODB, DB, Classes, Contnrs, XMLDoc, XMLIntf, Variants, SysUtils,
  Forms, U_DBEngineInterface;

type
{ TDBConfig 数据库相关参数信息 }
  TDBConfig = class
  private
    FDBShownName: string;
    FsUserName: string;
    FsPassword: string;
    FDataSource: string;   // access和oracle 都可使用  DataSource不为空且oracle专有参数有一个为空则表示使用DataSource属性
    function GetUserName: string;
    procedure SetUserName(value: String);
    function GetPassword: string;
    procedure SetPassword(value: String);
  public
    DBType: TDBType;
    DBEngineType: TDBEngineType;
    // for access
    AcsDataSource: string;
    AcsSecuredDB: string;
    AcsUser: string;
    AcsPwd: string;
    LastDir: string;

    // for oracle
    SID: string;
    Host: string;
    Port: string;
    Protocol: string;
    OraUser: string;
    OraPwd: string;
    IsLocalTns: Boolean;

    // Sybase数据库
    SybServerName: string;
    SybDatabaseName: string;
    SybPort: string;
    SybUser: string;
    SybPwd: string;

    // MySQL
    mslHost: string;
    mslDataBase: string;
    mslCharSet: string;
    mslUser:string;
    mslPwd: string;

    // SQLite
    sltDB: String;

    IsFixDB: Boolean;
    IsPasswordEncoded: Boolean;

    property UserName: string read GetUserName write SetUserName;
    property Password: string read GetPassword write SetPassword;
    property DBShownName: string read FDBShownName write FDBShownName;
    property DataSource: string read FDataSource write FDataSource;
  public
    constructor Create;
    destructor Destroy;override;
    procedure Assign(source: TDBConfig);

    function CheckIntegrity(var ErrorMsg: string):Boolean;
    function BuildConnectionString: string;
    function IsSameDB(Adbcfg: TDBConfig):Boolean;
    procedure ParseDataSource(dbt:TDBType; sDataSource: string);
//    procedure GetConnectionDataSource;
    function GetSeparatedDataSource: string;
    
    function GetCurrUserName: string;
    function getEncodePwd(ADBType: TDBType): string;
    procedure setEncodePwd(ADBType: TDBType; sPwd: string);
  end;

{ TDBConfigList 多个数据库的信息}
  TDBConfigList = class(TObjectList)
  private
    FRecentDBCount: Integer;

  protected
    function GetItem(Index: Integer): TDBConfig;
    procedure SetItem(Index: Integer; const Value: TDBConfig);

  public
    property Items[Index: Integer]: TDBConfig read GetItem write SetItem;
    property RecentDBCount: Integer read FRecentDBCount write FRecentDBCount;
  public
    function FindDBConfig(DBShownName: string): TDBConfig;overload;
    function FindDBConfig(sDataSource, sUserName: string): TDBConfig;overload;
    function FixDBCount: Integer;

    procedure LoadFromFile(sFileName: string);
    procedure SaveToFile(sFileName: string);
  end;

implementation

uses
  U_Base64Codec;

const
   C_sDBConfig_Node_Document = 'dbnameSet';  
   C_sDBConfig_Node_DBS      = 'dbs';
   C_sDBConfig_Node_DB       = 'db';
   C_sDBConfig_Node_DBShownName  = 'dbshownname';
   C_sDBConfig_Node_DBType       = 'dbtype';
   C_sDBConfig_Node_DBEngineType = 'dbenginetype';
   C_sDBConfig_Node_DataSource   = 'datasource';
   C_sDBConfig_Node_UserName     = 'username';
   C_sDBConfig_Node_Password     = 'password';

   C_sDBConfig_NodeParam_IsLocalTns = 'IsLocalTns';
   C_sDBConfig_NodeAttr_Fix = 'fix';

   C_sDBConfig_NodeAttr_UseDBMaxCount = 'usedbmaxcount';
   C_nDBConfig_UsedDB_MaxCount = 10;

function GetChildNodeAttr(const Anode: IXMLNode; sChildNode, sAttr: string): string;
var
  node: IXMLNode;
begin
  node := Anode.ChildNodes.FindNode(sChildNode);
  if (node <> nil) and (node.HasChildNodes) and
    (node.ChildNodes[0].NodeType = ntText) then
  begin
    if node.Attributes[sAttr] <> Null then
      Result := node.Attributes[sAttr]
  end
  else
    Result := '';
end;

procedure SetChildNodeText(const Anode: IXMLNode; sChildNode: string; value: string);
var
  node: IXMLNode;
begin
  node := Anode.ChildNodes.FindNode(sChildNode);
  if node = nil then
    node := Anode.AddChild(sChildNode);
  node.Text := value;
end;    

procedure SetChildNodeAttr(const Anode: IXMLNode; sChildNode, sAttr: string;
  value: string);
var
  node: IXMLNode;
begin
  node := Anode.ChildNodes.FindNode(sChildNode);
  if node = nil then
    node := Anode.AddChild(sChildNode);
  node.Attributes[sAttr] := value;
end;


{ TDBConfig }

function TDBConfig.BuildConnectionString: string;
var
 sAcsSec: string;
 sOraTNS: string;
 slst: TStrings;
begin
  case DBType of
    dbtAccess:
      begin
        sAcsSec := AcsSecuredDB;
        if '' <> sAcsSec then
          sAcsSec := Format(SYSDB_ACCESS, [sAcsSec]);
        Result := Format(CONNSTR_ACCESS, [FDataSource, sAcsSec,
                                                UserName, Password]);
      end;
    dbtAccess2007:
      begin
        sAcsSec := AcsSecuredDB;
        if '' <> sAcsSec then
          sAcsSec := Format(SYSDB_ACCESS2007, [sAcsSec]);
        Result := Format(CONNSTR_ACCESS2007, [FDataSource, sAcsSec,
                                                    UserName, Password]);
      end;
    dbtOracle:
      begin
        if ('' <> SID ) and ('' <> Host) and
           ('' <> Port) and ('' <> Protocol) then // 使用oracle专有参数
        begin
          sOraTNS := Format(TNSNAME_ORACLE, [SID, Host, Port, Protocol]);
        end
        else if '' <> FDataSource then            // 使用DataSource
        begin
          if Pos(C_sSEPARATOR_DATASOURCE, FDataSource) = 0 then    // 使用本地数据库别名时传入的参数不带逗号
            sOraTNS := FDataSource
          else                                                 // 远程服务名
          begin
            // 构造Oracle的DataSource
            slst := TStringList.Create;
            try
              ExtractStrings([C_sSEPARATOR_DATASOURCE], [], PChar(FDataSource), slst);
              if slst.Count < 4 then
                raise Exception.Create(ERROR_ORACLE_DATASOURCE)
              else
                sOraTNS := Format(TNSNAME_ORACLE, [slst.Strings[3], {SID}
                                                   slst.Strings[1], {Host}
                                                   slst.Strings[2], {端口}
                                                   slst.Strings[0]  {协议}]);
            finally
              slst.Free;
            end;
          end;
        end
        else
          raise Exception.Create('无法成功分析参数！');

        Result := Format(CONNSTR_ORACLE, [Password, UserName, sOraTNS]);
      end;
    else
      begin
        raise Exception.Create(ERROR_DBTYPE);
        Result := sOraTNS;
      end;
  end;
end; 

function TDBConfig.CheckIntegrity(var ErrorMsg: string): Boolean;
begin
  ErrorMsg := '';
  Result := False;
  case DBType of
    dbtAccess:
    begin
      if '' = FDataSource then
        ErrorMsg := 'Access数据源不能为空！'
      else
        Result := True;
    end;
    dbtOracle:
    begin
      if '' = SID then
        ErrorMsg := '数据库服务名不合法！'
      else if '' = Host then
        ErrorMsg := '主机地址不能为空'
      else if '' = Port then
        ErrorMsg := '端口不能为空！'
      else if '' = Protocol then
        ErrorMsg := '协议名称不能为空！'
    end;
    else
      Result := True;
  end;
end;

constructor TDBConfig.Create;
begin
end;

destructor TDBConfig.Destroy;
begin
  inherited;
end;

function TDBConfig.IsSameDB(Adbcfg: TDBConfig): Boolean;
begin
  Result := False;
  case DBType of
    dbtAccess, dbtAccess2007:
    begin
      Result := SameText(AcsDataSource, Adbcfg.AcsDataSource) and
                SameText(AcsSecuredDB, Adbcfg.AcsSecuredDB);
    end;
    dbtOracle:
    begin
      Result := SameText(SID, Adbcfg.SID) and
                SameText(Host, Adbcfg.Host) and
                SameText(Port, Adbcfg.Port) and
                SameText(Protocol, Adbcfg.Protocol);
    end;
    dbtSyBase:
    begin
      Result := SameText(FDataSource, Adbcfg.FDataSource);
    end;
  end;
end;

procedure TDBConfig.ParseDataSource(dbt: TDBType; sDataSource: string);
var
  slst: TStrings;
begin
  FDataSource := sDataSource;
  slst := TStringList.Create;
  ExtractStrings([','], [], PChar(sDataSource), slst);
  try
    case dbt of
      dbtOracle:
      begin
        while slst.Count < 4 do
          slst.Add('');
        if slst[2] = '' then slst[2] := '1521';
        if slst[3] = '' then slst[3] := 'tcp';
        SID := slst[0];
        Host := slst[1];
        Port := slst[2];
        Protocol := slst[3];
      end;
      dbtAccess, dbtAccess2007:
      begin
        while slst.Count < 2
          do slst.Add('');
        AcsDataSource := slst[0];
        AcsSecuredDB := slst[1];
      end;
      dbtSybase:
      begin
        while slst.Count < 3 do
          slst.Add('');
        SybServerName := slst[0];
        SybPort := slst[1];
        SybDatabaseName := slst[2];
      end;
      dbtMySql:
      begin
        while slst.Count < 3 do
          slst.Add('');
        mslHost := slst[0];
        mslDataBase := slst[1];
        mslCharSet := slst[2];
      end;
      dbtSqlite:
      begin
        sltDB := slst[0];
      end;
      else
//        DataSource := sDataSource;
      end;
  finally
    slst.Free;
  end;
end;   

function TDBConfig.GetSeparatedDataSource: string;
begin
  case DBType of
    dbtOracle:
    begin
      if Host <> '' then
        Result := BuildOracleDataSource(SID, Host, Port, Protocol, IsLocalTns)
      else
        Result := SID;
    end;
    dbtAccess, dbtAccess2007:
    begin
      Result := BuildAccessDataSource(AcsDataSource, AcsSecuredDB);
    end;
    dbtSybase, dbtSybaseASA:
    begin
      Result := BuildSybaseDataSource(SybServerName, SybPort, SybDatabaseName);
    end;  
    dbtMySql:
    begin
      Result := BuildMySqlDataSource(mslHost, mslDataBase, mslCharSet);
    end;
    dbtSqlite:
    begin
      Result := sltDB;
    end;
    else
      Result := FDataSource;
  end;
end;

procedure TDBConfig.Assign(source: TDBConfig);
begin
  Self.LastDir            := source.LastDir        ;

  Self.DBType             := source.DBType         ;
  Self.DBEngineType       := source.DBEngineType   ;
  Self.FDataSource         := source.FDataSource     ;

  Self.AcsDataSource      := source.AcsDataSource  ;
  Self.AcsSecuredDB       := source.AcsSecuredDB   ;
  Self.AcsUser            := source.AcsUser        ;
  Self.AcsPwd             := source.AcsPwd         ;

  Self.SID                := source.SID            ;
  Self.Host               := source.Host           ;
  Self.Port               := source.Port           ;
  Self.Protocol           := source.Protocol       ;
  Self.OraUser            := source.OraUser        ;
  Self.OraPwd             := source.OraPwd         ;
  Self.IsLocalTns         := source.IsLocalTns     ;

  Self.SybServerName      := source.SybServerName  ;
  Self.SybDatabaseName    := source.SybDatabaseName;
  Self.SybPort            := source.SybPort        ;
  Self.SybUser            := source.SybUser        ;
  Self.SybPwd             := source.SybPwd         ;

  Self.mslHost            := source.mslHost        ;
  Self.mslDataBase        := source.mslDataBase    ;
  Self.mslCharSet         := source.mslCharSet     ;
  Self.mslUser            := source.mslUser        ;
  Self.mslPwd             := source.mslPwd         ;

  Self.sltDB              := source.sltDB;

  Self.IsFixDB            := source.IsFixDB        ;
  Self.IsPasswordEncoded  := source.IsPasswordEncoded;
end;

//procedure TDBConfig.GetConnectionDataSource;
//var
//  sDBFile: String;
//  function GetValidFullPath(Str: string):string;
//  begin
//    if (Trim(Str) = '') or (Pos(':', Str) <> 0)
//       or (Pos('\\', Str) = 1) then
//      Result := Str
//    else
//      Result := ExtractFilePath(Application.ExeName)+Str;
//  end;
//begin
//  case DBType of
//  dbtAccess, dbtAccess2007: begin
//    sDBFile := GetValidFullPath(AcsDataSource);
//    FDataSource := BuildAccessDataSource(sDBFile, GetValidFullPath(AcsSecuredDB));
////    if FileExists(sDBFile) then begin
////      FileSetAttr(sDBFile, FileGetAttr(sDBFile) and (not SysUtils.faReadOnly));
////    end;
//    UserName := AcsUser;
//    Password := AcsPwd;
//  end;
//  dbtOracle: begin
//    FDataSource := BuildOracleDataSource(SID, Host, Port, Protocol, IsLocalTns or (Host = ''));
//    UserName := OraUser;
//    Password := OraPwd;
//  end;
//  dbtSyBase, dbtSybaseASA: begin
//    FDataSource := BuildSybaseDataSource(SybServerName, SybPort, SybDatabaseName);
//    UserName := SybUser;
//    Password := SybPwd;
//  end;
//  dbtMySQL: begin
//    FDataSource := BuildMySqlDataSource(mslHost, mslDataBase, mslCharset);
//    UserName := mslUser;
//    Password := mslPwd;
//  end;
//  dbtSqlite: begin
//    FDataSource := sltDB;
//    UserName := '';
//    Password := '';
//  end
//  else begin
//  end;
//  end;
//end;

function TDBConfig.GetCurrUserName: string;
begin
  case DBType of
  dbtAccess, dbtAccess2007:
    Result := AcsUser;
  dbtOracle:
    Result := OraUser;
  dbtSyBase:
    Result := SybUser;
  dbtMySQL:
    Result := mslUser;
  else
    Result := UserName;
  end;
end;

function TDBConfig.getEncodePwd(ADBType: TDBType): string;
begin
  case ADBType of
  dbtAccess, dbtAccess2007:
    Result := Base64Encode(AcsPwd);
  dbtOracle:
    Result := Base64Encode(OraPwd);
  dbtSyBase:
    Result := Base64Encode(SybPwd);
  dbtMySQL:
    Result := Base64Encode(mslPwd);
  else
    Result := '';
  end;
end;

procedure TDBConfig.setEncodePwd(ADBType: TDBType; sPwd: string);
begin   
  case ADBType of
  dbtAccess, dbtAccess2007:
     if ('' = Base64Decode(sPwd)) or
        (AcsPwd <> Base64Decode(sPwd)) then
       AcsPwd := sPwd;
  dbtOracle:
     if ('' = Base64Decode(sPwd)) or
        (OraPwd <> Base64Decode(sPwd)) then
       OraPwd := sPwd;
  dbtSyBase:
     if ('' = Base64Decode(sPwd)) or
        (SybPwd <> Base64Decode(sPwd)) then
       SybPwd := sPwd;
  dbtMySQL:   
     if ('' = Base64Decode(sPwd)) or
        (mslPwd <> Base64Decode(sPwd)) then
       mslPwd := sPwd;
  end;
end;

function TDBConfig.GetPassword: string;
begin
  case DBType of
  dbtAccess, dbtAccess2007:
    Result := AcsPwd;
  dbtOracle:
    Result := OraPwd;
  dbtSybase:
    Result := SybPwd;
  dbtMySql:
    Result := mslPwd;
  else
    Result := FsPassword;
  end;
end;

function TDBConfig.GetUserName: string;
begin
  case DBType of
  dbtAccess, dbtAccess2007:
    Result := AcsUser;
  dbtOracle:
    Result := OraUser;
  dbtSybase:
    Result := SybUser;
  dbtMySql:
    Result := mslUser;
  else
    Result := FsUserName;
  end;
end;

procedure TDBConfig.SetPassword(value: String);
begin
  case DBType of
  dbtAccess, dbtAccess2007:
    AcsPwd := value;
  dbtOracle:
    OraPwd := value;
  dbtSybase:
    SybPwd := value;
  dbtMySql:
    mslPwd := value;
  else
    FsPassword := value;
  end;
end;

procedure TDBConfig.SetUserName(value: String);
begin
  case DBType of
  dbtAccess, dbtAccess2007:
    AcsUser := value;
  dbtOracle:
    OraUser := value;
  dbtSybase:
    SybUser := value;
  dbtMySql:
    mslUser := value;
  else
    FsUserName := value;
  end;
end;

{ TDBConfigList }


function TDBConfigList.FindDBConfig(DBShownName: string): TDBConfig;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Count - 1 do
  begin
    if SameText(Items[i].DBShownName, DBShownName) then
    begin
      Result := Items[i];
      Break;
    end;
  end;
end;        

function TDBConfigList.FindDBConfig(sDataSource,
  sUserName: string): TDBConfig;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Count - 1 do
  begin
    if SameText(Items[i].FDataSource, sDataSource)
       and SameText(Items[i].UserName, sUserName) then
    begin
      Result := Items[i];
      Break;
    end;
  end; 
end;

function TDBConfigList.GetItem(Index: Integer): TDBConfig;
begin
  Result := inherited GetItem(Index) as TDBConfig;
end;

procedure TDBConfigList.SetItem(Index: Integer; const Value: TDBConfig);
begin
  inherited SetItem(Index, Value);
end;

procedure TDBConfigList.LoadFromFile(sFileName: string);
var
  node: IXmlNode;
  nodeDoc: IXMLNode;
  xml: IXMLDocument;
  dbcfg: TDBConfig;
  
  function GetChildNodeText(const Anode: IXMLNode; sChildNode: string): string;
  var
    node: IXMLNode;
  begin
    node := Anode.ChildNodes.FindNode(sChildNode);
    if (node <> nil) and (node.HasChildNodes) and
      (node.ChildNodes[0].NodeType = ntText) then
      Result := node.Text
    else
      Result := '';
  end;
begin
  if ExtractFileExt(sFileName) <> '.xml' then Exit;
  xml := NewXMLDocument();
  try
    xml.Encoding := 'GBK';
    xml.LoadFromFile(sFileName);
    nodeDoc := xml.DocumentElement;
    if nodeDoc.HasAttribute(C_sDBConfig_NodeAttr_UseDBMaxCount) then
      FRecentDBCount := StrToInt(nodeDoc.Attributes[
        C_sDBConfig_NodeAttr_UseDBMaxCount])
    else
      FRecentDBCount := C_nDBConfig_UsedDB_MaxCount;       
                                         
    node := xml.DocumentElement.ChildNodes.FindNode('dbs');
    if node <> nil then
      node := node.ChildNodes.First
    else
      Exit;  // no db config
//    node := xml.DocumentElement.ChildNodes.First;

    while node <> nil do
    begin
      dbcfg := TDBConfig.Create();
      dbcfg.DBShownName := GetChildNodeText(node,
        C_sDBConfig_Node_DBShownName);

      // 数据库类型
      dbcfg.DBType := StrToDBType(GetChildNodeText(node,
        C_sDBConfig_Node_DBType));
      // 引擎类型
      dbcfg.DBEngineType := StrToDBEngineType(GetChildNodeText(node,
        C_sDBConfig_Node_DBEngineType));
      // 数据源字符串
      dbcfg.FDataSource := GetChildNodeText(node, C_sDBConfig_Node_DataSource);
      dbcfg.ParseDataSource(dbcfg.DBType, dbcfg.FDataSource);
      dbcfg.UserName := GetChildNodeText(node,
        C_sDBConfig_Node_UserName);
      dbcfg.Password := GetChildNodeText(node,
        C_sDBConfig_Node_Password);
      // 默认加密方式，除非指定Encoded="false"
      dbcfg.IsPasswordEncoded := not SameText(GetChildNodeAttr(
        node,C_sDBConfig_Node_Password,'Encoded'), '0');

      dbcfg.IsFixDB := False;
      if node.HasAttribute('fix') then
        dbcfg.IsFixDB := SameText(node.Attributes['fix'], '1');

      Self.Add(dbcfg);
      node := node.NextSibling;
    end;
  finally
    xml := nil;
  end;
end;   

procedure TDBConfigList.SaveToFile(sFileName: string);  
var
  i: Integer;
  nodeDoc, nodeDBS, node: IXmlNode;
  xml: TXMLDocument;
  dbcfg: TDBConfig;
  nMaxCount: Integer;
  nFixDBCount: Integer;
  nNodeIndex: Integer;
  bLoadSucc: Boolean;
  function GetChildNodeText(const Anode: IXMLNode; sChildNode: string): string;
  var
    node: IXMLNode;
  begin
    node := Anode.ChildNodes.FindNode(sChildNode);
    if (node <> nil) and (node.HasChildNodes) and
      (node.ChildNodes[0].NodeType = ntText) then
      Result := node.Text
    else
      Result := '';
  end;
  // 比较数据源和用户名
  function FindDBNode(ndFirst:IXMLNode; sDataSource, sUser: string): IXMLNode;
  var
    nd: IXMLNode;
  begin
    Result := nil;
    nd := ndFirst;
    while nd <> nil do
    begin
      if SameText(GetChildNodeText(nd, C_sDBConfig_Node_DataSource), sDataSource)
         and SameText(GetChildNodeText(nd, C_sDBConfig_Node_UserName), sUser) then
      begin
        Result := nd;
        Break;
      end;
      nd := nd.NextSibling;    
    end;
  end;    
  function IndexDBNode(ndFirst:IXMLNode; sDataSource, sUser: string): Integer;
  var
    nd: IXMLNode;
    nIndex: Integer;
  begin
    Result := -1;
    nIndex := -1;
    nd := ndFirst;
    while nd <> nil do
    begin
      Inc(nIndex);
      if SameText(GetChildNodeText(nd, C_sDBConfig_Node_DataSource), sDataSource)
         and SameText(GetChildNodeText(nd, C_sDBConfig_Node_UserName), sUser) then
      begin
        Result := nIndex;
        Break;
      end;
      nd := nd.NextSibling;    
    end;
  end; 
begin
  xml := TXMLDocument.Create(Application);
  xml.Active := True;
  xml.Encoding := 'GBK';
  try
    bLoadSucc := False;
    try
      if FileExists(sFileName) then
      begin
        xml.LoadFromFile(sFileName);
        nodeDoc := xml.DocumentElement;
        nodeDBS := nodeDoc.ChildNodes.FindNode(C_sDBConfig_Node_DBS);
        node := nodeDBS.ChildNodes.First;
        bLoadSucc := True;
      end;
    except
      xml.Free;       
      xml := TXMLDocument.Create(Application);
      xml.Active := True;
      xml.Encoding := 'GBK';
      bLoadSucc := False;
    end;
    if not bLoadSucc then
    begin
      nodeDoc := xml.AddChild(C_sDBConfig_Node_Document);
      nodeDBS := nodeDoc.AddChild(C_sDBConfig_Node_DBS);
      node := nodeDBS.ChildNodes.First;
    end;
    nMaxCount := FRecentDBCount;
//    if nodeDoc.HasAttribute(C_sDBConfig_NodeAttr_UseDBMaxCount) then
//      nMaxCount := StrToIntDef(nodeDoc.Attributes[
//        C_sDBConfig_NodeAttr_UseDBMaxCount], nMaxCount);

    nodeDBS.ChildNodes.Clear;
    for i := 0 to Count - 1 do
    begin
      dbcfg := Items[i];
      if (dbcfg.FDataSource = '') then
        Continue;

      node := FindDBNode(nodeDBS.ChildNodes.First,
        dbcfg.FDataSource, dbcfg.UserName);
      if node = nil then
        node := nodeDBS.AddChild(C_sDBConfig_Node_DB);

      SetChildNodeText(node, C_sDBConfig_Node_DBShownName,
        dbcfg.DBShownName);  
      SetChildNodeText(node, C_sDBConfig_Node_DBType,
        DBTypeToStr(dbcfg.DBType, False));
      SetChildNodeText(node, C_sDBConfig_Node_DBEngineType,
        DBEngineTypeToStr(dbcfg.DBEngineType, False));
      SetChildNodeText(node, C_sDBConfig_Node_DataSource,
        dbcfg.FDataSource);
      SetChildNodeText(node, C_sDBConfig_Node_UserName, dbcfg.UserName);
      SetChildNodeText(node, C_sDBConfig_Node_Password, dbcfg.Password);
      if dbcfg.IsFixDB then
        node.Attributes[C_sDBConfig_NodeAttr_Fix] := '1'
      else
        node.Attributes[C_sDBConfig_NodeAttr_Fix] := '0';
    end;

    nFixDBCount := FixDBCount;
    while Count - nFixDBCount > nMaxCount do
    begin
      for i := Count - 1 downto 0 do
      begin
        if not Items[i].IsFixDB then
        begin
          nNodeIndex := IndexDBNode(nodeDBS.ChildNodes.First,
            Items[i].FDataSource, Items[i].UserName);
          if nNodeIndex <> -1 then
          begin
            nodeDBS.ChildNodes.Delete(nNodeIndex);
          end;  
          Delete(i);
          Break;
        end;  
      end;  
    end;  
    xml.SaveToFile(sFileName);
  finally
    xml.Free;
  end;
end;

function TDBConfigList.FixDBCount: Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
  begin
    if Items[i].IsFixDB then
      Inc(Result);
  end;  
end;

end.
