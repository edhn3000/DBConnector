unit U_Recent;

interface

uses
  U_DBUtil, U_DBEngine, Contnrs;

type

  { 最近打开 }
  TRecent = class
  private

  public
    DB: string;
    Secured: string;
    User: string;
    Pwd: string;
    OpenTime: TDateTime;
  public
    constructor Create(sDB, sSecured, sUser, sPwd: string;
      date: TDateTime);
    destructor Destroy; override;
  end;

  { 最近打开的列表 }
  TRecentList = class
  private
    FslobjRecents: TObjectList;

    function GetCount: Integer;
    function GetItem(index: Integer): TRecent;
  protected
    procedure FitSize;
  public
    constructor Create;
    destructor Destroy; override;

    function InitRecents(sDBFile: string): Boolean;
    function SaveRecents(sDBFile: string): Boolean;
    procedure AddItem(sDBFile, sSecuredFile, sUser, sPwd: string);
    procedure RemoveItem(sDBFile, sSecuredFile, sUser, sPwd: string);

    property Items[index: Integer]: TRecent read GetItem;
    property Count: Integer read GetCount;
  end;
           
const
  FILENAME_RECENTDB = 'DB_Recent.mdb';

implementation

uses
  SysUtils, Forms, U_Base64Codec;


{ TRecentList }

procedure TRecentList.AddItem(sDBFile, sSecuredFile, sUser, sPwd: string);
begin
  FslobjRecents.Insert(0, TRecent.Create(sDBFile, sSecuredFile, sUser, sPwd,
    Now));
  FitSize;
end;

procedure TRecentList.RemoveItem(sDBFile, sSecuredFile, sUser, sPwd: string);
var
  i: Integer;
begin
  for i := Count - 1 downto 0 do
  begin
    if (Items[i].DB = sDBFile) and (Items[i].Secured = sSecuredFile) and
      (Items[i].User = sUser) and (Items[i].Pwd = sPwd) then
      FslobjRecents.Delete(i);
  end;
end;

constructor TRecentList.Create;
begin
  FslobjRecents := TObjectList.Create;
end;

destructor TRecentList.Destroy;
begin
  FslobjRecents.Free;
  inherited;
end;

procedure TRecentList.FitSize;
var
  i: Integer;
begin
  if FslobjRecents.Count > 10 then
    for i := FslobjRecents.Count - 1 downto 10 do
    begin
      FslobjRecents.Delete(i);
    end;
end;

function TRecentList.GetCount: Integer;
begin
  Result := FslobjRecents.Count;
end;

function TRecentList.GetItem(index: Integer): TRecent;
begin
  Result := TRecent(FslobjRecents.Items[index]);
end;

function TRecentList.InitRecents(sDBFile: string): Boolean;
const
  C_sGetRecents = 'select * from T_Recent order by D_OpenTime desc';
begin
  Result := True;
  with TDBConnect.Create do
  begin
    try
      try
        if sDBFile = '' then
          sDBFile := ExtractFilePath(Application.ExeName) + FILENAME_RECENTDB;
        if not OpenDB(sDBFile, '', '', dbtAccess, dbetADO) then
          Exit;
        ExecQuery(C_sGetRecents);
        with Qry do
        begin
          while not Eof do
          begin
            FslobjRecents.Add(
              TRecent.Create(FieldByName('C_DBPath').AsString,
              FieldByName('C_SecPath').AsString,
              FieldByName('C_User').AsString,
              Base64Decode(FieldByName('C_Pwd').AsString),
              FieldByName('D_OpenTime').AsDateTime));
            Next;
          end;
        end;
        FitSize;
      except
        Result := False;
      end;
    finally
      Free;
    end;
  end;
end;

function TRecentList.SaveRecents(sDBFile: string): Boolean;
const
  C_sDelAll = 'delete from T_Recent;';
  C_sInsert = 'insert into T_Recent(C_DBPath,C_SecPath,C_User,C_Pwd,D_OpenTime) '
    +
    'values(''%s'',''%s'',''%s'',''%s'',datevalue("%s"));';
  {C_sInsert = 'insert into T_Recent(C_DBPath,C_SecPath,C_User,C_Pwd,D_OpenTime) ' +
      'values(:DBPATH,:SECPATH,:USER,:PWD,:OPENTIME);'; }
var
  i: Integer;
  r: TRecent;
  s: string;
begin
  Result := True;
  with TDBConnect.Create do
  begin
    try
      try
        if sDBFile = '' then
          sDBFile := ExtractFilePath(Application.ExeName) + FILENAME_RECENTDB;
        if not OpenDB(sDBFile, '', '', dbtAccess, dbetADO) then
          Exit;
        s := C_sDelAll;
        for i := 0 to FslobjRecents.Count - 1 do
        begin
          r := TRecent(FslobjRecents.Items[i]);
          s := s + Format(C_sInsert, [r.DB, r.Secured, r.User,
            Base64Encode(r.Pwd), FormatDateTime('yyyy-mm-dd', r.OpenTime)]);
              // hh:mm:ss
        end;
        ExecSqlText(s, ';');
      except
        Result := False;
      end;
    finally
      Free;
    end;
  end;

end;

{ TRecent }

constructor TRecent.Create(sDB, sSecured, sUser, sPwd: string;
  date: TDateTime);
begin
  DB := sDB;
  Secured := sSecured;
  User := sUser;
  Pwd := sPwd;
  OpenTime := date;
end;

destructor TRecent.Destroy;
begin

  inherited;
end;

end.
