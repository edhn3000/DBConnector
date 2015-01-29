{*
 * U_PagingCounter
 * 分页计数器
 * @date 2010-08-15
 * @author: fengyq
 * <p>Description:  只用于分页计数，保持与分页业务处理无关
 * 对TotalCount NumPerPage进行设置后即可使用其他功能
 *}
unit U_PagingCounter;

interface

type
  TPagingCounter = class
  private     
    FTotalCount: Integer;    // 总数
    FCurPage: Integer;       // 当前页号
    FNumPerPage: Integer;    // 每页数目
    FLoadCounter: Integer;   // 已加载行的计数器
    FEnabled: Boolean;       // 使分页计算是否生效

    // 不会超过页号的上下限
    procedure SetCurPage(value: Integer);
    // 不允许设置0以下的每页行数
    procedure SetNumPerPage(value: Integer);
    // 返回总页数
    function GetPageCount: Integer;
  public
    // 需要在使用其他功能钳设置
    property TotalCount: Integer read FTotalCount write FTotalCount;
    property NumPerPage: Integer read FNumPerPage write SetNumPerPage; 
    property Enabled: Boolean read FEnabled write FEnabled;

    // 以下为支持的功能
    property CurPage: Integer read FCurPage write SetCurPage;
    property PageCount: Integer read GetPageCount;

    constructor Create;
    destructor Destroy; override;
    procedure Reset;

    // 获取每页的起始行号
    function GetStartNum: Integer;
    // 获取每页的中止
    function GetEndNum: Integer;

    // 检查当前页是否首页
    function CurPageIsFirstPage: Boolean;
    // 检查当前页是否尾页
    function CurPageIsLastPage: Boolean;
                                      
    procedure ResetLoadCounter;
    // 内置加载数据的计数器，+1
    procedure IncLoadCounter;
    // 检查加载数据的计数器是否已经到达了每页的记录数上限
    function ReachedPageCount: Boolean;
  end;

implementation

uses
  Math;

{ TPagingCounter }

constructor TPagingCounter.Create;
begin                
  NumPerPage := 0;   // 0代表不分页
  FEnabled := True;  // false时相当于numperpage为0，用于在初始化后临时设置无效
  TotalCount := 0;
  Reset;
end;

procedure TPagingCounter.Reset;
begin
  FCurPage := 1;
//  NumPerPage := 0;
//  TotalCount := 0;
  ResetLoadCounter;
end; 

function TPagingCounter.CurPageIsFirstPage: Boolean;
begin           
  Result := FCurPage = 1;
end;

function TPagingCounter.CurPageIsLastPage: Boolean;
begin        
  Result := FCurPage = GetPageCount;
end;

destructor TPagingCounter.Destroy;
begin

  inherited;
end;  

function TPagingCounter.GetStartNum: Integer;
begin
  if not FEnabled then  // 失效时，始终从第一个数据开始读取
    Result := 1
  else
    Result := (FCurPage - 1) * NumPerPage + 1;
end;

function TPagingCounter.GetEndNum: Integer;
begin
  Result := FCurPage * NumPerPage;
  if Result > TotalCount then
    Result := TotalCount;
end;

function TPagingCounter.GetPageCount: Integer;
begin
  if (NumPerPage = 0) or (not FEnabled) then
  begin
    if TotalCount>0 then
      Result := 1
    else
      Result := 0;
  end
  else
    Result := Ceil(TotalCount / NumPerPage);
end;

procedure TPagingCounter.SetCurPage(value: Integer);
begin
  if value < 1 then
    FCurPage := 1
  else if value > GetPageCount then
    FCurPage := GetPageCount
  else
    FCurPage := value;
end;

procedure TPagingCounter.SetNumPerPage(value: Integer);
begin
  if value < 0 then
    FNumPerPage := 0
  else
    FNumPerPage := value;
end;

procedure TPagingCounter.IncLoadCounter;
begin
  FLoadCounter := FLoadCounter + 1;
end;

function TPagingCounter.ReachedPageCount: Boolean;
begin
  if not FEnabled then  // 失效时，此方法始终返回false
    Result := False
  else
    Result := FLoadCounter >= FNumPerPage;
end;

procedure TPagingCounter.ResetLoadCounter;
begin
  FLoadCounter := 0;
end;

end.
