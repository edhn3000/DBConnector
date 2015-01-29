{*
 * U_PagingCounter
 * ��ҳ������
 * @date 2010-08-15
 * @author: fengyq
 * <p>Description:  ֻ���ڷ�ҳ�������������ҳҵ�����޹�
 * ��TotalCount NumPerPage�������ú󼴿�ʹ����������
 *}
unit U_PagingCounter;

interface

type
  TPagingCounter = class
  private     
    FTotalCount: Integer;    // ����
    FCurPage: Integer;       // ��ǰҳ��
    FNumPerPage: Integer;    // ÿҳ��Ŀ
    FLoadCounter: Integer;   // �Ѽ����еļ�����
    FEnabled: Boolean;       // ʹ��ҳ�����Ƿ���Ч

    // ���ᳬ��ҳ�ŵ�������
    procedure SetCurPage(value: Integer);
    // ����������0���µ�ÿҳ����
    procedure SetNumPerPage(value: Integer);
    // ������ҳ��
    function GetPageCount: Integer;
  public
    // ��Ҫ��ʹ����������ǯ����
    property TotalCount: Integer read FTotalCount write FTotalCount;
    property NumPerPage: Integer read FNumPerPage write SetNumPerPage; 
    property Enabled: Boolean read FEnabled write FEnabled;

    // ����Ϊ֧�ֵĹ���
    property CurPage: Integer read FCurPage write SetCurPage;
    property PageCount: Integer read GetPageCount;

    constructor Create;
    destructor Destroy; override;
    procedure Reset;

    // ��ȡÿҳ����ʼ�к�
    function GetStartNum: Integer;
    // ��ȡÿҳ����ֹ
    function GetEndNum: Integer;

    // ��鵱ǰҳ�Ƿ���ҳ
    function CurPageIsFirstPage: Boolean;
    // ��鵱ǰҳ�Ƿ�βҳ
    function CurPageIsLastPage: Boolean;
                                      
    procedure ResetLoadCounter;
    // ���ü������ݵļ�������+1
    procedure IncLoadCounter;
    // ���������ݵļ������Ƿ��Ѿ�������ÿҳ�ļ�¼������
    function ReachedPageCount: Boolean;
  end;

implementation

uses
  Math;

{ TPagingCounter }

constructor TPagingCounter.Create;
begin                
  NumPerPage := 0;   // 0������ҳ
  FEnabled := True;  // falseʱ�൱��numperpageΪ0�������ڳ�ʼ������ʱ������Ч
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
  if not FEnabled then  // ʧЧʱ��ʼ�մӵ�һ�����ݿ�ʼ��ȡ
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
  if not FEnabled then  // ʧЧʱ���˷���ʼ�շ���false
    Result := False
  else
    Result := FLoadCounter >= FNumPerPage;
end;

procedure TPagingCounter.ResetLoadCounter;
begin
  FLoadCounter := 0;
end;

end.
