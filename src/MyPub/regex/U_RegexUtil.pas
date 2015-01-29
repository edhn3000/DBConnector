unit U_RegexUtil;
{
 @author  fengyq
 @date 2014-09-27
}

interface
uses
  Contnrs, PerlRegEx, SysUtils, XMLDoc, Generics.Collections;

type
  {TMatchResult}
  TMatchResult = class(TObject)
  private
    FMatchStr: string;
    FStart: Integer;
    FLen: Integer;

    FGroups: TObjectList<TMatchResult>;
  public
    property MatchStr: String read FMatchStr;
    property Start: Integer read FStart;
    property Len: Integer read FLen;
    property Groups: TObjectList<TMatchResult> read FGroups;

    constructor Create;
    destructor Destroy(); override;
  end;

  {TRegexUtil}
  TRegexUtil = class(TObject)
  private

  public
    function ConvertRegex(sPattern: string):string;
{-------------------------------------------------------------------------------
  ������:    MatchFirst  ֻƥ���һ��Match������ָ��index�Ľ��
  ����:      fengyq
  ����:      2014.09.27
  ����:      sPattern: String;
             index: Integer; group(subException) index
             sContent: String;
             matchResult: TMatchResult
  ����ֵ:    ƥ��ɹ�����TMatchResult�������ⲿFree�������򷵻�nil
-------------------------------------------------------------------------------}
    function MatchFirst(sPattern: String; index: Integer; sContent: String): TMatchResult;

{-------------------------------------------------------------------------------
  ������:    MatchAll ƥ������Match��ÿ��ƥ��ȡָ��index�Ľ��
  ����:      fengyq
  ����:      2014.09.27
  ����:      sPattern, sContent: String; resultList: TObjectList
  ����ֵ:    ����һ��ƥ��ɹ�����True������False
-------------------------------------------------------------------------------}
    function MatchAll(sPattern:string; index: Integer; sContent: String;
      resultList: TObjectList<TMatchResult>): Boolean;
  end;

  function RegexUtil: TRegexUtil;

implementation

var
  mRegexUtil: TRegexUtil;

function RegexUtil: TRegexUtil;
begin
  if mRegexUtil = nil then
  begin
    mRegexUtil := TRegexUtil.Create;
  end;
  Result := mRegexUtil;
end;


{ TRegexUtil }

function TRegexUtil.ConvertRegex(sPattern: string): string;
var
  mrList: TObjectList<TMatchResult>;
  mr: TMatchResult;
  i: Integer;
  sReplace: String;
begin
  mrList := TObjectList<TMatchResult>.Create;
  MatchAll('\\u[\dA-Za-z]{4}', 0, sPattern, mrList);
  try
    for i := 0 to mrList.Count - 1 do
    begin
      mr := TMatchResult(mrList[i]);
      sReplace := '\x{' + Copy(mr.MatchStr, 3, MaxInt) + '}';
      sPattern := StringReplace(sPattern, mr.MatchStr, sReplace,[rfReplaceAll]);
    end;
    Result := sPattern;
  finally
    mrList.Free;
  end;
end;

function TRegexUtil.MatchFirst(sPattern: String; index: Integer;
  sContent: String): TMatchResult;
var
  regex: TPerlRegEx;
  matchResult, mrSub: TMatchResult;
  i: Integer;
begin
  Result := nil;
  regex := TPerlRegEx.Create(nil);
  try
    regex.Subject := sContent;
    regex.RegEx := sPattern;
    regex.Options := [preMultiLine, preExtended];
    if regex.Match then
    begin
      if regex.SubExpressionCount >= index then
      begin
        matchResult := TMatchResult.Create;
        matchResult.FMatchStr := regex.SubExpressions[index];
        matchResult.FStart := regex.SubExpressionOffsets[index];
        matchResult.FLen := regex.SubExpressionLengths[index];
        Result := matchResult;
        for i := 0 to regex.SubExpressionCount do begin
          mrSub := TMatchResult.Create;
          mrSub.FMatchStr := regex.SubExpressions[i];
          mrSub.FStart := regex.SubExpressionOffsets[i];
          mrSub.FLen := regex.SubExpressionLengths[i];
          matchResult.Groups.Add(mrSub);
        end;
      end;
    end;
  finally
    regex.Free;
  end;
end;

function TRegexUtil.MatchAll(sPattern: String; index: Integer; sContent: String;
  resultList: TObjectList<TMatchResult>): Boolean;
var
  regex: TPerlRegEx;
  matchResult, mrSub: TMatchResult;
  i: Integer;
begin
  regex := TPerlRegEx.Create(nil);
  try
    regex.Subject := sContent;
    regex.RegEx := sPattern;
    regex.Options := [preMultiLine, preExtended];
    if regex.Match then
    begin
      repeat
        if regex.SubExpressionCount >= index then
        begin
          matchResult := TMatchResult.Create;
          matchResult.FMatchStr := regex.SubExpressions[index];
          matchResult.FStart := regex.SubExpressionOffsets[index];
          matchResult.FLen := regex.SubExpressionLengths[index];
          resultList.Add(matchResult);
          for i := 0 to regex.SubExpressionCount do begin
            mrSub := TMatchResult.Create;
            mrSub.FMatchStr := regex.SubExpressions[i];
            mrSub.FStart := regex.SubExpressionOffsets[i];
            mrSub.FLen := regex.SubExpressionLengths[i];
            matchResult.Groups.Add(mrSub);
          end;
        end;
      until not regex.MatchAgain;
    end;
    Result := resultList.Count > 0;
  finally
    regex.Free;
  end;
end;

{ TMatchResult }

constructor TMatchResult.Create;
begin
  FGroups := TObjectList<TMatchResult>.Create;
end;

destructor TMatchResult.Destroy;
begin
  FGroups.Free;
  inherited;
end;

initialization

finalization
  if Assigned(mRegexUtil) then
    FreeAndNil(mRegexUtil);

end.