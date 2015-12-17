unit U_RegexUtil;
{
 @author  fengyq
 @date 2014-09-27
}

{$WARNINGS OFF}
interface
uses
  Contnrs, PerlRegEx, SysUtils, XMLDoc, Generics.Collections, Generics.Defaults;

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
  public
    class function ConvertRegex(sPattern: string):string;
{-------------------------------------------------------------------------------
  过程名:    MatchFirst  只匹配第一次Match，返回指定index的结果
  作者:      fengyq
  日期:      2014.09.27
  参数:      sPattern: String;
             index: Integer; group(subException) index
             sContent: String;
             matchResult: TMatchResult
  返回值:    匹配成功返回TMatchResult对象（需外部Free），否则返回nil
-------------------------------------------------------------------------------}
    class function MatchFirst(sPattern: String; index: Integer; sContent: String): TMatchResult;
    class function MatchFirstStr(sPattern: String; index: Integer; sContent: String): String;

{-------------------------------------------------------------------------------
  过程名:    MatchAll 匹配所有Match，每次匹配取指定index的结果
  作者:      fengyq
  日期:      2014.09.27
  参数:      sPattern, sContent: String; resultList: TObjectList
  返回值:    至少一个匹配成功返回True，否则False
-------------------------------------------------------------------------------}
    class function MatchAll(sPattern:string; index: Integer; sContent: String;
      resultList: TObjectList<TMatchResult>): Boolean;

  end;

implementation

var
  FRegexCache: TDictionary<String,TPerlRegEx>;

function GetRegexObject(sPattern: String): TPerlRegEx;
var
  regex: TPerlRegEx;
begin
  if FRegexCache.ContainsKey(sPattern) then begin
    regex := FRegexCache.Items[sPattern];
  end else begin
    regex := TPerlRegEx.Create(nil);
    regex.RegEx := sPattern;
    regex.Options := [preMultiLine, preExtended];
    regex.Compile;
    FRegexCache.Add(sPattern, regex);
  end;
  Result := regex;
end;

{ TRegexUtil }

class function TRegexUtil.ConvertRegex(sPattern: string): string;
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

class function TRegexUtil.MatchFirst(sPattern: String; index: Integer;
  sContent: String): TMatchResult;
var
  regex: TPerlRegEx;
  matchResult, mrSub: TMatchResult;
  i: Integer;
begin
  Result := nil;
  regex := GetRegexObject(sPattern);
  regex.Subject := sContent;
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
end;

class function TRegexUtil.MatchFirstStr(sPattern: String; index: Integer;
  sContent: String): String;
var
  mr: TMatchResult;
begin
  Result := '';
  mr := MatchFirst(sPattern, index, sContent);
  if mr <> nil then
  try
    Result := mr.MatchStr;
  finally
    mr.Free;
  end;
end;

class function TRegexUtil.MatchAll(sPattern: String; index: Integer; sContent: String;
  resultList: TObjectList<TMatchResult>): Boolean;
var
  regex: TPerlRegEx;
  matchResult, mrSub: TMatchResult;
  i: Integer;
begin
  regex := GetRegexObject(sPattern);
  regex.Subject := sContent;
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
  FRegexCache := TObjectDictionary<String,TPerlRegEx>.Create([doOwnsValues]);

finalization
  if Assigned(FRegexCache) then
    FreeAndNil(FRegexCache);

end.
