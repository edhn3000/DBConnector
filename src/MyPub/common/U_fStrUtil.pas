{**
 * <p>Title: fStrUtil </p>
 * <p>Copyright: Copyright (c) 2006/10/14</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: 字符串处理单元 </p>
 *}
unit U_fStrUtil;

interface

uses
  Classes, U_RegexUtil, Generics.Collections;

type
  TfStrUtil = class
  private

  public
    function IsNumber(S: string):Boolean;
    function IsInteger(S: string):Boolean;

    // 对数字使用数值大小比较，对字符串使用ComareText；s1大于s2返回1， s1小于s2返回-1，相等返回0；
    function CompareNumberText(s1, s2: string): Integer;

    function CombineListToText(List: TStrings; sDelimiter: string;
      bQuote: Boolean): string;
    function Split(sText,sSeparator: String; ssParts: TStrings):TStrings;
    // 拆分字符串的时候越过指定类型的括号中的内容，用于字符串嵌套的情况
    function SplitEscapeBracket(sText,sSeparator,sBracketLeft, sBracketRight: String;
      ssParts: TStrings):TStrings;   
    function SplitEscapeQuote(sText,sSeparator,sQuote: String;sEscapeChar: string;
      ssParts: TStrings):TStrings;
    function RebuildDelimitatedText(sText, sOldDelimiter, sNewDelimiter: string;
      bQuote: Boolean): string;
    procedure StrToKeyValue(AStr, sSeparator:string; var Key, Value: string);

    // 在字符串中查找子串时越过指定类型的括号中的内容，用于字符串嵌套的情况                                   
    function PosEscapeBracket(sSub, S: String; sBracketLeft, sBracketRight: string;
      nFrom: Integer = 1; ignoreCase: Boolean = True): Integer;                                    
    function PosEscapeQuote(sSub, S: String; sQuote: string;sEscapeChar: string;
      nFrom: Integer = 1; ignoreCase: Boolean = True): Integer;
    {*
     * 可传入起始位置的Pos函数
     *}
    function PosFrom(sSub, S: String; nFrom: Integer = 1;
      ignoreCase: Boolean = True): Integer;
    function PosBack(sSub, S: String; nFrom: Integer = 1;
      ignoreCase: Boolean = True): Integer;

    // 定位到出现子串之前的位置
    function PosUntilNot(sSub, S: String; nFrom: Integer = 1;
      ignoreCase: Boolean = True): Integer;
    {*
     * 查找指定的字符串并可指定到第几个匹配位置
     *}
    function PosMatchIndex(sSub, sFull: String; nMatchIndex: Integer =1;
      nFrom: Integer = 1; ignoreCase: Boolean = True): Integer;
    function PosMatchCount(sSub, S: String; nFrom: Integer = 1;
         ignoreCase: Boolean = True): Integer;
         
    // 从后向前查找，没找到返回0
    function PosLast(sSub, S: String; nFrom: Integer = 0;
        ignoreCase: Boolean = True): Integer;
    function PosEndsText(sSub, S: string; var nPos:Integer;
      ignoreCase: Boolean = True): Boolean;
    function ParseMatchedBracket(nLeftBracketIndex:Integer; Astr: string;
        sBracketLeft: string = '('; sBracketRight: string = ')'): Integer; 
    function ParseMatchedQuote(nLeftQuoteIndex:Integer; Astr: string;
        sQuote: string = ''''; sEscapeChar: string = '\'): Integer;

    {* 查找数组中被包含在字符串并在最前的位置
     * @param S            母串
     * @param Subs         子串数组
     * @param nMatchIndex  数组中匹配的元素的脚标
     * @return             最靠前的位置，都没找到返回0
    }
    function PosArrayFrom(S: string; Subs: array of string;
      var nMatchIndex: Integer; nFrom: Integer = 1;
      ignoreCase: Boolean = True):Integer;overload; 
    function PosArrayFrom(S: string; Subs: TStrings;
      var nMatchIndex: Integer; nFrom: Integer = 1;
      ignoreCase: Boolean = True):Integer;overload;
    function PosArrayFromEscapeBracket(S: string; Subs: array of string;
      var nMatchIndex: Integer; sBracketLeft, sBracketRight: string;
      nFrom: Integer = 1;ignoreCase: Boolean = True):Integer;
    function PosArrayFromEscapeQuote(S: string; Subs: array of string;
      var nMatchIndex: Integer; sQuote: string; sEscapeChar: string;
      nFrom: Integer = 1;ignoreCase: Boolean = True):Integer;
      
    function SubStringBetween(S, sBegin, sEnd: string;
      ignoreCase: Boolean = True): string;
    function WildcardCompareText(const sSub: string; const S: string;
      ignoreCase: Boolean = True): Integer;

    // 判断一个位置是否在引号中
    function IsPosInQuotedStr(sMain: string; nSubPos: Integer;
      sQuoteStart, sQuoteEnd: string): Boolean;
      
    // 比较，完全相同返回0，否则返回有差别的index
    function CompareTextAtIndex(s1, s2: string): Integer;

    function MatchRegxSubStr(sSub, S: string): Boolean;
//    function GetSubStrWithRegx(sExpress, S: string): string;
    function GetSubStrListWithRegx(sExpress, S: string): TStrings;

  end;

var
  fStrUtil: TfStrUtil;

implementation
     
uses
  SysUtils, Math;

{ TfStrUtil }     

function TfStrUtil.IsNumber(S: string):Boolean;
var
  nCode: Integer;
  fdNum: Double;
begin
  Val(S, fdNum, nCode);
  Result := (nCode = 0) and (fdNum = fdNum);
end;

function TfStrUtil.IsInteger(S: string):Boolean;
var
  nCode: Integer;
  nNum: Integer;
begin
  Val(S, nNum, nCode);
  Result := (nCode = 0) and (nNum = nNum);
end;   

function TfStrUtil.CompareNumberText(s1, s2: string):Integer;
var
  tempResult:Double;
begin
  if IsNumber(s1) and IsNumber(s2) then
  begin
    tempResult := StrToFloat(s1) - StrToFloat(s2);
    if tempResult > 0 then
      Result := 1
    else if tempResult < 0 then
      Result := -1
    else
      Result := 0;
  end
  else
    Result := SysUtils.CompareText(s1, s2);
end;

function TfStrUtil.Split(sText, sSeparator: String; ssParts: TStrings):TStrings;
var
  nPos: Integer;
  sPart, sRemain: String;
begin
  Result := ssParts;
  if Result = nil then
    Result := TStringList.Create;
  Result.Clear;
  sRemain:= sText;
  nPos:= Pos(sSeparator, sRemain);

  while nPos > 0 do
  begin
    sPart:= Copy(sRemain, 1, nPos- 1);
    if Trim(sPart) <> '' then
      Result.Add(sPart);
    sRemain:= Copy(sRemain, nPos+ Length(sSeparator), Length(sRemain)- nPos+ 1);
    nPos:= Pos(sSeparator, sRemain);
  end;

  if sRemain <> '' then
    Result.Add(sRemain);
end; 
 
function TfStrUtil.SplitEscapeBracket(sText,sSeparator,sBracketLeft, sBracketRight: String;
  ssParts: TStrings):TStrings;
var
  nPos: Integer;
  sPart, sRemain: String;
begin
  Result := ssParts;
  if Result = nil then
    Result := TStringList.Create;
  Result.Clear;
  sRemain:= sText;
  nPos:= PosEscapeBracket(sSeparator, sRemain, sBracketLeft, sBracketRight);

  while nPos > 0 do
  begin
    sPart:= Copy(sRemain, 1, nPos- 1);
    if Trim(sPart) <> '' then
      Result.Add(sPart);
    sRemain:= Copy(sRemain, nPos+ Length(sSeparator), Length(sRemain)- nPos+ 1);
    nPos:= PosEscapeBracket(sSeparator, sRemain, sBracketLeft, sBracketRight);
  end;

  if sRemain <> '' then
    Result.Add(sRemain);
end;    

function TfStrUtil.SplitEscapeQuote(sText, sSeparator, sQuote: String; sEscapeChar: string;
  ssParts: TStrings): TStrings;
var
  nPos: Integer;
  sPart, sRemain: String;
begin
  Result := ssParts;
  if Result = nil then
    Result := TStringList.Create;
  Result.Clear;
  sRemain:= sText;
  nPos:= PosEscapeQuote(sSeparator, sRemain, sQuote, sEscapeChar);

  while nPos > 0 do
  begin
    sPart:= Copy(sRemain, 1, nPos- 1);
    if Trim(sPart) <> '' then
      Result.Add(sPart);
    sRemain:= Copy(sRemain, nPos+ Length(sSeparator), Length(sRemain)- nPos+ 1);
    nPos:= PosEscapeQuote(sSeparator, sRemain, sQuote, sEscapeChar);
  end;

  if sRemain <> '' then
    Result.Add(sRemain);
end;

function TfStrUtil.PosEscapeBracket(sSub, S, sBracketLeft,
  sBracketRight: string; nFrom: Integer; ignoreCase: Boolean): Integer;
var
  nIndexMain, nIndexSub: Integer;
  nIndexSave: Integer;
begin
  Result := 0;
  // 特殊情况
  if Length(sSub) > Length(S) then
    Exit;

  nIndexSub  := 1;
  nIndexMain := nFrom;  // 起始位置

  while nIndexSub <= Length(sSub) do
  begin
    // 跳过大括号
    if Copy(S, nIndexMain, Length(sBracketLeft)) = sBracketLeft then
    begin
      nIndexMain := ParseMatchedBracket(nIndexMain, S, sBracketLeft,
        sBracketRight) + Length(sBracketRight);
    end;
    
    if not ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
            or (sSub[nIndexSub]=S[nIndexMain])) then
    begin
      Inc(nIndexMain);
      if nIndexMain > Length(S) then
        Break;
    end
    else
    begin
      nIndexSave := nIndexMain;    // 保存原位置

      while (nIndexSub<= Length(sSub))
            and ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
                 or (sSub[nIndexSub]=S[nIndexMain])) do
      begin
        Inc(nIndexMain);
        Inc(nIndexSub);
      end;

      if nIndexSub = Length(sSub)+1 then      // 找到
      begin
        Result := nIndexSave;       // 保存下的位置就是结果
        Break;
      end
      else                          // 没找到
      begin
        nIndexMain := nIndexSave + 1;  // 向后搜索
        nIndexSub := 1;
      end;
    end;
  end;
end;   

function TfStrUtil.PosEscapeQuote(sSub, S, sQuote, sEscapeChar: string; nFrom: Integer;
  ignoreCase: Boolean): Integer;
var
  nIndexMain, nIndexSub: Integer;
  nIndexSave: Integer;
begin
  Result := 0;
  // 特殊情况
  if Length(sSub) > Length(S) then
    Exit;

  nIndexSub  := 1;
  nIndexMain := nFrom;  // 起始位置

  while nIndexSub <= Length(sSub) do
  begin
    // 跳过大括号
    if Copy(S, nIndexMain, Length(sQuote)) = sQuote then
    begin
      nIndexMain := ParseMatchedQuote(nIndexMain, S, sQuote, sEscapeChar);
      // 找不到合法的引号结尾
      if nIndexMain = 0 then
      begin
        Result := 0;
        Exit;
      end;
      nIndexMain := nIndexMain + Length(sQuote);
    end;
    
    if not ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
            or (sSub[nIndexSub]=S[nIndexMain])) then
    begin
      Inc(nIndexMain);
      if nIndexMain > Length(S) then
        Break;
    end
    else
    begin
      nIndexSave := nIndexMain;    // 保存原位置

      while (nIndexSub<= Length(sSub))
            and ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
                 or (sSub[nIndexSub]=S[nIndexMain])) do
      begin
        Inc(nIndexMain);
        Inc(nIndexSub);
      end;

      if nIndexSub = Length(sSub)+1 then      // 找到
      begin
        Result := nIndexSave;       // 保存下的位置就是结果
        Break;
      end
      else                          // 没找到
      begin
        nIndexMain := nIndexSave + 1;  // 向后搜索
        nIndexSub := 1;
      end;
    end;
  end;
end;

function TfStrUtil.PosFrom(sSub, S: String;
  nFrom: Integer;ignoreCase: Boolean): Integer;
var
  nIndexMain, nIndexSub: Integer;
  nIndexSave: Integer;
begin
  Result := 0;
  // 特殊情况
  if Length(sSub) > Length(S) then
    Exit;

  nIndexSub  := 1;
  nIndexMain := nFrom;  // 起始位置

  while nIndexSub <= Length(sSub) do
  begin
    if not ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
            or (sSub[nIndexSub]=S[nIndexMain])) then
    begin
      Inc(nIndexMain);
      if nIndexMain > Length(S) then
        Break;
    end
    else
    begin
      nIndexSave := nIndexMain;    // 保存原位置

      while (nIndexSub<= Length(sSub))
            and ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
                 or (sSub[nIndexSub]=S[nIndexMain])) do
      begin
        Inc(nIndexMain);
        Inc(nIndexSub);
      end;

      if nIndexSub = Length(sSub)+1 then      // 找到
      begin
        Result := nIndexSave;       // 保存下的位置就是结果
        Break;
      end
      else                          // 没找到
      begin
        nIndexMain := nIndexSave + 1;  // 向后搜索
        nIndexSub := 1;
      end;
    end;
  end;
end;  

function TfStrUtil.PosBack(sSub, S: String; nFrom: Integer;
  ignoreCase: Boolean): Integer;
var
  nIndexMain, nIndexSub: Integer;
  nIndexSave: Integer;
begin
  Result := 0;
  // 特殊情况
  if Length(sSub) > Length(S) then
    Exit;

  nIndexSub  := Length(sSub);
  nIndexMain := nFrom;  // 起始位置

  while nIndexSub > 0 do
  begin
    if not ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
            or (sSub[nIndexSub]=S[nIndexMain])) then
    begin
      Dec(nIndexMain);
      if nIndexMain <= 0 then
        Break;
    end
    else
    begin
      nIndexSave := nIndexMain;    // 保存原位置

      while (nIndexSub > 0)
            and ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
                 or (sSub[nIndexSub]=S[nIndexMain])) do
      begin
        Dec(nIndexMain);
        Dec(nIndexSub);
      end;

      if nIndexSub = 0 then      // 找到
      begin
        Result := nIndexSave;       // 保存下的位置就是结果
        Break;
      end
      else                          // 没找到
      begin
        nIndexMain := nIndexSave - 1;  // 向后搜索
        nIndexSub := Length(sSub);
      end;
    end;
  end;
end;

function TfStrUtil.PosUntilNot(sSub, S: String; nFrom: Integer;
  ignoreCase: Boolean): Integer;
begin
  Result := nFrom;
  while Result < Length(S) do
  begin
    if (Copy(S, Result, Length(sSub)) = sSub)
       or(ignoreCase and SameText(Copy(S, Result, Length(sSub)), sSub)) then
    begin
      Inc(Result);
    end
    else
    begin
      Break;
    end;
  end;
end;

function TfStrUtil.PosMatchIndex(sSub, sFull: String; nMatchIndex, nFrom: Integer;
  ignoreCase: Boolean): Integer;
var
  nMatch: Integer;
  bFind: Boolean;
begin
  nMatch := 0;
  bFind := False;
  Result := nFrom; // 初始的查找起始位置
  while nMatch < nMatchIndex do
  begin
    // 第一次无匹配，从传入位置找；以后每次查找的位置为上次匹配的位置+1
    if bFind then
      Result := PosFrom(sSub, sFull, Result+1, ignoreCase)
    else                                                  
      Result := PosFrom(sSub, sFull, Result, ignoreCase);
    bFind := Result <> 0;
    if bFind then
      Inc(nMatch)
    else  // 有一次找不到说明已经无匹配，退出
      Break;
  end;
end;

function TfStrUtil.PosMatchCount(sSub, S: String; nFrom: Integer;
     ignoreCase: Boolean): Integer;
var
  nIndexMain, nIndexSub: Integer;
  nIndexSave: Integer;
begin
  Result := 0;
  // 特殊情况
  if Length(sSub) > Length(S) then
    Exit;

  nIndexSub  := 1;
  nIndexMain := nFrom;  // 起始位置

  while nIndexSub <= Length(sSub) do
  begin
    if not ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
            or (sSub[nIndexSub]=S[nIndexMain])) then
    begin
      Inc(nIndexMain);
      if nIndexMain > Length(S) then
        Break;
    end
    else
    begin
      nIndexSave := nIndexMain;    // 保存原位置

      while (nIndexSub<= Length(sSub))
            and ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
                 or (sSub[nIndexSub]=S[nIndexMain])) do      
      begin
        Inc(nIndexMain);
        Inc(nIndexSub);
      end;

      if nIndexSub = Length(sSub)+1 then      // 找到
      begin
        Inc(Result);
      end;

      nIndexMain := nIndexSave + 1;  // 向后搜索
      nIndexSub := 1;
    end;
  end;
end;   

function TfStrUtil.PosArrayFrom(S: string; Subs: array of string;
  var nMatchIndex: Integer; nFrom: Integer; ignoreCase: Boolean): Integer;
var
  nPos, i: Integer; 
  nLastPos: Integer;
begin  
  nLastPos := 0;
  for i := Low(Subs) to High(Subs) do
  begin
    nPos := PosFrom(Subs[i], S, nFrom, ignoreCase);
    if (nPos > 0) and ((nLastPos = 0) or (nPos < nLastPos)) then
    begin
      nLastPos := nPos;
      nMatchIndex := i;
    end;
    if nLastPos = 1 then  // 第一位置提前结束
    begin
      Break;
    end;  
  end;
  Result := nLastPos;
end;  

function TfStrUtil.PosArrayFrom(S: string; Subs: TStrings;
  var nMatchIndex: Integer; nFrom: Integer; ignoreCase: Boolean): Integer;
var
  nPos, i: Integer; 
  nLastPos: Integer;
begin  
  nLastPos := 0;
  for i := 0 to Subs.Count - 1 do
  begin
    nPos := PosFrom(Subs[i], S, nFrom, ignoreCase);
    if (nPos > 0) and ((nLastPos = 0) or (nPos < nLastPos)) then
    begin
      nLastPos := nPos;
      nMatchIndex := i;
    end;
    if nLastPos = 1 then  // 第一位置提前结束
    begin
      Break;
    end;  
  end;
  Result := nLastPos;
end;

function TfStrUtil.PosArrayFromEscapeBracket(S: string;
  Subs: array of string; var nMatchIndex: Integer; sBracketLeft,
  sBracketRight: string; nFrom: Integer; ignoreCase: Boolean): Integer;
var
  nPos, i: Integer; 
  nLastPos: Integer;
begin  
  nLastPos := 0;
  for i := Low(Subs) to High(Subs) do
  begin
    nPos := PosEscapeBracket(Subs[i], S, sBracketLeft, sBracketRight, nFrom,
      ignoreCase);
    if (nPos > 0) and ((nLastPos = 0) or (nPos < nLastPos)) then
    begin
      nLastPos := nPos;
      nMatchIndex := i;
    end;
    if nLastPos = 1 then  // 第一位置提前结束
    begin
      Break;
    end;  
  end;
  Result := nLastPos;
end;

function TfStrUtil.PosArrayFromEscapeQuote(S: string;
  Subs: array of string; var nMatchIndex: Integer; sQuote: string;
  sEscapeChar: string; nFrom: Integer; ignoreCase: Boolean): Integer;
var
  nPos, i: Integer; 
  nLastPos: Integer;
begin  
  nLastPos := 0;
  for i := Low(Subs) to High(Subs) do
  begin
    nPos := PosEscapeQuote(Subs[i], S, sQuote, sEscapeChar, nFrom,
      ignoreCase);
    if (nPos > 0) and ((nLastPos = 0) or (nPos < nLastPos)) then
    begin
      nLastPos := nPos;
      nMatchIndex := i;
    end;
    if nLastPos = 1 then  // 第一位置提前结束
    begin
      Break;
    end;  
  end;
  Result := nLastPos;
end;

function TfStrUtil.PosEndsText(sSub, S: string; var nPos: Integer;
  ignoreCase: Boolean): Boolean;
var
  n: Integer;
begin
  Result := False;
  n := PosLast(sSub, S, 0, ignoreCase);
  if (n > 0) and (n+Length(sSub)-1=Length(S)) then
  begin
    Result := True;
    nPos := n;
  end;
end;

function TfStrUtil.PosLast(sSub, S: String;
  nFrom: Integer;ignoreCase: Boolean): Integer;
var
  nIndexMain, nIndexSub: Integer;
  nIndexSave: Integer;
begin
  Result := 0;
  if nFrom = 0 then
    nIndexMain := Length(S)
  else
    nIndexMain := Min(nFrom, Length(sSub));
  nIndexSub  := Length(sSub);

  // 特殊情况
  if nIndexSub > nIndexMain then
    Exit
  else if nIndexSub = nIndexMain then begin
    if sSub = Copy(S, 1, nIndexMain) then
      Result := 1;
    Exit;
  end;

  while (nIndexSub > 0)
        and (nIndexMain > 0) do
  begin
    if not ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
            or (sSub[nIndexSub]=S[nIndexMain])) then
      Dec(nIndexMain)
    else
    begin
      nIndexSave := nIndexMain;    // 保存原位置

      while (nIndexSub > 0)
            and ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
                 or (sSub[nIndexSub]=S[nIndexMain])) do
      begin
        Dec(nIndexMain);
        Dec(nIndexSub);
      end;

      if nIndexSub = 0 then      // 找到
      begin
        Result := nIndexMain + 1;       // +1后为正确起始位置
        Break;
      end
      else                      // 没找到
      begin
        nIndexMain := nIndexSave - 1;  // 向前搜索
        nIndexSub := Length(sSub);;
      end;
    end;
  end;
end;

function TfStrUtil.SubStringBetween(S, sBegin, sEnd: string;
    ignoreCase: Boolean): string;
var
  nPosBegin, nPosEnd: Integer;
begin
  if sBegin = '' then            // 为空表示从头开始
    nPosBegin := 1
  else
    nPosBegin := PosFrom(sBegin, S, 1, ignoreCase);

  if sEnd = '' then              // 为空表示截取到尾
    nPosEnd := Length(S)
  else
    nPosEnd := PosFrom(sEnd, S, nPosBegin+1, ignoreCase);

  // 开头和结尾必须都找到，或者明确给定为''分别表示起始和结尾，否则返回空
  if (nPosBegin = 0) or (nPosEnd = 0) then
    Result := ''
  else
    Result := Copy(S, nPosBegin+Length(sBegin),nPosEnd-nPosBegin-Length(sBegin));
end;

//function TfStrUtil.SameStrCustomCase(s1, s2: string;
//  ignoreCase: Boolean): Boolean;
//begin
//  if ignoreCase then
//    Result := SameText(s1, s2)
//  else
//    Result := s1=s2;
//end;    

function TfStrUtil.CombineListToText(List: TStrings; sDelimiter: string;
  bQuote: Boolean): string;
var
  i: Integer;
begin
  for i := 0 to List.Count - 1 do
  begin
    if bQuote then
      List[i] := QuotedStr(List[i]);
    if (Result = '')
       or (bQuote and (Result = QuotedStr(''))) then
      Result := List[i]
    else
      Result := Result + sDelimiter + List[i];
  end;
end;

function TfStrUtil.RebuildDelimitatedText(sText, sOldDelimiter,
  sNewDelimiter: string; bQuote: Boolean): string;
var
  slst: TStrings;
begin
  slst:= TStringList.Create;
  try
    Split(sText, sOldDelimiter, slst);
    Result := CombineListToText(slst,sNewDelimiter, bQuote);
  finally
    slst.Free;
  end;   
end;  

procedure TfStrUtil.StrToKeyValue(AStr, sSeparator: string; var Key,
  Value: string);
var
  nIndex: Integer;
begin
  nIndex := Pos(sSeparator, AStr);
  if nIndex = 0 then
  begin
    Key := '';
    Value := Trim(AStr);
  end
  else
  begin
    Key := Trim(Copy(AStr, 1, nIndex-Length(sSeparator)));
    Value := Trim(Copy(AStr, nIndex+Length(sSeparator),MaxInt));
  end;
end;

function TfStrUtil.WildcardCompareText(const sSub, S: string;
  ignoreCase: Boolean): Integer;
var
  nIndexMain, nIndexSub: Integer;
//  nValidLengthSub: Integer;
  nPos: Integer;
  nPos2: Integer;
  nLengthSub, nLengthS: Integer;
  nMatchWildCard: Integer;
  sNotWildCardInSub: string;
  nMatchIndex: Integer;
begin
  Result := 0;
  nIndexSub  := 1;
  nIndexMain := 1;  // 起始位置
  nLengthSub := Length(sSub);
  nLengthS := Length(S);
  nMatchIndex := 0;
//  nValidLengthSub := nLengthSub - PosMatchCount('*', sSub, 1, ignoreCase);

  while (nIndexSub <= nLengthSub)
        and (nIndexMain <= nLengthS) do
  begin
    if ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
        or (sSub[nIndexSub]=S[nIndexMain]))
       or (sSub[nIndexSub]='?') then
    begin
      // 记录第一次匹配的索引
      if nMatchIndex = 0 then
        nMatchIndex := nIndexMain;
      Inc(nIndexSub);
      Inc(nIndexMain);
      Continue;
    end else if sSub[nIndexSub]='*' then
    begin
      if nLengthSub = nIndexSub then begin
        nIndexMain := nLengthS+1;  // succ
      end else begin
//        nPos := PosLast(sSub[nIndexSub+1], S, 0, ignoreCase);
                                       
        // 母串1234567890abcd
        // 子串1234*6789?abcd    从 * 取到 ? 即中间的sNotWildCardInSub取出6789
        // 子串123*6789          从 * 取到结尾
        nPos2 := PosArrayFrom(sSub, ['*','?'], nMatchWildCard, nIndexSub+1, ignoreCase); // 看是否还有通配符
        if nPos2 = 0 then
          nPos2 := Length(sSub)+1;
        sNotWildCardInSub := Copy(sSub, nIndexSub+1, nPos2-nIndexSub - 1);
        // sNotWildCardInSub 在母串中的开始位置
        nPos := PosFrom(sNotWildCardInSub, S, nIndexMain);
        if nPos > 0 then
        begin
          nIndexMain := nPos + Length(sNotWildCardInSub);
//          nMatchIndex := nPos - (nLengthSub - Length(sNotWildCardInSub));
          nMatchIndex := nIndexSub;
          nIndexSub := nPos2;
        end
        else
          Break;     // 后面的字符没有 说明是匹配失败的情况
      end;  
    end else
    begin
      nMatchIndex := 0;
      Break;
    end;
  end;   
  if (nIndexSub > nLengthSub) then // 子串分析完毕，取匹配索引返回
    Result := nMatchIndex;
end;

function TfStrUtil.ParseMatchedBracket(nLeftBracketIndex: Integer; Astr,
  sBracketLeft, sBracketRight: string): Integer;
var
  nLeftBracketsBetween: Integer;
  i: Integer;
//  isQuote: Boolean;
begin
  Result := 0;
//  isQuote := sBracketLeft = sBracketRight;
//  if isQuote then
//  begin
////    Result := 0;
////    i := nLeftBracketIndex + 1;      // 从下一个开始
////    while i <= Length(Astr) do
////    begin
////      if Astr[i] = sBracketLeft then     // 找到引号
////      begin
////        if (i > 0) and (Astr[i-1] <> sEscapeChar) then    //检查转义符
////        begin
////          Result := i;
////          Break;
////        end;
////      end;
////      Inc(i);
////    end
//  end
//  else
//  begin
    nLeftBracketsBetween := 0;       // 左括号和与之匹配的右括号之间有多少个左括号
    i := nLeftBracketIndex + 1;      // 从下一个开始
    while i <= Length(Astr) do
    begin
      if Astr[i] = sBracketLeft then
        Inc(nLeftBracketsBetween)
      else if Astr[i] = sBracketRight then begin
        if nLeftBracketsBetween = 0 then begin
          Result := i;
          Break;
        end else
          Dec(nLeftBracketsBetween);
      end;
      Inc(i);
    end;
//  end;
end;   

function TfStrUtil.ParseMatchedQuote(nLeftQuoteIndex: Integer; Astr,
  sQuote: string; sEscapeChar: string): Integer;
var
  i: Integer;
  nTotalLen: Integer;
  nQuoteLen, nEscapeCharLen: Integer;
begin
  Result := 0;
  i := nLeftQuoteIndex + 1;      // 从下一个开始
  nTotalLen := Length(Astr);
  nQuoteLen := Length(sQuote);
  nEscapeCharLen := Length(sEscapeChar);
  while i <= nTotalLen do
  begin
    if sEscapeChar <> '' then
    begin
      // 先检查sEscapeChar 适合sEscapeChar=sQuote的情况
      if Copy(Astr, i, nEscapeCharLen) = sEscapeChar then
      begin
        if (i<nTotalLen) and (Copy(Astr, i+1, nQuoteLen) = sQuote) then
        begin
          Inc(i, nQuoteLen + nEscapeCharLen);
        end;
      end;
    end;
    if Copy(Astr, i, nQuoteLen) = sQuote then     // 找到引号
    begin
      if i > 0 then
      begin
        Result := i;
        Break;
      end;
    end;
    Inc(i);
  end;
end;

function TfStrUtil.IsPosInQuotedStr(sMain: string; nSubPos: Integer;
  sQuoteStart, sQuoteEnd: string): Boolean;
var
  nInQuotedCount: Integer;
  nCurrPos: Integer;
begin
  Result := False;
  nInQuotedCount := 0;
  nCurrPos := 1;
  while nCurrPos < Length(sMain)do
  begin
    // 引号开头 增加计数
    if Copy(sMain, nCurrPos, Length(sQuoteStart)) = sQuoteStart then
      Inc(nInQuotedCount)
    // 引号结尾 减少计数，不会小于0
    else if (Copy(sMain, nCurrPos, Length(sQuoteEnd)) = sQuoteEnd)
            and (nInQuotedCount > 0) then
      Dec(nInQuotedCount);
    // 看计数判断是否在引号中
    if nCurrPos = nSubPos then begin
      if nInQuotedCount > 0 then
        Result := True
      else
        Result := False;
      Break;
    end;
    Inc(nCurrPos);
  end;  
end;

function TfStrUtil.CompareTextAtIndex(s1, s2: string): Integer;
var
  nIndex: Integer;
begin
  nIndex := 1;
  while (nIndex <= Length(s1)) and (nIndex <= Length(s2)) do
  begin
    if SameText(s1[nIndex], s2[nIndex]) then
      Inc(nIndex)
    else
      Break;
  end;
  if (nIndex = Length(s1)+1) and (nIndex = Length(s2)+1) then
    Result := 0
  else
    Result := nIndex;
end;

function TfStrUtil.MatchRegxSubStr(sSub, S: string): Boolean;
var
  mr: TMatchResult;
begin
  Result := False;
  mr := TRegexUtil.MatchFirst(sSub, 0, S);
  if mr <> nil then
  try
    Result := '' <> mr.MatchStr;
  finally
    mr.Free;
  end;
end;
//
//function TfStrUtil.GetSubStrWithRegx(sExpress, S: string): string;
//var
//  expr: TRegExpr;
//begin
//  expr := TRegExpr.Create;
//  try
//    expr.InputString := S;
//    expr.Expression := sExpress;
//    if expr.Expression = '' then
//      Exit;
//    if expr.Exec then
//    begin
//      Result := expr.Match[0];
//    end;
//  finally
//    expr.Free;
//  end;
//end;
//
function TfStrUtil.GetSubStrListWithRegx(sExpress, S: string): TStrings;
var
  mrList: TObjectList<TMatchResult>;
  i: Integer;
begin
  Result := TStringList.Create;
  mrList := TObjectList<TMatchResult>.Create;
  try
    TRegexUtil.MatchAll(sExpress, 0, S, mrList);
    i := 0;
    while i < mrList.Count do
    begin
      Result.Add(mrList[i].MatchStr);
      Inc(i);
    end;
  finally
    mrList.Free;
  end;
end;

initialization
  fStrUtil := TfStrUtil.Create;
finalization
  if Assigned(fStrUtil) then
    FreeAndNil(fStrUtil);

end.
