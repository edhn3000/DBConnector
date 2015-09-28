{**
 * <p>Title: fStrUtil </p>
 * <p>Copyright: Copyright (c) 2006/10/14</p>
 * @author fengyq
 * @version 1.0
 * <p>Description: �ַ�������Ԫ </p>
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

    // ������ʹ����ֵ��С�Ƚϣ����ַ���ʹ��ComareText��s1����s2����1�� s1С��s2����-1����ȷ���0��
    function CompareNumberText(s1, s2: string): Integer;

    function CombineListToText(List: TStrings; sDelimiter: string;
      bQuote: Boolean): string;
    function Split(sText,sSeparator: String; ssParts: TStrings):TStrings;
    // ����ַ�����ʱ��Խ��ָ�����͵������е����ݣ������ַ���Ƕ�׵����
    function SplitEscapeBracket(sText,sSeparator,sBracketLeft, sBracketRight: String;
      ssParts: TStrings):TStrings;   
    function SplitEscapeQuote(sText,sSeparator,sQuote: String;sEscapeChar: string;
      ssParts: TStrings):TStrings;
    function RebuildDelimitatedText(sText, sOldDelimiter, sNewDelimiter: string;
      bQuote: Boolean): string;
    procedure StrToKeyValue(AStr, sSeparator:string; var Key, Value: string);

    // ���ַ����в����Ӵ�ʱԽ��ָ�����͵������е����ݣ������ַ���Ƕ�׵����                                   
    function PosEscapeBracket(sSub, S: String; sBracketLeft, sBracketRight: string;
      nFrom: Integer = 1; ignoreCase: Boolean = True): Integer;                                    
    function PosEscapeQuote(sSub, S: String; sQuote: string;sEscapeChar: string;
      nFrom: Integer = 1; ignoreCase: Boolean = True): Integer;
    {*
     * �ɴ�����ʼλ�õ�Pos����
     *}
    function PosFrom(sSub, S: String; nFrom: Integer = 1;
      ignoreCase: Boolean = True): Integer;
    function PosBack(sSub, S: String; nFrom: Integer = 1;
      ignoreCase: Boolean = True): Integer;

    // ��λ�������Ӵ�֮ǰ��λ��
    function PosUntilNot(sSub, S: String; nFrom: Integer = 1;
      ignoreCase: Boolean = True): Integer;
    {*
     * ����ָ�����ַ�������ָ�����ڼ���ƥ��λ��
     *}
    function PosMatchIndex(sSub, sFull: String; nMatchIndex: Integer =1;
      nFrom: Integer = 1; ignoreCase: Boolean = True): Integer;
    function PosMatchCount(sSub, S: String; nFrom: Integer = 1;
         ignoreCase: Boolean = True): Integer;
         
    // �Ӻ���ǰ���ң�û�ҵ�����0
    function PosLast(sSub, S: String; nFrom: Integer = 0;
        ignoreCase: Boolean = True): Integer;
    function PosEndsText(sSub, S: string; var nPos:Integer;
      ignoreCase: Boolean = True): Boolean;
    function ParseMatchedBracket(nLeftBracketIndex:Integer; Astr: string;
        sBracketLeft: string = '('; sBracketRight: string = ')'): Integer; 
    function ParseMatchedQuote(nLeftQuoteIndex:Integer; Astr: string;
        sQuote: string = ''''; sEscapeChar: string = '\'): Integer;

    {* ���������б��������ַ���������ǰ��λ��
     * @param S            ĸ��
     * @param Subs         �Ӵ�����
     * @param nMatchIndex  ������ƥ���Ԫ�صĽű�
     * @return             �ǰ��λ�ã���û�ҵ�����0
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

    // �ж�һ��λ���Ƿ���������
    function IsPosInQuotedStr(sMain: string; nSubPos: Integer;
      sQuoteStart, sQuoteEnd: string): Boolean;
      
    // �Ƚϣ���ȫ��ͬ����0�����򷵻��в���index
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
  // �������
  if Length(sSub) > Length(S) then
    Exit;

  nIndexSub  := 1;
  nIndexMain := nFrom;  // ��ʼλ��

  while nIndexSub <= Length(sSub) do
  begin
    // ����������
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
      nIndexSave := nIndexMain;    // ����ԭλ��

      while (nIndexSub<= Length(sSub))
            and ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
                 or (sSub[nIndexSub]=S[nIndexMain])) do
      begin
        Inc(nIndexMain);
        Inc(nIndexSub);
      end;

      if nIndexSub = Length(sSub)+1 then      // �ҵ�
      begin
        Result := nIndexSave;       // �����µ�λ�þ��ǽ��
        Break;
      end
      else                          // û�ҵ�
      begin
        nIndexMain := nIndexSave + 1;  // �������
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
  // �������
  if Length(sSub) > Length(S) then
    Exit;

  nIndexSub  := 1;
  nIndexMain := nFrom;  // ��ʼλ��

  while nIndexSub <= Length(sSub) do
  begin
    // ����������
    if Copy(S, nIndexMain, Length(sQuote)) = sQuote then
    begin
      nIndexMain := ParseMatchedQuote(nIndexMain, S, sQuote, sEscapeChar);
      // �Ҳ����Ϸ������Ž�β
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
      nIndexSave := nIndexMain;    // ����ԭλ��

      while (nIndexSub<= Length(sSub))
            and ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
                 or (sSub[nIndexSub]=S[nIndexMain])) do
      begin
        Inc(nIndexMain);
        Inc(nIndexSub);
      end;

      if nIndexSub = Length(sSub)+1 then      // �ҵ�
      begin
        Result := nIndexSave;       // �����µ�λ�þ��ǽ��
        Break;
      end
      else                          // û�ҵ�
      begin
        nIndexMain := nIndexSave + 1;  // �������
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
  // �������
  if Length(sSub) > Length(S) then
    Exit;

  nIndexSub  := 1;
  nIndexMain := nFrom;  // ��ʼλ��

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
      nIndexSave := nIndexMain;    // ����ԭλ��

      while (nIndexSub<= Length(sSub))
            and ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
                 or (sSub[nIndexSub]=S[nIndexMain])) do
      begin
        Inc(nIndexMain);
        Inc(nIndexSub);
      end;

      if nIndexSub = Length(sSub)+1 then      // �ҵ�
      begin
        Result := nIndexSave;       // �����µ�λ�þ��ǽ��
        Break;
      end
      else                          // û�ҵ�
      begin
        nIndexMain := nIndexSave + 1;  // �������
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
  // �������
  if Length(sSub) > Length(S) then
    Exit;

  nIndexSub  := Length(sSub);
  nIndexMain := nFrom;  // ��ʼλ��

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
      nIndexSave := nIndexMain;    // ����ԭλ��

      while (nIndexSub > 0)
            and ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
                 or (sSub[nIndexSub]=S[nIndexMain])) do
      begin
        Dec(nIndexMain);
        Dec(nIndexSub);
      end;

      if nIndexSub = 0 then      // �ҵ�
      begin
        Result := nIndexSave;       // �����µ�λ�þ��ǽ��
        Break;
      end
      else                          // û�ҵ�
      begin
        nIndexMain := nIndexSave - 1;  // �������
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
  Result := nFrom; // ��ʼ�Ĳ�����ʼλ��
  while nMatch < nMatchIndex do
  begin
    // ��һ����ƥ�䣬�Ӵ���λ���ң��Ժ�ÿ�β��ҵ�λ��Ϊ�ϴ�ƥ���λ��+1
    if bFind then
      Result := PosFrom(sSub, sFull, Result+1, ignoreCase)
    else                                                  
      Result := PosFrom(sSub, sFull, Result, ignoreCase);
    bFind := Result <> 0;
    if bFind then
      Inc(nMatch)
    else  // ��һ���Ҳ���˵���Ѿ���ƥ�䣬�˳�
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
  // �������
  if Length(sSub) > Length(S) then
    Exit;

  nIndexSub  := 1;
  nIndexMain := nFrom;  // ��ʼλ��

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
      nIndexSave := nIndexMain;    // ����ԭλ��

      while (nIndexSub<= Length(sSub))
            and ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
                 or (sSub[nIndexSub]=S[nIndexMain])) do      
      begin
        Inc(nIndexMain);
        Inc(nIndexSub);
      end;

      if nIndexSub = Length(sSub)+1 then      // �ҵ�
      begin
        Inc(Result);
      end;

      nIndexMain := nIndexSave + 1;  // �������
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
    if nLastPos = 1 then  // ��һλ����ǰ����
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
    if nLastPos = 1 then  // ��һλ����ǰ����
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
    if nLastPos = 1 then  // ��һλ����ǰ����
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
    if nLastPos = 1 then  // ��һλ����ǰ����
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

  // �������
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
      nIndexSave := nIndexMain;    // ����ԭλ��

      while (nIndexSub > 0)
            and ((ignoreCase and SameText(sSub[nIndexSub], S[nIndexMain]))
                 or (sSub[nIndexSub]=S[nIndexMain])) do
      begin
        Dec(nIndexMain);
        Dec(nIndexSub);
      end;

      if nIndexSub = 0 then      // �ҵ�
      begin
        Result := nIndexMain + 1;       // +1��Ϊ��ȷ��ʼλ��
        Break;
      end
      else                      // û�ҵ�
      begin
        nIndexMain := nIndexSave - 1;  // ��ǰ����
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
  if sBegin = '' then            // Ϊ�ձ�ʾ��ͷ��ʼ
    nPosBegin := 1
  else
    nPosBegin := PosFrom(sBegin, S, 1, ignoreCase);

  if sEnd = '' then              // Ϊ�ձ�ʾ��ȡ��β
    nPosEnd := Length(S)
  else
    nPosEnd := PosFrom(sEnd, S, nPosBegin+1, ignoreCase);

  // ��ͷ�ͽ�β���붼�ҵ���������ȷ����Ϊ''�ֱ��ʾ��ʼ�ͽ�β�����򷵻ؿ�
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
  nIndexMain := 1;  // ��ʼλ��
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
      // ��¼��һ��ƥ�������
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
                                       
        // ĸ��1234567890abcd
        // �Ӵ�1234*6789?abcd    �� * ȡ�� ? ���м��sNotWildCardInSubȡ��6789
        // �Ӵ�123*6789          �� * ȡ����β
        nPos2 := PosArrayFrom(sSub, ['*','?'], nMatchWildCard, nIndexSub+1, ignoreCase); // ���Ƿ���ͨ���
        if nPos2 = 0 then
          nPos2 := Length(sSub)+1;
        sNotWildCardInSub := Copy(sSub, nIndexSub+1, nPos2-nIndexSub - 1);
        // sNotWildCardInSub ��ĸ���еĿ�ʼλ��
        nPos := PosFrom(sNotWildCardInSub, S, nIndexMain);
        if nPos > 0 then
        begin
          nIndexMain := nPos + Length(sNotWildCardInSub);
//          nMatchIndex := nPos - (nLengthSub - Length(sNotWildCardInSub));
          nMatchIndex := nIndexSub;
          nIndexSub := nPos2;
        end
        else
          Break;     // ������ַ�û�� ˵����ƥ��ʧ�ܵ����
      end;  
    end else
    begin
      nMatchIndex := 0;
      Break;
    end;
  end;   
  if (nIndexSub > nLengthSub) then // �Ӵ�������ϣ�ȡƥ����������
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
////    i := nLeftBracketIndex + 1;      // ����һ����ʼ
////    while i <= Length(Astr) do
////    begin
////      if Astr[i] = sBracketLeft then     // �ҵ�����
////      begin
////        if (i > 0) and (Astr[i-1] <> sEscapeChar) then    //���ת���
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
    nLeftBracketsBetween := 0;       // �����ź���֮ƥ���������֮���ж��ٸ�������
    i := nLeftBracketIndex + 1;      // ����һ����ʼ
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
  i := nLeftQuoteIndex + 1;      // ����һ����ʼ
  nTotalLen := Length(Astr);
  nQuoteLen := Length(sQuote);
  nEscapeCharLen := Length(sEscapeChar);
  while i <= nTotalLen do
  begin
    if sEscapeChar <> '' then
    begin
      // �ȼ��sEscapeChar �ʺ�sEscapeChar=sQuote�����
      if Copy(Astr, i, nEscapeCharLen) = sEscapeChar then
      begin
        if (i<nTotalLen) and (Copy(Astr, i+1, nQuoteLen) = sQuote) then
        begin
          Inc(i, nQuoteLen + nEscapeCharLen);
        end;
      end;
    end;
    if Copy(Astr, i, nQuoteLen) = sQuote then     // �ҵ�����
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
    // ���ſ�ͷ ���Ӽ���
    if Copy(sMain, nCurrPos, Length(sQuoteStart)) = sQuoteStart then
      Inc(nInQuotedCount)
    // ���Ž�β ���ټ���������С��0
    else if (Copy(sMain, nCurrPos, Length(sQuoteEnd)) = sQuoteEnd)
            and (nInQuotedCount > 0) then
      Dec(nInQuotedCount);
    // �������ж��Ƿ���������
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
