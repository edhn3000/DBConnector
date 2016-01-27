unit U_SqlUtils;

interface

uses
  Classes, StrUtils, SysUtils;

type
  TSqlUtils = class
  private
  public                                                     
    class function IsSingleLineComment(const ASql: string): Boolean;
    class function IsSqlBeginCommand(const sSqlCommand: string): Boolean;
    class procedure CheckSqlScript(var slstSqlList: TStringList);
    class procedure RemoveNewLine(var Text: string; sReplace : string = ' ');
    class procedure RemoveComment(var slstSqlList: TStringList);
    class function DeleteSQLComment(const sSql: string;
      bOnlyRemoveHead: Boolean = False): string;
    class function GetSqlCommand(Asql: string): string;overload;
    class function GetSqlCommand(Asql: string;
      var sRelData: string): string;overload;
    class function IsQuerySql(ASql: string): Boolean; 
    class function GetTableNameBySQL(sSql: string): string;
    // 将insert语句分隔
    class procedure SplitInsert(const sInsert: string; slstNameValues: TStringList);

  end;

implementation

uses
  U_fStrUtil;

const     
  C_sComment_MultiBegin = '/*';
  C_sComment_MultiEnd   = '*/';
  
  // 数组中全部为小写关键字
  C_as_SQLBeginCommands: array[0..12] of string = (
    'insert', 'update', 'delete', 'select', 'create', 'drop', 'alter',
    'grant', 'commit', 'rollback', 'set', 'open', 'close');

{ TSqlUtils }

class function TSqlUtils.IsSqlBeginCommand(const sSqlCommand: string):Boolean;
  function StrInArray(str: string; strArray: array of string): Boolean;
  var
    i: Integer;
  begin
    Result := False;
    for i := Low(strArray) to High(strArray) do
      if str = strArray[i] then
      begin
        Result := True;
        Break;
      end;
  end;
begin
  Result := StrInArray(LowerCase(sSqlCommand), C_as_SQLBeginCommands);
end;

class procedure TSqlUtils.CheckSqlScript(var slstSqlList: TStringList);
var
  i: Integer;
  sSql: string;
  nFirstSpacePos: Integer;
begin
  RemoveComment(slstSqlList);
  for i := slstSqlList.Count - 1 downto 0 do
  begin
    sSql := Trim(slstSqlList.Strings[i]);
    nFirstSpacePos := Pos(' ',sSql);
    if nFirstSpacePos = 0 then  // 至少关键字之后应该有一个空格，否则不合法
      slstSqlList.Delete(i)
    else if not IsSqlBeginCommand(Copy(sSql, 1, nFirstSpacePos)) then
      slstSqlList.Delete(i);
  end;
end;

class function TSqlUtils.DeleteSQLComment(const sSql: string;
  bOnlyRemoveHead: Boolean): string;
var
  nPosComments, nPos1, nPos2: Integer;
  nMatch: Integer;
//// DeleteComment Begin
begin
  Result := sSql;
  nPos1 := Pos(C_sComment_MultiBegin, Result);
  nPos2 := Pos(C_sComment_MultiEnd, Result);
  if (nPos1 > 0) and (nPos1 < nPos2) then             // 判断多行注释开头和结尾在同行时谁在前
    Result := Copy(Result,1, nPos1-1) + Copy(Result, nPos2+Length(
      C_sComment_MultiEnd), Length(Result))
  else if nPos2 > 0 then
    Result := Copy(Result, nPos2+Length(C_sComment_MultiEnd), Length(Result));

  nPosComments := fStrUtil.PosArrayFromEscapeQuote(Result,['--', '//', '/*'],
    nMatch, '''', '''');
  if nPosComments <> 0 then
  begin
    if not bOnlyRemoveHead then
      Result := Copy(Result, 1, nPosComments - 1)
    else
      Result := Copy(Result, 3, MaxInt)
  end;
end;

class function TSqlUtils.GetSqlCommand(Asql: string): string;
var
  s: string;
begin
  Result := GetSqlCommand(Asql, s);
end;

class function TSqlUtils.GetSqlCommand(Asql: string;
  var sRelData: string): string;
  function getTableNameBySQL(sSql: string): string;
  const                  
    C_ARY_KEWWORDS_OBJ: array[0..6] of string =
      ('table', 'procedure', 'function', 'view', 'trigger', 'index', 'from');
  var
    nPos1, nPos2: Integer;
    sTable: string;
    nMatch: Integer;
  begin
    sSql := StringReplace(sSql, #$D, ' ', [rfReplaceAll]);  
    sSql := StringReplace(sSql, #$A, ' ', [rfReplaceAll]);
    nPos1  := fStrUtil.PosArrayFrom(sSql, C_ARY_KEWWORDS_OBJ, nMatch);
    if (nPos1 <> 0) then
    begin
      sTable := Trim(Copy(sSql, nPos1+LENGTH(C_ARY_KEWWORDS_OBJ[nMatch]), MaxInt));
      nPos2  := fStrUtil.PosFrom('(', sTable);
      if nPos2 = 0 then
        nPos2 := MaxInt;
      sTable := Trim(Copy(sTable, 1, nPos2-1));
    end
    else
      sTable := 'UnknownTable';
    Result := sTable;
  end;
begin
  Result := Copy(Asql, 1, Pos(' ', Asql)-1);
  if SameText(Result, 'create')
     or SameText(Result, 'drop') then
    sRelData := getTableNameBySQL(Asql);
end;

class function TSqlUtils.IsSingleLineComment(const ASql: string):Boolean;
var
  s: string;
begin
  s := Trim(ASql);
  if AnsiStartsText('--', s) or
     AnsiStartsText('//', s) then
    Result := True
  else
    Result := False;
end;

class procedure TSqlUtils.RemoveComment(var slstSqlList: TStringList);
var
  i: Integer;
  sSql: string;
begin
  for i := slstSqlList.Count - 1 downto 0 do
  begin
    sSql := Trim(slstSqlList.Strings[i]);
    if (sSql = '') or
       IsSingleLineComment(sSql) then
      slstSqlList.Delete(i)
  end;
end;

class procedure TSqlUtils.RemoveNewLine(var Text: string;
  sReplace: string);
begin 
  Text := StringReplace( Text, #$D#$A, sReplace, [rfReplaceAll, rfIgnoreCase] );
end;  

class function TSqlUtils.IsQuerySql(ASql: string): Boolean;
begin
  Result := AnsiStartsText('select ', Asql);
end;

class function TSqlUtils.GetTableNameBySQL(sSql: string): string;
var
  nPos1, nPos2: Integer;
  nMatchIndex: Integer;
  sTable: string;
  sCommand: string;
begin
  sSql := Trim(sSql);    
  sTable := '';
  sCommand := fStrUtil.SubStringBetween(LowerCase(Trim(sSql)), '', ' ');
  if (sCommand = 'select')
     or (sCommand = 'delete') then
  begin
    // 取from后 一个单词
    nPos1  := fStrUtil.PosFrom('from', sSql);
    if (nPos1 <> 0) then
    begin
      sTable := Trim(Copy(sSql, nPos1+4, MaxInt));
      nPos2  := fStrUtil.PosFrom(' ', sTable);   // 再找下一个空格
      if nPos2 = 0 then
        nPos2 := MaxInt;
      sTable := Trim(Copy(sTable, 1, nPos2));
    end;
  end
  else if (sCommand = 'insert') then
  begin
    // 取insert into 后一个单词
    nPos1 := Pos('into', sSql);
    if nPos1 <> 0 then
    begin
      // into后定位到非空格的位置 就是表名
      nPos1 := fStrUtil.PosUntilNot(' ', sSql, nPos1+4);
      // 表名开始位置 一直找到有 ( 或values
      nPos2 := fStrUtil.PosArrayFrom(sSql, ['(', 'values'], nMatchIndex,
        nPos1);
      if nPos2 = 0 then begin
        // not match?
        sTable := Copy(sSql, nPos1, MaxInt);
      end else if nMatchIndex = 0 then begin
        // 说明括号在values 前 取到括号 要trim
        sTable := Trim(Copy(sSql, nPos1, nPos2-nPos1-1));
      end else // if nMatchIndex = 1 then begin
        // 说明values在括号前 即表后未跟字段
        sTable := Trim(Copy(sSql, nPos1, nPos2-nPos1-1));
//      end;
    end;
  end
  else if (sCommand = 'update') then
  begin
    // 取update 后一个单词
    nPos1 := fStrUtil.PosUntilNot(' ', sSql, 7);
    nPos2 := fStrUtil.PosFrom(' ', sSql, nPos1+1);
    if nPos2 = 0 then
      nPos2 := MaxInt;
    sTable := Copy(sSql, nPos1, nPos2-nPos1-1);
  end;
  Result := sTable;
end;    

class procedure TSqlUtils.SplitInsert(const sInsert: string;
  slstNameValues: TStringList);
var  
  sFields, sValues: string;
  slstFields, slstValues: TStringList;
  i: Integer;
  nPos1, nPos2, nPos3: Integer;
begin
  nPos1 := fStrUtil.PosFrom('(', sInsert, 1, True);
  nPos2 := fStrUtil.ParseMatchedBracket(nPos1, sInsert);
  nPos3 := fStrUtil.PosFrom('values', sInsert, 1, True);
  if (nPos1<=nPos2) and (nPos3 >= nPos2) then
  begin
    sFields := Copy(sInsert, nPos1+1, nPos2-nPos1-1);
    nPos1 := fStrUtil.PosFrom('(', sInsert, nPos3);
    nPos2 := fStrUtil.ParseMatchedBracket(nPos1, sInsert);
    sValues := Copy(sInsert, nPos1+1, nPos2-nPos1-1);
    slstFields := TStringList.Create;
    slstValues := TStringList.Create;
    try
      fStrUtil.Split(sFields, ',', slstFields);  
      fStrUtil.Split(sValues, ',', slstValues);
      for i := 0 to slstFields.Count - 1 do
      begin
        if i <= slstValues.Count - 1 then
          slstNameValues.Add(Format('%s=%s',[slstFields[i], Trim(slstValues[i])]))
        else
          slstNameValues.Add(Format('%s=%s',[slstFields[i], '']));
      end;
    finally
      slstFields.Free;
      slstValues.Free;
    end;   
  end;  
end;

end.
