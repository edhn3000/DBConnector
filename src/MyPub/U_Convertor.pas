unit U_Convertor;

interface   
uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, U_CommonFunc, IdCoderMIME, IdCoderUUE, IdCoderXXE,
  U_fStrUtil, uSHA1, uMD5;
  
const
  C_colModulus = 0.4374;
  C_rowModulus = 2.7682;
  C_MM2Inch = 1/25.4;

type
  // TCALL_ID (呼叫标识)
  TCALL_ID = packed record //呼叫标识结构
    ulTime:ULONG;          //呼叫进入时间
    usDsn:Word;            //呼叫进入的任务号
    ucHandle:Byte;         //呼叫进入一个任务的次数
    ucServer:Byte;         //唯一标识一个服务器的标识
  end;
  PCALL_ID = ^TCALL_ID;

  TMM2PelModule = (mpmCol, mpmRow, mpmInch);
  TConvertDataType = (cdtNone, cdtMM, cdtExcelPels, cdtInch,
                      cdtHZ, cdtUnicode, cdtPyHead, cdtChar,
                      cdtBase64, cdtXUE{XXEUUEMME加密}, cdtMD5, cdtDES,
                      cdtSHA1, cdtHMAC_MD5,
                      cdtICDBilllogCallId,
                      cdtICDRecordInfoCallid);
  TConvertor = class
  private
  public
    MM2PelModule: TMM2PelModule;

    UniCodePrefix: string;
    UniCodeHex: Boolean;
    NormalKey: string;
    
    LastConvertResult: string;

    function MM2Pels(fdMM: Double; mpmMoudle: TMM2PelModule): Double;
    function Pels2MM(fdPels: Double; mpmMoudle: TMM2PelModule): Double;
    
    function Dec2Hex(Dec: string): string;
    function Hex2Dec(Hex: string): string;

    function Hz2Unicode(Hz:string;sPrefix: string = '\u'; bHex: Boolean = True): string;
    function Unicode2Hz(UniCode:string;sPrefix: string = '\u'; bHex: Boolean = True): string;

    function Hz2PySingle(const AHzStr: string): string;
    function GetPYIndexChar(hzchar:string): Char;

    function Normal2Base64(const s: string): string;
    function Base642Normal(const s: string): string;

    function XXEUUEMIMEEncode(str:string):string;
    function XXEUUEMIMEDecode(str:string):string;

    function MD5Encode(str:string):string;
    function MD5Decode(str:string):string;

    function DESEncode(str, key:string):string;
    function DESDecode(str, key:string):string;

    function getCallIDAtPlatFormVer31(nPointerCallID: Integer): string; // 3.0/3.1版本的平台
    function getCallIDAtPlatFormVer32(nPointerCallID: Integer): string; // 3.2及以上版本的平台
    function getRecordInfoCallID(nPointerCallID: Integer): string;    
    function BillLogCallIDVer31ToRecordInfoCallID(sBillLogCallID: string): string;
    function BillLogCallIDVer32ToRecordInfoCallID(sBillLogCallID: string): string;
    function RecordInfoCallIDToBillLogCallIDVer31(sRecordInfoCallID: string): string;
    function RecordInfoCallIDToBillLogCallIDVer32(sRecordInfoCallID: string): string;

    function ConvertNormalData(sData: string; cdtSource, cdtDest: TConvertDataType): Boolean;
  end;

{ TSupportConvert }  
  TSupportConvert = record
    sSourceName: WideString;
    sDestName: WideString;
    cdtSource: TConvertDataType;
    cdtDest: TConvertDataType;
    bSourceToDest: Boolean;
    bDestToSource: Boolean;
  end;

const
  C_as_SupportedConvert: array[0..9] of TSupportConvert =
      ((sSourceName: '毫米    ';  sDestName:'Excel像素 '; cdtSource: cdtMM;   cdtDest: cdtExcelPels; bSourceToDest: True; bDestToSource: True),
       (sSourceName: '毫米    ';  sDestName:'英寸      '; cdtSource: cdtMM;   cdtDest: cdtInch;      bSourceToDest: True; bDestToSource: True),
       (sSourceName: '一般字符';  sDestName:'UniCode   '; cdtSource: cdtChar; cdtDest: cdtUnicode;   bSourceToDest: True; bDestToSource: True),
       (sSourceName: '汉字    ';  sDestName:'拼音首字母'; cdtSource: cdtHZ;   cdtDest: cdtPyHead;    bSourceToDest: True; bDestToSource: False),
       (sSourceName: '一般字符';  sDestName:'Base64编码'; cdtSource: cdtChar; cdtDest: cdtBase64;    bSourceToDest: True; bDestToSource: True),
       (sSourceName: '一般字符';  sDestName:'XUM编码   '; cdtSource: cdtChar; cdtDest: cdtXUE;       bSourceToDest: True; bDestToSource: True),
       (sSourceName: '一般字符';  sDestName:'MD5加密   '; cdtSource: cdtChar; cdtDest: cdtMD5;       bSourceToDest: True; bDestToSource: False),
       (sSourceName: '一般字符';  sDestName:'DES加密   '; cdtSource: cdtChar; cdtDest: cdtDES;       bSourceToDest: True; bDestToSource: True),
       (sSourceName: '一般字符';  sDestName:'SHA1      '; cdtSource: cdtChar; cdtDest: cdtSHA1;      bSourceToDest: True; bDestToSource: True),
       (sSourceName: '一般字符';  sDestName:'HMAC_MD5  '; cdtSource: cdtChar; cdtDest: cdtHMAC_MD5;  bSourceToDest: True; bDestToSource: False)
      );

var
  Convertor: TConvertor;
implementation

uses
  Math, U_Base64Codec, U_md5, CnDES;

function TConvertor.MM2Pels(fdMM: Double; mpmMoudle: TMM2PelModule): Double;
begin
  case mpmMoudle of
    mpmCol:
      Result := fdMM * C_colModulus;
    mpmRow:
      Result := fdMM * C_rowModulus;
    mpmInch:
      Result := fdMM * C_MM2Inch;
    else
      Result := fdMM;
  end;
end;  

function TConvertor.Pels2MM(fdPels: Double;
  mpmMoudle: TMM2PelModule): Double;
begin
  case mpmMoudle of
    mpmCol:
      Result := fdPels / C_colModulus;
    mpmRow:
      Result := fdPels / C_rowModulus;
    mpmInch:
      Result := fdPels / C_MM2Inch;
    else
      Result := fdPels;
  end;
end;     

function TConvertor.Hex2Dec(Hex: string): string;
var
  i, nPower, nOneDec, nResult: Integer;
  function getDecByOneHex(cHex: Char): Integer;
  begin
    case cHex of
      '0'..'9': Result := StrToInt(cHex);
      'A','a': Result := 10;
      'B','b': Result := 11;
      'C','c': Result := 12;
      'D','d': Result := 13;
      'E','e': Result := 14;
      'F','f': Result := 15;
      else
        Result := -1;
    end;
  end;
  // 由于给delphi的Power方法传整数时老提示Integer和Extended不兼容所以写这个
  function MyPower(base, thepower: Integer): Integer;
  begin
    case thepower of
    0: Result := 1;
    1: Result := base;
    else
      Result := base * MyPower(base, thepower -1);
    end;
  end;
begin
  nResult := 0;
  i := Length(Hex);
  nPower := 0;
  while i>0 do
  begin
    nOneDec := getDecByOneHex(Hex[i]);
    nResult := nResult + nOneDec * MyPower(16, nPower);
    Dec(i);
    Inc(nPower);
  end;
  Result := IntToStr(nResult);
end;   

function TConvertor.Dec2Hex(Dec: string): string;
var
  nDec: Int64; // double?
  nMod: Integer;
  sResult: string;

  function getHexByOneDec(nDec: Integer): string;
  begin
    case nDec of
    15: Result := 'F';
    14: Result := 'E';
    13: Result := 'D';
    12: Result := 'C';
    11: Result := 'B';
    10: Result := 'A';
    else
      Result := IntToStr(nDec);
    end;
  end;
begin
  nDec := StrToInt64(Dec);
  while nDec <> 0 do
  begin
    nMod :=(nDec mod 16);
    sResult := getHexByOneDec(nMod) + sResult;
    nDec := Trunc(nDec / 16);
  end;
  Result := sResult;
end;

function TConvertor.Hz2Unicode(Hz:string;sPrefix: string; bHex: Boolean):string;
var
  s: string;
  i, j, k: integer;
  a: array [1..1000] of char;
  // 某些字符编码16进制不足2位，需要补全到2位再连接成字符串。在Format中指定成2位再填充
  function fillzero(s: string): string;
  var
    index: Integer;
  begin
    Result := s;
    for index := 0 to Length(Result) do
      if Result[index] = ' ' then
        Result[index] := '0';
  end;  
begin
  s:='';
  StringToWideChar(Hz, @(a[1]), 1000);
//  Result := a;
  i:=1;
  while ((a[i]<>#0) or (a[i+1]<>#0)) do
  begin
    j:=Integer(a[i]);
    k:=Integer(a[i+1]);
    // 显示为16进制
    if bHex then
      s:=s +sPrefix+fillzero(Format('%0:2X%1:2X',[k, j])) //Copy(Format('%X',[k*$100+j+$10000]) ,2,4);
    else
      s:=s +sPrefix+Hex2Dec(fillzero(Format('%0:2X%1:2X',[k, j]))); //Copy(Format('%X',[k*$100+j+$10000]) ,2,4);

    i:=i+2;
  end;
  Result:=s;
//  Result := IntToStr(Ord(WideString(Hz)[1]));    // 此方式可实现转换
end;

function TConvertor.Unicode2Hz(UniCode: string;sPrefix: string; bHex: Boolean): string;
var
  sOneUnicode: string;
//  bWithPrefix: Boolean;
  acOneUnicode: array[1..4] of char;
  sLeftUnicode: string;
//  i: Integer;
  nIndex, nOneCodeLen: Integer;
//  slstUnicode: TStrings;
begin
//  if Copy(UniCode, 1, Length(sPrefix)) = sPrefix then
//    bWithPrefix := True
//  else
//    bWithPrefix := False;

//  slstUnicode := TStringList.Create;
//  try
//    commonFunc.Split(UniCode, sPrefix, slstUnicode);
    if bHex then
      nOneCodeLen := 4
    else
      nOneCodeLen := 5;
//    i := 0;
    sLeftUnicode := UniCode;
    while True do
    begin
      nIndex := fStrUtil.PosFrom(sPrefix, sLeftUnicode);
      if nIndex = 0 then
      begin
        Break;
      end;

      sOneUnicode := Copy(sLeftUnicode, nIndex+Length(sPrefix), nOneCodeLen);  //slstUnicode[i]; //getOneUnicode(i, bWithPrefix);
      // 十进制 先转换到十六进制....
      if not bHex then
      begin
        sOneUnicode := Dec2Hex(sOneUnicode);
      end;  

      Result := Result + Copy(sLeftUnicode, 1, nIndex-1);
      sLeftUnicode := Copy(sLeftUnicode, nIndex+Length(sPrefix) + nOneCodeLen, MaxInt);
      // 后2位在前，前2位在后
      acOneUnicode[1] := Char(StrToInt(Hex2Dec(Copy(sOneUnicode, 3, 2))));
      acOneUnicode[2] := Char(StrToInt(Hex2Dec(Copy(sOneUnicode, 1, 2))));
      acOneUnicode[3] := #0;
      acOneUnicode[4] := #0;
      Result := Result + WideCharToString(@acOneUnicode[1]);
//      Inc(i);
    end;
//  finally
//    slstUnicode.Free;
//  end;
end;

// 使"中文名称"输入翻译成"简码"  如"工资处"译成"GZC"
function TConvertor.Hz2PySingle(const AHzStr: string): string;
const
  ChinaCode: array[0..25, 0..1] of Integer = ((1601, 1636), (1637, 1832), (1833, 2077),
    (2078, 2273), (2274, 2301), (2302, 2432), (2433, 2593), (2594, 2786), (9999, 0000),
    (2787, 3105), (3106, 3211), (3212, 3471), (3472, 3634), (3635, 3722), (3723, 3729),
    (3730, 3857), (3858, 4026), (4027, 4085), (4086, 4389), (4390, 4557), (9999, 0000),
    (9999, 0000), (4558, 4683), (4684, 4924), (4925, 5248), (5249, 5589));
var
  i, j, HzOrd: integer;
  bFind: Boolean;
//  Hz: string[2]; 
begin
  i := 1;
  while i <= Length(AHzStr) do
  begin
    if (AHzStr[i] >= #160) and (AHzStr[i + 1] >= #160) then
    begin 
      HzOrd := (Ord(AHzStr[i]) - 160) * 100 + Ord(AHzStr[i + 1]) - 160;
      bFind := False;
      for j := 0 to 25 do 
      begin 
        if (HzOrd >= ChinaCode[j][0]) and (HzOrd <= ChinaCode[j][1]) then 
        begin 
          Result := Result + char(byte('A') + j);
          bFind := True;
          Break;
        end; 
      end;
      // 如果无法找到，则返回原来字的字符  比如汉字逗号 适用这种情况
      if not bFind then
        Result := Result + AHzStr[i] + AHzStr[i+1];
      Inc(i, 2);
    end
    else
    begin  
      // 英文字符 直接返回
      Result := Result + AHzStr[i];  
      Inc(i,1);
    end;
  end; 
end;

  /////////////////////////////////////// 
// 这个函数用户识别单独汉字的简码 字符串的简码函数请自行制作
function TConvertor.GetPYIndexChar(hzchar:string):char;
begin 
  case WORD(hzchar[1]) shl 8 + WORD(hzchar[2]) of
    $B0A1..$B0C4 : result := 'A';
    $B0C5..$B2C0 : result := 'B'; 
    $B2C1..$B4ED : result := 'C';
    $B4EE..$B6E9 : result := 'D';
    $B6EA..$B7A1 : result := 'E'; 
    $B7A2..$B8C0 : result := 'F';
    $B8C1..$B9FD : result := 'G'; 
    $B9FE..$BBF6 : result := 'H'; 
    $BBF7..$BFA5 : result := 'J'; 
    $BFA6..$C0AB : result := 'K'; 
    $C0AC..$C2E7 : result := 'L'; 
    $C2E8..$C4C2 : result := 'M'; 
    $C4C3..$C5B5 : result := 'N'; 
    $C5B6..$C5BD : result := 'O'; 
    $C5BE..$C6D9 : result := 'P'; 
    $C6DA..$C8BA : result := 'Q'; 
    $C8BB..$C8F5 : result := 'R';
    $C8F6..$CBF9 : result := 'S'; 
    $CBFA..$CDD9 : result := 'T'; 
    $CDDA..$CEF3 : result := 'W'; 
    $CEF4..$D188 : result := 'X'; 
    $D1B9..$D4D0 : result := 'Y'; 
    $D4D1..$D7F9 : result := 'Z'; 
  else
    result := char(0);
  end;
end;


function TConvertor.Base642Normal(const s: string): string;
begin
  Result := Base64Decode(s);
end;

function TConvertor.Normal2Base64(const s: string): string;
begin
  Result := Base64Encode(s);
end;   

//对str进行简单编码，编码口令
function TConvertor.XXEUUEMIMEEncode(str:string):string;
var
  enMime:TidEncoderMime;
  enUUE:TidEncoderUUE;
  enXXE:TidEncoderXXE;
  sRes:string;
begin
  sRes:='';
  enMime:=TidEncoderMime.Create(nil);
  enUUE:=TidEncoderUUE.Create(nil);
  enXXE:=TidEncoderXXE.Create(nil);
  try
    sRes:=enXXE.EncodeString(str);
    sRes:=enUUE.EncodeString(sRes);
    sRes:=enMime.EncodeString(sRes);
    //sRes:=enMime.EncodeString(enUUE.EncodeString(enXXE.EncodeString(str)));
  finally
    enMime.Free;
    enUUE.Free;
    enXXE.Free;
  end;
  result:=sRes;
end;

//对str进行解码
function TConvertor.XXEUUEMIMEDecode(str:string):string;
var
  deMime:TidDecoderMime;
  deUUE:TidDecoderUUE;
  deXXE:TidDecoderXXE;
  sRes:string;
begin
  sRes:='';
  deMime:=TidDecoderMime.Create(nil);
  deUUE:=TidDecoderUUE.Create(nil);
  deXXE:=TidDecoderXXE.Create(nil);
  try
    sRes:=deMime.DecodeString(str);
    sRes:=deUUE.DecodeString(sRes);
    sRes:=deXXE.DecodeString(sRes);
  finally
    deMime.Free;
    deUUE.Free;
    deXXE.Free;
  end;
  result:=sRes;
end;

function TConvertor.getRecordInfoCallID(nPointerCallID: Integer): string;
var
  pCallid: PCALL_ID;
//  n64bitResult: Int64;
begin
  pCallid := Pointer(nPointerCallID);
//  n64bitResult := Int64(pCallid^.ulTime) * 65536;
//  n64bitResult := n64bitResult + pCallid^.usDsn;
//  n64bitResult := n64bitResult * 256;
//  n64bitResult := n64bitResult + pCallid^.ucHandle;
//  n64bitResult := n64bitResult * 256;
//  n64bitResult := n64bitResult + pCallid^.ucServer;
  Result := IntToStr((((Int64(pCallid^.ulTime) * 65536) + pCallid^.usDsn) * 256
      + pCallid^.ucHandle) * 256 + pCallid^.ucServer);
end;

function TConvertor.getCallIDAtPlatFormVer31(nPointerCallID: Integer): string;
var
  pCallid: PCALL_ID;
begin
  pCallid := Pointer(nPointerCallID);
  Result := IntToStr(((pCallid^.ucServer * 256) + pCallid^.ucHandle) * 65536 +
      (pCallid^.usDsn Mod 256) * 256 + (pCallid^.usDsn Div 256));
end;

function TConvertor.getCallIDAtPlatFormVer32(nPointerCallID: Integer): string;
var
  pCallid: PCALL_ID;
begin
  pCallid := Pointer(nPointerCallID);
  Result := IntToStr(((pCallid^.usDsn * 256) + pCallid^.ucHandle) * 256 +
      pCallid^.ucServer);
end;

function TConvertor.BillLogCallIDVer31ToRecordInfoCallID(
  sBillLogCallID: string): string;
var
  nCallID: Int64;
  pCallID: PCALL_ID;
begin
{*  BillLogCallID的组成方式 ((ucServer * 256) + ucHandle) * 65536 + (usDsn Mod 256) * 256 + (usDsn Div 256)，参照getCallIDAtPlatFormVer31
 * 高4位为 ucServer 和ucHandle  低4位为 usDsn
 }
  New(pCallID);
  try
    pCallID^.ulTime := StrToInt64(Copy(sBillLogCallID, 1, Pos('-', sBillLogCallID)-1));
    nCallID := StrToInt64(Copy(sBillLogCallID, Pos('-', sBillLogCallID)+1, Length(sBillLogCallID))); // 取到-后边的部分
    // 计算usDsn
    pCallID^.usDsn := nCallID mod 65536;   // 取低4位
//    pCallID^.usDsn := (pCallID^.usDsn mod 256) * 256 + pCallID^.usDsn div 256;  // 4位16进制数字前两位和后两位掉换
    // 计算ucServer 和ucHandle
    nCallID := Ceil(nCallID div 65536);
    pCallID^.ucHandle := nCallID mod 256;  // ucHandle后两位
    pCallID^.ucServer := nCallID div 256;  // ucServer前两位

//    pCallID^.ulTime := 1142388732;
//    pCallID^.usDsn := 11608;
//    pCallID^.ucHandle := 157;
//    pCallID^.ucServer := 0;

    Result := getRecordInfoCallID(Integer(pCallID));
  finally
    Dispose(pCallID)
  end;
end;

function TConvertor.BillLogCallIDVer32ToRecordInfoCallID(
  sBillLogCallID: string): string;   
var
  nCallID: Integer;
  pCallID: PCALL_ID;
begin
{*  BillLogCallID的组成方式 ((usDsn * 256) + ucHandle) * 256 + ucServer，参照getCallIDAtPlatFormVer32
 *  高4位为 usDsn 低4位为 ucHandle 和  ucServer
 }
  New(pCallID);
  try            
    pCallID^.ulTime := StrToInt64(Copy(sBillLogCallID, 1, Pos('-', sBillLogCallID)-1));
    nCallID := StrToInt(Copy(sBillLogCallID, Pos('-', sBillLogCallID)+1, Length(sBillLogCallID))); // 取到-后边的部分
    // 计算usDsn
    pCallID^.usDsn := nCallID div 65536;      // 取高4位
    pCallID^.usDsn := (pCallID^.usDsn mod 256) * 256 + pCallID^.usDsn div 256;  // 4位16进制数字前两位和后两位掉换
    // 计算ucHandle 和 ucServer
    nCallID := nCallID mod 65536;
    pCallID^.ucHandle := nCallID div 256;  // ucHandle前两位
    pCallID^.ucServer := nCallID mod 256;  // ucServer后两位
    Result := getRecordInfoCallID(Integer(pCallID));
  finally
    Dispose(pCallID)
  end;
end;

function TConvertor.RecordInfoCallIDToBillLogCallIDVer31(
  sRecordInfoCallID: string): string;
var
  nCallID: Int64;
  pCallID: PCALL_ID;
begin
  New(pCallID);
  try
    nCallID := StrToInt64(sRecordInfoCallID);
    // ucServer
    pCallID^.ucServer := nCallID mod 256;
    nCallID := nCallID div 256;

    pCallID^.ucHandle := nCallID mod 256;
    nCallID := nCallID div 256;

    pCallID^.ucHandle := nCallID mod 65536;
    nCallID := nCallID div 65536;

    pCallID^.ulTime := nCallID;
    Result := getCallIDAtPlatFormVer31(Integer(pCallID));
  finally
    Dispose(pCallID)
  end;
end;

function TConvertor.RecordInfoCallIDToBillLogCallIDVer32(
  sRecordInfoCallID: string): string;
var
  nCallID: Int64;
  pCallID: PCALL_ID;
begin
  New(pCallID);
  try
    nCallID := StrToInt64(sRecordInfoCallID);
    // ucServer
    pCallID^.ucServer := nCallID mod 256;
    nCallID := nCallID div 256;

    pCallID^.ucHandle := nCallID mod 256;
    nCallID := nCallID div 256;

    pCallID^.ucHandle := nCallID mod 65536;
    nCallID := nCallID div 65536;

    pCallID^.ulTime := nCallID;    
    Result := getCallIDAtPlatFormVer32(Integer(pCallID));
  finally
    Dispose(pCallID)
  end;
end;

function TConvertor.MD5Decode(str: string): string;
begin
  Result := MD5Print(MD5String(str));
end;

function TConvertor.MD5Encode(str: string): string;
begin
  Result := MD5Print(MD5String(str));
//  Result := MD5.MD5Encode(str);
end;

function TConvertor.DESEncode(str, key:string):string;
begin
  Result := CnDES.DESEncryptStr(str, key);
end;

function TConvertor.DESDecode(str, key:string):string;
begin
  Result := CnDES.DESDecryptStr(str, key);
end;  

function TConvertor.ConvertNormalData(sData: string; cdtSource,
  cdtDest: TConvertDataType): Boolean;
begin
  LastConvertResult := '';
  Result := False;
  case cdtSource of
  cdtMM:
  begin
    case cdtDest of
    cdtExcelPels:
    begin
      LastConvertResult := FloatToStr(MM2Pels(StrToFloat(sData), MM2PelModule));
      Result := True;
    end;
    cdtInch:
    begin
      LastConvertResult := FloatToStr(MM2Pels(StrToFloat(sData), mpmInch));
      Result := True;
    end;
    end;
  end;
  cdtExcelPels:
  begin 
    case cdtDest of
    cdtMM:
    begin
      LastConvertResult := FloatToStr(Pels2MM(StrToFloat(sData), MM2PelModule));
      Result := True;
    end;
    end;
  end;
  cdtInch:
  begin
    case cdtDest of
    cdtMM:
    begin  
      LastConvertResult := FloatToStr(Pels2MM(StrToFloat(sData), mpmInch));
      Result := True;
    end;
    end;  
  end;  
  cdtHZ, cdtChar:
  begin
    case cdtSource of
    cdtHZ:
      case cdtDest of
//      cdtUnicode:
//      begin
//        LastConvertResult := Hz2Unicode(sData, UniCodePrefix, UniCodeHex);
//        Result := True;
//      end;
      cdtPyHead:
      begin
        LastConvertResult := Hz2PySingle(sData);
        Result := True;
      end;
      end;
    cdtChar:
      case cdtDest of
      cdtUnicode:
      begin
        LastConvertResult := Hz2Unicode(sData, UniCodePrefix, UniCodeHex);
        Result := True;
      end;
      cdtBase64:
      begin
        LastConvertResult := Normal2Base64(sData);
        Result := True;
      end;
      cdtXUE:
      begin
        LastConvertResult := XXEUUEMIMEEncode(sData);
        Result := True;
      end;
      cdtMD5:
      begin
        LastConvertResult := MD5Encode(sData);
        Result := True;
      end; 
      cdtDES:
      begin
        LastConvertResult := DESEncode(sData, NormalKey);
        Result := True;
      end;
      cdtSHA1:
      begin
        LastConvertResult := uSHA1.HMAC_SHA1(sData, NormalKey);
        Result := True;
      end;
      cdtHMAC_MD5:
      begin
        LastConvertResult := uMD5.HMAC_MD5(sData, NormalKey);
        Result := True;
      end;    
      end;
    end;
  end;
  cdtUnicode:
  begin
    case cdtDest of
    cdtChar:
    begin
      LastConvertResult := Unicode2Hz(sData, UniCodePrefix, UniCodeHex);
      Result := True;
    end;
    end;
  end;
  cdtPyHead:  
  begin
    // 拼音首字母不能转
  end;
  cdtBase64:  
  begin
    case cdtDest of
    cdtHZ, cdtChar:
    begin
      LastConvertResult := Base642Normal(sData);
      Result := True;
    end;
    end;
  end;
  cdtXUE: 
  begin
    case cdtDest of
    cdtHZ, cdtChar:
    begin
      LastConvertResult := XXEUUEMIMEDecode(sData);
      Result := True;
    end;
    end;
  end;
  cdtMD5:   
  begin
    // md5到一般字符，不明转法，虽然据说md5不可逆，但已知一些md5可被转换回
  end;
  cdtICDBilllogCallId:   
  begin
    // billlog的callid
  end;
  cdtICDRecordInfoCallid:  
  begin
    // recordinfo的callid
  end;
  cdtDES:
  begin
    case cdtDest of
      cdtChar:
      begin    
        LastConvertResult := DESDecode(sData, NormalKey);
        Result := True;
      end;
    end;
  end;    
  end;
end;

initialization
  Convertor := TConvertor.Create;

finalization
  if Assigned(Convertor) then
  FreeAndNil(Convertor);

end.
