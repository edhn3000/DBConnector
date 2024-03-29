unit HMAC;

{General HMAC unit}


interface

(*************************************************************************

 DESCRIPTION     :  General HMAC (hash message authentication) unit

 REQUIREMENTS    :  TP5-7, D1-D7/9, FPC, VP

 EXTERNAL DATA   :  ---

 MEMORY USAGE    :  ---

 DISPLAY MODE    :  ---

 REFERENCES      :  - HMAC: Keyed-Hashing for Message Authentication
                      (http://www.faqs.org/rfcs/rfc2104.html)
                    - The Keyed-Hash Message Authentication Code (HMAC)
                      http://csrc.nist.gov/publications/fips/fips198/fips-198a.pdf


 Version  Date      Author      Modification
 -------  --------  -------     ------------------------------------------
 0.10     15.01.06  W.Ehrhardt  Initial version based on HMACWHIR
**************************************************************************)

(*-------------------------------------------------------------------------
 (C) Copyright 2006 Wolfgang Ehrhardt

 This software is provided 'as-is', without any express or implied warranty.
 In no event will the authors be held liable for any damages arising from
 the use of this software.

 Permission is granted to anyone to use this software for any purpose,
 including commercial applications, and to alter it and redistribute it
 freely, subject to the following restrictions:

 1. The origin of this software must not be misrepresented; you must not
    claim that you wrote the original software. If you use this software in
    a product, an acknowledgment in the product documentation would be
    appreciated but is not required.

 2. Altered source versions must be plainly marked as such, and must not be
    misrepresented as being the original software.

 3. This notice may not be removed or altered from any source distribution.
----------------------------------------------------------------------------*)

uses hash;

type
  THMAC_Context = record
                    hashctx: THashContext;
                    hmacbuf: THashBuffer;
                    phashd : PHashDesc;
                  end;
type
  THMAC_string  = string[255];

procedure hmac_init(var ctx: THMAC_Context; phash: PHashDesc; key: pointer; klen: word);
  {-initialize HMAC context with hash descr phash^ and key}
  {$ifdef DLL} stdcall; {$endif}

procedure hmac_inits(var ctx: THMAC_Context; phash: PHashDesc; skey: THMAC_string);
  {-initialize HMAC context with hash descr phash^ and skey}
  {$ifdef DLL} stdcall; {$endif}

procedure hmac_update(var ctx: THMAC_Context; data: pointer; dlen: word);
  {-HMAC data input, may be called more than once}
  {$ifdef DLL} stdcall; {$endif}

procedure hmac_updateXL(var ctx: THMAC_Context; data: pointer; dlen: longint);
  {-HMAC data input, may be called more than once}
  {$ifdef DLL} stdcall; {$endif}

procedure hmac_final(var ctx: THMAC_Context; var mac: THashDigest);
  {-end data input, calculate HMAC digest}
  {$ifdef DLL} stdcall; {$endif}

  function DoHMAC(sHash: string; nHashLen: Integer; sKey, sData: string): string;
  function HMAC_MD5(sKey, sData: string): string;


implementation

uses
  SysUtils;

function HexByte(b: byte): string;
  {-byte as hex string}
const
  nib: array[0..15] of char = '0123456789abcdef';
begin
  Result := nib[b div 16] + nib[b and 15];
end;

function HexStr(psrc: pointer; L: integer): string;
  {-hex string of memory block of length L pointed by psrc}
var
  i: integer;
  s: string;
begin
  s := '';
  if psrc<>nil then begin
    for i:=0 to L-1 do begin
      s := s + HexByte(pByte(psrc)^);
      inc(longint(psrc));
    end;
  end;
  Result := s;
end;

function DoHMAC(sHash: string; nHashLen: Integer; sKey, sData: string): string;
var
  ctx  : THMAC_Context;
  mac  : THashDigest;
  phash: PHashDesc;
begin
  New(phash);
  RegisterHash(_MD5, phash);
  phash := FindHash_by_Name(sHash);
  if phash = nil then
  begin
    raise Exception.Create(Format('hash not supported  [%s]', [sHash]));
  end;  
  hmac_inits(ctx, phash, sKey);
  hmac_update(ctx, @sData, sizeof(sData));
  hmac_final(ctx, mac);
  Result := HexStr(@mac, nHashLen);
end;  

function HMAC_MD5(sKey, sData: string): string;
var
  ctx  : THMAC_Context;
  mac  : THashDigest;
  phash: PHashDesc;
begin
  phash := FindHash_by_ID(_MD5);
  if phash = nil then
  begin
    raise Exception.Create(Format('hash not supported ', []));
  end;  
  hmac_inits(ctx, phash, sKey);
  hmac_update(ctx, @sData, sizeof(sData));
  hmac_final(ctx, mac);
  Result := HexStr(@mac, sizeof(TMD5Digest));
end;

{---------------------------------------------------------------------------}
procedure hmac_init(var ctx: THMAC_Context; phash: PHashDesc; key: pointer; klen: word);
  {-initialize HMAC context with hash descr phash^ and key}
var
  i,lk,bl: word;
  kb: THashDigest;
begin
  fillchar(ctx, sizeof(ctx),0);
  if phash<>nil then with ctx do begin
    phashd := phash;
    lk := klen;
    bl := phash^.HBlockLen;
    if lk > bl then begin
      {Hash if key length > block length}
      HashFullXL(phash, kb, key, lk);
      lk := phash^.HDigestLen;
      move(kb, hmacbuf, lk);
    end
    else move(key^, hmacbuf, lk);
    {XOR with ipad}
    for i:=0 to bl-1 do hmacbuf[i] := hmacbuf[i] xor $36;
    {start inner hash}
    phash^.HInit(hashctx);
    phash^.HUpdateXL(hashctx, @hmacbuf, bl);
  end;
end;


{---------------------------------------------------------------------------}
procedure hmac_inits(var ctx: THMAC_Context; phash: PHashDesc; skey: THMAC_string);
  {-initialize HMAC context with hash descr phash^ and skey}
begin
  if phash<>nil then hmac_init(ctx, phash, @skey[1], length(skey));
end;


{---------------------------------------------------------------------------}
procedure hmac_update(var ctx: THMAC_Context; data: pointer; dlen: word);
  {-HMAC data input, may be called more than once}
begin
  with ctx do begin
    if phashd<>nil then phashd^.HUpdateXL(hashctx, data, dlen);
  end;
end;


{---------------------------------------------------------------------------}
procedure hmac_updateXL(var ctx: THMAC_Context; data: pointer; dlen: longint);
  {-HMAC data input, may be called more than once}
begin
  with ctx do begin
    if phashd<>nil then phashd^.HUpdateXL(hashctx, data, dlen);
  end;
end;


{---------------------------------------------------------------------------}
procedure hmac_final(var ctx: THMAC_Context; var mac: THashDigest);
  {-end data input, calculate HMAC digest}
var
  i: integer;
  bl: word;
begin
  with ctx do if phashd<>nil then begin
    bl := phashd^.HBlockLen;
    {complete inner hash}
    phashd^.HFinal(hashctx, mac);
    {remove ipad from buf, XOR opad}
    for i:=0 to bl-1 do hmacbuf[i] := hmacbuf[i] xor ($36 xor $5c);
    {outer hash}
    phashd^.HInit(hashctx);
    phashd^.HUpdateXL(hashctx, @hmacbuf, bl);
    phashd^.HUpdateXL(hashctx, @mac, phashd^.HDigestLen);
    phashd^.HFinal(hashctx, mac);
  end;
end;

end.

