// ============================================================================
// D3-implementation of "RSA Data Security, Inc. MD5 Message-Digest Algorithm"
// Copyright (c) 1999, Juergen Haible.
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to
// deal in the Software without restriction, including without limitation the
// rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
// sell copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
// IN THE SOFTWARE.
// ============================================================================

unit uMD5; // The MD5 Message-Digest Algorithm (RFC1321)

{------------------------------------------------------------------------------

   RSA Data Security, Inc., MD5 message-digest algorithm

   Copyright (C) 1991-2, RSA Data Security, Inc. Created 1991. All rights
   reserved.

   License to copy and use this software is granted provided that it is
   identified as the "RSA Data Security, Inc. MD5 Message-Digest Algorithm"
   in all material mentioning or referencing this software or this function.

   License is also granted to make and use derivative works provided that
   such works are identified as "derived from the RSA Data Security, Inc.
   MD5 Message-Digest Algorithm" in all material mentioning or referencing
   the derived work.

   RSA Data Security, Inc. makes no representations concerning either the
   merchantability of this software or the suitability of this software for
   any particular purpose. It is provided "as is" without express or implied
   warranty of any kind.

   These notices must be retained in any copies of any part of this
   documentation and/or software.

------------------------------------------------------------------------------}

interface

//{$INCLUDE Compiler.inc}

{$R-}
{$Q-}

uses SysUtils, Classes;

const
   MD5HashSize = 16;

type
   Chars64      = array[0..63] of Char;
   LongInts4    = array[0.. 3] of LongInt;
   LongInts2    = array[0.. 2] of LongInt;

   // MD5 context.
   MD5_CTX = record
      state : LongInts4; // state (ABCD)
      count : LongInts2; // number of bits, modulo 2^64 (lsb first)
      buffer: Chars64;   // input buffer
   end;

   MD5_DIGEST = array[0..MD5HashSize-1] of Char;
   MD5DigestString = String;  // string containing 16 chars

procedure MD5Init  ( var context: MD5_CTX);
procedure MD5Update( var context: MD5_CTX; const input; inputLen: LongInt );
procedure MD5Final ( var digest: MD5_DIGEST; var context: MD5_CTX );

function MD5ofStr   ( const s: String            ): MD5DigestString;
function MD5ofBuf   ( const buf; buflen: Integer ): MD5DigestString;
function MD5ofStream( const strm: TStream        ): MD5DigestString;

function MD5toHex( const digest: MD5DigestString ): String;

procedure HMAC_MD5( const Data; DataLen: Integer;
                    const Key;  KeyLen : Integer;
                    out   Digest       : MD5_DIGEST ); overload;
procedure HMAC_MD5( const Data: String;
                    const Key : String;
                    out   DigestStr: MD5DigestString ); overload;
function  HMAC_MD5( const Data: String;
                    const Key : String ): String; overload;

// ----------------------------------------------------------------------------

implementation

const
   // Constants for MD5Transform routine.
   S11 = 7; S12 = 12; S13 = 17; S14 = 22;
   S21 = 5; S22 =  9; S23 = 14; S24 = 20;
   S31 = 4; S32 = 11; S33 = 16; S34 = 23;
   S41 = 6; S42 = 10; S43 = 15; S44 = 21;

   PADDING: Chars64 = (
      #$80, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0,
        #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0,
        #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0,
        #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0, #0
   );

// ----------------------------------------------------------------------------

// F, G, H and I are basic MD5 functions.

function F(x, y, z: LongInt): LongInt;
begin
   Result := (((x) and (y)) or ((not x) and (z)))
end;

function G(x, y, z: LongInt): LongInt;
begin
   Result := (((x) and (z)) or ((y) and (not z)))
end;

function H(x, y, z: LongInt): LongInt;
begin
   Result := ((x) xor (y) xor (z))
end;

function I(x, y, z: LongInt): LongInt;
begin
   Result := ((y) xor ((x) or (not z)))
end;


// ROTATE_LEFT rotates x left n bits.

function ROTATE_LEFT(x, n: LongInt): LongInt;
begin
   Result := (((x) shl (n)) or ((x) shr (32-(n))))
end;


// FF, GG, HH, and II transformations for rounds 1, 2, 3, and 4.
// Rotation is separate from addition to prevent recomputation.

procedure FF(var a: LongInt; b, c, d, x, s, ac: LongInt);
begin
   a := a + F ((b), (c), (d)) + (x) + ac;
   a := ROTATE_LEFT ((a), (s));
   a := a + (b);
end;

procedure GG(var a: LongInt; b, c, d, x, s, ac: LongInt);
begin
   a := a + G ((b), (c), (d)) + (x) + (ac);
   a := ROTATE_LEFT ((a), (s));
   a := a + (b);
end;

procedure HH(var a: LongInt; b, c, d, x, s, ac: LongInt);
begin
   a := a + H ((b), (c), (d)) + (x) + (ac);
   a := ROTATE_LEFT ((a), (s));
   a := a + (b);
end;

procedure II(var a: LongInt; b, c, d, x, s, ac: LongInt);
begin
   a := a + I ((b), (c), (d)) + (x) + (ac);
   a := ROTATE_LEFT ((a), (s));
   a := a + (b);
end;


// Decodes input (unsigned char) into output (UINT4). Assumes len is a multiple of 4.

procedure Decode( var output: array of LongInt; const input: array of Char; len: Integer );
var  i, j: LongInt;
begin
   i := 0;
   j := 0;
   while j < len do begin
      output[i] := (Ord(input[j  ])       )
                or (Ord(input[j+1]) shl  8)
                or (Ord(input[j+2]) shl 16)
                or (Ord(input[j+3]) shl 24);
      inc( i );
      inc( j, 4 );
   end;
end;


// Encodes input (UINT4) into output (unsigned char). Assumes len is a multiple of 4.

procedure Encode( var output: array of Char; const input: array of LongInt; len: LongInt );
var  i, j: Integer;
begin
   i := 0;
   j := 0;
   while ( j < len ) do begin
      output[j]   := chr( ((input[i]       ) and $ff) );
      output[j+1] := chr( ((input[i] shr  8) and $ff) );
      output[j+2] := chr( ((input[i] shr 16) and $ff) );
      output[j+3] := chr( ((input[i] shr 24) and $ff) );
      inc( i );
      inc( j, 4 );
   end;
end;

// ----------------------------------------------------------------------------

// MD5 initialization. Begins an MD5 operation, writing a new context.

procedure MD5Init(var context: MD5_CTX);
begin
  context.count[0] := 0;
  context.count[1] := 0;

  // Load magic initialization constants.
  context.state[0] := LongInt( $67452301 );
  context.state[1] := LongInt( $efcdab89 );
  context.state[2] := LongInt( $98badcfe );
  context.state[3] := LongInt( $10325476 );
end;


// MD5 basic transformation. Transforms state based on block.

procedure MD5Transform( var state: LongInts4; const block: array of Char );
var  a, b, c, d: LongInt;
     x: array[0..15] of LongInt;
begin
   a := state[0];
   b := state[1];
   c := state[2];
   d := state[3];

   Decode( x, block, 64 );

   // Round 1
   FF (a, b, c, d, x[ 0], S11, LongInt( $d76aa478 )); // 1
   FF (d, a, b, c, x[ 1], S12, LongInt( $e8c7b756 )); // 2
   FF (c, d, a, b, x[ 2], S13, LongInt( $242070db )); // 3
   FF (b, c, d, a, x[ 3], S14, LongInt( $c1bdceee )); // 4
   FF (a, b, c, d, x[ 4], S11, LongInt( $f57c0faf )); // 5
   FF (d, a, b, c, x[ 5], S12, LongInt( $4787c62a )); // 6
   FF (c, d, a, b, x[ 6], S13, LongInt( $a8304613 )); // 7
   FF (b, c, d, a, x[ 7], S14, LongInt( $fd469501 )); // 8
   FF (a, b, c, d, x[ 8], S11, LongInt( $698098d8 )); // 9
   FF (d, a, b, c, x[ 9], S12, LongInt( $8b44f7af )); // 10
   FF (c, d, a, b, x[10], S13, LongInt( $ffff5bb1 )); // 11
   FF (b, c, d, a, x[11], S14, LongInt( $895cd7be )); // 12
   FF (a, b, c, d, x[12], S11, LongInt( $6b901122 )); // 13
   FF (d, a, b, c, x[13], S12, LongInt( $fd987193 )); // 14
   FF (c, d, a, b, x[14], S13, LongInt( $a679438e )); // 15
   FF (b, c, d, a, x[15], S14, LongInt( $49b40821 )); // 16

   // Round 2
   GG (a, b, c, d, x[ 1], S21, LongInt( $f61e2562 )); // 17
   GG (d, a, b, c, x[ 6], S22, LongInt( $c040b340 )); // 18
   GG (c, d, a, b, x[11], S23, LongInt( $265e5a51 )); // 19
   GG (b, c, d, a, x[ 0], S24, LongInt( $e9b6c7aa )); // 20
   GG (a, b, c, d, x[ 5], S21, LongInt( $d62f105d )); // 21
   GG (d, a, b, c, x[10], S22, LongInt(  $2441453 )); // 22
   GG (c, d, a, b, x[15], S23, LongInt( $d8a1e681 )); // 23
   GG (b, c, d, a, x[ 4], S24, LongInt( $e7d3fbc8 )); // 24
   GG (a, b, c, d, x[ 9], S21, LongInt( $21e1cde6 )); // 25
   GG (d, a, b, c, x[14], S22, LongInt( $c33707d6 )); // 26
   GG (c, d, a, b, x[ 3], S23, LongInt( $f4d50d87 )); // 27
   GG (b, c, d, a, x[ 8], S24, LongInt( $455a14ed )); // 28
   GG (a, b, c, d, x[13], S21, LongInt( $a9e3e905 )); // 29
   GG (d, a, b, c, x[ 2], S22, LongInt( $fcefa3f8 )); // 30
   GG (c, d, a, b, x[ 7], S23, LongInt( $676f02d9 )); // 31
   GG (b, c, d, a, x[12], S24, LongInt( $8d2a4c8a )); // 32

   // Round 3
   HH (a, b, c, d, x[ 5], S31, LongInt( $fffa3942 )); // 33
   HH (d, a, b, c, x[ 8], S32, LongInt( $8771f681 )); // 34
   HH (c, d, a, b, x[11], S33, LongInt( $6d9d6122 )); // 35
   HH (b, c, d, a, x[14], S34, LongInt( $fde5380c )); // 36
   HH (a, b, c, d, x[ 1], S31, LongInt( $a4beea44 )); // 37
   HH (d, a, b, c, x[ 4], S32, LongInt( $4bdecfa9 )); // 38
   HH (c, d, a, b, x[ 7], S33, LongInt( $f6bb4b60 )); // 39
   HH (b, c, d, a, x[10], S34, LongInt( $bebfbc70 )); // 40
   HH (a, b, c, d, x[13], S31, LongInt( $289b7ec6 )); // 41
   HH (d, a, b, c, x[ 0], S32, LongInt( $eaa127fa )); // 42
   HH (c, d, a, b, x[ 3], S33, LongInt( $d4ef3085 )); // 43
   HH (b, c, d, a, x[ 6], S34, LongInt(  $4881d05 )); // 44
   HH (a, b, c, d, x[ 9], S31, LongInt( $d9d4d039 )); // 45
   HH (d, a, b, c, x[12], S32, LongInt( $e6db99e5 )); // 46
   HH (c, d, a, b, x[15], S33, LongInt( $1fa27cf8 )); // 47
   HH (b, c, d, a, x[ 2], S34, LongInt( $c4ac5665 )); // 48

   // Round 4
   II (a, b, c, d, x[ 0], S41, LongInt( $f4292244 )); // 49
   II (d, a, b, c, x[ 7], S42, LongInt( $432aff97 )); // 50
   II (c, d, a, b, x[14], S43, LongInt( $ab9423a7 )); // 51
   II (b, c, d, a, x[ 5], S44, LongInt( $fc93a039 )); // 52
   II (a, b, c, d, x[12], S41, LongInt( $655b59c3 )); // 53
   II (d, a, b, c, x[ 3], S42, LongInt( $8f0ccc92 )); // 54
   II (c, d, a, b, x[10], S43, LongInt( $ffeff47d )); // 55
   II (b, c, d, a, x[ 1], S44, LongInt( $85845dd1 )); // 56
   II (a, b, c, d, x[ 8], S41, LongInt( $6fa87e4f )); // 57
   II (d, a, b, c, x[15], S42, LongInt( $fe2ce6e0 )); // 58
   II (c, d, a, b, x[ 6], S43, LongInt( $a3014314 )); // 59
   II (b, c, d, a, x[13], S44, LongInt( $4e0811a1 )); // 60
   II (a, b, c, d, x[ 4], S41, LongInt( $f7537e82 )); // 61
   II (d, a, b, c, x[11], S42, LongInt( $bd3af235 )); // 62
   II (c, d, a, b, x[ 2], S43, LongInt( $2ad7d2bb )); // 63
   II (b, c, d, a, x[ 9], S44, LongInt( $eb86d391 )); // 64

   inc( state[0], a );
   inc( state[1], b );
   inc( state[2], c );
   inc( state[3], d );

   // Zeroize sensitive information.
   FillChar( x, sizeof(x), 0 );
end;


// MD5 block update operation. Continues an MD5 message-digest operation,
// processing another message block, and updating the context.

procedure MD5Update( var context: MD5_CTX; const input; inputLen: LongInt );
var  i, index, partLen: LongInt;
begin
   // Compute number of bytes mod 64
   index := ( (context.count[0] shr 3) and $3F );

   // Update number of bits
   inc( context.count[0], inputLen shl 3 );
   if context.count[0] < (inputLen shl 3) then inc( context.count[1] );
   inc( context.count[1], inputLen shr 29 );

   partLen := 64 - index;

   // Transform as many times as possible.
   if inputLen >= partLen then begin
      Move( input, context.buffer[index], partLen );
      MD5Transform( context.state, context.buffer );

      i := partLen;
      while ( i + 63 < inputLen ) do begin
         MD5Transform( context.state, (PChar(@input)+i)^ );
         inc( i, 64 );
      end;

      index := 0;
   end else
      i := 0;

   // Buffer remaining input
   Move( (PChar(@input)+i)^, context.buffer[index], inputLen-i );
end;


// MD5 finalization. Ends an MD5 message-digest operation, writing the
// message digest and zeroizing the context.

procedure MD5Final (var digest: MD5_DIGEST; var context: MD5_CTX);
var  bits: array[0..7] of Char;
     index, padLen: LongInt;
begin
   // Save number of bits
   Encode( bits, context.count, 8 );

   // Pad out to 56 mod 64.
   index := ( (context.count[0] shr 3) and $3f );
   if (index < 56) then padLen := ( 56 - index)
                   else padLen := (120 - index);
   MD5Update( context, PADDING, padLen );

   // Append length (before padding)
   MD5Update( context, bits, 8 );

   // Store state in digest
   Encode( digest, context.state, 16 );

   // Zeroize sensitive information.
   FillChar( context, sizeof(context), 0 );
end;

// ----------------------------------------------------------------------------

// returns MD5 digest of given string

function MD5ofStr( const s: String ): MD5DigestString;
var  context: MD5_CTX;
     digest : MD5_DIGEST;
begin
   MD5Init  ( context );
   MD5Update( context, s[1], length(s) );
   MD5Final ( digest, context );
   SetLength( Result, sizeof(digest) );
   Move( digest, Result[1], sizeof(digest) );
end;


// returns MD5 digest of given buffer

function MD5ofBuf( const buf; buflen: Integer ): MD5DigestString;
var  context: MD5_CTX;
     digest : MD5_DIGEST;
begin
   MD5Init  ( context );
   MD5Update( context, buf, buflen );
   MD5Final ( digest, context );
   SetLength( Result, sizeof(digest) );
   Move( digest, Result[1], sizeof(digest) );
end;


// returns MD5 digest of given stream

function MD5ofStream( const strm: TStream ): MD5DigestString;
var  context: MD5_CTX;
     digest : MD5_DIGEST;
     buf: array[0..4095] of char;
     buflen: Integer;
begin
   MD5Init  ( context );
   strm.Position := 0;
   repeat
      buflen := strm.Read( buf[0], 4096 );
      if buflen>0 then MD5Update( context, buf[0], buflen );
   until buflen<4096;
   MD5Final ( digest, context );
   SetLength( Result, sizeof(digest) );
   Move( digest, Result[1], sizeof(digest) );
end;


// converts MD5 digest into a hex-string

function MD5toHex( const digest: MD5DigestString ): String;
var  i: Integer;
begin
   Result := '';
   for i:=1 to length(digest) do Result := Result + inttohex( ord( digest[i] ), 2 );
   Result := LowerCase( Result );
end;

// ----------------------------------------------------------------------------

// Keyed MD5 (HMAC-MD5), RFC 2104

procedure HMAC_MD5( const Data; DataLen: Integer;
                    const Key;  KeyLen : Integer;
                    out   Digest       : MD5_DIGEST );
var  k_ipad, k_opad: array[0..64] of Byte;
     Context: MD5_CTX;
     i      : Integer;
begin
   // clear pads
   FillChar( k_ipad, sizeof(k_ipad), 0 );
   FillChar( k_opad, sizeof(k_ipad), 0 );

   if KeyLen > 64 then begin
      // if key is longer than 64 bytes reset it to key=MD5(key)
      MD5Init  ( Context );
      MD5Update( Context, Key, KeyLen );
      MD5Final ( Digest, Context );
      // store key in pads
      Move( Digest, k_ipad, MD5HashSize );
      Move( Digest, k_opad, MD5HashSize );
   end else begin
      // store key in pads
      Move( Key, k_ipad, KeyLen );
      Move( Key, k_opad, KeyLen );
   end;

   // XOR key with ipad and opad values
   for i:=0 to 63 do begin
      k_ipad[i] := k_ipad[i] xor $36;
      k_opad[i] := k_opad[i] xor $5c;
   end;

   // perform inner MD5
   MD5Init  ( Context );
   MD5Update( Context, k_ipad, 64 );
   MD5Update( Context, Data, DataLen );
   MD5Final ( Digest, context );

   // perform outer MD5
   MD5Init  ( Context );
   MD5Update( Context, k_opad, 64 );
   MD5Update( Context, Digest, MD5HashSize );
   MD5Final ( Digest, Context );
end;

procedure HMAC_MD5( const Data: String;
                    const Key : String;
                    out   DigestStr: MD5DigestString ); overload;
var  Digest: MD5_DIGEST;
begin
   HMAC_MD5( Data[1], length(Data), Key[1], length(Key), Digest );
   SetLength( DigestStr, MD5HashSize );
   Move( Digest[0], DigestStr[1], MD5HashSize );
end;

function HMAC_MD5( const Data: String;
                   const Key : String ): String;
begin
   HMAC_MD5( Data, Key, Result );
   Result := MD5toHex( Result );
end;

// ----------------------------------------------------------------------------

{
MD5 test suite:
MD5 ("") = d41d8cd98f00b204e9800998ecf8427e
MD5 ("a") = 0cc175b9c0f1b6a831c399e269772661
MD5 ("abc") = 900150983cd24fb0d6963f7d28e17f72
MD5 ("message digest") = f96b697d7cb7938d525a2f31aaf161d0
MD5 ("abcdefghijklmnopqrstuvwxyz") = c3fcd3d76192e4007dfb496cca67e13b
MD5 ("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789")
    = d174ab98d277d9f5a5611c2c9f419d9f
MD5 ("12345678901234567890123456789012345678901234567890123456789
     012345678901234567890") = 57edf4a22be3c955ac49da2e2107b67a
}

{
HMAC-MD5 test suite of RFC 2202:
procedure TForm1.Button2Click(Sender: TObject);
var  Data, Key, Wanted: String;
     i: Integer;
begin
   Key    := StringOfChar( #$0b, 16 );
   Data   := 'Hi There';
   Wanted := '9294727a3638bb1c13f48ef8158bfc9d';
   ListBox1.Items.Add( 'Wanted: ' + Wanted );
   ListBox1.Items.Add( 'Result: ' + HMAC_MD5( Data, Key ) );

   Key    := 'Jefe';
   Data   := 'what do ya want for nothing?';
   Wanted := '750c783e6ab0b503eaa86e310a5db738';
   ListBox1.Items.Add( 'Wanted: ' + Wanted );
   ListBox1.Items.Add( 'Result: ' + HMAC_MD5( Data, Key ) );

   Key    := StringOfChar( #$aa, 16 );
   Data   := StringOfChar( #$dd, 50 );
   Wanted := '56be34521d144c88dbb8c733f0e8b3f6';
   ListBox1.Items.Add( 'Wanted: ' + Wanted );
   ListBox1.Items.Add( 'Result: ' + HMAC_MD5( Data, Key ) );

   SetLength( Key, 25 );
   for i:=1 to 25 do Key[i]:=chr(i);
   Data   := StringOfChar( #$cd, 50 );
   Wanted := '697eaf0aca3a3aea3a75164746ffaa79';
   ListBox1.Items.Add( 'Wanted: ' + Wanted );
   ListBox1.Items.Add( 'Result: ' + HMAC_MD5( Data, Key ) );

   Key    := StringOfChar( #$0c, 16 );
   Data   := 'Test With Truncation';
   Wanted := '56461ef2342edc00f9bab995690efd4c';
   ListBox1.Items.Add( 'Wanted: ' + Wanted );
   ListBox1.Items.Add( 'Result: ' + HMAC_MD5( Data, Key ) );

   Key    := StringOfChar( #$aa, 80 );
   Data   := 'Test Using Larger Than Block-Size Key - Hash Key First';
   Wanted := '6b1ab7fe4bd7bf8f0b62e6ce61b9d0cd';
   ListBox1.Items.Add( 'Wanted: ' + Wanted );
   ListBox1.Items.Add( 'Result: ' + HMAC_MD5( Data, Key ) );

   Key    := StringOfChar( #$aa, 80 );
   Data   := 'Test Using Larger Than Block-Size Key and Larger Than One Block-Size Data';
   Wanted := '6f630fad67cda0ee1fb1f562db3aa53e';
   ListBox1.Items.Add( 'Wanted: ' + Wanted );
   ListBox1.Items.Add( 'Result: ' + HMAC_MD5( Data, Key ) );
end;
}

end.
