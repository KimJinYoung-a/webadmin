
<script language="javascript" runat="server">
	//'#############################################################################
	//'#####
	//'#####	한국신용정보주식회사				본인확인서비스 스크립트
	//'#####
	//'#####	=====================================================================
	//'#####
	//'#####	Descriptions
	//'#####		- 한신정에서 제공하는 서비스에 대한 확인 작업을 처리한다.
	//'#####
	//'#####	---------------------------------------------------------------------
	//'#####
	//'#####	작성자 		: 한국신용정보 (www.nice.co.kr)
	//'#####	원본참조	:
	//'#####	원본파일	:
	//'#####	작성일자	: 2006-08-28
	//'#####
	//'#############################################################################
	var strDelimeter				= "!#!";		//	데이터 분리 Delimeter
	var CRNDSIZE 					= 24;			//	보안 처리에 사용할 키 길이

	//	-----------------------------------------------------
	//	-----	보안 처리를 위한 Object 처리 정보
	//	-----------------------------------------------------

	var oivsObject = new Object();
	oivsObject.KeyStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
	oivsObject.AthKeyStr = "";


	oivsObject.clientData = "";
	oivsObject.niceId = "";
	oivsObject.pingInfo = "";
	oivsObject.message = "";
	oivsObject.retCd = "";
	oivsObject.resCd = "";
	oivsObject.skData = "";
	oivsObject.ckData = "";
	oivsObject.identifier = "";
	oivsObject.residNo = "";
	oivsObject.ordNo = "";
	oivsObject.trNo = "";
	oivsObject.phoneNo = "";

	//
	//	Client Random Key 생성
	//

	oivsObject.getRandomKey = function( digits )
	{
		var rndKey, nIndex;

		rndKey = "";

		do
		{
			nIndex = Math.floor( Math.random() * oivsObject.KeyStr.length ) + 1;
			rndKey = rndKey + oivsObject.KeyStr.substr( nIndex, 1 );
		}
		while ( rndKey.length < digits )

		return rndKey;
	}
	//
	//	Base 64 Decode
	//
	oivsObject.decode64 = function( input )
	{
		var output = "";
		var chr1, chr2, chr3;
		var enc1, enc2, enc3, enc4;
		var i = 0;
		var strValue = new String( input );

		// remove all characters that are not A-Z, a-z, 0-9, +, /, or =
		var re = /[^A-Za-z0-9\+\/\=]/g;
		strValue = strValue.replace( re, "" );

		do
		{
			enc1 = oivsObject.KeyStr.indexOf(strValue.charAt(i++));
			enc2 = oivsObject.KeyStr.indexOf(strValue.charAt(i++));
			enc3 = oivsObject.KeyStr.indexOf(strValue.charAt(i++));
			enc4 = oivsObject.KeyStr.indexOf(strValue.charAt(i++));

			chr1 = (enc1 << 2) | (enc2 >> 4);
			chr2 = ((enc2 & 15) << 4) | (enc3 >> 2);
			chr3 = ((enc3 & 3) << 6) | enc4;

			output = output + String.fromCharCode(chr1);

			if (enc3 != 64)
			{
				output = output + String.fromCharCode(chr2);
			}

			if (enc4 != 64)
			{
				output = output + String.fromCharCode(chr3);
			}
		}
		while (i < strValue.length);

	   return output;
	}

	//
	//	Base 64 Encode
	//
	oivsObject.encode64 = function( input )
	{
		var output = "";
		var chr1, chr2, chr3;
		var enc1, enc2, enc3, enc4;
		var i = 0;

		do
		{
			chr1 = input.charCodeAt(i++);
			chr2 = input.charCodeAt(i++);
			chr3 = input.charCodeAt(i++);

			enc1 = chr1 >> 2;
			enc2 = ((chr1 & 3) << 4) | (chr2 >> 4);
			enc3 = ((chr2 & 15) << 2) | (chr3 >> 6);
			enc4 = chr3 & 63;

			if (isNaN(chr2))
			{
				enc3 = enc4 = 64;
			}
			else if (isNaN(chr3))
			{
				enc4 = 64;
			}

			output = output + oivsObject.KeyStr.charAt(enc1) + oivsObject.KeyStr.charAt(enc2) +
				oivsObject.KeyStr.charAt(enc3) + oivsObject.KeyStr.charAt(enc4);
		}
		while (i < input.length);

		return output;
	}

	//
	//	MD5 Hash
	//
	oivsObject.md5 = function( sMessage )
	{
		function RotateLeft(lValue, iShiftBits) { return (lValue<<iShiftBits) | (lValue>>>(32-iShiftBits)); }
		function AddUnsigned(lX,lY) {
			var lX4,lY4,lX8,lY8,lResult;
			lX8 = (lX & 0x80000000);
			lY8 = (lY & 0x80000000);
			lX4 = (lX & 0x40000000);
			lY4 = (lY & 0x40000000);
			lResult = (lX & 0x3FFFFFFF)+(lY & 0x3FFFFFFF);
			if (lX4 & lY4) return (lResult ^ 0x80000000 ^ lX8 ^ lY8);
			if (lX4 | lY4)
			{
				if (lResult & 0x40000000) return (lResult ^ 0xC0000000 ^ lX8 ^ lY8);
				else return (lResult ^ 0x40000000 ^ lX8 ^ lY8);
			}
			else return (lResult ^ lX8 ^ lY8);
		}
		function F(x,y,z) { return (x & y) | ((~x) & z); }
		function G(x,y,z) { return (x & z) | (y & (~z)); }
		function H(x,y,z) { return (x ^ y ^ z); }
		function I(x,y,z) { return (y ^ (x | (~z))); }
		function FF(a,b,c,d,x,s,ac) {
			a = AddUnsigned(a, AddUnsigned(AddUnsigned(F(b, c, d), x), ac));
			return AddUnsigned(RotateLeft(a, s), b);
		}
		function GG(a,b,c,d,x,s,ac) {
			a = AddUnsigned(a, AddUnsigned(AddUnsigned(G(b, c, d), x), ac));
			return AddUnsigned(RotateLeft(a, s), b);
		}
		function HH(a,b,c,d,x,s,ac) {
			a = AddUnsigned(a, AddUnsigned(AddUnsigned(H(b, c, d), x), ac));
			return AddUnsigned(RotateLeft(a, s), b);
		}
		function II(a,b,c,d,x,s,ac) {
			a = AddUnsigned(a, AddUnsigned(AddUnsigned(I(b, c, d), x), ac));
			return AddUnsigned(RotateLeft(a, s), b);
		}
		function ConvertToWordArray(sMessage) {
			var lWordCount;
			var lMessageLength = sMessage.length;
			var lNumberOfWords_temp1=lMessageLength + 8;
			var lNumberOfWords_temp2=(lNumberOfWords_temp1-(lNumberOfWords_temp1 % 64))/64;
			var lNumberOfWords = (lNumberOfWords_temp2+1)*16;
			var lWordArray=Array(lNumberOfWords-1);
			var lBytePosition = 0;
			var lByteCount = 0;
			while ( lByteCount < lMessageLength ) {
				lWordCount = (lByteCount-(lByteCount % 4))/4;
				lBytePosition = (lByteCount % 4)*8;
				lWordArray[lWordCount] = (lWordArray[lWordCount] | (sMessage.charCodeAt(lByteCount)<<lBytePosition));
				lByteCount++;
			}
			lWordCount = (lByteCount-(lByteCount % 4))/4;
			lBytePosition = (lByteCount % 4)*8;
			lWordArray[lWordCount] = lWordArray[lWordCount] | (0x80<<lBytePosition);
			lWordArray[lNumberOfWords-2] = lMessageLength<<3;
			lWordArray[lNumberOfWords-1] = lMessageLength>>>29;
			return lWordArray;
		}
		function WordToHex(lValue) {
			var WordToHexValue="",WordToHexValue_temp="",lByte,lCount;
			for (lCount = 0;lCount<=3;lCount++) {
				lByte = (lValue>>>(lCount*8)) & 255;
				WordToHexValue_temp = "0" + lByte.toString(16);
				WordToHexValue = WordToHexValue + WordToHexValue_temp.substr(WordToHexValue_temp.length-2,2);
			}
			return WordToHexValue;
		}

		var x=Array();
		var k,AA,BB,CC,DD,a,b,c,d
		var S11=7, S12=12, S13=17, S14=22;
		var S21=5, S22=9 , S23=14, S24=20;
		var S31=4, S32=11, S33=16, S34=23;
		var S41=6, S42=10, S43=15, S44=21;
		// Steps 1 and 2.  Append padding bits and length and convert to words
		x = ConvertToWordArray(sMessage);
		// Step 3.  Initialise
		a = 0x67452301; b = 0xEFCDAB89; c = 0x98BADCFE; d = 0x10325476;
		// Step 4.  Process the message in 16-word blocks
		for (k=0;k<x.length;k+=16) {
			AA=a; BB=b; CC=c; DD=d;
			a=FF(a,b,c,d,x[k+0], S11,0xD76AA478);
			d=FF(d,a,b,c,x[k+1], S12,0xE8C7B756);
			c=FF(c,d,a,b,x[k+2], S13,0x242070DB);
			b=FF(b,c,d,a,x[k+3], S14,0xC1BDCEEE);
			a=FF(a,b,c,d,x[k+4], S11,0xF57C0FAF);
			d=FF(d,a,b,c,x[k+5], S12,0x4787C62A);
			c=FF(c,d,a,b,x[k+6], S13,0xA8304613);
			b=FF(b,c,d,a,x[k+7], S14,0xFD469501);
			a=FF(a,b,c,d,x[k+8], S11,0x698098D8);
			d=FF(d,a,b,c,x[k+9], S12,0x8B44F7AF);
			c=FF(c,d,a,b,x[k+10],S13,0xFFFF5BB1);
			b=FF(b,c,d,a,x[k+11],S14,0x895CD7BE);
			a=FF(a,b,c,d,x[k+12],S11,0x6B901122);
			d=FF(d,a,b,c,x[k+13],S12,0xFD987193);
			c=FF(c,d,a,b,x[k+14],S13,0xA679438E);
			b=FF(b,c,d,a,x[k+15],S14,0x49B40821);
			a=GG(a,b,c,d,x[k+1], S21,0xF61E2562);
			d=GG(d,a,b,c,x[k+6], S22,0xC040B340);
			c=GG(c,d,a,b,x[k+11],S23,0x265E5A51);
			b=GG(b,c,d,a,x[k+0], S24,0xE9B6C7AA);
			a=GG(a,b,c,d,x[k+5], S21,0xD62F105D);
			d=GG(d,a,b,c,x[k+10],S22,0x2441453);
			c=GG(c,d,a,b,x[k+15],S23,0xD8A1E681);
			b=GG(b,c,d,a,x[k+4], S24,0xE7D3FBC8);
			a=GG(a,b,c,d,x[k+9], S21,0x21E1CDE6);
			d=GG(d,a,b,c,x[k+14],S22,0xC33707D6);
			c=GG(c,d,a,b,x[k+3], S23,0xF4D50D87);
			b=GG(b,c,d,a,x[k+8], S24,0x455A14ED);
			a=GG(a,b,c,d,x[k+13],S21,0xA9E3E905);
			d=GG(d,a,b,c,x[k+2], S22,0xFCEFA3F8);
			c=GG(c,d,a,b,x[k+7], S23,0x676F02D9);
			b=GG(b,c,d,a,x[k+12],S24,0x8D2A4C8A);
			a=HH(a,b,c,d,x[k+5], S31,0xFFFA3942);
			d=HH(d,a,b,c,x[k+8], S32,0x8771F681);
			c=HH(c,d,a,b,x[k+11],S33,0x6D9D6122);
			b=HH(b,c,d,a,x[k+14],S34,0xFDE5380C);
			a=HH(a,b,c,d,x[k+1], S31,0xA4BEEA44);
			d=HH(d,a,b,c,x[k+4], S32,0x4BDECFA9);
			c=HH(c,d,a,b,x[k+7], S33,0xF6BB4B60);
			b=HH(b,c,d,a,x[k+10],S34,0xBEBFBC70);
			a=HH(a,b,c,d,x[k+13],S31,0x289B7EC6);
			d=HH(d,a,b,c,x[k+0], S32,0xEAA127FA);
			c=HH(c,d,a,b,x[k+3], S33,0xD4EF3085);
			b=HH(b,c,d,a,x[k+6], S34,0x4881D05);
			a=HH(a,b,c,d,x[k+9], S31,0xD9D4D039);
			d=HH(d,a,b,c,x[k+12],S32,0xE6DB99E5);
			c=HH(c,d,a,b,x[k+15],S33,0x1FA27CF8);
			b=HH(b,c,d,a,x[k+2], S34,0xC4AC5665);
			a=II(a,b,c,d,x[k+0], S41,0xF4292244);
			d=II(d,a,b,c,x[k+7], S42,0x432AFF97);
			c=II(c,d,a,b,x[k+14],S43,0xAB9423A7);
			b=II(b,c,d,a,x[k+5], S44,0xFC93A039);
			a=II(a,b,c,d,x[k+12],S41,0x655B59C3);
			d=II(d,a,b,c,x[k+3], S42,0x8F0CCC92);
			c=II(c,d,a,b,x[k+10],S43,0xFFEFF47D);
			b=II(b,c,d,a,x[k+1], S44,0x85845DD1);
			a=II(a,b,c,d,x[k+8], S41,0x6FA87E4F);
			d=II(d,a,b,c,x[k+15],S42,0xFE2CE6E0);
			c=II(c,d,a,b,x[k+6], S43,0xA3014314);
			b=II(b,c,d,a,x[k+13],S44,0x4E0811A1);
			a=II(a,b,c,d,x[k+4], S41,0xF7537E82);
			d=II(d,a,b,c,x[k+11],S42,0xBD3AF235);
			c=II(c,d,a,b,x[k+2], S43,0x2AD7D2BB);
			b=II(b,c,d,a,x[k+9], S44,0xEB86D391);
			a=AddUnsigned(a,AA); b=AddUnsigned(b,BB); c=AddUnsigned(c,CC); d=AddUnsigned(d,DD);
		}
		// Step 5.  Output the 128 bit digest
		var temp= WordToHex(a)+WordToHex(b)+WordToHex(c)+WordToHex(d);
		return temp.toLowerCase();
	}

	//
	//	DES Symmetric Key
	//
	oivsObject.des = function( key, message, encrypt, mode, iv )
	{
		//des_createKeys
		//this takes as input a 64 bit key (even though only 56 bits are used)
		//as an array of 2 integers, and returns 16 48 bit keys
		function des_createKeys( key )
		{
			//declaring this locally speeds things up a bit
			pc2bytes0  = new Array (0,0x4,0x20000000,0x20000004,0x10000,0x10004,0x20010000,0x20010004,0x200,0x204,0x20000200,0x20000204,0x10200,0x10204,0x20010200,0x20010204);
			pc2bytes1  = new Array (0,0x1,0x100000,0x100001,0x4000000,0x4000001,0x4100000,0x4100001,0x100,0x101,0x100100,0x100101,0x4000100,0x4000101,0x4100100,0x4100101);
			pc2bytes2  = new Array (0,0x8,0x800,0x808,0x1000000,0x1000008,0x1000800,0x1000808,0,0x8,0x800,0x808,0x1000000,0x1000008,0x1000800,0x1000808);
			pc2bytes3  = new Array (0,0x200000,0x8000000,0x8200000,0x2000,0x202000,0x8002000,0x8202000,0x20000,0x220000,0x8020000,0x8220000,0x22000,0x222000,0x8022000,0x8222000);
			pc2bytes4  = new Array (0,0x40000,0x10,0x40010,0,0x40000,0x10,0x40010,0x1000,0x41000,0x1010,0x41010,0x1000,0x41000,0x1010,0x41010);
			pc2bytes5  = new Array (0,0x400,0x20,0x420,0,0x400,0x20,0x420,0x2000000,0x2000400,0x2000020,0x2000420,0x2000000,0x2000400,0x2000020,0x2000420);
			pc2bytes6  = new Array (0,0x10000000,0x80000,0x10080000,0x2,0x10000002,0x80002,0x10080002,0,0x10000000,0x80000,0x10080000,0x2,0x10000002,0x80002,0x10080002);
			pc2bytes7  = new Array (0,0x10000,0x800,0x10800,0x20000000,0x20010000,0x20000800,0x20010800,0x20000,0x30000,0x20800,0x30800,0x20020000,0x20030000,0x20020800,0x20030800);
			pc2bytes8  = new Array (0,0x40000,0,0x40000,0x2,0x40002,0x2,0x40002,0x2000000,0x2040000,0x2000000,0x2040000,0x2000002,0x2040002,0x2000002,0x2040002);
			pc2bytes9  = new Array (0,0x10000000,0x8,0x10000008,0,0x10000000,0x8,0x10000008,0x400,0x10000400,0x408,0x10000408,0x400,0x10000400,0x408,0x10000408);
			pc2bytes10 = new Array (0,0x20,0,0x20,0x100000,0x100020,0x100000,0x100020,0x2000,0x2020,0x2000,0x2020,0x102000,0x102020,0x102000,0x102020);
			pc2bytes11 = new Array (0,0x1000000,0x200,0x1000200,0x200000,0x1200000,0x200200,0x1200200,0x4000000,0x5000000,0x4000200,0x5000200,0x4200000,0x5200000,0x4200200,0x5200200);
			pc2bytes12 = new Array (0,0x1000,0x8000000,0x8001000,0x80000,0x81000,0x8080000,0x8081000,0x10,0x1010,0x8000010,0x8001010,0x80010,0x81010,0x8080010,0x8081010);
			pc2bytes13 = new Array (0,0x4,0x100,0x104,0,0x4,0x100,0x104,0x1,0x5,0x101,0x105,0x1,0x5,0x101,0x105);

			//how many iterations (1 for des, 3 for triple des)
			var iterations = key.length >= 24 ? 3 : 1;
			//stores the return keys
			var keys = new Array (32 * iterations);
			//now define the left shifts which need to be done
			var shifts = new Array (0, 0, 1, 1, 1, 1, 1, 1, 0, 1, 1, 1, 1, 1, 1, 0);
			//other variables
			var lefttemp, righttemp, m=0, n=0, temp;

			for (var j=0; j<iterations; j++)
			{ //either 1 or 3 iterations
				left = (key.charCodeAt(m++) << 24) | (key.charCodeAt(m++) << 16) | (key.charCodeAt(m++) << 8) | key.charCodeAt(m++);
				right = (key.charCodeAt(m++) << 24) | (key.charCodeAt(m++) << 16) | (key.charCodeAt(m++) << 8) | key.charCodeAt(m++);

				temp = ((left >>> 4) ^ right) & 0x0f0f0f0f; right ^= temp; left ^= (temp << 4);
				temp = ((right >>> -16) ^ left) & 0x0000ffff; left ^= temp; right ^= (temp << -16);
				temp = ((left >>> 2) ^ right) & 0x33333333; right ^= temp; left ^= (temp << 2);
				temp = ((right >>> -16) ^ left) & 0x0000ffff; left ^= temp; right ^= (temp << -16);
				temp = ((left >>> 1) ^ right) & 0x55555555; right ^= temp; left ^= (temp << 1);
				temp = ((right >>> 8) ^ left) & 0x00ff00ff; left ^= temp; right ^= (temp << 8);
				temp = ((left >>> 1) ^ right) & 0x55555555; right ^= temp; left ^= (temp << 1);

				//the right side needs to be shifted and to get the last four bits of the left side
				temp = (left << 8) | ((right >>> 20) & 0x000000f0);
				//left needs to be put upside down
				left = (right << 24) | ((right << 8) & 0xff0000) | ((right >>> 8) & 0xff00) | ((right >>> 24) & 0xf0);
				right = temp;

				//now go through and perform these shifts on the left and right keys
				for (i=0; i < shifts.length; i++)
				{
					//shift the keys either one or two bits to the left
					if (shifts[i]) {left = (left << 2) | (left >>> 26); right = (right << 2) | (right >>> 26);}
					else {left = (left << 1) | (left >>> 27); right = (right << 1) | (right >>> 27);}
					left &= 0xfffffff0; right &= 0xfffffff0;

					//now apply PC-2, in such a way that E is easier when encrypting or decrypting
					//this conversion will look like PC-2 except only the last 6 bits of each byte are used
					//rather than 48 consecutive bits and the order of lines will be according to
					//how the S selection functions will be applied: S2, S4, S6, S8, S1, S3, S5, S7
					lefttemp = pc2bytes0[left >>> 28] | pc2bytes1[(left >>> 24) & 0xf]
							| pc2bytes2[(left >>> 20) & 0xf] | pc2bytes3[(left >>> 16) & 0xf]
							| pc2bytes4[(left >>> 12) & 0xf] | pc2bytes5[(left >>> 8) & 0xf]
							| pc2bytes6[(left >>> 4) & 0xf];
					righttemp = pc2bytes7[right >>> 28] | pc2bytes8[(right >>> 24) & 0xf]
							  | pc2bytes9[(right >>> 20) & 0xf] | pc2bytes10[(right >>> 16) & 0xf]
							  | pc2bytes11[(right >>> 12) & 0xf] | pc2bytes12[(right >>> 8) & 0xf]
							  | pc2bytes13[(right >>> 4) & 0xf];
					temp = ((righttemp >>> 16) ^ lefttemp) & 0x0000ffff;
					keys[n++] = lefttemp ^ temp; keys[n++] = righttemp ^ (temp << 16);
				}
			} //for each iterations

			//return the keys we've created
			return keys;
		} //end of des_createKeys

		//declaring this locally speeds things up a bit
		var spfunction1 = new Array (0x1010400,0,0x10000,0x1010404,0x1010004,0x10404,0x4,0x10000,0x400,0x1010400,0x1010404,0x400,0x1000404,0x1010004,0x1000000,0x4,0x404,0x1000400,0x1000400,0x10400,0x10400,0x1010000,0x1010000,0x1000404,0x10004,0x1000004,0x1000004,0x10004,0,0x404,0x10404,0x1000000,0x10000,0x1010404,0x4,0x1010000,0x1010400,0x1000000,0x1000000,0x400,0x1010004,0x10000,0x10400,0x1000004,0x400,0x4,0x1000404,0x10404,0x1010404,0x10004,0x1010000,0x1000404,0x1000004,0x404,0x10404,0x1010400,0x404,0x1000400,0x1000400,0,0x10004,0x10400,0,0x1010004);
		var spfunction2 = new Array (0x80108020,0x80008000,0x8000,0x108020,0x100000,0x20,0x80100020,0x80008020,0x80000020,0x80108020,0x80108000,0x80000000,0x80008000,0x100000,0x20,0x80100020,0x108000,0x100020,0x80008020,0,0x80000000,0x8000,0x108020,0x80100000,0x100020,0x80000020,0,0x108000,0x8020,0x80108000,0x80100000,0x8020,0,0x108020,0x80100020,0x100000,0x80008020,0x80100000,0x80108000,0x8000,0x80100000,0x80008000,0x20,0x80108020,0x108020,0x20,0x8000,0x80000000,0x8020,0x80108000,0x100000,0x80000020,0x100020,0x80008020,0x80000020,0x100020,0x108000,0,0x80008000,0x8020,0x80000000,0x80100020,0x80108020,0x108000);
		var spfunction3 = new Array (0x208,0x8020200,0,0x8020008,0x8000200,0,0x20208,0x8000200,0x20008,0x8000008,0x8000008,0x20000,0x8020208,0x20008,0x8020000,0x208,0x8000000,0x8,0x8020200,0x200,0x20200,0x8020000,0x8020008,0x20208,0x8000208,0x20200,0x20000,0x8000208,0x8,0x8020208,0x200,0x8000000,0x8020200,0x8000000,0x20008,0x208,0x20000,0x8020200,0x8000200,0,0x200,0x20008,0x8020208,0x8000200,0x8000008,0x200,0,0x8020008,0x8000208,0x20000,0x8000000,0x8020208,0x8,0x20208,0x20200,0x8000008,0x8020000,0x8000208,0x208,0x8020000,0x20208,0x8,0x8020008,0x20200);
		var spfunction4 = new Array (0x802001,0x2081,0x2081,0x80,0x802080,0x800081,0x800001,0x2001,0,0x802000,0x802000,0x802081,0x81,0,0x800080,0x800001,0x1,0x2000,0x800000,0x802001,0x80,0x800000,0x2001,0x2080,0x800081,0x1,0x2080,0x800080,0x2000,0x802080,0x802081,0x81,0x800080,0x800001,0x802000,0x802081,0x81,0,0,0x802000,0x2080,0x800080,0x800081,0x1,0x802001,0x2081,0x2081,0x80,0x802081,0x81,0x1,0x2000,0x800001,0x2001,0x802080,0x800081,0x2001,0x2080,0x800000,0x802001,0x80,0x800000,0x2000,0x802080);
		var spfunction5 = new Array (0x100,0x2080100,0x2080000,0x42000100,0x80000,0x100,0x40000000,0x2080000,0x40080100,0x80000,0x2000100,0x40080100,0x42000100,0x42080000,0x80100,0x40000000,0x2000000,0x40080000,0x40080000,0,0x40000100,0x42080100,0x42080100,0x2000100,0x42080000,0x40000100,0,0x42000000,0x2080100,0x2000000,0x42000000,0x80100,0x80000,0x42000100,0x100,0x2000000,0x40000000,0x2080000,0x42000100,0x40080100,0x2000100,0x40000000,0x42080000,0x2080100,0x40080100,0x100,0x2000000,0x42080000,0x42080100,0x80100,0x42000000,0x42080100,0x2080000,0,0x40080000,0x42000000,0x80100,0x2000100,0x40000100,0x80000,0,0x40080000,0x2080100,0x40000100);
		var spfunction6 = new Array (0x20000010,0x20400000,0x4000,0x20404010,0x20400000,0x10,0x20404010,0x400000,0x20004000,0x404010,0x400000,0x20000010,0x400010,0x20004000,0x20000000,0x4010,0,0x400010,0x20004010,0x4000,0x404000,0x20004010,0x10,0x20400010,0x20400010,0,0x404010,0x20404000,0x4010,0x404000,0x20404000,0x20000000,0x20004000,0x10,0x20400010,0x404000,0x20404010,0x400000,0x4010,0x20000010,0x400000,0x20004000,0x20000000,0x4010,0x20000010,0x20404010,0x404000,0x20400000,0x404010,0x20404000,0,0x20400010,0x10,0x4000,0x20400000,0x404010,0x4000,0x400010,0x20004010,0,0x20404000,0x20000000,0x400010,0x20004010);
		var spfunction7 = new Array (0x200000,0x4200002,0x4000802,0,0x800,0x4000802,0x200802,0x4200800,0x4200802,0x200000,0,0x4000002,0x2,0x4000000,0x4200002,0x802,0x4000800,0x200802,0x200002,0x4000800,0x4000002,0x4200000,0x4200800,0x200002,0x4200000,0x800,0x802,0x4200802,0x200800,0x2,0x4000000,0x200800,0x4000000,0x200800,0x200000,0x4000802,0x4000802,0x4200002,0x4200002,0x2,0x200002,0x4000000,0x4000800,0x200000,0x4200800,0x802,0x200802,0x4200800,0x802,0x4000002,0x4200802,0x4200000,0x200800,0,0x2,0x4200802,0,0x200802,0x4200000,0x800,0x4000002,0x4000800,0x800,0x200002);
		var spfunction8 = new Array (0x10001040,0x1000,0x40000,0x10041040,0x10000000,0x10001040,0x40,0x10000000,0x40040,0x10040000,0x10041040,0x41000,0x10041000,0x41040,0x1000,0x40,0x10040000,0x10000040,0x10001000,0x1040,0x41000,0x40040,0x10040040,0x10041000,0x1040,0,0,0x10040040,0x10000040,0x10001000,0x41040,0x40000,0x41040,0x40000,0x10041000,0x1000,0x40,0x10040040,0x1000,0x41040,0x10001000,0x40,0x10000040,0x10040000,0x10040040,0x10000000,0x40000,0x10001040,0,0x10041040,0x40040,0x10000040,0x10040000,0x10001000,0x10001040,0,0x10041040,0x41000,0x41000,0x1040,0x1040,0x40040,0x10000000,0x10041000);

		//create the 16 or 48 subkeys we will need
		var keys = des_createKeys (key);
		var m=0, i, j, temp, temp2, right1, right2, left, right, looping;
		var cbcleft, cbcleft2, cbcright, cbcright2
		var endloop, loopinc;
		var len = message.length;
		var chunk = 0;
		//set up the loops for single and triple des
		var iterations = keys.length == 32 ? 3 : 9; //single or triple des
		if (iterations == 3) {looping = encrypt ? new Array (0, 32, 2) : new Array (30, -2, -2);}
		else {looping = encrypt ? new Array (0, 32, 2, 62, 30, -2, 64, 96, 2) : new Array (94, 62, -2, 32, 64, 2, 30, -2, -2);}

		message += "\0\0\0\0\0\0\0\0"; //pad the message out with null bytes
		//store the result here
		result = "";
		tempresult = "";

		if (mode == 1) { //CBC mode
		  cbcleft = (iv.charCodeAt(m++) << 24) | (iv.charCodeAt(m++) << 16) | (iv.charCodeAt(m++) << 8) | iv.charCodeAt(m++);
		  cbcright = (iv.charCodeAt(m++) << 24) | (iv.charCodeAt(m++) << 16) | (iv.charCodeAt(m++) << 8) | iv.charCodeAt(m++);
		  m=0;
		}

		//loop through each 64 bit chunk of the message
		while (m < len)
		{
			left = (message.charCodeAt(m++) << 24) | (message.charCodeAt(m++) << 16) | (message.charCodeAt(m++) << 8) | message.charCodeAt(m++);
			right = (message.charCodeAt(m++) << 24) | (message.charCodeAt(m++) << 16) | (message.charCodeAt(m++) << 8) | message.charCodeAt(m++);

			//for Cipher Block Chaining mode, xor the message with the previous result
			if (mode == 1) {if (encrypt) {left ^= cbcleft; right ^= cbcright;} else {cbcleft2 = cbcleft; cbcright2 = cbcright; cbcleft = left; cbcright = right;}}

			//first each 64 but chunk of the message must be permuted according to IP
			temp = ((left >>> 4) ^ right) & 0x0f0f0f0f; right ^= temp; left ^= (temp << 4);
			temp = ((left >>> 16) ^ right) & 0x0000ffff; right ^= temp; left ^= (temp << 16);
			temp = ((right >>> 2) ^ left) & 0x33333333; left ^= temp; right ^= (temp << 2);
			temp = ((right >>> 8) ^ left) & 0x00ff00ff; left ^= temp; right ^= (temp << 8);
			temp = ((left >>> 1) ^ right) & 0x55555555; right ^= temp; left ^= (temp << 1);

			left = ((left << 1) | (left >>> 31));
			right = ((right << 1) | (right >>> 31));

			//do this either 1 or 3 times for each chunk of the message
			for (j=0; j<iterations; j+=3)
			{
				endloop = looping[j+1];
				loopinc = looping[j+2];
				//now go through and perform the encryption or decryption
				for (i=looping[j]; i!=endloop; i+=loopinc)
				{ //for efficiency
					right1 = right ^ keys[i];
					right2 = ((right >>> 4) | (right << 28)) ^ keys[i+1];
					//the result is attained by passing these bytes through the S selection functions
					temp = left;
					left = right;
					right = temp ^ (spfunction2[(right1 >>> 24) & 0x3f] | spfunction4[(right1 >>> 16) & 0x3f]
						| spfunction6[(right1 >>>  8) & 0x3f] | spfunction8[right1 & 0x3f]
						| spfunction1[(right2 >>> 24) & 0x3f] | spfunction3[(right2 >>> 16) & 0x3f]
						| spfunction5[(right2 >>>  8) & 0x3f] | spfunction7[right2 & 0x3f]);
				}

				temp = left; left = right; right = temp; //unreverse left and right
			} //for either 1 or 3 iterations

			//move then each one bit to the right
			left = ((left >>> 1) | (left << 31));
			right = ((right >>> 1) | (right << 31));

			//now perform IP-1, which is IP in the opposite direction
			temp = ((left >>> 1) ^ right) & 0x55555555; right ^= temp; left ^= (temp << 1);
			temp = ((right >>> 8) ^ left) & 0x00ff00ff; left ^= temp; right ^= (temp << 8);
			temp = ((right >>> 2) ^ left) & 0x33333333; left ^= temp; right ^= (temp << 2);
			temp = ((left >>> 16) ^ right) & 0x0000ffff; right ^= temp; left ^= (temp << 16);
			temp = ((left >>> 4) ^ right) & 0x0f0f0f0f; right ^= temp; left ^= (temp << 4);

			//for Cipher Block Chaining mode, xor the message with the previous result
			if (mode == 1) {if (encrypt) {cbcleft = left; cbcright = right;} else {left ^= cbcleft2; right ^= cbcright2;}}
			tempresult += String.fromCharCode ((left>>>24), ((left>>>16) & 0xff), ((left>>>8) & 0xff), (left & 0xff), (right>>>24), ((right>>>16) & 0xff), ((right>>>8) & 0xff), (right & 0xff));

			chunk += 8;
			if (chunk == 512) {result += tempresult; tempresult = ""; chunk = 0;}
		} //for every 8 characters, or 64 bits in the message

		//return the result as an array
		return result + tempresult;
	} //end of des

	oivsObject.URLDecode = function( data )
	{
		var HEXCHARS = "0123456789ABCDEFabcdef";
		var plaintext = "";
		var i = 0;

		while (i < data.length)
		{
			var ch = data.charAt(i);
			if (ch == "+")
			{
				plaintext += " ";
				i++;
			}
			else if (ch == "%")
			{
				if (i < (data.length-2)
					&& HEXCHARS.indexOf(data.charAt(i+1)) != -1
					&& HEXCHARS.indexOf(data.charAt(i+2)) != -1 )
				{
					plaintext += unescape( data.substr(i,3) );
					i += 3;
				}
				else
				{
					alert( 'Bad escape combination near ...' + encoded.substr(i) );
					plaintext += "%[ERROR]";
					i++;
				}
			}
			else
			{
			   plaintext += ch;
			   i++;
			}
		} // while

		return plaintext;
	}

	oivsObject.resolveDatas = function( SendInfo )
	{
		var strCDESInfo = "";
		var aryValues;
		var dateObj;
		var keyPivot;
		var key;
		var encData; 

		strCDESInfo = oivsObject.decode64( SendInfo );
		aryValues = strCDESInfo.split( "!@!" );

		if (aryValues.length < 4) {
			this.retCd = "0";
			this.message = "전달받은 데이터에 오류가 있습니다.";
			return false;
		}
		this.ordNo = aryValues[0];
		this.resCd = aryValues[1];
		this.message = decodeURIComponent(aryValues[2]);
		strCDESInfo = aryValues[3];
		

		var i = 0;
		var hashCnt = aryValues[4];	
		for(i=0;i<hashCnt;i++){
			this.AthKeyStr = oivsObject.md5(this.AthKeyStr);
		}

		dateObj = new Date();
		keyPivot = (dateObj.getYear() % 100 * 15 + (dateObj.getMonth() + 1)) % 8;
		key = oivsObject.AthKeyStr.substring(keyPivot, keyPivot + 24);
		encData = oivsObject.decode64( strCDESInfo ); 

		strDecData = oivsObject.des( key, encData, 0, 1, key );
		aryValues = strDecData.split( "!@!" );

		if (aryValues.length < 5) {
			this.retCd = "0";
			this.message = "전달 받은 데이터에 오류가 있습니다.";
			return false;
		} else if (this.ordNo != aryValues[1]
				|| this.resCd != aryValues[3]) {
			this.retCd = "0";
			this.message = "전달 받은 데이터에 오류가 있습니다.";
			return false;
		} else {
			this.niceId = aryValues[0];
			this.retCd = aryValues[2];
			this.residNo = aryValues[4];
			this.phoneNo = aryValues[5];

			return true;
		}
	}

	function StringReplace(strTarget, strOld, strNew){
		while(strTarget.indexOf(strOld) >=0){
			strTarget = strTarget.replace(strOld, strNew);
		}
		return strTarget;
	}
	function makeSendInfo(NiceId, ResidNo, SIKey, ReturnURL, MobileNo, ConfirmMsg){
		
		var fnRslt = "";
		var keyHashineCnt = 0;
		var dateObj;
		var keyPivot;
		var key;
		
		if(ConfirmMsg!="") {
			keyHashineCnt = Math.round(Math.random()*100 + 1);
			for(i=0;i<keyHashineCnt;i++){
				oivsObject.AthKeyStr = oivsObject.md5(oivsObject.AthKeyStr);
			}

			dateObj = new Date();
			keyPivot = (dateObj.getYear() % 100 * 15 + (dateObj.getMonth() + 1)) % 8;

			key = oivsObject.AthKeyStr.substring(keyPivot, keyPivot + 24);
			encRslt = oivsObject.des( key, ConfirmMsg, 1, 1, key );

			fnRslt = oivsObject.encode64(encRslt);
		}
		
		var strCDESInfo = "";
		var ckData = "";
		var xmlhttp =  new ActiveXObject( "Microsoft.XMLHTTP" );
		xmlhttp.open("GET", "http://secure.nuguya.com/nuguya/rlnmPing.do", false );
		xmlhttp.setRequestHeader("Content-Type", "text/xml; charset=euc-kr");
		xmlhttp.send();
		pingInfo = xmlhttp.responseText;

		if ( pingInfo.indexOf( "500" ) >= 0 ){
			pingInfo =  "";
			mackSendInfo = "";
			return mackSendInfo;
		}
		xmlhttp = null;

		ckData = oivsObject.getRandomKey( CRNDSIZE );

		// var strCHashedInfo = oivsObject.md5( NiceId + pingInfo );
		var strEncData = oivsObject.des( ckData, NiceId + strDelimeter + pingInfo + strDelimeter + oivsObject.encode64(ResidNo+":"+MobileNo+":"+keyHashineCnt+":"+fnRslt) + strDelimeter + SIKey + strDelimeter + ReturnURL, 1, 1, ckData );

		strCDESInfo = oivsObject.encode64( ckData + strDelimeter + strEncData );
		return strCDESInfo;	
	}
</script>