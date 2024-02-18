<%
'+----------------------------------------------------------------------------------------------------------------------+
'|                                               HTML °ø Åë   ÇÔ ¼ö ¼± ¾ð                                               |
'+-------------------------------------------+--------------------------------------------------------------------------+
'|             ÇÔ ¼ö ¸í                      |                          ±â    ´É                                        |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| FormatDate(ddate, formatstring)          | ³¯Â¥Çü½ÄÀ» ÁöÁ¤µÈ ¹®ÀÚÇüÀ¸·Î º¯È¯                            |
'|                                          | »ç¿ë¿¹ : printdate = FormatDate(now(),"0000.00.00")          |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| GetImageSubFolderByItemid(byval iitemid)  | ÀÌ¹ÌÁöÆÄÀÏÀÇ ¼­ºê Æú´õ¸íÀ» ¹ÝÈ¯ÇÑ´Ù.                                     |
'|                                           | »ç¿ë¿¹ : SubFolder = GetImageSubFolderByItemid(1126)                     |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| db2html(checkvalue)                       | DBÀÇ ³»¿ëÀ» HTML¿¡ »ç¿ëÇÒ ¼ö ÀÖµµ·Ï º¯È¯                                 |
'|                                           | »ç¿ë¿¹ : Contents = db2html("DBÀÇ ³»¿ë")                                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| html2db(checkvalue)                       | »ç¿ëÀÚ°¡ ÀÔ·ÂÇÑ ³»¿ëÀ» DB¿¡ ³ÖÀ» ¼ö ÀÖµµ·Ï º¯È¯                          |
'|                                           | »ç¿ë¿¹ : Contents = html2db("ÀúÀåÇÒ ³»¿ë")                               |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| nl2br(checkvalue)                         | ³»¿ëÀÇ »õÁÙ(vbCrLf)À» "<br>"ÅÂ±×·Î Ä¡È¯ÇÏ¿© ¹ÝÈ¯                         |
'|                                           | »ç¿ë¿¹ : Contents = nl2br("³»¿ë")                                        |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| CurrFormat(byVal v)                       | ¼ýÀÚ¸¦ 3ÀÚ¸® ±¸ºÐÀÇ ¹®ÀÚ¿­·Î º¯È¯                                        |
'|                                           | »ç¿ë¿¹ : strNum = CurrFormat(1230) ¡æ "1,230"                            |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Format00(n,orgData)                       | ¼ýÀÚ¸¦ 0À¸·Î Ã¤¿öÁø ÁöÁ¤µÈ ±æÀÌÀÇ ¹®ÀÚ¿­·Î º¯È¯                          |
'|                                           | »ç¿ë¿¹ : strNum = Format00(5,123) ¡æ "00123"                             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| FormatCode(itemcode)                      | Á¦Ç° ÀÏ·Ã¹øÈ£¸¦ 6ÀÚ¸®ÀÇ ¹®ÀÚ¿­·Î º¯È¯                                    |
'|                                           | »ç¿ë¿¹ : itemCode = FormatCode(2654) ¡æ "002654"                         |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| GetCurrentTimeFormat()                    | ÇöÀç½Ã°£À» ¹®ÀÚ¿­·Î ¹ÝÈ¯ (yyyymmddhhmmss)                                |
'|                                           | »ç¿ë¿¹ : strNow = GetCurrentTimeFormat() ¡æ "20060508101833"             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| GetListImageUrl(byval itemid)             | Á¦Ç°¹øÈ£¿¡ ¸Â´Â ¸®½ºÆ® ÀÌ¹ÌÁö ¹× Æú´õ ¹ÝÈ¯                               |
'|                                           | »ç¿ë¿¹ : img = GetListImageUrl("53100") ¡æ "/image/list/L000053100.jpg"  |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| DDotFormat(byval str,byval n)             | ³»¿ëÀ» ÁöÁ¤ÇÑ ±æÀÌ·Î ÀÚ¸¥´Ù.                                             |
'|                                           | »ç¿ë¿¹ : strShort = DDotFormat("³»¿ëÀÔ´Ï´Ù.",3) ¡æ "³»¿ëÀÔ..."           |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| stripHTML(strng)                          | ³»¿ë Áß HTMLÅÂ±×¸¦ ¾ø¾Ø´Ù.                                               |
'|                                           | »ç¿ë¿¹ : Contents = stripHTML("<b>³»¿ë</b>") ¡æ " ³»¿ë "                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| getFileExtention(strFile)                 | ÆÄÀÏ¸íÀÇ È®ÀåÀÚ¸¦ ¹ÝÈ¯ÇÑ´Ù.                                              |
'|                                           | »ç¿ë¿¹ : ext = getFileExtention("123.jpg") ¡æ "jpg"                      |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Num2Str(inum,olen,cChr,oalign)   		 | ¼ýÀÚ¸¦ ÁöÁ¤ÇÑ ±æÀÌÀÇ ¹®ÀÚ¿­·Î º¯È¯ÇÑ´Ù.                      			|
'|                                   		 | »ç¿ë¿¹ : Num2Str(425,4,"0","R") ¡æ 0425                      			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ChkIIF(trueOrFalse, trueVal, falseVal)    | like iif function                                                        |
'|                                           | »ç¿ë¿¹ : ChkIIF(1>2,"a","b") ¡æ "b"                                       |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Alert_return(strMSG)                      | °æ°íÃ¢ ¶ç¿îÈÄ ÀÌÀüÀ¸·Î µ¹¾Æ°£´Ù.                            				|
'|                                           | »ç¿ë¿¹ : Call Alert_return("µÚ·Î µ¹¾Æ°©´Ï´Ù.")               			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Alert_close(strMSG)                       | °æ°íÃ¢ ¶ç¿îÈÄ ÇöÀçÃ¢À» ´Ý´Â´Ù.                               			|
'|                                           | »ç¿ë¿¹ : Call Alert_close("Ã¢À» ´Ý½À´Ï´Ù.")                  			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Alert_move(strMSG,targetURL)              | °æ°íÃ¢ ¶ç¿îÈÄ ÁöÁ¤ÆäÀÌÁö·Î ÀÌµ¿ÇÑ´Ù.                         			|
'|                                           | »ç¿ë¿¹ : Call Alert_move("ÀÌµ¿ÇÕ´Ï´Ù.","/index.asp")         			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chrbyte(str,chrlen,dot)                   | ÁöÁ¤±æÀÌ·Î ¹®ÀÚ¿­ ÀÚ¸£±â                                                 |
'|                                           | »ç¿ë¿¹ : chrbyte("¾È³çÇÏ¼¼¿ä",3,"Y") ¡æ ¾È³ç...                           |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chkPasswordComplex(uid,pwd)               | ºñ¹Ð¹øÈ£ Á¤Ã¥ÀÇ º¹Àâ¼ºÀ» ¸¸Á·ÇÏ´ÂÁö °Ë»çÇÏ°í ±× ÀÌÀ¯¸¦ ¹ÝÈ¯              |
'|                                           | »ç¿ë¿¹ : chkPasswordComplex("kobula","abcd")                             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chkPasswordComplexNonid(pwd)         	 | ¾ÆÀÌµð°¡ ¾ø°íºñ¹Ð¹øÈ£ Á¤Ã¥ÀÇ º¹Àâ¼ºÀ»¸¸Á·ÇÏ´ÂÁö °Ë»çÇÏ°í ±× ÀÌÀ¯¸¦ ¹ÝÈ¯  |
'|                                           | »ç¿ë¿¹ : chkPasswordComplexNonid("abcd")                             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chkWord(str,patrn)                        | ¹®ÀÚ¿­ÀÇ Çü½ÄÀ» Á¤±Ô½ÄÀ¸·Î °Ë»ç                                          |
'|                                           | »ç¿ë¿¹ : chkWord("abcd","[^-a-zA-Z0-9/ ]") : ¿µ¾î¼ýÀÚ¸¸ Çã¿ë             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ParsingPhoneNumber(str,patrn)             | ÀüÈ­¹øÈ£¿¡ ´ë½Ã Ãß°¡                                                     |
'|                                           | »ç¿ë¿¹ : ParsingPhoneNumber("0112223333") :                              |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceBracket(strng)                     | ²©Àº°ýÈ£ ÅÂ±×·Î Ä¡È¯('<', '>')                                           |
'|                                           | »ç¿ë¿¹ : ReplaceBracket("<>") ¡æ &lt;&gt;                                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceBracketOther(strng)                | ²©Àº°ýÈ£ ´Ù¸¥ °ýÈ£·Î Ä¡È¯('<', '>')                                        |
'|                                           | »ç¿ë¿¹ : ReplaceBracketOther("<>") ¡æ []                                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceScript(strng)                      | Script Tag Ä¡È¯                                                          |
'|                                           | »ç¿ë¿¹ : ReplaceScript("<script") ¡æ &lt;script                           |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| getNumeric(strNum)                        | ¹®ÀÚ¿­¿¡¼­ ¼ýÀÚ¸¸ ÃßÃâ º¯È¯                                              |
'|                                           | »ç¿ë¿¹ : getNumeric("a45d61*124") -> 461124                              |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| RepWord(str,patrn,repval)                | Á¤±Ô½Ä ÆÐÅÏÀ» »ç¿ëÇÑ ¹®ÀÚ¿­ Ã³¸®                             				|
'|                                          | »ç¿ë¿¹ : RepWord(SearchText,"[^°¡-ÆRa-zA-Z0-9\s]","")      			  	|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceRequestSpecialChar(strng)        	| Æ¯¼ö ¹®ÀÚ Á¦°Å(' ,--)                                        				|
'|                                          | »ç¿ë¿¹ : cont = ReplaceRequestSpecialChar(Rs("strng"))       				|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| checkNotValidHTML(ostr)                  | ³»¿ë¿¡ ±ÝÁöµÈ HTMLÅÂ±×°¡ ÀÖ´ÂÁö °Ë»ç                         				|
'|                                          | »ç¿ë¿¹ : checkNotValidHTML("<script...") ¡æ true             				|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| minutechagehour(v)                 		| ºÐ´ÜÀ§¸¦ ½Ã°£´ÜÀ§À¸·Î Â©¶ó¼­ ¹ÝÈ¯                      					|
'|                                          | »ç¿ë¿¹ : minutechagehour(v)             									|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| BinaryToText(BinaryData, CharSet)         | ¹ÙÀÌ³Ê¸® µ¥ÀÌÅÍ TEXTÇüÅÂ·Î º¯È¯                                          |
'|                                           | »ç¿ë¿¹ : BinaryToText(objXML.ResponseBody, "euc-kr")                     |
'+------------------------------------------+---------------------------------------------------------------------------+
'| URLEncodeUTF8(byVal szSource)            | ASCIIÀ» UTF8 ¹®ÀÚ¿­·Î º¯È¯                                                |
'|                                          | »ç¿ë¿¹ : strUF8 = URLEncodeUTF8(STR)                                      |
'+------------------------------------------+---------------------------------------------------------------------------+
'| URLDecodeUTF8(byVal pURL)                | UTF8À» ASCII ¹®ÀÚ¿­·Î º¯È¯                                                |
'|                                          | »ç¿ë¿¹ : strASC = URLDecodeUTF8(URL)                                      |
'+------------------------------------------+---------------------------------------------------------------------------+
'| chkArrValue(aVal,cVal)                    | ÄÞ¸¶·Î ±¸ºÐµÈ ¹è¿­°ª¿¡ ÁöÁ¤µÈ °ªÀÌ ÀÖ´ÂÁö ¹ÝÈ¯                           |
'|                                           | »ç¿ë¿¹ : chkArrValue("A,B,C", "B") ¡æ true                                |
'+-------------------------------------------+--------------------------------------------------------------------------+

function G_IsLocalDev()
	G_IsLocalDev = (application("Svr_Info")="Dev") AND (request.ServerVariables("LOCAL_ADDR")="::1" or request.ServerVariables("LOCAL_ADDR")="127.0.0.1")
end function

''ÄíÅ°¿¡ ³ÖÀ»¶§ »ç¿ë / ¾ÆÀÌµð ´Ü¹æÇâ ÇØ½¬°ª : »ç¿ë½Ã md5 ÇÊ¿ä. (md5 ºÎÇÏ ½ÉÇÒ°æ¿ì component, db ÀÌ¿ë °¡´É)
function HashTenID(byval oid)
    dim orgid : orgid = LCASE(oid)
    dim hashid

    HashTenID = orgid
    if Len(orgid)<1 then Exit function      ''ºó°ªÀÎ°æ¿ì ¿ø·¡°ª
    if Len(orgid)<2 then orgid=orgid+"1"    ''±æÀÌ°¡1ÀÏ°æ¿ì ¿À·ùÇÇÇÔ.


    hashid = Right(orgid,4) + Left(orgid,Len(orgid)-1)
    hashid = Right(hashid,5) + Left(hashid,Len(hashid)-2)
    hashid = Right(hashid,6) + Left(hashid,Len(hashid)-3)
    hashid = Right(hashid,7) + Left(hashid,Len(hashid)-4)
    hashid = Right(hashid,8) + Left(hashid,Len(hashid)-5)
    HashTenID = MD5(hashid)

end function

'// ³¯Â¥¸¦ ÁöÁ¤µÈ ¹®ÀÚÇüÀ¸·Î º¯È¯ //
function FormatDate(ddate, formatstring)
	dim s
	Select Case formatstring
		Case "0000-00-00 00:00:00"
			s = CStr(year(ddate)) & "-" &_
				Num2Str(month(ddate),2,"0","R") & "-" &_
				Num2Str(day(ddate),2,"0","R") & " " &_
				Num2Str(hour(ddate),2,"0","R") & ":" &_
				Num2Str(minute(ddate),2,"0","R") & ":" &_
				Num2Str(Second(ddate),2,"0","R")
		Case "0000.00.00"
			s = CStr(year(ddate)) & "." &_
				Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R")
		Case "0000-00-00"
			s = CStr(year(ddate)) & "-" &_
				Num2Str(month(ddate),2,"0","R") & "-" &_
				Num2Str(day(ddate),2,"0","R")
		Case "00000000"
			s = CStr(year(ddate)) &_
				Num2Str(month(ddate),2,"0","R") &_
				Num2Str(day(ddate),2,"0","R")
		Case "00000000000000"
			s = CStr(year(ddate))  &_
				Num2Str(month(ddate),2,"0","R") &_
				Num2Str(day(ddate),2,"0","R")  &_
				Num2Str(hour(ddate),2,"0","R")  &_
				Num2Str(minute(ddate),2,"0","R") &_
				Num2Str(Second(ddate),2,"0","R")
		Case "000000000000"
			s = CStr(year(ddate))  &_
				Num2Str(month(ddate),2,"0","R") &_
				Num2Str(day(ddate),2,"0","R")  &_
				Num2Str(hour(ddate),2,"0","R")  &_
				Num2Str(minute(ddate),2,"0","R")
		Case "0000.00"
			s = CStr(year(ddate)) & "." &_
				Num2Str(month(ddate),2,"0","R")
		Case "0000.00.00-00:00:00"
			s = CStr(year(ddate)) & "." &_
				Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R") & "-" &_
				Num2Str(hour(ddate),2,"0","R") & ":" &_
				Num2Str(minute(ddate),2,"0","R") & ":" &_
				Num2Str(Second(ddate),2,"0","R")
		Case "0000.00.00 00:00:00"
			s = CStr(year(ddate)) & "." &_
				Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R") & " " &_
				Num2Str(hour(ddate),2,"0","R") & ":" &_
				Num2Str(minute(ddate),2,"0","R") & ":" &_
				Num2Str(Second(ddate),2,"0","R")
		Case "0000/00/00"
			s = CStr(year(ddate)) & "/" &_
				Num2Str(month(ddate),2,"0","R") & "/" &_
				Num2Str(day(ddate),2,"0","R")
		Case "00/00/00"
			s = Num2Str(year(ddate),2,"0","R") & "/" &_
				Num2Str(month(ddate),2,"0","R") & "/" &_
				Num2Str(day(ddate),2,"0","R")
		Case "00.00.00"
			s = Num2Str(year(ddate),2,"0","R") & "." &_
				Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R")
		Case "00/00"
			s = Num2Str(month(ddate),2,"0","R") & "/" &_
				Num2Str(day(ddate),2,"0","R")
		Case "00.00"
			s = Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R")
		Case "0000.00.00-00:00"
			s = CStr(year(ddate)) & "." &_
				Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R") & "-" &_
				Num2Str(hour(ddate),2,"0","R") & ":" &_
				Num2Str(minute(ddate),2,"0","R")
		Case "0000/00/00/00:00"
			s = CStr(year(ddate)) & "/" &_
				Num2Str(month(ddate),2,"0","R") & "/" &_
				Num2Str(day(ddate),2,"0","R") & "/" &_
				Num2Str(hour(ddate),2,"0","R") & ":" &_
				Num2Str(minute(ddate),2,"0","R")
		Case "0000-00-00T00:00Z"
			s = CStr(year(ddate)) & "-" &_
				Num2Str(month(ddate),2,"0","R") & "-" &_
				Num2Str(day(ddate),2,"0","R") & "T" &_
				Num2Str(hour(ddate),2,"0","R") & ":" &_
				Num2Str(minute(ddate),2,"0","R") & "Z"
		Case Else
			s = CStr(ddate)
	End Select

	FormatDate = s
end function

function GetImageSubFolderByItemid(byval iitemid)
	IF iitemid<>"" THEN
	GetImageSubFolderByItemid = Num2Str(CStr(Clng(iitemid) \ 10000),2,"0","R")
	END IF
end function

'' ±âÁ¸ µðºñ¿¡ ÀÌÀü Çü½Ä ÀÖÀ½.. Â÷ÈÄ »èÁ¦
function db2html(checkvalue)
	dim v
	v = checkvalue
	if Isnull(v) then Exit function

    On Error resume Next
    v = replace(v, "&amp;", "&")
    v = replace(v, "&lt;", "<")
    v = replace(v, "&gt;", ">")
    v = replace(v, "&quot;", "'")
    v = Replace(v, "", "<br>")
    v = Replace(v, "\0x5C", "\")
    v = Replace(v, "\0x22", "'")
    v = Replace(v, "\0x25", "'")
    v = Replace(v, "\0x27", "%")
    v = Replace(v, "\0x2F", "/")
    v = Replace(v, "\0x5F", "_")
    ''checkvalue = Replace(checkvalue, vbcrlf,"<br>")
    db2html = v
end function

'' 2008 03 ¼öÁ¤ - Eastone
function html2db(checkvalue)
	html2db = Newhtml2db(checkvalue)
end function

function Newhtml2db(checkvalue)
	dim v
	v = checkvalue
	if Isnull(v) then Exit function
	v = Replace(v, "'", "''")
	Newhtml2db = v
end function

function html2db2017(checkvalue)
	dim v
	v = checkvalue
	if Isnull(v) then Exit function
	v = replace(v, "'", "`")
	v = replace(v, """", "")
	html2db2017 = v
end function

function nl2br(v)
	if IsNull(v) then
		nl2br = ""
		Exit function
	end if

    nl2br = Replace(v, vbcrlf,"<br />")
    nl2br = Replace(v, vbCr,"<br />")
    nl2br = Replace(v, vbLf,"<br />")
end function

'// ¹®ÀÚ¿­³» CR/LF¸¦ °ø¹éÀ¸·Î Ä¡È¯ //
function nl2blank(v)
	if IsNull(v) then
		nl2blank = ""
		Exit function
	end if

    nl2blank = Replace(v, vbcrlf,"")
end function

function CurrFormat(byVal v)
        if ((v = "") or (isnull(v))) then
                CurrFormat = 0
        else
                CurrFormat = FormatNumber(FormatCurrency(v),0)
        end if
end function


function Format00(n,orgData)
    dim tmp

    if IsNULL(orgData) then Exit function

	if (n-Len(CStr(orgData))) < 0 then
		Format00 = CStr(orgData)
		Exit Function
	end if

	tmp = String(n-Len(CStr(orgData)), "0") & CStr(orgData)
	Format00 = tmp
end function


function FormatCode(itemcode)
    if isNULL(itemcode) then
        FormatCode = itemcode
        Exit function
    end if

    if (itemcode>=1000000) then
        FormatCode = Format00(8,itemcode)
    else
	    FormatCode = Format00(6,itemcode)
    end if
end function


function GetCurrentTimeFormat()
	dim d
	d = now
	GetCurrentTimeFormat = replace(Left(FormatDateTime(d,2),7),"-","") + Format00(2,Day(d)) + Format00(2,Hour(d)) + Format00(2,Minute(d))  +  Format00(2,Second(d))

end function


function GetListImageUrl(byval itemid)
	GetListImageUrl = "/image/list/L" + Format00(9,itemid) + ".jpg"
end function


function DDotFormat(byval str,byval n)
	DDotFormat = str
	if Len(str)> n then
		DDotFormat = Left(str,n) + "..."
	end if
end function


function stripHTML(strng)
   Dim regEx
   Set regEx = New RegExp
   regEx.Pattern = "[<][^>]*[>]"
   regEx.IgnoreCase = True
   regEx.Global = True
   stripHTML = regEx.Replace(strng, " ")
   Set regEx = nothing
End Function

function Format00(n,orgData)
    dim tmp

    if IsNULL(orgData) then Exit function

	if (n-Len(CStr(orgData))) < 0 then
		Format00 = CStr(orgData)
		Exit Function
	end if

	tmp = String(n-Len(CStr(orgData)), "0") & CStr(orgData)
	Format00 = tmp
end function

function getFileExtention(strFile)
	Dim file_length, file_point, ext_len

	if Not(strFile="" or isNull(strFile)) then
		file_length = LEN(strFile)
		file_point = inStrRev(strFile,".") + 1
		ext_len = file_length - file_point + 1

		getFileExtention = Lcase(MID(strFile,file_point,ext_len))
	end if
End Function

function adminColor(v)
	adminColor = "#FFFFFF"

	if v="menubar" then
		adminColor = "#DEDFFF"
	elseif v="menubar_left" then
		adminColor = "#CCCCCC"
	elseif v="topbar" then
		adminColor = "#F4F4F4"
	elseif v="tabletop" then
		adminColor = "#E6E6E6"
	elseif v="tablebg" then
		adminColor = "#999999"

	elseif v="pink" then
		adminColor = "#FFDDDD"
	elseif v="green" then
		adminColor = "#DDFFDD"
	elseif v="sky" then
		adminColor = "#DDDDFF"
	elseif v="gray" then
		adminColor = "#EEEEEE"
	elseif v="dgray" then
		adminColor = "#CCCCCC"

	else

	end if
end function

	'// ¼ýÀÚ¸¦ ÁöÁ¤ÇÑ ±æÀÌÀÇ ¹®ÀÚ¿­·Î ¹ÝÈ¯ //
	Function Num2Str(inum,olen,cChr,oalign)
		Dim i, ilen, strChr

		ilen = len(Cstr(inum))
		strChr = ""

		if ilen < olen then
			for i=1 to olen-ilen
				strChr = strChr & cChr
			next
		end if

		'°áÇÕ¹æ¹ý¿¡µû¸¥ °á°ú ºÐ±â
		if oalign="L" then
			'¿ÞÂÊ±âÁØ
			Num2Str = inum & strChr
		else
			'¿À¸¥ÂÊ ±âÁØ (±âº»°ª)
			Num2Str = strChr & inum
		end if

    End Function


'// ¹®ÀÚ¿­À» Àß¶ó ¿øÇÏ´Â À§Ä¡ÀÇ °ªÀ» ¹ÝÈ¯ //
function SplitValue(orgStr,delim,pos)
    dim buf
    SplitValue = ""
    if IsNULL(orgStr) then Exit function
    if (Len(delim)<1) then Exit function
    buf = split(orgStr,delim)

    if UBound(buf)<pos then Exit function

    SplitValue = buf(pos)
end function


'// ÆÄ¶ó¸ÞÅÍ ±æÀÌ Ã¼Å© ÈÄ Maxlen ÀÌÇÏ·Î µ¹·ÁÁÜ Code, id µîÀÇ Param ¿¡ »ç¿ë //
function requestCheckVar(orgval,maxlen)
	requestCheckVar = trim(orgval)
	requestCheckVar = replace(requestCheckVar,"'","")
	requestCheckVar = replace(requestCheckVar,"--","")
	requestCheckVar = Left(requestCheckVar,maxlen)
end function

function requestCheckVarNoTrim(orgval,maxlen)
	requestCheckVarNoTrim = orgval
	requestCheckVarNoTrim = replace(requestCheckVarNoTrim,"'","")
	requestCheckVarNoTrim = replace(requestCheckVarNoTrim,"--","")
	requestCheckVarNoTrim = Left(requestCheckVarNoTrim,maxlen)
end function


'// °ªºñ±³ ÈÄ Return °ª like iif function
Function ChkIIF(trueOrFalse, trueVal, falseVal)
	if (trueOrFalse) then
	    ChkIIF = trueVal
	else
	    ChkIIF = falseVal
	end if
End Function

'// °æ°í¹® Ãâ·ÂÈÄ µÚ·Î°¡±â //
Sub Alert_return(strMSG)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"history.back();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub


'// °æ°í¹® Ãâ·ÂÈÄ Ã¢´Ý±â //
Sub Alert_close(strMSG)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"self.close();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub


'// °æ°í¹® Ãâ·ÂÈÄ ÁöÁ¤ ÆäÀÌÁö·Î ÀÌµ¿ //
Sub Alert_move(strMSG,targetURL)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"self.location.replace('" & targetURL & "');" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub

'// ÁöÁ¤±æÀÌ·Î ¹®ÀÚ¿­ ÀÚ¸£±â //
Function chrbyte(str,chrlen,dot)

    Dim charat, wLen, cut_len, ext_chr, cblp

    if IsNULL(str) then Exit function

    for cblp=1 to len(str)
        charat=mid(str, cblp, 1)
        if asc(charat)>0 and asc(charat)<255 then
            wLen=wLen+1
        else
            wLen=wLen+2
        end if

        if wLen >= cint(chrlen) then
           cut_len = cblp
           exit for
        end if
    next

    if len(cut_len) = 0 then
        cut_len = len(str)
    end if

	if len(str)>cut_len and dot="Y" then
		ext_chr = "..."
	else
		ext_chr = ""
	end if

    chrbyte = Trim(left(str,cut_len)) & ext_chr

end function

'// ÆÐ½º¿öµå º¹Àâ¼º °Ë»ç ÇÔ¼ö(±âÁ¸¹öÀü ±æÀÌ6, ·Î±×ÀÎ½Ã Ã¼Å©ÇÔ.)		//2017.09.25 ÇÑ¿ë¹Î »ý¼º
Function chkPasswordComplex_Len6Ver(uid,pwd)
	dim msg, i, sT, sN, numAlpha, numNums, numSpecials, buf, index
    numAlpha = 0
    numNums = 0
    numSpecials = 0
	msg = ""

	'ºñ¹Ð¹øÈ£ ±æÀÌ °Ë»ç
	if len(pwd)<8 then
		msg = msg & "- ºñ¹Ð¹øÈ£´Â ÃÖ¼Ò 8ÀÚ¸®ÀÌ»óÀ¸·Î ÀÔ·ÂÇØÁÖ¼¼¿ä.\n"
	end if

	'¾ÆÀÌµð¿Í µ¿ÀÏ ¶Ç´Â Æ÷ÇÔÇÏ°í ÀÖ´Â°¡?
	if instr(lcase(pwd),lcase(uid))>0 then
		msg = msg & "- ¾ÆÀÌµð¿Í µ¿ÀÏÇÏ°Å³ª ¾ÆÀÌµð¸¦ Æ÷ÇÔÇÏ°í ÀÖ´Â ºñ¹Ð¹øÈ£ÀÔ´Ï´Ù.\n"
	end if

	'## º¹Àâ¼ºÀ» ¸¸Á·ÇÏ´Â°¡?
	'°°Àº¹®ÀÚ 3¹ø ¿¬¼Ó ±ÝÁö
	sT=""
	sN=0
	for i=1 to len(pwd)
		if st=mid(pwd,i,1) then
			sN = sN +1
		else
			sN = 0
		end if
		st = mid(pwd,i,1)
		if sN>=2 then
			msg = msg & "- °°Àº¹®ÀÚ°¡ 3¹ø ¿¬¼ÓÀ¸·Î ¾²¿´½À´Ï´Ù.\n"
			exit for
		end if
	next

'Á¤±Ô½Ä ¶È¹Ù·Î ¾È¹¬³×. ¸Ó²¿
'	if chkWord(pwd,"[^-a-zA-Z]") then
'		numAlpha = numAlpha + 1
'	end if
'	if chkWord(pwd,"[^-0-9 ]") then
'		numNums = numNums + 1
'	end if
'	if chkWord(pwd,"[~!@\#$%<>^&*\()\-=+_\¡¯]") then
'		numSpecials = numSpecials + 1
'	end if

	index = 1
	do until index > len(pwd)
	    buf = mid(pwd, index, cint(1))
	    if (lcase(buf) >= "a" and lcase(buf) <= "z") then
			numAlpha = numAlpha + 1
	    elseif (buf >= "0" and buf <= "9") then
			numNums = numNums + 1
	    else
			numSpecials = numSpecials + 1
	    end if
	    index = index + 1
	loop

	'// 3°¡Áö Á¶ÇÕ
    if (numAlpha>0 and numNums>0 and numSpecials>0) then
    	if (len(pwd) >= 8) then
    	else
    		msg = msg & "- »õ·Î¿î ÆÐ½º¿öµå´Â ¿µ¹®/¼ýÀÚ/Æ¯¼ö¹®ÀÚ µî µÎ°¡Áö ÀÌ»óÀÇ Á¶ÇÕÀ¸·Î ÀÔ·ÂÇÏ¼¼¿ä. ÃÖ¼Ò±æÀÌ 10ÀÚ(2Á¶ÇÕ) , 8ÀÚ(3Á¶ÇÕ)\n"
    	end if

	'// 2°¡Áö Á¶ÇÕ
    elseif ((numAlpha>0 and numNums>0) or (numAlpha>0 and numSpecials>0) or (numNums>0 and numSpecials>0)) then
    	if (len(pwd) >= 10) then
    	else
    		msg = msg & "- »õ·Î¿î ÆÐ½º¿öµå´Â ¿µ¹®/¼ýÀÚ/Æ¯¼ö¹®ÀÚ µî µÎ°¡Áö ÀÌ»óÀÇ Á¶ÇÕÀ¸·Î ÀÔ·ÂÇÏ¼¼¿ä. ÃÖ¼Ò±æÀÌ 10ÀÚ(2Á¶ÇÕ) , 8ÀÚ(3Á¶ÇÕ)\n"
    	end if

    else
    	msg = msg & "- »õ·Î¿î ÆÐ½º¿öµå´Â ¿µ¹®/¼ýÀÚ/Æ¯¼ö¹®ÀÚ µî µÎ°¡Áö ÀÌ»óÀÇ Á¶ÇÕÀ¸·Î ÀÔ·ÂÇÏ¼¼¿ä. ÃÖ¼Ò±æÀÌ 10ÀÚ(2Á¶ÇÕ) , 8ÀÚ(3Á¶ÇÕ)\n"
    end if

	'°á°ú ¹ÝÈ¯
	chkPasswordComplex_Len6Ver = msg
end Function

'// ÆÐ½º¿öµå º¹Àâ¼º °Ë»ç ÇÔ¼ö		//2017.09.25 ÇÑ¿ë¹Î »ý¼º
Function chkPasswordComplex(uid,pwd)
	dim msg, i, sT, sN, numAlpha, numNums, numSpecials, buf, index
    numAlpha = 0
    numNums = 0
    numSpecials = 0
	msg = ""

	'ºñ¹Ð¹øÈ£ ±æÀÌ °Ë»ç
	if len(pwd)<8 then
		msg = msg & "- ºñ¹Ð¹øÈ£´Â ÃÖ¼Ò 8ÀÚ¸®ÀÌ»óÀ¸·Î ÀÔ·ÂÇØÁÖ¼¼¿ä.\n"
	end if

	'¾ÆÀÌµð¿Í µ¿ÀÏ ¶Ç´Â Æ÷ÇÔÇÏ°í ÀÖ´Â°¡?
	if instr(lcase(pwd),lcase(uid))>0 then
		msg = msg & "- ¾ÆÀÌµð¿Í µ¿ÀÏÇÏ°Å³ª ¾ÆÀÌµð¸¦ Æ÷ÇÔÇÏ°í ÀÖ´Â ºñ¹Ð¹øÈ£ÀÔ´Ï´Ù.\n"
	end if

	'## º¹Àâ¼ºÀ» ¸¸Á·ÇÏ´Â°¡?
	'°°Àº¹®ÀÚ 3¹ø ¿¬¼Ó ±ÝÁö
	sT=""
	sN=0
	for i=1 to len(pwd)
		if st=mid(pwd,i,1) then
			sN = sN +1
		else
			sN = 0
		end if
		st = mid(pwd,i,1)
		if sN>=2 then
			msg = msg & "- °°Àº¹®ÀÚ°¡ 3¹ø ¿¬¼ÓÀ¸·Î ¾²¿´½À´Ï´Ù.\n"
			exit for
		end if
	next

'Á¤±Ô½Ä ¶È¹Ù·Î ¾È¹¬³×. ¸Ó²¿
'	if chkWord(pwd,"[^-a-zA-Z]") then
'		numAlpha = numAlpha + 1
'	end if
'	if chkWord(pwd,"[^-0-9 ]") then
'		numNums = numNums + 1
'	end if
'	if chkWord(pwd,"[~!@\#$%<>^&*\()\-=+_\¡¯]") then
'		numSpecials = numSpecials + 1
'	end if

	index = 1
	do until index > len(pwd)
	    buf = mid(pwd, index, cint(1))
	    if (lcase(buf) >= "a" and lcase(buf) <= "z") then
			numAlpha = numAlpha + 1
	    elseif (buf >= "0" and buf <= "9") then
			numNums = numNums + 1
	    else
			numSpecials = numSpecials + 1
	    end if
	    index = index + 1
	loop

	'// 3°¡Áö Á¶ÇÕ
    if (numAlpha>0 and numNums>0 and numSpecials>0) then
    	if (len(pwd) >= 8) then
    	else
    		msg = msg & "- »õ·Î¿î ÆÐ½º¿öµå´Â ¿µ¹®/¼ýÀÚ/Æ¯¼ö¹®ÀÚ µî µÎ°¡Áö ÀÌ»óÀÇ Á¶ÇÕÀ¸·Î ÀÔ·ÂÇÏ¼¼¿ä. ÃÖ¼Ò±æÀÌ 10ÀÚ(2Á¶ÇÕ) , 8ÀÚ(3Á¶ÇÕ)\n"
    	end if

	'// 2°¡Áö Á¶ÇÕ
    elseif ((numAlpha>0 and numNums>0) or (numAlpha>0 and numSpecials>0) or (numNums>0 and numSpecials>0)) then
    	if (len(pwd) >= 10) then
    	else
    		msg = msg & "- »õ·Î¿î ÆÐ½º¿öµå´Â ¿µ¹®/¼ýÀÚ/Æ¯¼ö¹®ÀÚ µî µÎ°¡Áö ÀÌ»óÀÇ Á¶ÇÕÀ¸·Î ÀÔ·ÂÇÏ¼¼¿ä. ÃÖ¼Ò±æÀÌ 10ÀÚ(2Á¶ÇÕ) , 8ÀÚ(3Á¶ÇÕ)\n"
    	end if

    else
    	msg = msg & "- »õ·Î¿î ÆÐ½º¿öµå´Â ¿µ¹®/¼ýÀÚ/Æ¯¼ö¹®ÀÚ µî µÎ°¡Áö ÀÌ»óÀÇ Á¶ÇÕÀ¸·Î ÀÔ·ÂÇÏ¼¼¿ä. ÃÖ¼Ò±æÀÌ 10ÀÚ(2Á¶ÇÕ) , 8ÀÚ(3Á¶ÇÕ)\n"
    end if

	'°á°ú ¹ÝÈ¯
	chkPasswordComplex = msg
end Function

'// ÆÐ½º¿öµå º¹Àâ¼º °Ë»ç ÇÔ¼ö
Function chkPasswordComplexNonID(pwd)
	dim msg, i, sT, sN
	msg = ""

	'ºñ¹Ð¹øÈ£ ±æÀÌ °Ë»ç
	if len(pwd)<8 then
		msg = msg & "- ºñ¹Ð¹øÈ£´Â ÃÖ¼Ò 8ÀÚ¸®ÀÌ»óÀ¸·Î ÀÔ·ÂÇØÁÖ¼¼¿ä.\n"
	end if


	'## º¹Àâ¼ºÀ» ¸¸Á·ÇÏ´Â°¡?
	'°°Àº¹®ÀÚ 3¹ø ¿¬¼Ó ±ÝÁö
	sT=""
	sN=0
	for i=1 to len(pwd)
		if st=mid(pwd,i,1) then
			sN = sN +1
		else
			sN = 0
		end if
		st = mid(pwd,i,1)
		if sN>=2 then
			msg = msg & "- °°Àº¹®ÀÚ°¡ 3¹ø ¿¬¼ÓÀ¸·Î ¾²¿´½À´Ï´Ù.\n"
			exit for
		end if
	next
	'¿µ¹®/¼ýÀÚÀÇ Á¶ÇÕ
	if chkWord(pwd,"[^-a-zA-Z]") or chkWord(pwd,"[^-0-9 ]") then
		msg = msg & "- ºñ¹Ð¹øÈ£´Â ¹Ýµå½Ã ¾ËÆÄºª°ú ¼ýÀÚ¸¦ Á¶ÇÕÇØ¼­ ¸¸µé¾î¾ßÇÕ´Ï´Ù.\n"
	end if

	'°á°ú ¹ÝÈ¯
	chkPasswordComplexNonID = msg
end Function

'//Á¤±Ô½Ä ¹®ÀÚ¿­ °Ë»ç
Function chkWord(str,patrn)
    Dim regEx, match, matches

    SET regEx = New RegExp

    regEx.Pattern = patrn            ' ÆÐÅÏÀ» ¼³Á¤ÇÕ´Ï´Ù.
    regEx.IgnoreCase = True      ' ´ë/¼Ò¹®ÀÚ¸¦ ±¸ºÐÇÏÁö ¾Êµµ·Ï ÇÕ´Ï´Ù.
    regEx.Global = True             ' ÀüÃ¼ ¹®ÀÚ¿­À» °Ë»öÇÏµµ·Ï ¼³Á¤ÇÕ´Ï´Ù.

    SET Matches = regEx.Execute(str)

    if 0 < Matches.count then
        chkWord = false
    Else
        chkWord = true
    end if
End Function

'// ÀüÈ­¹øÈ£¿¡ ´ë½Ã Ãß°¡
function ParsingPhoneNumber(orgnum)
    dim noDashNum, PreNum, CuttedNum
    noDashNum = Replace(orgnum,"-","")

    ParsingPhoneNumber = noDashNum

    if Len(noDashNum)<7 then
        exit function
    end if

    if Len(noDashNum)=7 then
        ParsingPhoneNumber = Left(noDashNum,3) & "-" & Right(noDashNum,4)
        Exit function
    end if

    if Len(noDashNum)=8 then
        ParsingPhoneNumber = Left(noDashNum,4) & "-" & Right(noDashNum,4)
        Exit function
    end if

    if (Left(noDashNum,1)<>"0") then
        Exit function
    end if

    PreNum = Left(noDashNum,2)
    if (PreNum="02") then
        CuttedNum = Mid(noDashNum,3,255)
    else
        PreNum = Left(noDashNum,3)
        if (PreNum="010") or (PreNum="011") or (PreNum="016") or (PreNum="017") or (PreNum="019") then
            CuttedNum = Mid(noDashNum,4,255)
        else
            CuttedNum = Mid(noDashNum,4,255)
        end if
    end if

    if Len(CuttedNum)=7 then
        ParsingPhoneNumber = PreNum & "-" & Left(CuttedNum,3) & "-" & Right(CuttedNum,4)
    elseif Len(CuttedNum)=8 then
        ParsingPhoneNumber = PreNum & "-" & Left(CuttedNum,4) & "-" & Right(CuttedNum,4)
    else
        exit function
    end if
end function


'''''==================  2009 Ãß°¡

' response.write ÇÔ¼ö
Function rw(ByVal str)
	response.write str & "<br>"
End Function

' NullÀ» °ø¹éÀ¸·Î Ä¡È¯
Function null2blank(ByVal v)
	If IsNull(v) Then
		null2blank = ""
	Else
		null2blank = v
	End If
End Function

'// Å«µû¿ÈÇ¥ input ¹Ú½º value=""¿¡ »ç¿ëÇÒ¶§ Ä¡È¯
Function doubleQuote(ByVal v)
	If IsNull(v) Then
		doubleQuote = ""
	Else
		doubleQuote = Replace(v, """","&quot;")
	End If
End Function


' request ´ëÃ¼ ÇÔ¼ö(ÆÄ¶ó¹ÌÅÍ¸í, µðÆúÆ®°ª)
Function req(ByVal param, ByVal value)
'	VarType Return °ª
'	0 (°ø¹é)
'	1 (³Î)
'	2 integer
'	3 Long
'	4 Single
'	5 Double
'	6 Currency
'	7 Date
'	8 String
'	9 OLE Object
'	10 Error
'	11 Boolean
'	12 Variant
'	13 Non-OLE Object
'	17 Byte
'	8192 Array

	Dim tmpValue

	If VarType(value) = 2 Or VarType(value) = 3 Or VarType(value) = 4 Or VarType(value) = 5 Or VarType(value) = 6 Then
		tmpValue = Replace(Trim(Request(param)),",","")
		If Not IsNumeric(tmpValue) Then	' ¼ýÀÚ°¡ ¾Æ´Ï¸é
			tmpValue = value
		End If
		tmpValue = CDbl(tmpValue)
	Else
		tmpValue = Trim(Request(param))
		If tmpValue = "" Then			' Request°ªÀÌ ¾øÀ¸¸é
			tmpValue = value
		End If
	End If
	req = tmpValue

End Function

Sub sbDisplayPaging(ByVal strCurrentPage, ByVal intTotalRecord, ByVal intRecordPerPage, ByVal intBlockPerPage)

	'º¯¼ö ¼±¾ð
	Dim intCurrentPage, strCurrentPath
	Dim intStartBlock, intEndBlock, intTotalPage
	Dim strParamName, intLoop

	'ÇöÀç ÆäÀÌÁö ¼³Á¤
	intCurrentPage = Mid(strCurrentPage, InStr(strCurrentPage, "=")+1)		'ÇöÀç ÆäÀÌÁö °ª
	strCurrentPage = Left(strCurrentPage, InStr(strCurrentPage, "=")-1)		'ÆäÀÌÁö Æû°ª º¯¼ö¸í

	'ÇöÀç ÆäÀÌÁö ¸í
	strCurrentPath = Request.ServerVariables("Script_Name")

	'ÇØ´çÆäÀÌÁö¿¡ Ç¥½ÃµÇ´Â ½ÃÀÛÆäÀÌÁö¿Í ¸¶Áö¸·ÆäÀÌÁö ¼³Á¤
	intStartBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + 1
	intEndBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + intBlockPerPage

	'ÃÑ ÆäÀÌÁö ¼ö ¼³Á¤
	intTotalPage =  -(int(-(intTotalRecord/intRecordPerPage)))

	'Æû ¼³Á¤ & hidden ÆÄ¶ó¹ÌÅÍ ¼³Á¤
	Response.Write	"<form name='frmPaging' method='get' action ='" & strCurrentPath & "'>" &_
							"<input type='hidden' name='" & strCurrentPage & "'>"			'ÇöÀç ÆäÀÌÁö

	'ÆÄ¶ó¹ÌÅÍ °ªµé(¿¹: °Ë»ö¾î)À» hidden ÆÄ¶ó¹ÌÅÍ·Î ÀúÀåÇÑ´Ù
	strParamName = ""
	For Each strParamName In Request.Form
		If strParamName <> strCurrentPage Then

			'hidden ÆÄ¶ó¹ÌÅÍ °ªµµ ÆÄ¶ó¹ÌÅÍ °Ë¿­
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.Form(strParamName),50) & "'>"
		End If
	Next
	strParamName = ""

	For Each strParamName In Request.Querystring
		If strParamName <> strCurrentPage Then
			'hidden ÆÄ¶ó¹ÌÅÍ °ªµµ ÆÄ¶ó¹ÌÅÍ °Ë¿­
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.QueryString(strParamName),50) & "'>"
		END IF
	Next

	Response.Write "<table border='0' cellpadding='0' cellspacing='0' class=a><tr align='center'><td>"

	'ÀÌÀü ÆäÀÌÁö ÀÌ¹ÌÁö ¼³Á¤
	If intStartBlock > 1 Then
		Response.Write "<img src='http://fiximage.10x10.co.kr/web2008/designfingers/btn_pageprev01.gif' border='0' style='cursor:hand' alt='ÀÌÀü " & intBlockPerPage & " ÆäÀÌÁö'" &_
							   "onClick='javascript:document.frmPaging." & strCurrentPage & ".value=" & intStartBlock - intBlockPerPage & ";document.frmPaging.submit();'>"
	Else
		Response.Write "<img src='http://fiximage.10x10.co.kr/web2009/common/btn_pageprev01.gif' border='0' >"
	End If

	Response.Write "</td><td>&nbsp;"

	'ÆäÀÌÂ¡ Ãâ·Â
	If intTotalPage > 1 Then
		For intLoop = intStartBlock To intEndBlock
			If intLoop > intTotalPage Then Exit For

			If Int(intLoop) <> Int(intStartBlock) Then Response.Write "|"

			If Int(intLoop) = Int(intCurrentPage) Then		'ÇöÀç ÆäÀÌÁö
				Response.Write "&nbsp;<span class='text01'><strong>" & intLoop & "</strong></span>&nbsp;"
			Else															'±× ¿Ü ÆäÀÌÁö
				Response.Write "&nbsp;<a href='javascript:document.frmPaging." & strCurrentPage & ".value=" & intLoop & ";document.frmPaging.submit();'><font class='text01'>" & intLoop & "</font></a>&nbsp;"
			End If

		Next
	Else		'ÇÑ ÆäÀÌÁö¸¸ Á¸Àç ÇÒ¶§
		Response.Write "&nbsp;<span class='text01'><strong>1</strong></span>&nbsp;"
	End If

	Response.Write "&nbsp;</td><td>"

	'´ÙÀ½ ÆäÀÌÁö ÀÌ¹ÌÁö ¼³Á¤
	If Int(intEndBlock) < Int(intTotalPage) Then
		Response.Write "<img src='http://fiximage.10x10.co.kr/web2008/designfingers/btn_pagenext01.gif' border='0' style='cursor:hand' alt='´ÙÀ½ " & intBlockPerPage & " ÆäÀÌÁö'" &_
							   "onClick='javascript:document.frmPaging." & strCurrentPage & ".value=" & intEndBlock+1 & ";document.frmPaging.submit();'>"
	Else
	    Response.Write "<img src='http://fiximage.10x10.co.kr/web2009/common/btn_pagenext01.gif' border='0' >"
	End If

	Response.Write "</td></tr></form></table>"

End Sub



' µî·Ï,¼öÁ¤,»èÁ¦ ¸ðµå ÅØ½ºÆ® ¸®ÅÏ
Function getModeName(ByVal mode)
    Select Case mode
        Case "INS"	: getModeName = "µî·Ï"
        Case "UPD"	: getModeName = "¼öÁ¤"
        Case "DEL"	: getModeName = "»èÁ¦"
        Case "FIN"	: getModeName = "¿Ï·á"
        Case Else	: getModeName = "¹ÌÁ¤"
    End Select
End Function

'// ²©Àº°ýÈ£ HTMLÄÚµå·Î Ä¡È¯ //
' db2html ÀÌ¶û Ãæµ¹³ª¼­ »ç¿ë°¡´ÉÇÑ°÷¸¸ Àû¿ëÇÏ¼¼¿ä.
Function ReplaceBracket(strng)
	if isnull(strng) then exit Function

	strng = Replace(strng,"<","&lt;")
	strng = Replace(strng,">","&gt;")
	ReplaceBracket = strng
end Function

'// ²©Àº°ýÈ£ ´Ù¸¥ °ýÈ£·Î Ä¡È¯ //
Function ReplaceBracketOther(strng)
	if isnull(strng) then exit Function

	strng = Replace(strng,"<","[")
	strng = Replace(strng,">","]")
	ReplaceBracketOther = strng
end Function

'// Script TagÄ¡È¯ //
Function ReplaceScript(strng)
	if isnull(strng) then exit Function

	strng = Replace(strng,"<script","[script")
	strng = Replace(strng,"</script","[/script")
	strng = Replace(strng,"<iframe","[iframe")
	strng = Replace(strng,"</iframe","[/iframe")
	ReplaceScript = strng
end Function


' Á¤±Ô½Ä ÇÔ¼ö
Function ReplaceText(str, patrn, repStr)
	Dim regEx
	Set regEx = New RegExp
	with regEx
		.Pattern = patrn
		.IgnoreCase = True
		.Global = True
	End with
	ReplaceText = regEx.Replace(str, repStr)
End Function

Function TwoNumber(number)
	Dim vNumber
	If len(number) = 1 Then
		vNumber = "0" & number
	Else
		vNumber = number
	End If
	TwoNumber = vNumber
End Function

'// ¹®ÀÚ¿­¿¡¼­ ¼ýÀÚ¸¸ ÃßÃâ º¯È¯
Function getNumeric(strNum)
	Dim lp, tmpNo, strRst
	For lp=1 to len(strNum)
		tmpNo = mid(strNum, lp, 1)
		if (asc(tmpNo)>47 and asc(tmpNo)<58) or (asc(tmpNo)=45) or (asc(tmpNo)=46) then	'0~9,-,. Çã¿ë
			strRst = strRst & tmpNo
		end if
	Next
	getNumeric = strRst
End Function

'// Á¤±Ô½Ä ÆÐÅÏÁöÁ¤ ¹®ÀÚ¿­ Ã³¸®/¹ÝÈ¯
Function RepWord(str,patrn,repval)
	Dim regEx

	SET regEx = New RegExp
	regEx.Pattern = patrn			' ÆÐÅÏÀ» ¼³Á¤.
	regEx.IgnoreCase = True			' ´ë/¼Ò¹®ÀÚ¸¦ ±¸ºÐÇÏÁö ¾Êµµ·Ï .
	regEx.Global = True				' ÀüÃ¼ ¹®ÀÚ¿­À» °Ë»öÇÏµµ·Ï ¼³Á¤.
	RepWord = regEx.Replace(str,repval)
End Function

'/»ç¿ë±ÝÁö		'/lib/function.asp ¿¡ getUserLevelColor °ø¿ëÇÔ¼ö »ç¿ëÇÒ°Í. font color ·Î ¸ÔÀÏ°Í.		'/2016.07.20 ÇÑ¿ë¹Î
Function getUserLevelCSS(iuserLevel)
    if IsNULL(iuserLevel) then
        getUserLevelCSS = "member_no"
        exit function
    end if

    Select Case CStr(iuserLevel)
		Case "5"
			getUserLevelCSS = "member_orange"
		Case "0"
			getUserLevelCSS = "member_yellow"
		Case "1"
			getUserLevelCSS = "member_green"
		Case "2"
			getUserLevelCSS = "member_blue"
		Case "3"
			getUserLevelCSS = "member_vipsilver"
            ''getUserLevelCSS = "member_vip"
		Case "4"
			getUserLevelCSS = "member_vipgold"
		Case "7"
			getUserLevelCSS = "member_staff"
		Case "6"
			getUserLevelCSS = "member_red"
		Case "8"
			getUserLevelCSS = "member_red"
		Case "9"
			getUserLevelCSS = "member_red"
		Case Else
			getUserLevelCSS = "member_orange"
	end Select
end function

'//¹®ÀÚ¿­³» Æ¯¼ö¹®ÀÚ Á¦°Å
function ReplaceRequestSpecialChar(v)
	ReplaceRequestSpecialChar = replace(v,"'","")
	ReplaceRequestSpecialChar = replace(ReplaceRequestSpecialChar,"--","")
end function

'//¿Ã¸² ÇÔ¼ö
function ceil(Pnanum,nanum)
Dim result1, result2, variant_return

 result1 = Pnanum/nanum
 result2 = round(Pnanum/nanum)

 if result1 <> result2 then
  variant_return = fix(result1) + 1
 else
  variant_return = result1
 end if
ceil = variant_return
end function

'//¿Ã¸² ÇÔ¼ö
function ceilValue(iValue)
 if iValue <>  round(iValue) then
  ceilValue = fix(iValue) + 1
 else
  ceilValue = iValue
 end if
end function

'// ÁöÁ¤¼ö¸¸Å­ ÁöÁ¤ÇÑ ¹®ÀÚ·Î ¹Ù²Þ)
Function printUserId(strID,lng,chr)
	dim le, te

	if strID="" or isnull(strID) then
		exit Function
	end if

	le = len(strID)
	if(le<lng) Then
		printUserId = String(lng, le)
		Exit Function
	end if

	te = left(strID,le-lng) & String(lng, chr)
	printUserId = te

End Function

'// ³»¿ë¿¡ ±ÝÁöµÈ HTMLÅÂ±×°¡ ÀÖ´ÂÁö °Ë»ç //
function checkNotValidHTML(ostr)
	checkNotValidHTML = false

	dim LcaseStr
	LcaseStr = Lcase(ostr)
	LcaseStr = Replace(LcaseStr," ","")

	if InStr(LcaseStr,"<script")>0 or InStr(LcaseStr,"<object")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"</iframe>")>0 or InStr(LcaseStr,"<iframe>")>0 or InStr(LcaseStr,"iframe")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"<body")>0 or InStr(LcaseStr,"<input")>0 or InStr(LcaseStr,"<select")>0 or InStr(LcaseStr,"<textarea")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"onload=")>0 or InStr(LcaseStr,"onunload=")>0 or InStr(LcaseStr,"onclick=")>0 or InStr(LcaseStr,"onscroll=")>0 or InStr(LcaseStr,"onblur=")>0 or InStr(LcaseStr,"onerror=")>0 or InStr(LcaseStr,"onfocus=")>0 or InStr(LcaseStr,".href=")>0 or InStr(LcaseStr,".replace")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"onkeyup=")>0 or InStr(LcaseStr,"onkeydown=")>0 or InStr(LcaseStr,"onkeypress=")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"onmouseover=")>0 or InStr(LcaseStr,"onmouseout=")>0 or InStr(LcaseStr,"onmousedown=")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,".wmf")>0 or (InStr(LcaseStr,".js")>0 and Not(InStr(LcaseStr,".jsp")>0)) then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"window.")>0 then
		checkNotValidHTML = true
	end if

end function

'' 2015/10/06 checkNotValidHTML ahref, imgsrc ´Ù ¸·Èû;; »õ·Î ¸¸µë.
function checkNotValidHTMLcritical(ostr)
	checkNotValidHTMLcritical = false

	dim LcaseStr
	LcaseStr = Lcase(ostr)
	LcaseStr = Replace(LcaseStr," ","")

	if InStr(LcaseStr,"<script")>0 then
		checkNotValidHTMLcritical = true
	end if

	if InStr(LcaseStr,"<object")>0 then
		checkNotValidHTMLcritical = true
	end if

	if InStr(LcaseStr,"</iframe>")>0 then
		checkNotValidHTMLcritical = true
	end if

	if InStr(LcaseStr,"<iframe>")>0 then
		checkNotValidHTMLcritical = true
	end if

	if InStr(LcaseStr,"iframe")>0 then
		checkNotValidHTMLcritical = true
	end if

	'if InStr(LcaseStr,"imgsrc")>0 then
	'	checkNotValidHTMLcritical = true
	'end if

	'if InStr(LcaseStr,"ahref")>0 then
	'	checkNotValidHTMLcritical = true
	'end if

	if InStr(LcaseStr,".wmf")>0 then
		checkNotValidHTMLcritical = true
	end if

	if InStr(LcaseStr,".js")>0 then
		checkNotValidHTMLcritical = true
	end if
end function



'// °æ°í¹® Ãâ·ÂÈÄ Ã¢´Ý°í ¿ÀÇÂÃ¢ ¸®·Îµå -2011.02.23 Á¤À±Á¤Ãß°¡ //
Sub Alert_closenreload(strMSG)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"window.opener.location.reload();"& vbCrLf &_
			"self.close();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub

'// °æ°í¹® Ãâ·ÂÈÄ Ã¢´Ý°í ¿ÀÇÂÃ¢ Å¸°ÙÁÖ¼Ò·Î ÀÌµ¿ -2011.02.23 Á¤À±Á¤Ãß°¡ //
Sub Alert_closenmove(strMSG,targetURL)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"window.opener.location.href ='" & targetURL & "';" & vbCrLf &_
			"self.close();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub

'//ºÐ´ÜÀ§¸¦ ½Ã°£´ÜÀ§À¸·Î Â©¶ó¼­ ¹ÝÈ¯	'/2011.03.31 ÇÑ¿ë¹Î »ý¼º
function minutechagehour(v)
	dim tmpval , tmph , tmpm

	if v = "" or isnull(v) or v = 0 then
		minutechagehour = ""
	else
		tmph = int(v / 60)	'½Ã°£´ÜÀ§
		tmpm = v - (tmph * 60)	'ºÐ´ÜÀ§

		if tmph <> 0 then tmpval = tmpval & tmph & "½Ã°£ "
		if tmpm <> 0 then tmpval = tmpval & tmpm & "ºÐ"

		minutechagehour = tmpval
	end if
end function

'//¹ÙÀÌ³Ê¸® µ¥ÀÌÅÍ TEXTÇüÅÂ·Î º¯È¯
Function  BinaryToText(BinaryData, CharSet)
	 Const adTypeText = 2
	 Const adTypeBinary = 1

	 Dim BinaryStream
	 Set BinaryStream = CreateObject("ADODB.Stream")

	'¿øº» µ¥ÀÌÅÍ Å¸ÀÔ
	 BinaryStream.Type = adTypeBinary

	 BinaryStream.Open
	 BinaryStream.Write BinaryData
	 ' binary -> text
	 BinaryStream.Position = 0
	 BinaryStream.Type = adTypeText

	' º¯È¯ÇÒ µ¥ÀÌÅÍ Ä³¸¯ÅÍ¼Â
	 BinaryStream.CharSet = CharSet

	'º¯È¯ÇÑ µ¥ÀÌÅÍ ¹ÝÈ¯
	 BinaryToText = BinaryStream.ReadText

	 Set BinaryStream = Nothing
End Function

'// UTF8À» ASCII ¹®ÀÚ¿­·Î º¯È¯ //
Function URLDecodeUTF8(byVal pURL)
	Dim i, s1, s2, s3, u1, u2, result
	pURL = Replace(pURL,"+"," ")

	For i = 1 to Len(pURL)
		if Mid(pURL, i, 1) = "%" then
			s1 = CLng("&H" & Mid(pURL, i + 1, 2))

			'1¹ÙÀÌÆ®ÀÏ °æ¿ì(Pass)
			if (s1 < &H80) then
				result = result & Mid(pURL, i, 3)
				i = i + 2
			'2¹ÙÀÌÆ®ÀÏ °æ¿ì
			elseif ((s1 AND &HC0) = &HC0) AND ((s1 AND &HE0) <> &HE0) then
				s2 = CLng("&H" & Mid(pURL, i + 4, 2))

				u1 = (s1 AND &H1C) / &H04
				u2 = ((s1 AND &H03) * &H04 + ((s2 AND &H30) / &H10)) * &H10
				u2 = u2 + (s2 AND &H0F)
				result = result & ChrW((u1 * &H100) + u2)
				i = i + 5

			'3¹ÙÀÌÆ®ÀÏ °æ¿ì
			elseif (s1 AND &HE0 = &HE0) then
				s2 = CLng("&H" & Mid(pURL, i + 4, 2))
				s3 = CLng("&H" & Mid(pURL, i + 7, 2))

				u1 = ((s1 AND &H0F) * &H10)
				u1 = u1 + ((s2 AND &H3C) / &H04)
				u2 = ((s2 AND &H03) * &H04 +  (s3 AND &H30) / &H10) * &H10
				u2 = u2 + (s3 AND &H0F)
				result = result & ChrW((u1 * &H100) + u2)
				i = i + 8
			end if
		else
			result = result & Mid(pURL, i, 1)
		end if

	Next
	URLDecodeUTF8 = result
End Function

'// ASCIIÀ» UTF8 ¹®ÀÚ¿­·Î º¯È¯ //
Public Function URLEncodeUTF8(byVal szSource)
	Dim szChar, WideChar, nLength, i, result
	nLength = Len(szSource)

	For i = 1 To nLength
		szChar = Mid(szSource, i, 1)

		If Asc(szChar) < 0 Then
			WideChar = CLng(AscB(MidB(szChar, 2, 1))) * 256 + AscB(MidB(szChar, 1, 1))

			If (WideChar And &HFF80) = 0 Then
				result = result & "%" & Hex(WideChar)
			ElseIf (WideChar And &HF000) = 0 Then
				result = result & _
					"%" & Hex(CInt((WideChar And &HFFC0) / 64) Or &HC0) & _
					"%" & Hex(WideChar And &H3F Or &H80)
			Else
				result = result & _
					"%" & Hex(CInt((WideChar And &HF000) / 4096) Or &HE0) & _
					"%" & Hex(CInt((WideChar And &HFFC0) / 64) And &H3F Or &H80) & _
					"%" & Hex(WideChar And &H3F Or &H80)
			End If
		Else
			if (Asc(szChar)>=48 and Asc(szChar)<=57) or (Asc(szChar)>=65 and Asc(szChar)<=90) or (Asc(szChar)>=97 and Asc(szChar)<=122) then
				result = result + szChar
			else
				if Asc(szChar)=32 then
					result = result & "+"
				else
					result = result & "%" & Hex(AscB(MidB(szChar, 1, 1)))
				end if
			end if
		End If
	Next
	URLEncodeUTF8 = result
End Function

'// ÄÞ¸¶·Î ±¸ºÐµÈ ¹è¿­°ª¿¡ ÁöÁ¤µÈ °ªÀÌ ÀÖ´ÂÁö ¹ÝÈ¯
function chkArrValue(aVal,cVal)
	dim arrV, i
	chkArrValue = false
	arrV = split(aVal,",")
	for i=0 to ubound(arrV)
		if cStr(arrV(i))=cStr(cVal) then
			chkArrValue = true
			exit function
		end if
	next
end function

'// »ç³» Á¢¼Ó¿©ºÎ
Function isTenbyTenConnect()
	Dim conIp, arrIp, tmpIp
	conIp = Request.ServerVariables("REMOTE_ADDR")
	if left(conIp,2)<>"::" then
		arrIp = split(conIp,".")
		tmpIp = Num2Str(arrIp(0),3,"0","R") & Num2Str(arrIp(1),3,"0","R") & Num2Str(arrIp(2),3,"0","R") & Num2Str(arrIp(3),3,"0","R")
	end if

	'121.78.103.60 : 15Ãþ À¯¼±
	'10.10.10.36 : m2¼­¹ö
	'192.168.1.x : 15Ãþ ¿î¿µ,°³¹ß,ÀÎ»ç,Àç¹«
	'192.168.6.x : 15Ãþ ÀÏ¹Ý¸Á
	'110.11.187.233 : 15Ãþ wireless6
	'110.93.128.x : IDC

	if tmpIp="121078103060" or tmpIp="110011187233" or (tmpIp=>"110093128001" and tmpIp<="110093128256") or (tmpIp=>"192168001001" and tmpIp<="192168001256") or (tmpIp=>"192168006001" and tmpIp<="192168006256") then
		isTenbyTenConnect = True
	else
		isTenbyTenConnect = False
	end if
End Function

'/¼­¹ö ÁÖ±âÀû ¾÷µ¥ÀÌÆ® À§ÇÑ °ø»çÁß Ã³¸® '2011.11.11 ÇÑ¿ë¹Î »ý¼º
'/¸®´º¾ó½Ã ÀÌÀüÇØ ÁÖ½Ã°í Áö¿ìÁö ¸»¾Æ ÁÖ¼¼¿ä
Sub serverupdate_underconstruction()
	dim isServerDown : isServerDown = false
		'isServerDown = true	' ¼­¹ö´Ù¿î
		isServerDown = false	' ¼­¹öÈ°¼ºÈ­
		if isTenbyTenConnect then isServerDown = false	'»ç³»Á¢¼Ó Çã¿ë

	if Not(isServerDown) then exit Sub

	Response.write "<html>"
	Response.write "<head><title>¼­ºñ½º Á¡°ËÁßÀÔ´Ï´Ù</title></head>"
	Response.write "<meta http-equiv='Content-Type' content='text/html;charset=euc-kr' />"
	Response.write "<body>"
	Response.write "<table width='100%' height='100%' cellpadding='0' cellspacing='0' border='0'>"
	Response.write "<tr>"
	Response.write "	<td align='center' valign='middle'><img src='http://fiximage.10x10.co.kr/web2015/common/2015_10x10_open_ready_PC.jpg' width='1104' border='0' ></td>"
	Response.write "</tr>"
	Response.write "</table>"
	Response.write "<input type='hidden' name='refip' value='" & Request.ServerVariables("REMOTE_ADDR") & "' />"
	Response.write "</body>"
	Response.write "</html>"
	response.End
End Sub

function getSCMSSLURL()
    IF application("Svr_Info")="Dev" THEN
        getSCMSSLURL = "https://testwebadmin.10x10.co.kr"
		if (G_IsLocalDev) then getSCMSSLURL = ""
    ELSE
        getSCMSSLURL = "https://webadmin.10x10.co.kr"
    END IF
end function

function getSCMURL()
    IF application("Svr_Info")="Dev" THEN
        getSCMURL = "http://testwebadmin.10x10.co.kr"
		if (G_IsLocalDev) then getSCMURL = ""
    ELSE
        getSCMURL = "http://webadmin.10x10.co.kr"
    END IF
end function

Function r_g()
	Dim i, key
	response.write "<table width=750 border=1 bordercolor='#cccccc' style='border-collapse:collapse;font:10pt'>" + vbcrlf
	response.write "<tr bgcolor='gold'>" + vbcrlf
	response.write "	<td align='center'>name</td>" + vbcrlf
	response.write "	<td align='center'>value</td>" + vbcrlf
	response.write "</tr>" + vbcrlf
	For Each key in Request.QueryString
		response.write  "<tr align='center' bgcolor='#FFFFFF' onmouseover=this.style.background='#f1f1f1'; onmouseout=this.style.background='#FFFFFF';>" + vbcrlf
		response.write  "<td>" & key & "</td>" + vbcrlf
		If IsArray(Request.Form(key)) Then
			response.write  "<td>" & r_g(Request.QueryString(key)) & "</td>" + vbcrlf
		Else
			response.write  "<td>" & Request.QueryString(key) & "</td>" + vbcrlf
		End If
		response.write  "</tr>" + vbcrlf
	Next
	response.write "</table>" + vbcrlf
END function

Function r_s()
	Dim i, key
	response.write "<table width=750 border=1 bordercolor='#cccccc' style='border-collapse:collapse;font:10pt'>" + vbcrlf
	response.write "<tr bgcolor='gold'>" + vbcrlf
	response.write "	<td align='center'>name</td>" + vbcrlf
	response.write "	<td align='center'>value</td>" + vbcrlf
	response.write "</tr>" + vbcrlf
	For Each key in Request.ServerVariables
		response.write  "<tr align='center' bgcolor='#FFFFFF' onmouseover=this.style.background='#f1f1f1'; onmouseout=this.style.background='#FFFFFF';>" + vbcrlf
		response.write  "<td>" & key & "</td>" + vbcrlf
		If IsArray(Request.Form(key)) Then
			response.write  "<td>" & r_s(Request.ServerVariables(key)) & "</td>" + vbcrlf
		Else
			response.write  "<td>" & Request.ServerVariables(key) & "</td>" + vbcrlf
		End If
		response.write  "</tr>" + vbcrlf
	Next
	response.write "</table>" + vbcrlf
END function

'// Æ÷Åä¼­¹ö ½æ³×ÀÏ Á¦ÀÛ(±âÁ¸ ÆÄÀÏ¸í)		'/2016.04.19 ÇÑ¿ë¹Î ÇÁ·ÐÆ®¿¡¼­ º¹»ç/ÀÌµ¿
function getThumbImgFromURL(furl,wd,ht,fit,ws)
	dim sCmd

	'µµ¸ÞÀÎ Ä¡È¯
    IF application("Svr_Info")="Dev" THEN
		if instr(furl,"imgstatic")>0 then
			furl = replace(furl,"imgstatic.10x10.co.kr/","thumbnail.10x10.co.kr/testimgstatic/")
		elseif instr(furl,"webimage")>0 then
			furl = replace(furl,"webimage.10x10.co.kr/","thumbnail.10x10.co.kr/testwebimage/")
		end if
    ELSE
		if instr(furl,"imgstatic")>0 then
			furl = replace(furl,"imgstatic.10x10.co.kr/","thumbnail.10x10.co.kr/imgstatic/")
		elseif instr(furl,"webimage")>0 then
			furl = replace(furl,"webimage.10x10.co.kr/","thumbnail.10x10.co.kr/webimage/")
		end if
    END IF

	'½æ³×ÀÏ Ä¿¸Çµå
	sCmd = "?cmd=thumb"
	if wd<>"" then sCmd = sCmd & "&w=" & wd
	if ht<>"" then sCmd = sCmd & "&h=" & ht
	if fit<>"" then sCmd = sCmd & "&fit=" & fit
	if ws<>"" then sCmd = sCmd & "&ws=" & ws

	'º¯È¯ÁÖ¼Ò ¹ÝÈ¯
	getThumbImgFromURL = furl & sCmd
end function

'/°³ÀÎÁ¤º¸ ÀüÈ­¹øÈ£ Ã³¸®
function printtel(telNo)
	dim resultStr, tmpArr, i
    
	resultStr = telNo

	if IsNull(telno) then
		printtel = resultStr
		Exit Function
	end if
	if telno="" then
		printtel = resultStr
		Exit Function
	end if

	tmpArr = Split(telNo, "-")

	Select Case UBound(tmpArr)
		Case 1
			resultStr = treg_replace(tmpArr(0), ".", "*", True) & "-" & tmpArr(0)
		Case 2
			resultStr = tmpArr(0) & "-" & treg_replace(tmpArr(1), ".", "*", True) & "-" & tmpArr(2)
		Case Else
			resultStr = "ERR"
	End Select

	printtel = resultStr
end Function
function treg_replace(strOriginalString, strPattern, strReplacement, varIgnoreCase)
    ' Function replaces pattern with replacement
    ' varIgnoreCase must be TRUE (match is case insensitive) or FALSE (match is case sensitive)
    dim objRegExp : set objRegExp = new RegExp
    with objRegExp
        .Pattern = strPattern
        .IgnoreCase = varIgnoreCase
        .Global = True
    end with
    treg_replace = objRegExp.replace(strOriginalString, strReplacement)
    set objRegExp = nothing
end function
%>
