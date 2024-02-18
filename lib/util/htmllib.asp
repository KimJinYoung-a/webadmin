<%
'+----------------------------------------------------------------------------------------------------------------------+
'|                                               HTML �� ��   �� �� �� ��                                               |
'+-------------------------------------------+--------------------------------------------------------------------------+
'|             �� �� ��                      |                          ��    ��                                        |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| FormatDate(ddate, formatstring)          | ��¥������ ������ ���������� ��ȯ                            |
'|                                          | ��뿹 : printdate = FormatDate(now(),"0000.00.00")          |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| GetImageSubFolderByItemid(byval iitemid)  | �̹��������� ���� �������� ��ȯ�Ѵ�.                                     |
'|                                           | ��뿹 : SubFolder = GetImageSubFolderByItemid(1126)                     |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| db2html(checkvalue)                       | DB�� ������ HTML�� ����� �� �ֵ��� ��ȯ                                 |
'|                                           | ��뿹 : Contents = db2html("DB�� ����")                                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| html2db(checkvalue)                       | ����ڰ� �Է��� ������ DB�� ���� �� �ֵ��� ��ȯ                          |
'|                                           | ��뿹 : Contents = html2db("������ ����")                               |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| nl2br(checkvalue)                         | ������ ����(vbCrLf)�� "<br>"�±׷� ġȯ�Ͽ� ��ȯ                         |
'|                                           | ��뿹 : Contents = nl2br("����")                                        |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| CurrFormat(byVal v)                       | ���ڸ� 3�ڸ� ������ ���ڿ��� ��ȯ                                        |
'|                                           | ��뿹 : strNum = CurrFormat(1230) �� "1,230"                            |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Format00(n,orgData)                       | ���ڸ� 0���� ä���� ������ ������ ���ڿ��� ��ȯ                          |
'|                                           | ��뿹 : strNum = Format00(5,123) �� "00123"                             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| FormatCode(itemcode)                      | ��ǰ �Ϸù�ȣ�� 6�ڸ��� ���ڿ��� ��ȯ                                    |
'|                                           | ��뿹 : itemCode = FormatCode(2654) �� "002654"                         |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| GetCurrentTimeFormat()                    | ����ð��� ���ڿ��� ��ȯ (yyyymmddhhmmss)                                |
'|                                           | ��뿹 : strNow = GetCurrentTimeFormat() �� "20060508101833"             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| GetListImageUrl(byval itemid)             | ��ǰ��ȣ�� �´� ����Ʈ �̹��� �� ���� ��ȯ                               |
'|                                           | ��뿹 : img = GetListImageUrl("53100") �� "/image/list/L000053100.jpg"  |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| DDotFormat(byval str,byval n)             | ������ ������ ���̷� �ڸ���.                                             |
'|                                           | ��뿹 : strShort = DDotFormat("�����Դϴ�.",3) �� "������..."           |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| stripHTML(strng)                          | ���� �� HTML�±׸� ���ش�.                                               |
'|                                           | ��뿹 : Contents = stripHTML("<b>����</b>") �� " ���� "                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| getFileExtention(strFile)                 | ���ϸ��� Ȯ���ڸ� ��ȯ�Ѵ�.                                              |
'|                                           | ��뿹 : ext = getFileExtention("123.jpg") �� "jpg"                      |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Num2Str(inum,olen,cChr,oalign)   		 | ���ڸ� ������ ������ ���ڿ��� ��ȯ�Ѵ�.                      			|
'|                                   		 | ��뿹 : Num2Str(425,4,"0","R") �� 0425                      			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ChkIIF(trueOrFalse, trueVal, falseVal)    | like iif function                                                        |
'|                                           | ��뿹 : ChkIIF(1>2,"a","b") �� "b"                                       |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Alert_return(strMSG)                      | ���â ����� �������� ���ư���.                            				|
'|                                           | ��뿹 : Call Alert_return("�ڷ� ���ư��ϴ�.")               			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Alert_close(strMSG)                       | ���â ����� ����â�� �ݴ´�.                               			|
'|                                           | ��뿹 : Call Alert_close("â�� �ݽ��ϴ�.")                  			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Alert_move(strMSG,targetURL)              | ���â ����� ������������ �̵��Ѵ�.                         			|
'|                                           | ��뿹 : Call Alert_move("�̵��մϴ�.","/index.asp")         			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chrbyte(str,chrlen,dot)                   | �������̷� ���ڿ� �ڸ���                                                 |
'|                                           | ��뿹 : chrbyte("�ȳ��ϼ���",3,"Y") �� �ȳ�...                           |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chkPasswordComplex(uid,pwd)               | ��й�ȣ ��å�� ���⼺�� �����ϴ��� �˻��ϰ� �� ������ ��ȯ              |
'|                                           | ��뿹 : chkPasswordComplex("kobula","abcd")                             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chkPasswordComplexNonid(pwd)         	 | ���̵� �����й�ȣ ��å�� ���⼺�������ϴ��� �˻��ϰ� �� ������ ��ȯ  |
'|                                           | ��뿹 : chkPasswordComplexNonid("abcd")                             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chkWord(str,patrn)                        | ���ڿ��� ������ ���Խ����� �˻�                                          |
'|                                           | ��뿹 : chkWord("abcd","[^-a-zA-Z0-9/ ]") : ������ڸ� ���             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ParsingPhoneNumber(str,patrn)             | ��ȭ��ȣ�� ��� �߰�                                                     |
'|                                           | ��뿹 : ParsingPhoneNumber("0112223333") :                              |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceBracket(strng)                     | ������ȣ �±׷� ġȯ('<', '>')                                           |
'|                                           | ��뿹 : ReplaceBracket("<>") �� &lt;&gt;                                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceBracketOther(strng)                | ������ȣ �ٸ� ��ȣ�� ġȯ('<', '>')                                        |
'|                                           | ��뿹 : ReplaceBracketOther("<>") �� []                                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceScript(strng)                      | Script Tag ġȯ                                                          |
'|                                           | ��뿹 : ReplaceScript("<script") �� &lt;script                           |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| getNumeric(strNum)                        | ���ڿ����� ���ڸ� ���� ��ȯ                                              |
'|                                           | ��뿹 : getNumeric("a45d61*124") -> 461124                              |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| RepWord(str,patrn,repval)                | ���Խ� ������ ����� ���ڿ� ó��                             				|
'|                                          | ��뿹 : RepWord(SearchText,"[^��-�Ra-zA-Z0-9\s]","")      			  	|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceRequestSpecialChar(strng)        	| Ư�� ���� ����(' ,--)                                        				|
'|                                          | ��뿹 : cont = ReplaceRequestSpecialChar(Rs("strng"))       				|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| checkNotValidHTML(ostr)                  | ���뿡 ������ HTML�±װ� �ִ��� �˻�                         				|
'|                                          | ��뿹 : checkNotValidHTML("<script...") �� true             				|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| minutechagehour(v)                 		| �д����� �ð��������� ©�� ��ȯ                      					|
'|                                          | ��뿹 : minutechagehour(v)             									|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| BinaryToText(BinaryData, CharSet)         | ���̳ʸ� ������ TEXT���·� ��ȯ                                          |
'|                                           | ��뿹 : BinaryToText(objXML.ResponseBody, "euc-kr")                     |
'+------------------------------------------+---------------------------------------------------------------------------+
'| URLEncodeUTF8(byVal szSource)            | ASCII�� UTF8 ���ڿ��� ��ȯ                                                |
'|                                          | ��뿹 : strUF8 = URLEncodeUTF8(STR)                                      |
'+------------------------------------------+---------------------------------------------------------------------------+
'| URLDecodeUTF8(byVal pURL)                | UTF8�� ASCII ���ڿ��� ��ȯ                                                |
'|                                          | ��뿹 : strASC = URLDecodeUTF8(URL)                                      |
'+------------------------------------------+---------------------------------------------------------------------------+
'| chkArrValue(aVal,cVal)                    | �޸��� ���е� �迭���� ������ ���� �ִ��� ��ȯ                           |
'|                                           | ��뿹 : chkArrValue("A,B,C", "B") �� true                                |
'+-------------------------------------------+--------------------------------------------------------------------------+

function G_IsLocalDev()
	G_IsLocalDev = (application("Svr_Info")="Dev") AND (request.ServerVariables("LOCAL_ADDR")="::1" or request.ServerVariables("LOCAL_ADDR")="127.0.0.1")
end function

''��Ű�� ������ ��� / ���̵� �ܹ��� �ؽ��� : ���� md5 �ʿ�. (md5 ���� ���Ұ�� component, db �̿� ����)
function HashTenID(byval oid)
    dim orgid : orgid = LCASE(oid)
    dim hashid

    HashTenID = orgid
    if Len(orgid)<1 then Exit function      ''���ΰ�� ������
    if Len(orgid)<2 then orgid=orgid+"1"    ''���̰�1�ϰ�� ��������.


    hashid = Right(orgid,4) + Left(orgid,Len(orgid)-1)
    hashid = Right(hashid,5) + Left(hashid,Len(hashid)-2)
    hashid = Right(hashid,6) + Left(hashid,Len(hashid)-3)
    hashid = Right(hashid,7) + Left(hashid,Len(hashid)-4)
    hashid = Right(hashid,8) + Left(hashid,Len(hashid)-5)
    HashTenID = MD5(hashid)

end function

'// ��¥�� ������ ���������� ��ȯ //
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

'' ���� ��� ���� ���� ����.. ���� ����
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

'' 2008 03 ���� - Eastone
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

'// ���ڿ��� CR/LF�� �������� ġȯ //
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

	'// ���ڸ� ������ ������ ���ڿ��� ��ȯ //
	Function Num2Str(inum,olen,cChr,oalign)
		Dim i, ilen, strChr

		ilen = len(Cstr(inum))
		strChr = ""

		if ilen < olen then
			for i=1 to olen-ilen
				strChr = strChr & cChr
			next
		end if

		'���չ�������� ��� �б�
		if oalign="L" then
			'���ʱ���
			Num2Str = inum & strChr
		else
			'������ ���� (�⺻��)
			Num2Str = strChr & inum
		end if

    End Function


'// ���ڿ��� �߶� ���ϴ� ��ġ�� ���� ��ȯ //
function SplitValue(orgStr,delim,pos)
    dim buf
    SplitValue = ""
    if IsNULL(orgStr) then Exit function
    if (Len(delim)<1) then Exit function
    buf = split(orgStr,delim)

    if UBound(buf)<pos then Exit function

    SplitValue = buf(pos)
end function


'// �Ķ���� ���� üũ �� Maxlen ���Ϸ� ������ Code, id ���� Param �� ��� //
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


'// ���� �� Return �� like iif function
Function ChkIIF(trueOrFalse, trueVal, falseVal)
	if (trueOrFalse) then
	    ChkIIF = trueVal
	else
	    ChkIIF = falseVal
	end if
End Function

'// ��� ����� �ڷΰ��� //
Sub Alert_return(strMSG)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"history.back();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub


'// ��� ����� â�ݱ� //
Sub Alert_close(strMSG)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"self.close();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub


'// ��� ����� ���� �������� �̵� //
Sub Alert_move(strMSG,targetURL)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"self.location.replace('" & targetURL & "');" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub

'// �������̷� ���ڿ� �ڸ��� //
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

'// �н����� ���⼺ �˻� �Լ�(�������� ����6, �α��ν� üũ��.)		//2017.09.25 �ѿ�� ����
Function chkPasswordComplex_Len6Ver(uid,pwd)
	dim msg, i, sT, sN, numAlpha, numNums, numSpecials, buf, index
    numAlpha = 0
    numNums = 0
    numSpecials = 0
	msg = ""

	'��й�ȣ ���� �˻�
	if len(pwd)<8 then
		msg = msg & "- ��й�ȣ�� �ּ� 8�ڸ��̻����� �Է����ּ���.\n"
	end if

	'���̵�� ���� �Ǵ� �����ϰ� �ִ°�?
	if instr(lcase(pwd),lcase(uid))>0 then
		msg = msg & "- ���̵�� �����ϰų� ���̵� �����ϰ� �ִ� ��й�ȣ�Դϴ�.\n"
	end if

	'## ���⼺�� �����ϴ°�?
	'�������� 3�� ���� ����
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
			msg = msg & "- �������ڰ� 3�� �������� �������ϴ�.\n"
			exit for
		end if
	next

'���Խ� �ȹٷ� �ȹ���. �Ӳ�
'	if chkWord(pwd,"[^-a-zA-Z]") then
'		numAlpha = numAlpha + 1
'	end if
'	if chkWord(pwd,"[^-0-9 ]") then
'		numNums = numNums + 1
'	end if
'	if chkWord(pwd,"[~!@\#$%<>^&*\()\-=+_\��]") then
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

	'// 3���� ����
    if (numAlpha>0 and numNums>0 and numSpecials>0) then
    	if (len(pwd) >= 8) then
    	else
    		msg = msg & "- ���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)\n"
    	end if

	'// 2���� ����
    elseif ((numAlpha>0 and numNums>0) or (numAlpha>0 and numSpecials>0) or (numNums>0 and numSpecials>0)) then
    	if (len(pwd) >= 10) then
    	else
    		msg = msg & "- ���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)\n"
    	end if

    else
    	msg = msg & "- ���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)\n"
    end if

	'��� ��ȯ
	chkPasswordComplex_Len6Ver = msg
end Function

'// �н����� ���⼺ �˻� �Լ�		//2017.09.25 �ѿ�� ����
Function chkPasswordComplex(uid,pwd)
	dim msg, i, sT, sN, numAlpha, numNums, numSpecials, buf, index
    numAlpha = 0
    numNums = 0
    numSpecials = 0
	msg = ""

	'��й�ȣ ���� �˻�
	if len(pwd)<8 then
		msg = msg & "- ��й�ȣ�� �ּ� 8�ڸ��̻����� �Է����ּ���.\n"
	end if

	'���̵�� ���� �Ǵ� �����ϰ� �ִ°�?
	if instr(lcase(pwd),lcase(uid))>0 then
		msg = msg & "- ���̵�� �����ϰų� ���̵� �����ϰ� �ִ� ��й�ȣ�Դϴ�.\n"
	end if

	'## ���⼺�� �����ϴ°�?
	'�������� 3�� ���� ����
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
			msg = msg & "- �������ڰ� 3�� �������� �������ϴ�.\n"
			exit for
		end if
	next

'���Խ� �ȹٷ� �ȹ���. �Ӳ�
'	if chkWord(pwd,"[^-a-zA-Z]") then
'		numAlpha = numAlpha + 1
'	end if
'	if chkWord(pwd,"[^-0-9 ]") then
'		numNums = numNums + 1
'	end if
'	if chkWord(pwd,"[~!@\#$%<>^&*\()\-=+_\��]") then
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

	'// 3���� ����
    if (numAlpha>0 and numNums>0 and numSpecials>0) then
    	if (len(pwd) >= 8) then
    	else
    		msg = msg & "- ���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)\n"
    	end if

	'// 2���� ����
    elseif ((numAlpha>0 and numNums>0) or (numAlpha>0 and numSpecials>0) or (numNums>0 and numSpecials>0)) then
    	if (len(pwd) >= 10) then
    	else
    		msg = msg & "- ���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)\n"
    	end if

    else
    	msg = msg & "- ���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)\n"
    end if

	'��� ��ȯ
	chkPasswordComplex = msg
end Function

'// �н����� ���⼺ �˻� �Լ�
Function chkPasswordComplexNonID(pwd)
	dim msg, i, sT, sN
	msg = ""

	'��й�ȣ ���� �˻�
	if len(pwd)<8 then
		msg = msg & "- ��й�ȣ�� �ּ� 8�ڸ��̻����� �Է����ּ���.\n"
	end if


	'## ���⼺�� �����ϴ°�?
	'�������� 3�� ���� ����
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
			msg = msg & "- �������ڰ� 3�� �������� �������ϴ�.\n"
			exit for
		end if
	next
	'����/������ ����
	if chkWord(pwd,"[^-a-zA-Z]") or chkWord(pwd,"[^-0-9 ]") then
		msg = msg & "- ��й�ȣ�� �ݵ�� ���ĺ��� ���ڸ� �����ؼ� �������մϴ�.\n"
	end if

	'��� ��ȯ
	chkPasswordComplexNonID = msg
end Function

'//���Խ� ���ڿ� �˻�
Function chkWord(str,patrn)
    Dim regEx, match, matches

    SET regEx = New RegExp

    regEx.Pattern = patrn            ' ������ �����մϴ�.
    regEx.IgnoreCase = True      ' ��/�ҹ��ڸ� �������� �ʵ��� �մϴ�.
    regEx.Global = True             ' ��ü ���ڿ��� �˻��ϵ��� �����մϴ�.

    SET Matches = regEx.Execute(str)

    if 0 < Matches.count then
        chkWord = false
    Else
        chkWord = true
    end if
End Function

'// ��ȭ��ȣ�� ��� �߰�
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


'''''==================  2009 �߰�

' response.write �Լ�
Function rw(ByVal str)
	response.write str & "<br>"
End Function

' Null�� �������� ġȯ
Function null2blank(ByVal v)
	If IsNull(v) Then
		null2blank = ""
	Else
		null2blank = v
	End If
End Function

'// ū����ǥ input �ڽ� value=""�� ����Ҷ� ġȯ
Function doubleQuote(ByVal v)
	If IsNull(v) Then
		doubleQuote = ""
	Else
		doubleQuote = Replace(v, """","&quot;")
	End If
End Function


' request ��ü �Լ�(�Ķ���͸�, ����Ʈ��)
Function req(ByVal param, ByVal value)
'	VarType Return ��
'	0 (����)
'	1 (��)
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
		If Not IsNumeric(tmpValue) Then	' ���ڰ� �ƴϸ�
			tmpValue = value
		End If
		tmpValue = CDbl(tmpValue)
	Else
		tmpValue = Trim(Request(param))
		If tmpValue = "" Then			' Request���� ������
			tmpValue = value
		End If
	End If
	req = tmpValue

End Function

Sub sbDisplayPaging(ByVal strCurrentPage, ByVal intTotalRecord, ByVal intRecordPerPage, ByVal intBlockPerPage)

	'���� ����
	Dim intCurrentPage, strCurrentPath
	Dim intStartBlock, intEndBlock, intTotalPage
	Dim strParamName, intLoop

	'���� ������ ����
	intCurrentPage = Mid(strCurrentPage, InStr(strCurrentPage, "=")+1)		'���� ������ ��
	strCurrentPage = Left(strCurrentPage, InStr(strCurrentPage, "=")-1)		'������ ���� ������

	'���� ������ ��
	strCurrentPath = Request.ServerVariables("Script_Name")

	'�ش��������� ǥ�õǴ� ������������ ������������ ����
	intStartBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + 1
	intEndBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + intBlockPerPage

	'�� ������ �� ����
	intTotalPage =  -(int(-(intTotalRecord/intRecordPerPage)))

	'�� ���� & hidden �Ķ���� ����
	Response.Write	"<form name='frmPaging' method='get' action ='" & strCurrentPath & "'>" &_
							"<input type='hidden' name='" & strCurrentPage & "'>"			'���� ������

	'�Ķ���� ����(��: �˻���)�� hidden �Ķ���ͷ� �����Ѵ�
	strParamName = ""
	For Each strParamName In Request.Form
		If strParamName <> strCurrentPage Then

			'hidden �Ķ���� ���� �Ķ���� �˿�
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.Form(strParamName),50) & "'>"
		End If
	Next
	strParamName = ""

	For Each strParamName In Request.Querystring
		If strParamName <> strCurrentPage Then
			'hidden �Ķ���� ���� �Ķ���� �˿�
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.QueryString(strParamName),50) & "'>"
		END IF
	Next

	Response.Write "<table border='0' cellpadding='0' cellspacing='0' class=a><tr align='center'><td>"

	'���� ������ �̹��� ����
	If intStartBlock > 1 Then
		Response.Write "<img src='http://fiximage.10x10.co.kr/web2008/designfingers/btn_pageprev01.gif' border='0' style='cursor:hand' alt='���� " & intBlockPerPage & " ������'" &_
							   "onClick='javascript:document.frmPaging." & strCurrentPage & ".value=" & intStartBlock - intBlockPerPage & ";document.frmPaging.submit();'>"
	Else
		Response.Write "<img src='http://fiximage.10x10.co.kr/web2009/common/btn_pageprev01.gif' border='0' >"
	End If

	Response.Write "</td><td>&nbsp;"

	'����¡ ���
	If intTotalPage > 1 Then
		For intLoop = intStartBlock To intEndBlock
			If intLoop > intTotalPage Then Exit For

			If Int(intLoop) <> Int(intStartBlock) Then Response.Write "|"

			If Int(intLoop) = Int(intCurrentPage) Then		'���� ������
				Response.Write "&nbsp;<span class='text01'><strong>" & intLoop & "</strong></span>&nbsp;"
			Else															'�� �� ������
				Response.Write "&nbsp;<a href='javascript:document.frmPaging." & strCurrentPage & ".value=" & intLoop & ";document.frmPaging.submit();'><font class='text01'>" & intLoop & "</font></a>&nbsp;"
			End If

		Next
	Else		'�� �������� ���� �Ҷ�
		Response.Write "&nbsp;<span class='text01'><strong>1</strong></span>&nbsp;"
	End If

	Response.Write "&nbsp;</td><td>"

	'���� ������ �̹��� ����
	If Int(intEndBlock) < Int(intTotalPage) Then
		Response.Write "<img src='http://fiximage.10x10.co.kr/web2008/designfingers/btn_pagenext01.gif' border='0' style='cursor:hand' alt='���� " & intBlockPerPage & " ������'" &_
							   "onClick='javascript:document.frmPaging." & strCurrentPage & ".value=" & intEndBlock+1 & ";document.frmPaging.submit();'>"
	Else
	    Response.Write "<img src='http://fiximage.10x10.co.kr/web2009/common/btn_pagenext01.gif' border='0' >"
	End If

	Response.Write "</td></tr></form></table>"

End Sub



' ���,����,���� ��� �ؽ�Ʈ ����
Function getModeName(ByVal mode)
    Select Case mode
        Case "INS"	: getModeName = "���"
        Case "UPD"	: getModeName = "����"
        Case "DEL"	: getModeName = "����"
        Case "FIN"	: getModeName = "�Ϸ�"
        Case Else	: getModeName = "����"
    End Select
End Function

'// ������ȣ HTML�ڵ�� ġȯ //
' db2html �̶� �浹���� ��밡���Ѱ��� �����ϼ���.
Function ReplaceBracket(strng)
	if isnull(strng) then exit Function

	strng = Replace(strng,"<","&lt;")
	strng = Replace(strng,">","&gt;")
	ReplaceBracket = strng
end Function

'// ������ȣ �ٸ� ��ȣ�� ġȯ //
Function ReplaceBracketOther(strng)
	if isnull(strng) then exit Function

	strng = Replace(strng,"<","[")
	strng = Replace(strng,">","]")
	ReplaceBracketOther = strng
end Function

'// Script Tagġȯ //
Function ReplaceScript(strng)
	if isnull(strng) then exit Function

	strng = Replace(strng,"<script","[script")
	strng = Replace(strng,"</script","[/script")
	strng = Replace(strng,"<iframe","[iframe")
	strng = Replace(strng,"</iframe","[/iframe")
	ReplaceScript = strng
end Function


' ���Խ� �Լ�
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

'// ���ڿ����� ���ڸ� ���� ��ȯ
Function getNumeric(strNum)
	Dim lp, tmpNo, strRst
	For lp=1 to len(strNum)
		tmpNo = mid(strNum, lp, 1)
		if (asc(tmpNo)>47 and asc(tmpNo)<58) or (asc(tmpNo)=45) or (asc(tmpNo)=46) then	'0~9,-,. ���
			strRst = strRst & tmpNo
		end if
	Next
	getNumeric = strRst
End Function

'// ���Խ� �������� ���ڿ� ó��/��ȯ
Function RepWord(str,patrn,repval)
	Dim regEx

	SET regEx = New RegExp
	regEx.Pattern = patrn			' ������ ����.
	regEx.IgnoreCase = True			' ��/�ҹ��ڸ� �������� �ʵ��� .
	regEx.Global = True				' ��ü ���ڿ��� �˻��ϵ��� ����.
	RepWord = regEx.Replace(str,repval)
End Function

'/������		'/lib/function.asp �� getUserLevelColor �����Լ� ����Ұ�. font color �� ���ϰ�.		'/2016.07.20 �ѿ��
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

'//���ڿ��� Ư������ ����
function ReplaceRequestSpecialChar(v)
	ReplaceRequestSpecialChar = replace(v,"'","")
	ReplaceRequestSpecialChar = replace(ReplaceRequestSpecialChar,"--","")
end function

'//�ø� �Լ�
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

'//�ø� �Լ�
function ceilValue(iValue)
 if iValue <>  round(iValue) then
  ceilValue = fix(iValue) + 1
 else
  ceilValue = iValue
 end if
end function

'// ��������ŭ ������ ���ڷ� �ٲ�)
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

'// ���뿡 ������ HTML�±װ� �ִ��� �˻� //
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

'' 2015/10/06 checkNotValidHTML ahref, imgsrc �� ����;; ���� ����.
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



'// ��� ����� â�ݰ� ����â ���ε� -2011.02.23 �������߰� //
Sub Alert_closenreload(strMSG)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"window.opener.location.reload();"& vbCrLf &_
			"self.close();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub

'// ��� ����� â�ݰ� ����â Ÿ���ּҷ� �̵� -2011.02.23 �������߰� //
Sub Alert_closenmove(strMSG,targetURL)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"window.opener.location.href ='" & targetURL & "';" & vbCrLf &_
			"self.close();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub

'//�д����� �ð��������� ©�� ��ȯ	'/2011.03.31 �ѿ�� ����
function minutechagehour(v)
	dim tmpval , tmph , tmpm

	if v = "" or isnull(v) or v = 0 then
		minutechagehour = ""
	else
		tmph = int(v / 60)	'�ð�����
		tmpm = v - (tmph * 60)	'�д���

		if tmph <> 0 then tmpval = tmpval & tmph & "�ð� "
		if tmpm <> 0 then tmpval = tmpval & tmpm & "��"

		minutechagehour = tmpval
	end if
end function

'//���̳ʸ� ������ TEXT���·� ��ȯ
Function  BinaryToText(BinaryData, CharSet)
	 Const adTypeText = 2
	 Const adTypeBinary = 1

	 Dim BinaryStream
	 Set BinaryStream = CreateObject("ADODB.Stream")

	'���� ������ Ÿ��
	 BinaryStream.Type = adTypeBinary

	 BinaryStream.Open
	 BinaryStream.Write BinaryData
	 ' binary -> text
	 BinaryStream.Position = 0
	 BinaryStream.Type = adTypeText

	' ��ȯ�� ������ ĳ���ͼ�
	 BinaryStream.CharSet = CharSet

	'��ȯ�� ������ ��ȯ
	 BinaryToText = BinaryStream.ReadText

	 Set BinaryStream = Nothing
End Function

'// UTF8�� ASCII ���ڿ��� ��ȯ //
Function URLDecodeUTF8(byVal pURL)
	Dim i, s1, s2, s3, u1, u2, result
	pURL = Replace(pURL,"+"," ")

	For i = 1 to Len(pURL)
		if Mid(pURL, i, 1) = "%" then
			s1 = CLng("&H" & Mid(pURL, i + 1, 2))

			'1����Ʈ�� ���(Pass)
			if (s1 < &H80) then
				result = result & Mid(pURL, i, 3)
				i = i + 2
			'2����Ʈ�� ���
			elseif ((s1 AND &HC0) = &HC0) AND ((s1 AND &HE0) <> &HE0) then
				s2 = CLng("&H" & Mid(pURL, i + 4, 2))

				u1 = (s1 AND &H1C) / &H04
				u2 = ((s1 AND &H03) * &H04 + ((s2 AND &H30) / &H10)) * &H10
				u2 = u2 + (s2 AND &H0F)
				result = result & ChrW((u1 * &H100) + u2)
				i = i + 5

			'3����Ʈ�� ���
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

'// ASCII�� UTF8 ���ڿ��� ��ȯ //
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

'// �޸��� ���е� �迭���� ������ ���� �ִ��� ��ȯ
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

'// �系 ���ӿ���
Function isTenbyTenConnect()
	Dim conIp, arrIp, tmpIp
	conIp = Request.ServerVariables("REMOTE_ADDR")
	if left(conIp,2)<>"::" then
		arrIp = split(conIp,".")
		tmpIp = Num2Str(arrIp(0),3,"0","R") & Num2Str(arrIp(1),3,"0","R") & Num2Str(arrIp(2),3,"0","R") & Num2Str(arrIp(3),3,"0","R")
	end if

	'121.78.103.60 : 15�� ����
	'10.10.10.36 : m2����
	'192.168.1.x : 15�� �,����,�λ�,�繫
	'192.168.6.x : 15�� �Ϲݸ�
	'110.11.187.233 : 15�� wireless6
	'110.93.128.x : IDC

	if tmpIp="121078103060" or tmpIp="110011187233" or (tmpIp=>"110093128001" and tmpIp<="110093128256") or (tmpIp=>"192168001001" and tmpIp<="192168001256") or (tmpIp=>"192168006001" and tmpIp<="192168006256") then
		isTenbyTenConnect = True
	else
		isTenbyTenConnect = False
	end if
End Function

'/���� �ֱ��� ������Ʈ ���� ������ ó�� '2011.11.11 �ѿ�� ����
'/������� ������ �ֽð� ������ ���� �ּ���
Sub serverupdate_underconstruction()
	dim isServerDown : isServerDown = false
		'isServerDown = true	' �����ٿ�
		isServerDown = false	' ����Ȱ��ȭ
		if isTenbyTenConnect then isServerDown = false	'�系���� ���

	if Not(isServerDown) then exit Sub

	Response.write "<html>"
	Response.write "<head><title>���� �������Դϴ�</title></head>"
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

'// ���伭�� ����� ����(���� ���ϸ�)		'/2016.04.19 �ѿ�� ����Ʈ���� ����/�̵�
function getThumbImgFromURL(furl,wd,ht,fit,ws)
	dim sCmd

	'������ ġȯ
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

	'����� Ŀ�ǵ�
	sCmd = "?cmd=thumb"
	if wd<>"" then sCmd = sCmd & "&w=" & wd
	if ht<>"" then sCmd = sCmd & "&h=" & ht
	if fit<>"" then sCmd = sCmd & "&fit=" & fit
	if ws<>"" then sCmd = sCmd & "&ws=" & ws

	'��ȯ�ּ� ��ȯ
	getThumbImgFromURL = furl & sCmd
end function

'/�������� ��ȭ��ȣ ó��
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
