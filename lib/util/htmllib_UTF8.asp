<%
'+----------------------------------------------------------------------------------------------------------------------+
'|                                               HTML 공 통   함 수 선 언                                               |
'+-------------------------------------------+--------------------------------------------------------------------------+
'|             함 수 명                      |                          기    능                                        |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| FormatDate(ddate, formatstring)          | 날짜형식을 지정된 문자형으로 변환                            |
'|                                          | 사용예 : printdate = FormatDate(now(),"0000.00.00")          |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| GetImageSubFolderByItemid(byval iitemid)  | 이미지파일의 서브 폴더명을 반환한다.                                     |
'|                                           | 사용예 : SubFolder = GetImageSubFolderByItemid(1126)                     |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| db2html(checkvalue)                       | DB의 내용을 HTML에 사용할 수 있도록 변환                                 |
'|                                           | 사용예 : Contents = db2html("DB의 내용")                                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| html2db(checkvalue)                       | 사용자가 입력한 내용을 DB에 넣을 수 있도록 변환                          |
'|                                           | 사용예 : Contents = html2db("저장할 내용")                               |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| nl2br(checkvalue)                         | 내용의 새줄(vbCrLf)을 "<br>"태그로 치환하여 반환                         |
'|                                           | 사용예 : Contents = nl2br("내용")                                        |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| CurrFormat(byVal v)                       | 숫자를 3자리 구분의 문자열로 변환                                        |
'|                                           | 사용예 : strNum = CurrFormat(1230) → "1,230"                            |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Format00(n,orgData)                       | 숫자를 0으로 채워진 지정된 길이의 문자열로 변환                          |
'|                                           | 사용예 : strNum = Format00(5,123) → "00123"                             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| FormatCode(itemcode)                      | 제품 일련번호를 6자리의 문자열로 변환                                    |
'|                                           | 사용예 : itemCode = FormatCode(2654) → "002654"                         |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| GetCurrentTimeFormat()                    | 현재시간을 문자열로 반환 (yyyymmddhhmmss)                                |
'|                                           | 사용예 : strNow = GetCurrentTimeFormat() → "20060508101833"             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| GetListImageUrl(byval itemid)             | 제품번호에 맞는 리스트 이미지 및 폴더 반환                               |
'|                                           | 사용예 : img = GetListImageUrl("53100") → "/image/list/L000053100.jpg"  |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| DDotFormat(byval str,byval n)             | 내용을 지정한 길이로 자른다.                                             |
'|                                           | 사용예 : strShort = DDotFormat("내용입니다.",3) → "내용입..."           |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| stripHTML(strng)                          | 내용 중 HTML태그를 없앤다.                                               |
'|                                           | 사용예 : Contents = stripHTML("<b>내용</b>") → " 내용 "                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| getFileExtention(strFile)                 | 파일명의 확장자를 반환한다.                                              |
'|                                           | 사용예 : ext = getFileExtention("123.jpg") → "jpg"                      |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Num2Str(inum,olen,cChr,oalign)   		 | 숫자를 지정한 길이의 문자열로 변환한다.                      			|
'|                                   		 | 사용예 : Num2Str(425,4,"0","R") → 0425                      			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ChkIIF(trueOrFalse, trueVal, falseVal)    | like iif function                                                        |
'|                                           | 사용예 : ChkIIF(1>2,"a","b") → "b"                                       |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Alert_return(strMSG)                      | 경고창 띄운후 이전으로 돌아간다.                            				|
'|                                           | 사용예 : Call Alert_return("뒤로 돌아갑니다.")               			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Alert_close(strMSG)                       | 경고창 띄운후 현재창을 닫는다.                               			|
'|                                           | 사용예 : Call Alert_close("창을 닫습니다.")                  			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| Alert_move(strMSG,targetURL)              | 경고창 띄운후 지정페이지로 이동한다.                         			|
'|                                           | 사용예 : Call Alert_move("이동합니다.","/index.asp")         			|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chrbyte(str,chrlen,dot)                   | 지정길이로 문자열 자르기                                                 |
'|                                           | 사용예 : chrbyte("안녕하세요",3,"Y") → 안녕...                           |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chkPasswordComplex(uid,pwd)               | 비밀번호 정책의 복잡성을 만족하는지 검사하고 그 이유를 반환              |
'|                                           | 사용예 : chkPasswordComplex("kobula","abcd")                             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chkPasswordComplexNonid(pwd)         	 | 아이디가 없고비밀번호 정책의 복잡성을만족하는지 검사하고 그 이유를 반환  |
'|                                           | 사용예 : chkPasswordComplexNonid("abcd")                             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| chkWord(str,patrn)                        | 문자열의 형식을 정규식으로 검사                                          |
'|                                           | 사용예 : chkWord("abcd","[^-a-zA-Z0-9/ ]") : 영어숫자만 허용             |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ParsingPhoneNumber(str,patrn)             | 전화번호에 대시 추가                                                     |
'|                                           | 사용예 : ParsingPhoneNumber("0112223333") :                              |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceBracket(strng)                     | 꺽은괄호 태그로 치환('<', '>')                                           |
'|                                           | 사용예 : ReplaceBracket("<>") → &lt;&gt;                                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceBracketOther(strng)                | 꺽은괄호 다른 괄호로 치환('<', '>')                                        |
'|                                           | 사용예 : ReplaceBracketOther("<>") → []                                 |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceScript(strng)                      | Script Tag 치환                                                          |
'|                                           | 사용예 : ReplaceScript("<script") → &lt;script                           |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| getNumeric(strNum)                        | 문자열에서 숫자만 추출 변환                                              |
'|                                           | 사용예 : getNumeric("a45d61*124") -> 461124                              |
'+-------------------------------------------+--------------------------------------------------------------------------+
'| RepWord(str,patrn,repval)                | 정규식 패턴을 사용한 문자열 처리                             				|
'|                                          | 사용예 : RepWord(SearchText,"[^가-힣a-zA-Z0-9\s]","")      			  	|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| ReplaceRequestSpecialChar(strng)        	| 특수 문자 제거(' ,--)                                        				|
'|                                          | 사용예 : cont = ReplaceRequestSpecialChar(Rs("strng"))       				|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| checkNotValidHTML(ostr)                  | 내용에 금지된 HTML태그가 있는지 검사                         				|
'|                                          | 사용예 : checkNotValidHTML("<script...") → true             				|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| minutechagehour(v)                 		| 분단위를 시간단위으로 짤라서 반환                      					|
'|                                          | 사용예 : minutechagehour(v)             									|
'+-------------------------------------------+--------------------------------------------------------------------------+
'| BinaryToText(BinaryData, CharSet)         | 바이너리 데이터 TEXT형태로 변환                                          |
'|                                           | 사용예 : BinaryToText(objXML.ResponseBody, "euc-kr")                     |
'+------------------------------------------+---------------------------------------------------------------------------+
'| URLEncodeUTF8(byVal szSource)            | ASCII을 UTF8 문자열로 변환                                                |
'|                                          | 사용예 : strUF8 = URLEncodeUTF8(STR)                                      |
'+------------------------------------------+---------------------------------------------------------------------------+
'| URLDecodeUTF8(byVal pURL)                | UTF8을 ASCII 문자열로 변환                                                |
'|                                          | 사용예 : strASC = URLDecodeUTF8(URL)                                      |
'+------------------------------------------+---------------------------------------------------------------------------+
'| chkArrValue(aVal,cVal)                    | 콤마로 구분된 배열값에 지정된 값이 있는지 반환                           |
'|                                           | 사용예 : chkArrValue("A,B,C", "B") → true                                |
'+-------------------------------------------+--------------------------------------------------------------------------+

function G_IsLocalDev()
	G_IsLocalDev = (application("Svr_Info")="Dev") AND (request.ServerVariables("LOCAL_ADDR")="::1" or request.ServerVariables("LOCAL_ADDR")="127.0.0.1")
end function

''쿠키에 넣을때 사용 / 아이디 단방향 해쉬값 : 사용시 md5 필요. (md5 부하 심할경우 component, db 이용 가능)
function HashTenID(byval oid)
    dim orgid : orgid = LCASE(oid)
    dim hashid

    HashTenID = orgid
    if Len(orgid)<1 then Exit function      ''빈값인경우 원래값
    if Len(orgid)<2 then orgid=orgid+"1"    ''길이가1일경우 오류피함.


    hashid = Right(orgid,4) + Left(orgid,Len(orgid)-1)
    hashid = Right(hashid,5) + Left(hashid,Len(hashid)-2)
    hashid = Right(hashid,6) + Left(hashid,Len(hashid)-3)
    hashid = Right(hashid,7) + Left(hashid,Len(hashid)-4)
    hashid = Right(hashid,8) + Left(hashid,Len(hashid)-5)
    HashTenID = MD5(hashid)

end function

'// 날짜를 지정된 문자형으로 변환 //
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

'' 기존 디비에 이전 형식 있음.. 차후 삭제
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

'' 2008 03 수정 - Eastone
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

'// 문자열내 CR/LF를 공백으로 치환 //
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

	'// 숫자를 지정한 길이의 문자열로 반환 //
	Function Num2Str(inum,olen,cChr,oalign)
		Dim i, ilen, strChr

		ilen = len(Cstr(inum))
		strChr = ""

		if ilen < olen then
			for i=1 to olen-ilen
				strChr = strChr & cChr
			next
		end if

		'결합방법에따른 결과 분기
		if oalign="L" then
			'왼쪽기준
			Num2Str = inum & strChr
		else
			'오른쪽 기준 (기본값)
			Num2Str = strChr & inum
		end if

    End Function


'// 문자열을 잘라 원하는 위치의 값을 반환 //
function SplitValue(orgStr,delim,pos)
    dim buf
    SplitValue = ""
    if IsNULL(orgStr) then Exit function
    if (Len(delim)<1) then Exit function
    buf = split(orgStr,delim)

    if UBound(buf)<pos then Exit function

    SplitValue = buf(pos)
end function


'// 파라메터 길이 체크 후 Maxlen 이하로 돌려줌 Code, id 등의 Param 에 사용 //
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


'// 값비교 후 Return 값 like iif function
Function ChkIIF(trueOrFalse, trueVal, falseVal)
	if (trueOrFalse) then
	    ChkIIF = trueVal
	else
	    ChkIIF = falseVal
	end if
End Function

'// 경고문 출력후 뒤로가기 //
Sub Alert_return(strMSG)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"history.back();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub


'// 경고문 출력후 창닫기 //
Sub Alert_close(strMSG)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"self.close();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub


'// 경고문 출력후 지정 페이지로 이동 //
Sub Alert_move(strMSG,targetURL)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"self.location='" & targetURL & "';" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub

'// 지정길이로 문자열 자르기 //
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

'// 패스워드 복잡성 검사 함수(기존버전 길이6, 로그인시 체크함.)		//2017.09.25 한용민 생성
Function chkPasswordComplex_Len6Ver(uid,pwd)
	dim msg, i, sT, sN, numAlpha, numNums, numSpecials, buf, index
    numAlpha = 0
    numNums = 0
    numSpecials = 0
	msg = ""

	'비밀번호 길이 검사
	if len(pwd)<8 then
		msg = msg & "- 비밀번호는 최소 8자리이상으로 입력해주세요.\n"
	end if

	'아이디와 동일 또는 포함하고 있는가?
	if instr(lcase(pwd),lcase(uid))>0 then
		msg = msg & "- 아이디와 동일하거나 아이디를 포함하고 있는 비밀번호입니다.\n"
	end if

	'## 복잡성을 만족하는가?
	'같은문자 3번 연속 금지
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
			msg = msg & "- 같은문자가 3번 연속으로 쓰였습니다.\n"
			exit for
		end if
	next

'정규식 똑바로 안묵네. 머꼬
'	if chkWord(pwd,"[^-a-zA-Z]") then
'		numAlpha = numAlpha + 1
'	end if
'	if chkWord(pwd,"[^-0-9 ]") then
'		numNums = numNums + 1
'	end if
'	if chkWord(pwd,"[~!@\#$%<>^&*\()\-=+_\’]") then
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

	'// 3가지 조합
    if (numAlpha>0 and numNums>0 and numSpecials>0) then
    	if (len(pwd) >= 8) then
    	else
    		msg = msg & "- 새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)\n"
    	end if

	'// 2가지 조합
    elseif ((numAlpha>0 and numNums>0) or (numAlpha>0 and numSpecials>0) or (numNums>0 and numSpecials>0)) then
    	if (len(pwd) >= 10) then
    	else
    		msg = msg & "- 새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)\n"
    	end if

    else
    	msg = msg & "- 새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)\n"
    end if

	'결과 반환
	chkPasswordComplex_Len6Ver = msg
end Function

'// 패스워드 복잡성 검사 함수		//2017.09.25 한용민 생성
Function chkPasswordComplex(uid,pwd)
	dim msg, i, sT, sN, numAlpha, numNums, numSpecials, buf, index
    numAlpha = 0
    numNums = 0
    numSpecials = 0
	msg = ""

	'비밀번호 길이 검사
	if len(pwd)<8 then
		msg = msg & "- 비밀번호는 최소 8자리이상으로 입력해주세요.\n"
	end if

	'아이디와 동일 또는 포함하고 있는가?
	if instr(lcase(pwd),lcase(uid))>0 then
		msg = msg & "- 아이디와 동일하거나 아이디를 포함하고 있는 비밀번호입니다.\n"
	end if

	'## 복잡성을 만족하는가?
	'같은문자 3번 연속 금지
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
			msg = msg & "- 같은문자가 3번 연속으로 쓰였습니다.\n"
			exit for
		end if
	next

'정규식 똑바로 안묵네. 머꼬
'	if chkWord(pwd,"[^-a-zA-Z]") then
'		numAlpha = numAlpha + 1
'	end if
'	if chkWord(pwd,"[^-0-9 ]") then
'		numNums = numNums + 1
'	end if
'	if chkWord(pwd,"[~!@\#$%<>^&*\()\-=+_\’]") then
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

	'// 3가지 조합
    if (numAlpha>0 and numNums>0 and numSpecials>0) then
    	if (len(pwd) >= 8) then
    	else
    		msg = msg & "- 새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)\n"
    	end if

	'// 2가지 조합
    elseif ((numAlpha>0 and numNums>0) or (numAlpha>0 and numSpecials>0) or (numNums>0 and numSpecials>0)) then
    	if (len(pwd) >= 10) then
    	else
    		msg = msg & "- 새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)\n"
    	end if

    else
    	msg = msg & "- 새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)\n"
    end if

	'결과 반환
	chkPasswordComplex = msg
end Function

'// 패스워드 복잡성 검사 함수
Function chkPasswordComplexNonID(pwd)
	dim msg, i, sT, sN
	msg = ""

	'비밀번호 길이 검사
	if len(pwd)<8 then
		msg = msg & "- 비밀번호는 최소 8자리이상으로 입력해주세요.\n"
	end if


	'## 복잡성을 만족하는가?
	'같은문자 3번 연속 금지
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
			msg = msg & "- 같은문자가 3번 연속으로 쓰였습니다.\n"
			exit for
		end if
	next
	'영문/숫자의 조합
	if chkWord(pwd,"[^-a-zA-Z]") or chkWord(pwd,"[^-0-9 ]") then
		msg = msg & "- 비밀번호는 반드시 알파벳과 숫자를 조합해서 만들어야합니다.\n"
	end if

	'결과 반환
	chkPasswordComplexNonID = msg
end Function

'//정규식 문자열 검사
Function chkWord(str,patrn)
    Dim regEx, match, matches

    SET regEx = New RegExp

    regEx.Pattern = patrn            ' 패턴을 설정합니다.
    regEx.IgnoreCase = True      ' 대/소문자를 구분하지 않도록 합니다.
    regEx.Global = True             ' 전체 문자열을 검색하도록 설정합니다.

    SET Matches = regEx.Execute(str)

    if 0 < Matches.count then
        chkWord = false
    Else
        chkWord = true
    end if
End Function

'// 전화번호에 대시 추가
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


'''''==================  2009 추가

' response.write 함수
Function rw(ByVal str)
	response.write str & "<br>"
End Function

' Null을 공백으로 치환
Function null2blank(ByVal v)
	If IsNull(v) Then
		null2blank = ""
	Else
		null2blank = v
	End If
End Function

'// 큰따옴표 input 박스 value=""에 사용할때 치환
Function doubleQuote(ByVal v)
	If IsNull(v) Then
		doubleQuote = ""
	Else
		doubleQuote = Replace(v, """","&quot;")
	End If
End Function


' request 대체 함수(파라미터명, 디폴트값)
Function req(ByVal param, ByVal value)
'	VarType Return 값
'	0 (공백)
'	1 (널)
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
		If Not IsNumeric(tmpValue) Then	' 숫자가 아니면
			tmpValue = value
		End If
		tmpValue = CDbl(tmpValue)
	Else
		tmpValue = Trim(Request(param))
		If tmpValue = "" Then			' Request값이 없으면
			tmpValue = value
		End If
	End If
	req = tmpValue

End Function

Sub sbDisplayPaging(ByVal strCurrentPage, ByVal intTotalRecord, ByVal intRecordPerPage, ByVal intBlockPerPage)

	'변수 선언
	Dim intCurrentPage, strCurrentPath
	Dim intStartBlock, intEndBlock, intTotalPage
	Dim strParamName, intLoop

	'현재 페이지 설정
	intCurrentPage = Mid(strCurrentPage, InStr(strCurrentPage, "=")+1)		'현재 페이지 값
	strCurrentPage = Left(strCurrentPage, InStr(strCurrentPage, "=")-1)		'페이지 폼값 변수명

	'현재 페이지 명
	strCurrentPath = Request.ServerVariables("Script_Name")

	'해당페이지에 표시되는 시작페이지와 마지막페이지 설정
	intStartBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + 1
	intEndBlock = Int((intCurrentPage - 1) / intBlockPerPage) * intBlockPerPage + intBlockPerPage

	'총 페이지 수 설정
	intTotalPage =  -(int(-(intTotalRecord/intRecordPerPage)))

	'폼 설정 & hidden 파라미터 설정
	Response.Write	"<form name='frmPaging' method='get' action ='" & strCurrentPath & "'>" &_
							"<input type='hidden' name='" & strCurrentPage & "'>"			'현재 페이지

	'파라미터 값들(예: 검색어)을 hidden 파라미터로 저장한다
	strParamName = ""
	For Each strParamName In Request.Form
		If strParamName <> strCurrentPage Then

			'hidden 파라미터 값도 파라미터 검열
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.Form(strParamName),50) & "'>"
		End If
	Next
	strParamName = ""

	For Each strParamName In Request.Querystring
		If strParamName <> strCurrentPage Then
			'hidden 파라미터 값도 파라미터 검열
			Response.Write "<input type='hidden' name='" & strParamName & "' value='" & requestCheckVar(Request.QueryString(strParamName),50) & "'>"
		END IF
	Next

	Response.Write "<table border='0' cellpadding='0' cellspacing='0' class=a><tr align='center'><td>"

	'이전 페이지 이미지 설정
	If intStartBlock > 1 Then
		Response.Write "<img src='http://fiximage.10x10.co.kr/web2008/designfingers/btn_pageprev01.gif' border='0' style='cursor:hand' alt='이전 " & intBlockPerPage & " 페이지'" &_
							   "onClick='javascript:document.frmPaging." & strCurrentPage & ".value=" & intStartBlock - intBlockPerPage & ";document.frmPaging.submit();'>"
	Else
		Response.Write "<img src='http://fiximage.10x10.co.kr/web2009/common/btn_pageprev01.gif' border='0' >"
	End If

	Response.Write "</td><td>&nbsp;"

	'페이징 출력
	If intTotalPage > 1 Then
		For intLoop = intStartBlock To intEndBlock
			If intLoop > intTotalPage Then Exit For

			If Int(intLoop) <> Int(intStartBlock) Then Response.Write "|"

			If Int(intLoop) = Int(intCurrentPage) Then		'현재 페이지
				Response.Write "&nbsp;<span class='text01'><strong>" & intLoop & "</strong></span>&nbsp;"
			Else															'그 외 페이지
				Response.Write "&nbsp;<a href='javascript:document.frmPaging." & strCurrentPage & ".value=" & intLoop & ";document.frmPaging.submit();'><font class='text01'>" & intLoop & "</font></a>&nbsp;"
			End If

		Next
	Else		'한 페이지만 존재 할때
		Response.Write "&nbsp;<span class='text01'><strong>1</strong></span>&nbsp;"
	End If

	Response.Write "&nbsp;</td><td>"

	'다음 페이지 이미지 설정
	If Int(intEndBlock) < Int(intTotalPage) Then
		Response.Write "<img src='http://fiximage.10x10.co.kr/web2008/designfingers/btn_pagenext01.gif' border='0' style='cursor:hand' alt='다음 " & intBlockPerPage & " 페이지'" &_
							   "onClick='javascript:document.frmPaging." & strCurrentPage & ".value=" & intEndBlock+1 & ";document.frmPaging.submit();'>"
	Else
	    Response.Write "<img src='http://fiximage.10x10.co.kr/web2009/common/btn_pagenext01.gif' border='0' >"
	End If

	Response.Write "</td></tr></form></table>"

End Sub



' 등록,수정,삭제 모드 텍스트 리턴
Function getModeName(ByVal mode)
    Select Case mode
        Case "INS"	: getModeName = "등록"
        Case "UPD"	: getModeName = "수정"
        Case "DEL"	: getModeName = "삭제"
        Case "FIN"	: getModeName = "완료"
        Case Else	: getModeName = "미정"
    End Select
End Function

'// 꺽은괄호 HTML코드로 치환 //
' db2html 이랑 충돌나서 사용가능한곳만 적용하세요.
Function ReplaceBracket(strng)
	if isnull(strng) then exit Function

	strng = Replace(strng,"<","&lt;")
	strng = Replace(strng,">","&gt;")
	ReplaceBracket = strng
end Function

'// 꺽은괄호 다른 괄호로 치환 //
Function ReplaceBracketOther(strng)
	if isnull(strng) then exit Function

	strng = Replace(strng,"<","[")
	strng = Replace(strng,">","]")
	ReplaceBracketOther = strng
end Function

'// Script Tag치환 //
Function ReplaceScript(strng)
	if isnull(strng) then exit Function

	strng = Replace(strng,"<script","[script")
	strng = Replace(strng,"</script","[/script")
	strng = Replace(strng,"<iframe","[iframe")
	strng = Replace(strng,"</iframe","[/iframe")
	ReplaceScript = strng
end Function


' 정규식 함수
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

'// 문자열에서 숫자만 추출 변환
Function getNumeric(strNum)
	Dim lp, tmpNo, strRst
	For lp=1 to len(strNum)
		tmpNo = mid(strNum, lp, 1)
		if asc(tmpNo)>47 and asc(tmpNo)<58 then
			strRst = strRst & tmpNo
		end if
	Next
	getNumeric = strRst
End Function

'// 정규식 패턴지정 문자열 처리/반환
Function RepWord(str,patrn,repval)
	Dim regEx

	SET regEx = New RegExp
	regEx.Pattern = patrn			' 패턴을 설정.
	regEx.IgnoreCase = True			' 대/소문자를 구분하지 않도록 .
	regEx.Global = True				' 전체 문자열을 검색하도록 설정.
	RepWord = regEx.Replace(str,repval)
End Function

'/사용금지		'/lib/function.asp 에 getUserLevelColor 공용함수 사용할것. font color 로 먹일것.		'/2016.07.20 한용민
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

'//문자열내 특수문자 제거
function ReplaceRequestSpecialChar(v)
	ReplaceRequestSpecialChar = replace(v,"'","")
	ReplaceRequestSpecialChar = replace(ReplaceRequestSpecialChar,"--","")
end function

'//올림 함수
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

'//올림 함수
function ceilValue(iValue)
 if iValue <>  round(iValue) then
  ceilValue = fix(iValue) + 1
 else
  ceilValue = iValue
 end if
end function

'// 지정수만큼 지정한 문자로 바꿈)
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

'// 내용에 금지된 HTML태그가 있는지 검사 //
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
end function

'' 2015/10/06 checkNotValidHTML ahref, imgsrc 다 막힘;; 새로 만듬.
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



'// 경고문 출력후 창닫고 오픈창 리로드 -2011.02.23 정윤정추가 //
Sub Alert_closenreload(strMSG)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"window.opener.location.reload();"& vbCrLf &_
			"self.close();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub

'// 경고문 출력후 창닫고 오픈창 타겟주소로 이동 -2011.02.23 정윤정추가 //
Sub Alert_closenmove(strMSG,targetURL)
	dim strTemp
	strTemp = 	"<script language='javascript'>" & vbCrLf &_
			"alert('" & strMSG & "');" & vbCrLf &_
			"window.opener.location.href ='" & targetURL & "';" & vbCrLf &_
			"self.close();" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub

'//분단위를 시간단위으로 짤라서 반환	'/2011.03.31 한용민 생성
function minutechagehour(v)
	dim tmpval , tmph , tmpm

	if v = "" or isnull(v) or v = 0 then
		minutechagehour = ""
	else
		tmph = int(v / 60)	'시간단위
		tmpm = v - (tmph * 60)	'분단위

		if tmph <> 0 then tmpval = tmpval & tmph & "시간 "
		if tmpm <> 0 then tmpval = tmpval & tmpm & "분"

		minutechagehour = tmpval
	end if
end function

'//바이너리 데이터 TEXT형태로 변환
Function  BinaryToText(BinaryData, CharSet)
	 Const adTypeText = 2
	 Const adTypeBinary = 1

	 Dim BinaryStream
	 Set BinaryStream = CreateObject("ADODB.Stream")

	'원본 데이터 타입
	 BinaryStream.Type = adTypeBinary

	 BinaryStream.Open
	 BinaryStream.Write BinaryData
	 ' binary -> text
	 BinaryStream.Position = 0
	 BinaryStream.Type = adTypeText

	' 변환할 데이터 캐릭터셋
	 BinaryStream.CharSet = CharSet

	'변환한 데이터 반환
	 BinaryToText = BinaryStream.ReadText

	 Set BinaryStream = Nothing
End Function

'// UTF8을 ASCII 문자열로 변환 //
Function URLDecodeUTF8(byVal pURL)
	Dim i, s1, s2, s3, u1, u2, result
	pURL = Replace(pURL,"+"," ")

	For i = 1 to Len(pURL)
		if Mid(pURL, i, 1) = "%" then
			s1 = CLng("&H" & Mid(pURL, i + 1, 2))

			'1바이트일 경우(Pass)
			if (s1 < &H80) then
				result = result & Mid(pURL, i, 3)
				i = i + 2
			'2바이트일 경우
			elseif ((s1 AND &HC0) = &HC0) AND ((s1 AND &HE0) <> &HE0) then
				s2 = CLng("&H" & Mid(pURL, i + 4, 2))

				u1 = (s1 AND &H1C) / &H04
				u2 = ((s1 AND &H03) * &H04 + ((s2 AND &H30) / &H10)) * &H10
				u2 = u2 + (s2 AND &H0F)
				result = result & ChrW((u1 * &H100) + u2)
				i = i + 5

			'3바이트일 경우
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

'// ASCII을 UTF8 문자열로 변환 //
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

'// 콤마로 구분된 배열값에 지정된 값이 있는지 반환
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

'// 사내 접속여부
Function isTenbyTenConnect()
	Dim conIp, arrIp, tmpIp
	conIp = Request.ServerVariables("REMOTE_ADDR")
	if left(conIp,2)<>"::" then
		arrIp = split(conIp,".")
		tmpIp = Num2Str(arrIp(0),3,"0","R") & Num2Str(arrIp(1),3,"0","R") & Num2Str(arrIp(2),3,"0","R") & Num2Str(arrIp(3),3,"0","R")
	end if

	'121.78.103.60 : 15층 유선
	'10.10.10.36 : m2서버
	'192.168.1.x : 15층 운영,개발,인사,재무
	'192.168.6.x : 15층 일반망
	'110.11.187.233 : 15층 wireless6
	'110.93.128.x : IDC

	if tmpIp="121078103060" or tmpIp="110011187233" or (tmpIp=>"110093128001" and tmpIp<="110093128256") or (tmpIp=>"192168001001" and tmpIp<="192168001256") or (tmpIp=>"192168006001" and tmpIp<="192168006256") then
		isTenbyTenConnect = True
	else
		isTenbyTenConnect = False
	end if
End Function

'/서버 주기적 업데이트 위한 공사중 처리 '2011.11.11 한용민 생성
'/리뉴얼시 이전해 주시고 지우지 말아 주세요
Sub serverupdate_underconstruction()
	dim isServerDown : isServerDown = false
		'isServerDown = true	' 서버다운
		isServerDown = false	' 서버활성화
		if isTenbyTenConnect then isServerDown = false	'사내접속 허용

	if Not(isServerDown) then exit Sub

	Response.write "<html>"
	Response.write "<head><title>서비스 점검중입니다</title></head>"
	Response.write "<meta http-equiv='Content-Type' content='text/html;charset=utf-8' />"
	Response.write "<body>"
	Response.write "<table width='100%' height='100%' cellpadding='0' cellspacing='0' border='0'>"
	Response.write "<tr>"
	Response.write "	<td align='center' valign='middle'><img src='http://fiximage.10x10.co.kr/web2015/common/2015_10x10_open_ready_PC.jpg' width='1104' border='0' ></td>"
	Response.write "</tr>"
	Response.write "</table>"
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

'// 포토서버 썸네일 제작(기존 파일명)		'/2016.04.19 한용민 프론트에서 복사/이동
function getThumbImgFromURL(furl,wd,ht,fit,ws)
	dim sCmd

	'도메인 치환
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

	'썸네일 커맨드
	sCmd = "?cmd=thumb"
	if wd<>"" then sCmd = sCmd & "&w=" & wd
	if ht<>"" then sCmd = sCmd & "&h=" & ht
	if fit<>"" then sCmd = sCmd & "&fit=" & fit
	if ws<>"" then sCmd = sCmd & "&ws=" & ws

	'변환주소 반환
	getThumbImgFromURL = furl & sCmd
end function

'/개인정보 전화번호 처리
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
