<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<%
'+---------------------------------------------------------------------------------------------------------+
'|                                   문 자 열   공 통   함 수 선 언                                        |
'+------------------------------------------+--------------------------------------------------------------+
'|                함 수 명                  |                          기    능                            |
'+------------------------------------------+--------------------------------------------------------------+
'| Num2Str(inum,olen,cChr,oalign)           | 숫자를 지정한 길이의 문자열로 변환한다.                      |
'|                                          | 사용예 : Num2Str(425,4,"0","R") → 0425                      |
'+------------------------------------------+--------------------------------------------------------------+
'| SplitValue(orgStr,delim,pos)             | 문자열을 잘라 원하는 위치의 값을 반환                        |
'|                                          | 사용예 : SplitValue("A/B/C","/","2") → B                    |
'+------------------------------------------+--------------------------------------------------------------+
'| CurrFormat(byVal v)                      | 숫자열을 화폐형으로 변환                                     |
'|                                          | 사용예 : CurrFormat(1250) → 1,250                           |
'+------------------------------------------+--------------------------------------------------------------+
'| chrbyte(str,chrlen,dot)                  | 지정길이로 문자열 자르기                                     |
'|                                          | 사용예 : chrbyte("안녕하세요",3,"Y") → 안녕...              |
'+------------------------------------------+--------------------------------------------------------------+
'| URLDecodeUTF8(byVal pURL)                | UTF8을 ASCII 문자열로 변환                                   |
'|                                          | 사용예 : strASC = URLDecodeUTF8(URL)                         |
'+------------------------------------------+--------------------------------------------------------------+
'| URLEncodeUTF8(byVal szSource)            | ASCII을 UTF8 문자열로 변환                                   |
'|                                          | 사용예 : strUF8 = URLEncodeUTF8(STR)                         |
'+------------------------------------------+--------------------------------------------------------------+
'| cvtDBChkBoxData(strArr,mrk)              | CheckBox로 넘어온 배열을 DB에 쓸 수 있는 문자열로 변환       |
'|                                          | 사용예 : cvtDBChkBoxData("A, B","Y") -> 'A','B'              |
'+------------------------------------------+--------------------------------------------------------------+
'| printUserId(strID,lng,chr)               | 회원 아이디를 출력할 때 지정수만큼 문자 치환. 아이디 노출X   |
'|                                          | 사용예 : printUserId("kobula",2,"*") -> 'kobu**'             |
'+------------------------------------------+--------------------------------------------------------------+
'| getNumeric(strNum)                       | 문자열에서 숫자만 추출 변환                                  |
'|                                          | 사용예 : getNumeric("a45d61*124") -> 461124                  |
'+------------------------------------------+--------------------------------------------------------------+
'| RepWord(str,patrn,repval)                | 정규식 패턴을 사용한 문자열 처리                             |
'|                                          | 사용예 : RepWord(SearchText,"[^가-힣a-zA-Z0-9\s]","")        |
'+------------------------------------------+--------------------------------------------------------------+
'| chkWord(str,patrn)                       | 문자열의 형식을 정규식으로 검사                              |
'|                                          | 사용예 : chkWord("abcd","[^-a-zA-Z0-9/ ]") : 영어숫자만 허용 |
'+-------------------------------------------+-------------------------------------------------------------+

'+---------------------------------------------------------------------------------------------------------+
'|                                  날 짜 관 련   공 통   함 수 선 언                                      |
'+------------------------------------------+--------------------------------------------------------------+
'|                함 수 명                  |                          기    능                            |
'+------------------------------------------+--------------------------------------------------------------+
'| FormatDate(ddate, formatstring)          | 날짜형식을 지정된 문자형으로 변환                            |
'|                                          | 사용예 : printdate = FormatDate(now(),"0000.00.00")          |
'+------------------------------------------+--------------------------------------------------------------+
'| DayOfMonth(yymmdd)                       | 입력된 날짜에 해당하는 달의 날짜수를 반환                    |
'|                                          | 사용예 : date_count = DayOfMonth("2006-08-10")               |
'+------------------------------------------+--------------------------------------------------------------+
'| WeekOfMonth(yymmdd)                      | 입력된 날짜에 해당하는 달의 주 수를 반환                     |
'|                                          | 사용예 : week_count = WeekOfMonth("2006-08-10")              |
'+------------------------------------------+--------------------------------------------------------------+
'| StartDayOfWeek(yymmdd)                   | 입력된 날짜가 속한 주의 마지막날 반환                        |
'|                                          | 사용예 : week_first = StartDayOfWeek("2006-08-10")           |
'+------------------------------------------+--------------------------------------------------------------+
'| EndDayOfWeek(yymmdd)                     | 입력된 날짜가 속한 주의 마지막날 반환                        |
'|                                          | 사용예 : week_last = EndDayOfWeek("2006-08-10")              |
'+------------------------------------------+--------------------------------------------------------------+
'| DrawOneDateBox(yyyy,mm,dd,tt)            | 날짜 선택 셀렉트박스 출력 (년원일시)                         |
'|                                          | 사용예 : call DrawOneDateBox("2006","08","10","15")          |
'+------------------------------------------+--------------------------------------------------------------+

'+---------------------------------------------------------------------------------------------------------+
'|                                    H T M L   공 통   함 수 선 언                                        |
'+------------------------------------------+--------------------------------------------------------------+
'|                함 수 명                  |                          기    능                            |
'+------------------------------------------+--------------------------------------------------------------+
'| checkNotValidHTML(ostr)                  | 내용에 금지된 HTML태그가 있는지 검사                         |
'|                                          | 사용예 : checkNotValidHTML("<script...") → true             |
'+------------------------------------------+--------------------------------------------------------------+
'| checkNotValidTxt(ostr)                   | 내용에 금지어 및 html 태그가 있는지 검사 		               |
'|                                          | 사용예 : checkNotValidTxt("http://") → true                 |
'+------------------------------------------+--------------------------------------------------------------+
'| requestCheckVar(orgval,maxlen)           | 파라메터 길이 체크 후 Maxlen 이하로 돌려줌 Code, id 등의 Param 에 사용|
'|                                          | 사용예 : requestCheckVar(request("id"),32)                   |
'+------------------------------------------+--------------------------------------------------------------+
'| db2html(checkvalue)                      | DB저장된 구문을 사이트에 쓸 수 있도록 변환                   |
'|                                          | 사용예 : Response.Write db2html(Rs("title"))                 |
'+------------------------------------------+--------------------------------------------------------------+
'| html2db(checkvalue)                      | 사이트에서 입력받은 내용을 DB에 저장할 수 있도록 변환        |
'|                                          | 사용예 : strSQL = html2db("내용을 저장합니다")               |
'+------------------------------------------+--------------------------------------------------------------+
'| nl2br(v)                                 | 문자열내 CR/LF를 <BR>태그로 치환                             |
'|                                          | 사용예 : Response.Write nl2br(Rs("contents"))                |
'+------------------------------------------+--------------------------------------------------------------+
'| nl2li(v)                                 | 문자열내 CR/LF를 </li><li>태그로 치환                             |
'|                                          | 사용예 : Response.Write nl2li(Rs("contents"))                |
'+------------------------------------------+--------------------------------------------------------------+
'| stripHTML(strng)                         | HTML태그 제거                                                |
'|                                          | 사용예 : cont = stripHTML(Rs("content"))                     |
'+------------------------------------------+--------------------------------------------------------------+
'| ReplaceRequestSpecialChar(strng)        	| 특수 문자 제거(' ,--)                                        |
'|                                          | 사용예 : cont = ReplaceRequestSpecialChar(Rs("strng"))       |
'+------------------------------------------+--------------------------------------------------------------+
'| ReplaceRequest(strng)        			| 특수 문자및 쿼리문자제거(' ,--,select)                       |
'|                                          | 사용예 : cont = ReplaceRequest(Rs("strng"))      			   |
'+------------------------------------------+--------------------------------------------------------------+
'| ReplaceBracket(strng)        			| 꺽은괄호 태그로 치환('<', '>')                               |
'|                                          | 사용예 : ReplaceBracket("<>") → &lt;&gt;                    |
'+------------------------------------------+--------------------------------------------------------------+

'+---------------------------------------------------------------------------------------------------------+
'|                                사 이 트 관 련   공 통   함 수 선 언                                     |
'+------------------------------------------+--------------------------------------------------------------+
'|                함 수 명                  |                          기    능                            |
'+------------------------------------------+--------------------------------------------------------------+
'| GetUserLevelStr(iuserlevel)              | 사용자 등급의 해당명칭을 반환                                |
'|                                          | 사용예 : GetUserLevelStr(2) → 블루                          |
'+------------------------------------------+--------------------------------------------------------------+
'| GetImageSubFolderByItemid(byval iitemid) | 상품 이미지 경로를 계산하여 반환                             |
'|                                          | 사용예 : GetImageSubFolderByItemid(35285) → 03              |
'+------------------------------------------+--------------------------------------------------------------+
'| FormatCode(itemcode)                     | 제품번호를 문자열로 변환                                     |
'|                                          | 사용예 : FormatCode(69125) → 069125                         |
'+------------------------------------------+--------------------------------------------------------------+
'| Format00(totallength,orgData)            | 숫자 형식을 000NNNN 형식으로 변환                            |
'|                                          | 사용예 : Format00(7,69125) → 0069125                        |
'+------------------------------------------+--------------------------------------------------------------+
'| GetListImageUrl(byval itemid)            | 제품이미지 반환                                              |
'|                                          | 사용예 : ListImg = GetListImageUrl(69125)                    |
'+------------------------------------------+--------------------------------------------------------------+
'| executeFile(fnm)                         | 외부파일(HTML, ASP등) 실행 함수                              |
'|                                          | 사용예 : Call executeFile("leftmenu.asp")                    |
'+------------------------------------------+--------------------------------------------------------------+
'| GetPricePercent(Sprice,Oprice,pt)        | 할인율 계산                                                  |
'|                                          | 사용예 : GetPricePercent(800,1000,2) → 20.00%               |
'+------------------------------------------+--------------------------------------------------------------+
'| GetImgSwitchOnOff(skey, tkey)            | 이미지 on / off  문자열 반환                                 |
'|                                          | 사용예 : GetImgSwitchOnOff(aa,"aa") → "on"                   |
'+------------------------------------------+--------------------------------------------------------------+
'| ChkIIF(trueOrFalse, trueVal, falseVal)   | like iif function                                            |
'|                                          | 사용예 : ChkIIF(1>2,"a","b") → "b"                           |
'+------------------------------------------+--------------------------------------------------------------+
'| Alert_return(strMSG)                     | 경고창 띄운후 이전으로 돌아간다.                             |
'|                                          | 사용예 : Call Alert_return("뒤로 돌아갑니다.")               |
'+------------------------------------------+--------------------------------------------------------------+
'| Alert_close(strMSG)                      | 경고창 띄운후 현재창을 닫는다.                               |
'|                                          | 사용예 : Call Alert_close("창을 닫습니다.")                  |
'+------------------------------------------+--------------------------------------------------------------+
'| Alert_move(strMSG,targetURL)             | 경고창 띄운후 지정페이지로 이동한다.                         |
'|                                          | 사용예 : Call Alert_move("이동합니다.","/index.asp")         |
'+------------------------------------------+--------------------------------------------------------------+
'| getTopMenuId(pageName,param)             | 파일명을 메인메뉴 고유번호로 변환                            |
'|                                          | 사용예 : contMenu = getTopMenuId("leture","cdlarge=10")      |
'+------------------------------------------+--------------------------------------------------------------+
'| fnChkNumeric(iValue)                     | 숫자여부 확인                                                |
'|                                          | 사용예 : fnChkNumeric("123")→123 / "ABC"→Alert             |
'+------------------------------------------+--------------------------------------------------------------+

'+---------------------------------------------------------------------------------------------------------+
'|                                인 증 관 련   공 통   함 수 선 언                                        |
'+------------------------------------------+--------------------------------------------------------------+
'| IsUserLoginOK()                          | [아이디]로 로그인 했는지 여부 return Boolean                 |
'|                                          | 사용예 : bool = IsUserLoginOK()                              |
'+------------------------------------------+--------------------------------------------------------------+
'| IsGuestLoginOK()                         | [주문 번호]로 로그인 했는지 여부 return Boolean              |
'|                                          | 사용예 : bool = IsGuestLoginOK()                             |
'+------------------------------------------+--------------------------------------------------------------+
'| GetLoginUserID()                         | 로그인 한 UserID                                             |
'|                                          | 사용예 : ret = getLoginUserID()                              |
'+------------------------------------------+--------------------------------------------------------------+
'| GetLoginUserName()                       | 로그인 한 UserName                                           |
'|                                          | 사용예 : ret = getLoginUserName()                            |
'+------------------------------------------+--------------------------------------------------------------+
'| GetLoginUserEmail()                      | 로그인 한 UserUserEmail                                      |
'|                                          | 사용예 : ret = getLoginUserEmail()                           |
'+------------------------------------------+--------------------------------------------------------------+
'| GetLoginUserLevel()                      | 로그인 한 UserUserLevel                                      |
'|                                          | 사용예 : ret = getLoginUserLevel()                           |
'+------------------------------------------+--------------------------------------------------------------+
'| GetLoginUserDiv()                        | 로그인 한 UserUserDiv                                        |
'|                                          | 사용예 : ret = getLoginUserDiv()                             |
'+------------------------------------------+--------------------------------------------------------------+
'| GetLoginRealNameCheck()                  | 로그인 한 실명확인 여부 ('Y','N')                            |
'|                                          | 사용예 : ret = GetLoginRealNameCheck()                       |
'+------------------------------------------+--------------------------------------------------------------+
'| GetLoginCouponCount()                    | 로그인 당시 할인권 + 상품쿠푠  갯수   - 쿠폰 받았을때 세팅 필요|
'|                                          | 사용예 : ret = GetLoginCouponCount()                         |
'+------------------------------------------+--------------------------------------------------------------+
'| GetLoginCurrentMileage()                 | 로그인 당시 마일리지   - 마일리지 변경시 세팅 필요           |
'|                                          | 사용예 : ret = GetLoginCurrentMileage()                      |
'+------------------------------------------+--------------------------------------------------------------+
'| SetLoginCouponCount(couponcount)         | 로그인 당시 할인권 + 상품쿠푠 갯수 세팅                      |
'|                                          | 사용예 : call SetLoginCouponCount(couponcount)               |
'+------------------------------------------+--------------------------------------------------------------+
'| SetLoginCurrentMileage(currmileage)      | 로그인 당시 마일리지 세팅                                    |
'|                                          | 사용예 : call SetLoginCurrentMileage(currmileage)            |
'+------------------------------------------+--------------------------------------------------------------+
'| GetGuestLoginOrderserial()               | [주문 번호]로그인 한 주문번호                                |
'|                                          | 사용예 : Call GetGuestLoginOrderserial()                     |
'+------------------------------------------+--------------------------------------------------------------+
'| fnMakePostData()            				|  post형식의 데이터  get 스트링 형태로 변경                   |
'|                                          | 사용예 : Call fnMakePostData()                     		   |
'+------------------------------------------+--------------------------------------------------------------+
'| sbPostDataToHtml()             		    | get 스트링 형태로 넘어온 데이터를 post 형태로 변경           |
'|                                          | 사용예 : Call sbPostDataToHtml()                             |
'+------------------------------------------+--------------------------------------------------------------+
'| getRealNameErrMsg(DCd)          		    | 실명확인 상세결과 코드에 따른 메시지 반환                    |
'|                                          | 사용예 : msg = getRealNameErrMsg("A")                        |
'+------------------------------------------+--------------------------------------------------------------+
'| HashTenID(byval oid)          		    | 아이디 해시값 저장 md5.asp 필요**********                    |
'|                                          | 사용예 : response.cookies("uinfo")("shix") = HashTenID(userid)|
'+------------------------------------------+--------------------------------------------------------------+
'| getEncLoginUserID()          		    | 암호화된 아이디 검증 및 로그인된 아이디 가져옴 **md5.asp 필요|
'|                                          | 사용예 : userid = getEncLoginUserID()                        |
'+------------------------------------------+--------------------------------------------------------------+

'+---------------------------------------------------------------------------------------------------------+
'|                                2009 리뉴얼 추가 함수                                                    |
'+------------------------------------------+--------------------------------------------------------------+
'| rw(), rwe()                              | response.write 축약, rwe는 dbget.close()	:	response.End 포함                 |
'|                                          | 사용예 : rw 변수, rwe 변수                                   |
'+------------------------------------------+--------------------------------------------------------------+
'| null2blank()                             | Null을 Blank 공백으로 치환, 레코드셋에서 사용                |
'|                                          | 사용예 : 속성 = null2blank(rsget("컬럼"))                    |
'+------------------------------------------+--------------------------------------------------------------+
'| req()                                    | request 축약 + 디폴트                                        |
'|                                          | 사용예 : req("필드", 기본값)                                 |
'+------------------------------------------+--------------------------------------------------------------+
'| getThisFullURL()                         | 현재 페이지 URL + 모든 파라미터 QueryString                  |
'|                                          | 사용예 : 변수 = getThisFullURL()                             |
'+------------------------------------------+--------------------------------------------------------------+
'| fnPaging()                               | 2009 페이징 함수, 페이지값을 넘기는 파라미터명에 유의할 것   |
'| 사용예 : <%=fnPaging(페이지파라미터, 토탈레코드카운트, 현재페이지, 페이지사이즈, 블럭사이즈)%           |
'+------------------------------------------+--------------------------------------------------------------+
'| fnPaging2016()                           | 2016 페이징 함수, 페이지값을 넘기는 파라미터명에 유의할 것   |
'| 사용예 : <%=fnPaging(페이지파라미터, 토탈레코드카운트, 현재페이지, 페이지사이즈, 블럭사이즈)%           |
'+------------------------------------------+--------------------------------------------------------------+
'| fnPagingSSL()                            | 2009 SSL용 페이징 함수, 페이지값을 넘기는 파라미터명에 유의  |
'| 사용예 : <%=fnPagingSSL(페이지파라미터, 토탈레코드카운트, 현재페이지, 페이지사이즈, 블럭사이즈)%        |
'+------------------------------------------+--------------------------------------------------------------+
'|                                2016 리뉴얼 추가 함수                                                    |
'+------------------------------------------+--------------------------------------------------------------+
'| chkArrValue(aVal,cVal)                    | 콤마로 구분된 배열값에 지정된 값이 있는지 반환              |
'|                                           | 사용예 : chkArrValue("A,B,C", "B") → true                   |
'+-------------------------------------------+-------------------------------------------------------------+

'/서버 주기적 업데이트 위한 공사중 처리 '2011.11.11 한용민 생성
'/리뉴얼시 이전해 주시고 지우지 말아 주세요
Sub serverupdate_underconstruction()
	dim isServerDown : isServerDown = false
		'isServerDown = true	' 서버다운
		isServerDown = false	' 서버활성화

	if Not(isServerDown) then exit Sub

	If Response.Buffer Then
		Response.Clear
		Response.Expires = 0
	End If

	Response.write "<html>"
	Response.write "<head><title>더핑거스 - 서비스 점검중입니다</title></head>"
	Response.write "<meta http-equiv='Content-Type' content='text/html;charset=UTF-8' />"
	Response.write "<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />" & vbCrLf
	Response.write "<style type=""text/css"">" & vbCrLf
	Response.write "html, body, div, p {margin:0; padding:0;}" & vbCrLf
	Response.write "img {border:0;}" & vbCrLf
	Response.write ".prepare, .prepare #wrap {background:none;}" & vbCrLf
	Response.write ".prepare {padding:300px 0; text-align:center;}" & vbCrLf
	Response.write "</style>" & vbCrLf
	Response.write "</head>" & vbCrLf
	Response.write "<body class=""prepare"">" & vbCrLf
	Response.write "<div id=""wrap"">" & vbCrLf
	Response.write "	<p><img src=""http://image.thefingers.co.kr/2016/common/img_prepare.gif"" alt=""더 나은 서비스를 위한 사이트 점검중입니다."" /></p>" & vbCrLf
	Response.write "</div>" & vbCrLf
	Response.write "</body>" & vbCrLf
	Response.write "</html>" & vbCrLf
	response.End
End Sub

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

	if InStr(LcaseStr,"imgsrc")>0 or InStr(LcaseStr,"ahref")>0 or InStr(LcaseStr,"src=")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"<body")>0 or InStr(LcaseStr,"<input")>0 or InStr(LcaseStr,"<select")>0 or InStr(LcaseStr,"<textarea")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"onload=")>0 or InStr(LcaseStr,"onunload=")>0 or InStr(LcaseStr,"onclick=")>0 or InStr(LcaseStr,"onscroll=")>0 or InStr(LcaseStr,"onblur=")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"onkeyup=")>0 or InStr(LcaseStr,"onkeydown=")>0 or InStr(LcaseStr,"onkeypress=")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,"onmouseover=")>0 or InStr(LcaseStr,"onmouseout=")>0 then
		checkNotValidHTML = true
	end if

	if InStr(LcaseStr,".wmf")>0 or InStr(LcaseStr,".js")>0 then
		checkNotValidHTML = true
	end if
end function


'// 내용에 금지어가 있는지 검사 //
function checkNotValidTxt(ostr)
	dim LcaseStr, sNotValid, arrNotValid,i
	checkNotValidTxt = false
		
	' html 태그 검사
	IF (checkNotValidHTML(ostr)) THEN
		checkNotValidTxt = true
		exit function
	END IF	
	
	'금지어 정의
	sNotValid = "010.;010-;011.;011-;016.;016-;018.;018-;019.;019-"
	arrNotValid = split(sNotValid,";")
	
	LcaseStr = Lcase(ostr)
	LcaseStr = Replace(LcaseStr," ","")

	'금지어 검사
	for i =0 to uBound(arrNotValid)	
	if InStr(LcaseStr,trim(arrNotValid(i)))>0 then
		checkNotValidTxt = true	
		exit function
	end if
	next
	
end function	

'// 파라메터 길이 체크 후 Maxlen 이하로 돌려줌 Code, id 등의 Param 에 사용 //
function requestCheckVar(orgval,maxlen)
	requestCheckVar = trim(orgval)
	requestCheckVar = replace(requestCheckVar,"'","")
'	requestCheckVar = replace(requestCheckVar,"declare","")
'	requestCheckVar = replace(requestCheckVar,"DECLARE","")
'	requestCheckVar = replace(requestCheckVar,"Declare","")
	requestCheckVar = Left(requestCheckVar,maxlen)
end function


'// 사용자 등급의 해당명칭을 반환 //
function GetUserLevelStr(iuserlevel)
	Select Case CStr(iuserlevel)
		Case "1"
			GetUserLevelStr = "<span class='Seed'>Seed</span>"
		Case "2"
			GetUserLevelStr = "<span class='Bud'>Bud</span>"
		Case "3"
			GetUserLevelStr = "<span class='Leaf'>Leaf</span>"
		Case "4"
			GetUserLevelStr = "<span class='Bean'>Bean</span>"
		Case "5"
			GetUserLevelStr = "<span class='Tree'>Tree</span>"
		Case "6"
			GetUserLevelStr = "<span class='Staff'>STAFF</span>"
		Case Else
			GetUserLevelStr = "<span class='Seed'>Seed</span>"
	end Select
end function

'// 사용자 등급의 해당명칭의 CSS 클래스를 반환 //
Function GetUserLevelCSSClass()
	Select Case CStr(request.cookies("uinfo")("muserlevel"))
		Case "1"	GetUserLevelCSSClass = "sSeed"		''2016모바일에서 변경 Seed->sSeed
		Case "2"	GetUserLevelCSSClass = "sBud"		''2016모바일에서 변경 Bud->sBud
		Case "3"	GetUserLevelCSSClass = "sLeaf"		''2016모바일에서 변경 Leaf->sLeaf
		Case "4"	GetUserLevelCSSClass = "sBean"		''2016모바일에서 변경 Bean->sBean
		Case "5"	GetUserLevelCSSClass = "sTree"		''2016모바일에서 변경 Tree->sTree
		Case "6"	GetUserLevelCSSClass = "sStaff"		''2016모바일에서 변경 Staff->sStaff
		Case Else	GetUserLevelCSSClass = "sSeed"		''2016모바일에서 변경 Seed->sSeed
	End Select
End Function 

'// 로그인 레벨에 따른 색상 //
Function GetLoginUserColor()
    dim uselevel
    uselevel = request.cookies("uinfo")("muserlevel")
    
    Select Case Cstr(uselevel)
        Case "1"
            ''그린
            GetLoginUserColor = "#f0ca2c"
        Case "2"
            ''블루
            GetLoginUserColor = "#a3cf6c"
        Case "3"
            ''VIP
            GetLoginUserColor = "#6ca54e"
        Case "4"
            ''오렌지
            GetLoginUserColor = "#f68d3f"
        Case "5"
            ''옐로우
            GetLoginUserColor = "#865e25"
        Case "6"
            ''Staff
            GetLoginUserColor = "#B70606"
        Case Else
			GetLoginUserColor = "#f0ca2c"
	End Select
End Function


''// 장바구니 갯수 :
Function GetCartCount()
    dim tmp
    GetCartCount = 0
    
    tmp = request.cookies("etc")("cartCnt")

    if (Not IsNumeric(tmp)) then Exit function
    
    if tmp<1 then tmp = 0
    
    GetCartCount = tmp
End Function

'// 상품 이미지 경로를 계산하여 반환 //
function GetImageSubFolderByItemid(byval iitemid)
    if (iitemid <> "") then
	    GetImageSubFolderByItemid = Num2Str(CStr(Clng(iitemid) \ 10000),2,"0","R")
	else
	    GetImageSubFolderByItemid = ""
	end if
end function


'// DB저장된 구문을 사이트에 쓸 수 있도록 변환 //
function db2html(checkvalue)
	dim v
	v = checkvalue
	if Isnull(v) then Exit function

    On Error resume Next
    v = replace(v, "&amp;", "&")
    ''v = replace(v, "&lt;", "<")
    ''v = replace(v, "&gt;", ">")
    v = replace(v, "&quot;", "'")
    v = Replace(v, "", "<br>")
    v = Replace(v, "\0x5C", "\")
    v = Replace(v, "\0x22", "'")
    v = Replace(v, "\0x25", "'")
    v = Replace(v, "\0x27", "%")
    v = Replace(v, "\0x2F", "/")
    v = Replace(v, "\0x5F", "_")

    db2html = v
end function


'// 사이트에서 입력받은 내용을 DB에 저장할 수 있도록 변환 //
function html2db(checkvalue)
	dim v
	v = checkvalue
	if Isnull(v) then Exit function
	v = Replace(v, "'", "''")
	html2db = v
end function


'// 문자열내 CR/LF를 <BR>태그로 치환 //
function nl2br(v)
	if IsNull(v) then
		nl2br = ""
		Exit function
	end if

    nl2br = Replace(v, vbcrlf,"<br>")
end function

'// 문자열내 CR/LF를 </li><li>태그로 치환 //
function nl2li(v,g)
	if IsNull(v) then
		nl2li = ""
		Exit function
	end if
	
	if g = "a" then
    	nl2li = Replace(v, vbcrlf,"</li><li>")
    else
    	nl2li = Replace(v, "<BR>","</li><li>")
    end if
end function

'//문자열내 특수문자 제거
function ReplaceRequestSpecialChar(v)
	ReplaceRequestSpecialChar = replace(v,"'","")
	ReplaceRequestSpecialChar = replace(ReplaceRequestSpecialChar,"--","")
end function

'//문자열내 특수문자및 쿼리문자 제거
function ReplaceRequest(v)
	ReplaceRequest = replace(v,"'","")
	ReplaceRequest = replace(ReplaceRequest,"--","")
	ReplaceRequest = replace(ReplaceRequest,"select","")
	ReplaceRequest = replace(ReplaceRequest,"delete","")
	ReplaceRequest = replace(ReplaceRequest,"update","")
	ReplaceRequest = replace(ReplaceRequest,"union","")
	ReplaceRequest = replace(ReplaceRequest,"drop","")
end function

'// 날짜를 지정된 문자형으로 변환 //
function FormatDate(ddate, formatstring)
	dim s
	Select Case formatstring
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
		Case Else
			s = CStr(ddate)
	End Select

	FormatDate = s
end function


'// 해당월 날수 반환 //
function DayOfMonth(yymmdd)
        dim s

        s = CStr(year(yymmdd)) + "-" + CStr(month(yymmdd))

        if (isDate(s + "-31") = true) then
                DayOfMonth = 31
        elseif (isDate(s + "-30") = true) then
                DayOfMonth = 30
        elseif (isDate(s + "-29") = true) then
                DayOfMonth = 29
        else
                DayOfMonth = 28
        end if
end function


'// 해당월 주수 반환 //
function WeekOfMonth(yymmdd)
        dim buf

        buf = CStr(year(yymmdd)) + "-" + CStr(month(yymmdd))
        WeekOfMonth = 5
        if ((weekday(buf + "-01") = 1) and (isDate(buf + "-29") = false)) then
                WeekOfMonth = 4
        else
                if (isDate(buf + "-31") = false) then
                        if (weekday(buf + "-01") > weekday(buf + "-30")) then
                                WeekOfMonth = 6
                        end if
                else
                        if (weekday(buf + "-01") > weekday(buf + "-31")) then
                                WeekOfMonth = 6
                        end if
                end if
        end if
end function

'// 지정날짜가 속한 주의 첫날 반환 //
function StartDayOfWeek(yymmdd)
        StartDayOfWeek = dateadd("d", CDate(yymmdd), 1 - weekday(CDate(yymmdd)))
end function

'// 지정날짜가 속한 주의 마지막날 반환 //
function EndDayOfWeek(yymmdd)
        EndDayOfWeek = dateadd("d", CDate(yymmdd), 7 - weekday(CDate(yymmdd)))
end function

'// 날짜 선택상자 출력 - 플라워 지정일에만 쓰임 //
Sub DrawOneDateBox(byval yyyy,mm,dd,tt)
	dim buf,i

	buf = "<select name='yyyy' class='input_02'>"
    for i=Year(date()-1) to Year(date()+1)
		if (CStr(i)=CStr(yyyy)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + ">" + CStr(i) + "</option>"
		end if
	next
    buf = buf + "</select>년 "

    buf = buf + "<select name='mm' class='input_02'>"
    for i=1 to 12
		if (Num2Str(i,2,"0","R")=Num2Str(mm,2,"0","R")) then
			buf = buf + "<option value='" + Num2Str(i,2,"0","R") +"' selected>" + Num2Str(i,2,"0","R") + "</option>"
		else
    	    buf = buf + "<option value='" + Num2Str(i,2,"0","R") +"'>" + Num2Str(i,2,"0","R") + "</option>"
		end if
	next

    buf = buf + "</select>월 "

    buf = buf + "<select name='dd' class='input_02'>"
    for i=1 to 31
		if (Num2Str(i,2,"0","R")=Num2Str(dd,2,"0","R")) then
	    buf = buf + "<option value='" + Num2Str(i,2,"0","R") +"' selected>" + Num2Str(i,2,"0","R") + "</option>"
		else
        buf = buf + "<option value='" + Num2Str(i,2,"0","R") + "'>" + Num2Str(i,2,"0","R") + "</option>"
		end if
    next
    buf = buf + "</select>일 "


    buf = buf & "<select name='tt' class='input_02'>"
    for i=9 to 18
		if (Num2Str(i,2,"0","R")=Num2Str(tt,2,"0","R")) then
        buf = buf & "<option value='" & CStr(i) & "' selected>" & CStr(i) & "~" & CStr(i + 2) & "</option>"
		else
        buf = buf & "<option value='" & CStr(i) & "'>" & CStr(i) & "~" & CStr(i + 2) & "</option>"
		end if
    next
    buf = buf & "</select>시 "

    response.write buf
end Sub


'// 숫자를 지정한 길이의 문자열로 반환 //
Function Num2Str(inum,olen,cChr,oalign)
	dim i, ilen, strChr
    
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


'// 숫자열을 화폐형으로 변환 //
function CurrFormat(byVal v)
	CurrFormat = FormatNumber(FormatCurrency(v),0)
end function


'// 제품번호를 문자열로 변환 //
function FormatCode(itemcode)
	FormatCode = Num2Str(itemcode,6,"0","R")
end function


'// 제품이미지 반환 //
function GetListImageUrl(byval itemid)
	GetListImageUrl = "/image/list/L" + Num2Str(itemid,9,"0","R") + ".jpg"
end function

'// 숫자 형식을 000NNNN 형식으로 변환  //
function Format00(totallength,orgData)
    Format00 = ""
    
    if IsNULL(orgData) then Exit Function
    
    Format00 = Num2Str(orgData,totallength,"0","R")
end function

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


'// HTML태그 제거 //
function stripHTML(strng)
   Dim regEx
   Set regEx = New RegExp
   regEx.Pattern = "[<][^>]*[>]"
   regEx.IgnoreCase = True
   regEx.Global = True
   stripHTML = regEx.Replace(strng, " ")
   Set regEx = nothing
End Function


'// 외부파일 실행 함수 //
Sub executeFile(fnm)
	Dim fso 
	Set fso = Server.CreateObject("Scripting.FileSystemObject") 
	'지정한 파일이 존재할 때 실행
	If (fso.FileExists(Server.MapPath(fnm))) Then
		on Error resume Next
		Server.Execute(fnm)
		on Error goto 0
	end if
	Set fso = nothing
end Sub


'// 꺽은괄호 HTML코드로 치환 //
Function ReplaceBracket(strng)
	strng = Replace(strng,"<","&lt;")
	strng = Replace(strng,">","&gt;")
	ReplaceBracket = strng
end Function


'// 숫자여부 확인
Function fnChkNumeric(iValue)
	IF iValue <> "" THEN
		If Not IsNumeric(iValue) Then 
		 Call Alert_return("매개변수가 잘못되었습니다.") 
		 Exit Function
		END IF 
	END IF	
	fnChkNumeric = iValue
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

'// 아이디로 로그인 했는지 여부 //
Function IsUserLoginOK()
    IsUserLoginOK = (GetLoginUserID<>"")
End Function


'// 주문번호로 로그인 했는지 여부 //
Function IsGuestLoginOK()
    IsGuestLoginOK = (GetGuestLoginOrderserial<>"")
End Function


'// 로그인 아이디 - 암호화 필요 //
Function GetLoginUserID()
    GetLoginUserID = request.cookies("uinfo")("muserid")
End Function


'// 로그인 한 이름  //
Function GetLoginUserName()
    GetLoginUserName = request.cookies("uinfo")("musername")
End Function


'// 로그인 이메일 //
Function GetLoginUserEmail()
    GetLoginUserEmail = request.cookies("uinfo")("museremail")
End Function


'// 로그인 레벨 //
Function GetLoginUserLevel()
    dim uselevel
    uselevel = request.cookies("uinfo")("muserlevel")
    if (uselevel="") then
		GetLoginUserLevel = "5"
	else
		GetLoginUserLevel = uselevel
	end if
End Function

'// 로그인 회원구분 //
Function GetLoginUserDiv()
    dim userDiv
    userDiv = request.cookies("uinfo")("muserdiv")
    if (userDiv="") then
		GetLoginUserDiv = "01"
	else
		GetLoginUserDiv = userDiv
	end if
End Function

'// 로그인 실명확인여부 //
Function GetLoginRealNameCheck()
    dim RealNameCheck
    RealNameCheck = request.cookies("uinfo")("mrealnamecheck")
    if (RealNameCheck="") then
		GetLoginRealNameCheck = "N"
	else
		GetLoginRealNameCheck = RealNameCheck
	end if
End Function


''// 장바구니 갯수세팅  
Function SetCartCount(cartcount)
    dim tmp
    tmp = cartcount
    
    if (Not IsNumeric(tmp)) then Exit function
    if tmp<1 then tmp = 0
    
    response.Cookies("etc").domain = "thefingers.co.kr"
    response.Cookies("etc")("cartCnt") = tmp
End Function


''// 로그인 당시 쿠폰 + 상품 쿠폰 갯수 - 쿠폰 받았을때 /사용했을때 세팅 필요 :
Function GetLoginCouponCount()
    dim tmp
    GetLoginCouponCount = 0
    
    tmp = request.cookies("etc")("mcouponCnt")

    if (Not IsNumeric(tmp)) then Exit function
    
    if tmp<1 then tmp = 0
    
    GetLoginCouponCount = tmp
End Function


''// 로그인 당시 마일리지 - 변경시 세팅 필요/ Display에만 사용 :
Function GetLoginCurrentMileage()
    dim tmp
    GetLoginCurrentMileage = 0
    
    tmp = request.cookies("etc")("mcurrentmile")

    if (Not IsNumeric(tmp)) then Exit function
    
    if tmp<1 then tmp = 0
    
    GetLoginCurrentMileage = tmp
End Function

''// 로그인 당시 쿠폰 + 상품쿠폰세팅  
Function SetLoginCouponCount(couponcount)
    dim tmp
    tmp = couponcount
    
    if (Not IsNumeric(tmp)) then Exit function
    if tmp<1 then tmp = 0
    
    response.Cookies("uinfo").domain = "thefingers.co.kr"
    response.Cookies("uinfo")("mcouponCnt") = tmp
End Function


''// 로그인 당시 마일리지 세팅  
Function SetLoginCurrentMileage(currmileage)
    dim tmp
    tmp = currmileage
    
    if (Not IsNumeric(tmp)) then Exit function
    if tmp<1 then tmp = 0
    
    response.Cookies("uinfo").domain = "thefingers.co.kr"
    response.Cookies("uinfo")("mcurrentmile") = tmp
End Function

'// 로그인 아이콘 //
Function GetLoginUserICon()
    GetLoginUserICon = request.cookies("etc")("musericon")
End Function



'// 주문번호 로그인  //
Function GetGuestLoginOrderserial()
    GetGuestLoginOrderserial = session("userorderserial") 'request.cookies("guestinfo")("orderserial")
End Function


'// 가격 할인율 계산 //
Function GetPricePercent(Sprice,Oprice,pt)
	if Sprice="" or Oprice="" or isNull(Sprice) or isNull(Oprice) then Exit Function
	if Sprice < Oprice then
		GetPricePercent = FormatNumber(100-(Clng(Sprice)/Clng(Oprice)*100),pt) & "%"
	else
		GetPricePercent = FormatNumber(0,pt) & "%"
	end if
End Function


'// 값비교 On/Off 반환
Function GetImgSwitchOnOff(skey, tkey)
	if skey=tkey then
		GetImgSwitchOnOff = "on"
	else
		GetImgSwitchOnOff = "off"
	end if
End Function

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


'// CheckBox로 넘어온 값을 DB의 in문에서 사용할 수 있도록 변환 //
Function cvtDBChkBoxData(strArr,mrk)
	strArr = Trim(strArr)
	if strArr<>"" then
		if right(strArr,1)="," then
			strArr = Left(strArr,Len(strArr)-1)
		end if
		strArr = Replace(strArr,", ",",")
		if mrk="Y" then
			cvtDBChkBoxData = "'" & Replace(strArr,",","','") & "'"
		else
			cvtDBChkBoxData = strArr
		end if
	else
		cvtDBChkBoxData = ""
	end if
End Function



''//파일명을 메인메뉴 고유번호로 변환 
function getTopMenuId(pageName,param)
	Select Case pageName
		Case "lecturelist.asp", "lecturedetail.asp", "lecturegroup.asp"
			If inStr(param,"cate_large=10")>0 Then
				getTopMenuId = "10"
			ElseIf inStr(param,"cate_large=20")>0 Then
				getTopMenuId = "20"
			ElseIf inStr(param,"cate_large=30")>0 Then
				getTopMenuId = "30"
			ElseIf inStr(param,"cate_large=40")>0 Then
				getTopMenuId = "40"
			ElseIf inStr(param,"cate_large=50")>0 Then
				getTopMenuId = "50"
			ElseIf inStr(param,"cate_large=60")>0 Then
				getTopMenuId = "60"
			ElseIf inStr(param,"catecd1=10")>0 Then
				getTopMenuId = "110"
			ElseIf inStr(param,"catecd1=20")>0 Then
				getTopMenuId = "120"
			ElseIf inStr(param,"catecd3=20")>0 Then
				getTopMenuId = "220"
			Else
				getTopMenuId = "1"	'강좌 전체보기
			End If
		Case "artistroom.asp"
			getTopMenuId = "60"	'좋은강사

		Case ""
			getTopMenuId = "70"	'소문난 전시

		Case ""
			getTopMenuId = "80"	'생활 레시피

		Case Else
			getTopMenuId = ""
	End Select
end function


''// 무료배송 기준 금액

function getCommonFreeBeasongLimit()
    dim ulevel
    ulevel = CStr(GetLoginUserLevel())
    
    Select Case ulevel
	Case 5
		'오렌지 등급
		getCommonFreeBeasongLimit = 30000
	Case 0
		'옐로두 등급
		getCommonFreeBeasongLimit = 30000
	Case 1
		'그린 등급
		getCommonFreeBeasongLimit = 30000
	Case 2
		'블루 등급
		getCommonFreeBeasongLimit = 20000
	Case 3
		'VIP 등급 : 항상무료
		getCommonFreeBeasongLimit = 1
	Case 6
		'Friends 등급
		getCommonFreeBeasongLimit = 10000
	Case 7
		'Staff 등급 : 항상무료
		getCommonFreeBeasongLimit = 1
	Case 8
		'Family 등급
		getCommonFreeBeasongLimit = 10000
	Case Else
		'기타
		getCommonFreeBeasongLimit = 30000
    End Select
end function

''// 공사중일때 회사IP외에는 지정페이지로 이동
Sub Underconstruction()
	Dim conIp, arrIp, tmpIp
	conIp = Request.ServerVariables("REMOTE_ADDR")
	arrIp = split(conIp,".")
	tmpIp = Num2Str(arrIp(0),3,"0","R") & Num2Str(arrIp(1),3,"0","R") & Num2Str(arrIp(2),3,"0","R") & Num2Str(arrIp(3),3,"0","R")

	'//공사중
	if Not(tmpIp=>"115094163043" and tmpIp<="115094163045") and Not(tmpIp=>"061252133001" and tmpIp<="061252133127") and Not(tmpIp=>"061252143001" and tmpIp<="061252143127") then
		If Response.Buffer Then
			Response.Clear
			Response.Expires = 0
		End If

		Response.write "<html>"
		Response.write "<head><title>더핑거스 -서비스 점검중입니다</title></head>"
		Response.write "<body>"
		Response.write "<table width='100%' height='100%' cellpadding='0' cellspacing='0' border='0'>"
		Response.write "<tr>"
		Response.write "	<td align='center' valign='middle'><img src='http://www.thefingers.co.kr/fingersing.jpg'></td>"
		Response.write "</tr>"
		Response.write "</table>"
		Response.write "</body>"
		Response.write "</html>"
		response.End
	end if
End Sub

'// post형식의 데이타  스트링 형태로 변경
Function fnMakePostData()
	Dim strMethod			: strMethod			= Request.ServerVariables("REQUEST_METHOD")	' Form의 Method 정보
	
	'// 지역변수
	Dim strFormName
	Dim strPostData		: strPostData		= ""
	
	'// Post 형식일 경우 Form값을 String 형태로 취합한다.
	If Lcase(strMethod) = "post" Then
		For Each strFormName	 In Request.Form		
				strPostData = strPostData & strFormName & "=" & Request.Form(strFormName) & "&"			
		Next
	End If
	fnMakePostData =strPostData
End Function

'// get 스트링 형태로 넘어온 데이터를 post 형태로 변경
Sub sbPostDataToHtml(ByVal strPostData)
	If Trim(strPostData) = "" Then Exit Sub
	
	Dim arrTemp	: arrTemp = Split(strPostData, "&")
	Dim arrData	: arrData	= Null
	Dim intTemp
	
	If IsArray(arrTemp) Then
		For intTemp = 0 To Ubound(arrTemp) - 1	
			arrData = Split(arrTemp(intTemp), "=")
			%>
			<input type="hidden" name="<%= arrData(0)%>" value="<%= arrData(1)%>">
			<%
		Next
	End If	
End Sub


'// 사이트 출력용 회원ID 변환 함수(지정수만큼 지정한 문자로 바꿈)
Function printUserId(strID,lng,chr)
	dim le, te

	if GetLoginUserDiv()<>"01" then	'회원 구분이 일반회원이 아니라면 아이디 변환 안함(업체/직원 등 당첨자등 참고)
		printUserId = strID
		Exit Function
	else
		le = len(strID)
		if(le<lng) Then
			printUserId = String(lng, le)
			Exit Function
		end if

		te = left(strID,le-lng) & String(lng, chr)
		printUserId = te
	end if
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
	Dim regEx, match, matches
	
	SET regEx = New RegExp
	regEx.Pattern = patrn			' 패턴을 설정.
	regEx.IgnoreCase = True			' 대/소문자를 구분하지 않도록 .
	regEx.Global = True				' 전체 문자열을 검색하도록 설정.
	RepWord = regEx.Replace(str,repval)
End Function 

'// 패스워드 복잡성 검사 함수(웹용)
Function chkSimplePwdComplex(uid,pwd)
	dim msg, i, sT, sN
	msg = ""

	'비밀번호 길이 검사
	if len(pwd)<8 then
		msg = msg & "- 비밀번호는 최소 8자리이상으로 입력해주세요.\n"
	end if

	'아이디와 동일 또는 포함하고 있는가?
	''if instr(lcase(pwd),lcase(uid))>0 then
	''	msg = msg & "- 아이디와 동일하거나 아이디를 포함하고 있는 비밀번호입니다.\n"
	''end if
	if lcase(pwd)=lcase(uid) then
		msg = msg & "- 아이디와 동일한 비밀번호입니다.\n"
	end if

	'영문/숫자/특수문자 두가지 이상 조합
    dim aAlpha, aNumber, aSpecial, chkCnt
    chkCnt = 0
    aAlpha = "[a-zA-Z]"
    aNumber = "[0-9]"
    aSpecial = "[!|@|#|$|%|^|&|*|(|)|-|_|?]"

	if Not(chkWord(pwd,aAlpha)) then chkCnt = chkCnt+1
	if Not(chkWord(pwd,aNumber)) then chkCnt = chkCnt+1
	if Not(chkWord(pwd,aSpecial)) then chkCnt = chkCnt+1

	if chkCnt<2 then
		msg = msg & "- 패스워드는 영문/숫자/특수문자 중 두 가지 이상의 조합으로 입력해주세요.\n"
	end if

	'결과 반환
	chkSimplePwdComplex = msg
end Function

'//정규식 문자열 검사
Function chkWord(str,patrn)
    Dim regEx, match, matches

    SET regEx = New RegExp
    regEx.Pattern = patrn	' 패턴을 설정.
    regEx.IgnoreCase = True	' 대/소문자를 구분하지 않도록 .
    regEx.Global = True		' 전체 문자열을 검색하도록 설정.
    SET Matches = regEx.Execute(str)
	if 0 < Matches.count then
		chkWord= false
	Else
		chkWord= true
	end if

	'pattern0 = "[^가-힣]"  '한글만
	'pattern1 = "[^-0-9 ]"  '숫자만
	'pattern2 = "[^-a-zA-Z]"  '영어만
	'pattern3 = "[^-가-힣a-zA-Z0-9/ ]" '숫자와 영어 한글만
	'pattern4 = "<[^>]*>"   '태그만
	'pattern5 = "[^-a-zA-Z0-9/ ]"    '영어 숫자만
End Function

'//null 일때 대체값
Function NullFillWith(src , data )
	if isNULL(src) or src = "" then
		if Not isNull(data) or data = "" then
			NullFillWith = data
		 else
		 	NullFillWith = 0
		end if
	else
		If Not IsNumeric(src) then
			NullFillWith = Replace(Trim(src),"'","''")
		else
			NullFillWith = src
		End if		
	end if
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

Function CurrURL()
	CurrURL = Request.ServerVariables("PATH_INFO")
End Function

Function CurrURLQ()
	CurrURLQ = "http://" & Request.ServerVariables("Server_name") & Request.ServerVariables("PATH_INFO")
	If Request.ServerVariables("REQUEST_METHOD") = "POST" then
		CurrURLQ = Request.ServerVariables("PATH_INFO") & "?" & Request.Form
	 else
 		CurrURLQ = Request.ServerVariables("PATH_INFO") & "?" & Request.QueryString
	End if
End Function


'//실명확인 상세 에러메시지 반환
function getRealNameErrMsg(DCd)
	Select Case DCd
		Case "A"
			getRealNameErrMsg = "실명 확인"
		Case "B"
			getRealNameErrMsg = "성명 불일치\n\n실명확인이 실패하였습니다.\n입력하신 정보를 확인하시고 다시 시도해주세요."
		Case "C"
			getRealNameErrMsg = "명의도용 차단 신청중입니다.\n\n마이크레딧 명의보호관리 서비스에서\n명의도용 차단을 일시 해제 하신 후에 이용가능합니다."
		Case "D"
			getRealNameErrMsg = "주민등록 번호가 조합체계에 맞지 않습니다.\n\n입력하신 정보를 확인하시고 다시 시도해주세요."
		Case "E"
			getRealNameErrMsg = "일시적으로 통신장애가 발생했습니다.\n\n잠시 후에 다시 시도해주세요."
		Case "F"
			getRealNameErrMsg = "고객님의 성명이 두음법칙에 맞지 않게 입력되었습니다.\n(예: 류지선→유지선)\n\n입력하신 정보를 확인하시고 다시 시도해주세요."
		Case "Y"
			getRealNameErrMsg = "실명안심차단 대상자입니다.\n\n차단 해제화면에서 일시 해제 후 이용가능합니다."
		Case "G"
			getRealNameErrMsg = "주민등록 정보가 존재하지 않습니다.\n한국신용정보(1588-2486) 또는\nhttp://idcheck.co.kr/idcheck/sub3_02.jsp에서 개인정보를 등록해주세요."
		Case "H"
			getRealNameErrMsg = "실명확인 DB의 실명정보가 불완전한 상태입니다.\n한국신용정보(1588-2486) 또는\nhttp://idcheck.co.kr/idcheck/sub3_02.jsp에서 개인정보를 정정해주세요."
		Case Else
			getRealNameErrMsg = "실명확인을 할 수 없는 상태입니다.\n\n잠시 후에 다시 시도해주세요."
	End Select
end function





'''''''''''''''''''''''''''''' 20009 리뉴얼 추가함수 '''''''''''''''''''''''''''''''''''
' response.write 개행
Sub rw(ByVal str)
	response.write str & "<br>"
End Sub 

' response.write 개행 + dbget.close()	:	response.End
Sub rwe(ByVal str)
	response.write str & "<br>"
	dbget.close()	:	response.End 
End Sub 

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

Function getThisURL()
	getThisURL = Request.ServerVariables("URL")
End Function 

' 현재 페이지 URL + 모든 파라미터
Function getThisFullURL()
	Dim url
	url = Request.ServerVariables("URL")
	If Request.ServerVariables("QUERY_STRING") <> "" Then 
		url = url & "?" & Request.ServerVariables("QUERY_STRING")
	Else
		url = url & "?"
	End If 
	
	Dim objItem
	For Each objItem In Request.Form
		url = url & objItem & "=" 
		url = url & Request.Form(objItem) & "&" 
	Next

	getThisFullURL = url
End Function

' 페이징 함수 <%=fnPaging(페이지파라미터, 토탈레코드카운트, 현재페이지, 페이지사이즈, 블럭사이즈)
Function fnPaging(ByVal pageParam, ByVal iTotalCount, ByVal iCurrPage, ByVal iPageSize, ByVal iBlockSize)

	If iTotalCount = "" Then iTotalCount = 0
	Dim iTotalPage
	iTotalPage  = Int ( (iTotalCount - 1) / iPageSize ) + 1
	If iTotalCount = 0 Then	iTotalPage = 1

	Dim str, i, iStartPage
	Dim url, arr
	url = getThisFullURL()
	If InStr(url,pageParam) > 0 Then 
		arr = Split(url, pageParam&"=")
		If UBOUND(arr) > 0 Then
			If InStr(arr(1),"&") Then 
				url = arr(0) & Mid(arr(1),InStr(arr(1),"&")+1) & "&" & pageParam&"="
			Else 
				url = arr(0) & pageParam&"="
			End If 
		End If
	ElseIf InStr(url,"?") > 0 Then 
		url = url & "&" &  pageParam & "="
	Else
		url = url & "?" &  pageParam & "="
	End If 
	url = Replace(url,"?&","?")
	url = Replace(url,"&&","&")

	Dim imgPrev01, imgNext01, imgPrev02, imgNext02
	imgPrev01	= "<img src=""http://image.thefingers.co.kr/academy2012/common/paging_prev.gif"" alt=""이전으로 이동"" />"
	imgNext01	= "<img src=""http://image.thefingers.co.kr/academy2012/common/paging_next.gif"" alt=""다음으로 이동"" />"
	imgPrev02	= "<img src=""http://image.thefingers.co.kr/academy2012/common/paging_fist.gif"" alt=""맨처음으로 이동"" />"
	imgNext02	= "<img src=""http://image.thefingers.co.kr/academy2012/common/paging_end.gif"" alt=""맨 끝으로 이동"" />"

	' 시작페이지
	If (iCurrPage Mod iBlockSize) = 0 Then
		iStartPage = (iCurrPage - iBlockSize) + 1
	Else
		iStartPage = ((iCurrPage \ iBlockSize) * iBlockSize) + 1
	End If

	' 1 Page로 이동
	str = str & "<span><a href=""" & url & "1"">" & imgPrev02 & "</a></span>"

	' 이전 Block으로 이동
	If (iCurrPage / iBlockSize) > 1 Then
		str = str & "<span><a href=""" & url & "" & (iStartPage - iBlockSize) & """>" & imgPrev01 & "</a></span>"
	Else
		str = str & "<span><a>"& imgPrev01 &"</a></span>"
	End If

	' 페이지 Count 루프
	i = iStartPage

	Do While ((i < iStartPage + iBlockSize) And (i <= iTotalPage))
		If i > iStartPage Then str = str & " "
		If Int(i) = Int(iCurrPage) Then
			str = str & "<a href=""" & url & "" & i & """><strong>" & i & "</strong></a>"
		Else
			str = str & "<a href=""" & url & "" & i & """>" & i & "</a>"
		End If
		i = i + 1
	Loop
	
	' 다음 Block으로 이동

	If (iStartPage+iBlockSize) < iTotalPage+1 Then
			str = str & "<span><a href=""" & url & "" & i & """>" & imgNext01 & "</a></span>"
	Else
			str = str & "<span><a>"& imgNext01 &"</a></span>"
	End If

	' 마지막 Page로 이동
	str = str & "<span><a href=""" & url & "" & iTotalPage & """>" & imgNext02 & "</a></span>"

	fnPaging	= str

End Function

' 페이징 함수 <%=fnPaging2016(페이지파라미터, 토탈레코드카운트, 현재페이지, 페이지사이즈, 블럭사이즈)
Function fnPaging2016(ByVal pageParam, ByVal iTotalCount, ByVal iCurrPage, ByVal iPageSize, ByVal iBlockSize)

	If iTotalCount = "" Then iTotalCount = 0
	Dim iTotalPage
	iTotalPage  = Int ( (iTotalCount - 1) / iPageSize ) + 1
	If iTotalCount = 0 Then	iTotalPage = 1

	Dim str, i, iStartPage
	Dim url, arr
	url = getThisFullURL()
	If InStr(url,pageParam) > 0 Then 
		arr = Split(url, pageParam&"=")
		If UBOUND(arr) > 0 Then
			If InStr(arr(1),"&") Then 
				url = arr(0) & Mid(arr(1),InStr(arr(1),"&")+1) & "&" & pageParam&"="
			Else 
				url = arr(0) & pageParam&"="
			End If 
		End If
	ElseIf InStr(url,"?") > 0 Then 
		url = url & "&" &  pageParam & "="
	Else
		url = url & "?" &  pageParam & "="
	End If 
	url = Replace(url,"?&","?")
	url = Replace(url,"&&","&")

	' 시작페이지
	If (iCurrPage Mod iBlockSize) = 0 Then
		iStartPage = (iCurrPage - iBlockSize) + 1
	Else
		iStartPage = ((iCurrPage \ iBlockSize) * iBlockSize) + 1
	End If

	' 1 Page로 이동
	str = str & "<a href=""" & url & "1"" title=""처음 페이지"" class=""first arrow""><span>맨 처음 페이지로 이동</span></a>"

	' 이전 Block으로 이동
	If (iCurrPage / iBlockSize) > 1 Then
		str = str & "<a href=""" & url & "" & (iStartPage - iBlockSize) & """ title=""이전 페이지"" class=""prev arrow""><span>이전페이지로 이동</span></a>"
	Else
		str = str & "<a title=""이전 페이지"" class=""prev arrow""><span>이전페이지로 이동</span></a>"
	End If

	' 페이지 Count 루프
	i = iStartPage

	Do While ((i < iStartPage + iBlockSize) And (i <= iTotalPage))
		If i > iStartPage Then str = str & " "
		If Int(i) = Int(iCurrPage) Then
			str = str & "<a href=""" & url & "" & i & """ title="""& i &" 페이지"" class=""current""><span>"& i &"</span></a>"
		Else
			str = str & "<a href=""" & url & "" & i & """ title="""& i &" 페이지""><span>"& i &"</span></a>"
		End If
		i = i + 1
	Loop
	
	' 다음 Block으로 이동

	If (iStartPage+iBlockSize) < iTotalPage+1 Then
			str = str & "<a href=""" & url & "" & i & """ title=""다음 페이지"" class=""next arrow""><span>다음 페이지로 이동</span></a>"
	Else
			str = str & "<a title=""다음 페이지"" class=""next arrow""><span>다음 페이지로 이동</span></a>"
	End If

	' 마지막 Page로 이동
	str = str & "<a href=""" & url & "" & iTotalPage & """ title=""마지막 페이지"" class=""end arrow""><span>맨 마지막 페이지로 이동</span></a>"

	fnPaging2016	= str

End function

Function fnPagingSSL(ByVal pageParam, ByVal iTotalCount, ByVal iCurrPage, ByVal iPageSize, ByVal iBlockSize)

	If iTotalCount = "" Then iTotalCount = 0
	Dim iTotalPage
	iTotalPage  = Int ( (iTotalCount - 1) / iPageSize ) + 1
	If iTotalCount = 0 Then	iTotalPage = 1

	Dim str, i, iStartPage
	Dim url, arr
	url = getThisFullURL()
	If InStr(url,pageParam) > 0 Then 
		arr = Split(url, pageParam&"=")
		If UBOUND(arr) > 0 Then
			If InStr(arr(1),"&") Then 
				url = arr(0) & Mid(arr(1),InStr(arr(1),"&")+1) & "&" & pageParam&"="
			Else 
				url = arr(0) & pageParam&"="
			End If 
		End If
	ElseIf InStr(url,"?") > 0 Then 
		url = url & "&" &  pageParam & "="
	Else
		url = url & "?" &  pageParam & "="
	End If 
	url = Replace(url,"?&","?")

	Dim imgPrev01, imgNext01, imgPrev02, imgNext02
	imgPrev01	= "<img src=""/fiximage/web2009/common/btn_pageprev01.gif"" border=0 align=""absmiddle"">"
	imgNext01	= "<img src=""/fiximage/web2009/common/btn_pagenext01.gif"" border=0 align=""absmiddle"">"
	imgPrev02	= "<img src=""/fiximage/web2009/common/btn_pageprev02.gif"" border=0 align=""absmiddle"">"
	imgNext02	= "<img src=""/fiximage/web2009/common/btn_pagenext02.gif"" border=0 align=""absmiddle"">"

	' 시작페이지
	If (iCurrPage Mod iBlockSize) = 0 Then
		iStartPage = (iCurrPage - iBlockSize) + 1
	Else
		iStartPage = ((iCurrPage \ iBlockSize) * iBlockSize) + 1
	End If

	' 1 Page로 이동
	str = str & "<a href=""" & url & "1"">" & imgPrev02 & "</a>"
	str = str & "&nbsp; &nbsp;"

	' 이전 Block으로 이동
	If (iCurrPage / iBlockSize) > 1 Then
		str = str & "<a href=""" & url & "" & (iStartPage - iBlockSize) & """>" & imgPrev01 & "</a>"
	Else
		str = str & imgPrev01
	End If
	str = str & "&nbsp; &nbsp;"

	' 페이지 Count 루프
	i = iStartPage

	str = str & "<span class=""pagenum01"">"
	Do While ((i < iStartPage + iBlockSize) And (i <= iTotalPage))
		If i > iStartPage Then str = str & " "
		If Int(i) = Int(iCurrPage) Then
			str = str & "<strong>" & i & "</strong>"
		Else
			str = str & "<a href=""" & url & "" & i & """>" & i & "</a>"
		End If
		i = i + 1
	Loop
	str = str & "</span>"
	
	' 다음 Block으로 이동
	str = str & "&nbsp; &nbsp;"
	If (iStartPage+iBlockSize) < iTotalPage+1 Then
		str = str & "<a href=""" & url & "" & i & """>" & imgNext01 & "</a>"
	Else
		str = str & imgNext01
	End If

	' 마지막 Page로 이동
	str = str & "&nbsp; &nbsp;"
	str = str & "<a href=""" & url & "" & iTotalPage & """>" & imgNext02 & "</a>"

	fnPagingSSL	= str

End function

''EMail ComboBox
function DrawEamilBoxHTML(frmName,txBoxName, cbBoxName,emailVal)
    dim RetVal, i, isExists : isExists=false
    dim eArr : eArr = Array("naver.com","netian.com","paran.com","hanmail.net","dreamwiz.com","nate.com" _
                ,"empal.com","orgio.net","unitel.co.kr","chol.com","kornet.net","korea.com" _ 
                ,"freechal.com","hanafos.com","hitel.net","hanmir.com","hotmail.com")
	emailVal = LCase(emailVal)
	
    RetVal = "<input name='"&txBoxName&"' type='text' class='txtBasic tblInput' value='' style='width:120px;display:none;'/>&nbsp;"
    RetVal = RetVal & "<select name='"&cbBoxName&"' id='select3' class='select tblInput' style='width:120px;' onChange=""jsShowMailBox('"&frmName&"','"&cbBoxName&"','"&txBoxName&"');""\>"
    ''RetVal = RetVal & "<option value=''>메일선택</option>"
    for i=LBound(eArr) to UBound(eArr)
        if (eArr(i)=emailVal) then
            isExists = true
            RetVal = RetVal & "<option value='"&eArr(i)&"' selected>"&eArr(i)&"</option>"
        else
            RetVal = RetVal & "<option value='"&eArr(i)&"' >"&eArr(i)&"</option>"
        end if
    next
    
    if (Not isExists) and (emailVal<>"") then
        RetVal = RetVal & "<option value='"&emailVal&"' selected>"&emailVal&"</option>"
    end if
    RetVal = RetVal & "<option value='etc' >직접 입력</option>"
    RetVal = RetVal & "</select>"

    response.write RetVal
    
end Function

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

''쿠키 조작 검증이 필요한곳에서 기존 getLoginUserID 대신 사용 
function getEncLoginUserID()
    dim ret : ret=""
    dim planid : planid = getLoginUserID()
    dim encedID : encedID = request.cookies("uinfo")("shix")      ''암호화된쿠키값.
    getEncLoginUserID = ret
    
    if (planid="") then Exit function   '' 아이디 쿠키값없으면 로그인 안된것임
   
    ''if (HashTenID(planid)=encedID) then   ''해쉬된 값과 현재 아이디가 같으면 정상 아이디 리턴
    if (UCASE(HashTenID(planid))=UCASE(encedID)) then   ''해쉬된 값과 현재 아이디가 같으면 정상 아이디 리턴 UCASE 추가 2013.03.27
        getEncLoginUserID = planid
        Exit function
    end if
    
    'if (encedID="") then                '' 암호화된값이 없으면. 암호화전 운영인경우가 있으므로 일단 정상으로 판단. 차후 주석처리
    '    getEncLoginUserID = planid
    '    Exit function
    'end if
    
    ''다른경우 조작된 경우임.
    ''관리자에게 메세지발송
    On Error Resume Next
    call InfoMsgMailSend("planid="&planid&"<br>"&"encedID="&encedID)
    On Error Goto 0
    
    ''진행 계속 못함 (버퍼링 삭제 후 로그아웃!)
	If Response.Buffer Then
		Response.Clear
		Response.ContentType = "text/html"
		Response.Expires = 0
	End If
    response.write "<script>" & vbCrLf &_
    			   " alert('죄송합니다. 암호화 처리중 오류가 발생하였습니다. 다시 로그인후 이용해주세요.');" & vbCrLf &_
    			   " document.location = '/login/dologout.asp';" & vbCrLf &_
    			   "</script>"
    response.end
        
end function


''관리자에게 메세지 발송 (검증페이지에서 사용.)
function InfoMsgMailSend(paramMsg)
    dim strMsg, strMethod
    dim lngMaxFormBytes : lngMaxFormBytes =800
    strMsg = strMsg & "<li>서버:<br>"
	strMsg = strMsg & application("Svr_Info")
	strMsg = strMsg & "<br><br></li>"
	
	'// 접속자 브라우저 정보
	strMsg = strMsg & "<li>브라우저 종류:<br>"
	strMsg = strMsg & Server.HTMLEncode(Request.ServerVariables("HTTP_USER_AGENT"))
	strMsg = strMsg & "<br><br></li>"
	strMsg = strMsg & "<li>접속자 IP:<br>"
	strMsg = strMsg & Server.HTMLEncode(Request.ServerVariables("REMOTE_ADDR"))
	strMsg = strMsg & "<br><br></li>"
	strMsg = strMsg & "<li>경유페이지:<br>"
	strMsg = strMsg & request.ServerVariables("HTTP_REFERER")
	strMsg = strMsg & "<br><br></li>"
	'// 오류 페이지 정보
	strMsg = strMsg & "<li>페이지:<br>"
	strMethod = Request.ServerVariables("REQUEST_METHOD")
	strMsg = strMsg & "HOST : " & Request.ServerVariables("HTTP_HOST") & "<BR>"
	strMsg = strMsg & strMethod & " : "
	
	If strMethod = "POST" Then
		strMsg = strMsg & Request.TotalBytes & " bytes to "
	End If

	strMsg = strMsg & Request.ServerVariables("SCRIPT_NAME")
	strMsg = strMsg & "</li>"

	If strMethod = "POST" Then
		strMsg = strMsg & "<br><li>POST Data:<br>"

		'실행에 관련된 에러를 출력합니다.
		On Error Resume Next
		If Request.TotalBytes > lngMaxFormBytes Then
			strMsg = strMsg & Server.HTMLEncode(Left(Request.Form, lngMaxFormBytes)) & " . . ."'
		Else
			strMsg = strMsg & Server.HTMLEncode(Request.Form)
		End If
		On Error Goto 0
		strMsg = strMsg & "</li>"
	elseif strMethod = "GET" then
		strMsg = strMsg & "<br><li>GET Data:<br>"
		strMsg = strMsg & Request.QueryString
	End If
	strMsg = strMsg & "<br><br></li>"
	
    '### 시스템팀 구성원에게 오류 발생 내용 발송 ###
	dim cdoMessage,cdoConfig

	Set cdoConfig = CreateObject("CDO.Configuration")
	'-> 서버 접근방법을 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)
	'-> 서버 주소를 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="webadmin.10x10.co.kr"
	'-> 접근할 포트번호를 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	'-> 접속시도할 제한시간을 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
	'-> SMTP 접속 인증방법을 설정합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	'-> SMTP 서버에 인증할 ID를 입력합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "MailSendUser"
	'-> SMTP 서버에 인증할 암호를 입력합니다
	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wjddlswjddls"
	cdoConfig.Fields.Update

	Set cdoMessage = CreateObject("CDO.Message")
	Set cdoMessage.Configuration = cdoConfig

'	cdoMessage.To 		= "kobula@10x10.co.kr;tozzinet@10x10.co.kr;kjy8517@10x10.co.kr;errmail@10x10.co.kr;thensi7@10x10.co.kr;corpse2@10x10.co.kr;"
	cdoMessage.To 		= "errmail@10x10.co.kr"
	cdoMessage.From 	= "webserver@10x10.co.kr"
	cdoMessage.SubJect 	= "["&date()&"] theFingers페이지 메세지 발생"
	cdoMessage.HTMLBody	= strMsg & "<br><li>Message:<br>" & paramMsg &"</li>"
	
	cdoMessage.BodyPart.Charset="ks_c_5601-1987"         '/// 한글을 위해선 꼭 넣어 주어야 합니다.
    cdoMessage.HTMLBodyPart.Charset="ks_c_5601-1987"     '/// 한글을 위해선 꼭 넣어 주어야 합니다.

	cdoMessage.Send

	Set cdoMessage = nothing
	Set cdoConfig = nothing
end function

'// 자동로그인 확인(2011.12.09; 허진원 추가)
Sub chk_AutoLogin()
	if Not(IsUserLoginOK) and request.cookies("mSave")("SAVED_AUTO")<>"" then
		if tenDec(request.cookies("mSave")("SAVED_ID"))<>"" and tenDec(request.cookies("mSave")("SAVED_PW"))<>"" then
			on Error Resume Next
			'#HTTP통신으로 회원정보 확인
			dim objXML, xmlDOM, unm
			Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.Open "POST", www1Url & "/login/actLoginData.asp", false
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objXML.Send("userid=" & server.URLEncode(request.cookies("mSave")("SAVED_ID")) & "&userpass=" & server.URLEncode(request.cookies("mSave")("SAVED_PW")) & "&device=" & flgDevice)
			If objXML.Status = "200" Then
				'//전달받은 내용 확인
				'response.write BinaryToText(objXML.ResponseBody, "UTF-8")
				'response.End

				'XML을 담을 DOM 객체 생성
				Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 UTF-8로 변환(한글문제))
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "UTF-8")

				unm = xmlDOM.getElementsByTagName("username").item(0).text
				if Not(unm="" or isNull(unm)) then
					response.Cookies("uinfo").domain = "thefingers.co.kr"
					response.Cookies("uinfo")("muserid") = tenDec(request.cookies("mSave")("SAVED_ID"))
					response.Cookies("uinfo")("musername") = xmlDOM.getElementsByTagName("username").item(0).text
					response.Cookies("uinfo")("museremail") = xmlDOM.getElementsByTagName("useremail").item(0).text
					response.Cookies("uinfo")("muserdiv") = xmlDOM.getElementsByTagName("userdiv").item(0).text
					response.cookies("uinfo")("muserlevel") = xmlDOM.getElementsByTagName("userlevel").item(0).text
					response.cookies("uinfo")("mrealnamecheck") = xmlDOM.getElementsByTagName("realchk").item(0).text
					response.Cookies("uinfo")("misupche") = xmlDOM.getElementsByTagName("isupche").item(0).text
					response.cookies("uinfo")("shix") = xmlDOM.getElementsByTagName("shix").item(0).text        '''201212 추가.

					response.Cookies("etc").domain = "thefingers.co.kr"
					response.cookies("etc")("mcouponCnt") = xmlDOM.getElementsByTagName("coupon").item(0).text
					response.cookies("etc")("mcurrentmile") = xmlDOM.getElementsByTagName("mileage").item(0).text
					'response.cookies("etc")("currtencash") = xmlDOM.getElementsByTagName("currtencash").item(0).text
					'response.cookies("etc")("currtengiftcard") = xmlDOM.getElementsByTagName("currtengiftcard").item(0).text
					response.cookies("etc")("cartCnt") = xmlDOM.getElementsByTagName("cartCnt").item(0).text
					'response.Cookies("etc")("ordCnt") = xmlDOM.getElementsByTagName("ordCnt").item(0).text
					'response.Cookies("etc")("musericonNo") = xmlDOM.getElementsByTagName("usericonNo").item(0).text
					response.Cookies("etc")("logindate") = now()
					response.Cookies("etc")("ConfirmUser") = xmlDOM.getElementsByTagName("ConfirmUser").item(0).text
					
					''## 보안강화 세션 처리 2016/11/15
                    session("ssnuserid")  = LCase(Trim(request.Cookies("uinfo")("muserid")))
                    session("ssnlogindt") = Year(now())&Right("00"&Month(now()),2)&Right("00"&Day(now()),2)&Right("00"&Hour(now()),2)&Right("00"&Minute(now()),2)&Right("00"&Second(now()),2)
    
				else
					response.Cookies("mSave").domain = "thefingers.co.kr"
					response.cookies("mSave") = ""
					response.Cookies("mSave").Expires = Date - 1
				end if

				Set xmlDOM = Nothing
			else
				response.Cookies("mSave").domain = "thefingers.co.kr"
				response.cookies("mSave") = ""
				response.Cookies("mSave").Expires = Date - 1
			end if

			Set objXML= Nothing

			on Error Goto 0
		end if
	end if
end Sub

'// 로그인 유효기간 확인(2015.07.07; 허진원 추가)
Sub chk_ValidLogin()
	dim lgDt : lgDt = LEFT(request.Cookies("etc")("logindate"),10) ''left 추가 2015/07/16
	dim isChk : isChk=false

	if lgDt<>"" and IsUserLoginOK then
		if isDate(lgDt) then
			if datediff("m",lgDt,now)=0 then
				isChk = true
			end if
		else
			isChk = true
		end if
	end if

	// 로그아웃 처리
	if Not(isChk) and IsUserLoginOK then
		response.Cookies("uinfo").domain = "thefingers.co.kr"
		response.Cookies("uinfo") = ""
		response.Cookies("uinfo").Expires = Date - 1
		
		response.Cookies("etc").domain = "thefingers.co.kr"
		response.Cookies("etc") = ""
		response.Cookies("etc").Expires = Date - 1
		
		response.Cookies("mybadge").domain = "thefingers.co.kr"
		response.Cookies("mybadge") = ""
		response.Cookies("mybadge").Expires = Date - 1
	end if
end Sub

'// 무통장 입금 텐바이텐 계좌 //
Sub DrawTenBankAccount(accountnoName, accountno)
    dim buf
    buf = "<select name='" & accountnoName & "' id='bank'>"
    buf = buf & "<option value='국민 470301-01-014754' " & ChkIIF(accountno="국민 470301-01-014754","selected","") & " >국민은행 470301-01-014754</option>"
    buf = buf & "<option value='신한 100-016-523130' " & ChkIIF(accountno="신한 100-016-523130","selected","") & " >신한은행 100-016-523130</option>"
    buf = buf & "<option value='우리 092-275495-13-001' " & ChkIIF(accountno="우리 092-275495-13-001","selected","") & " >우리은행 092-275495-13-001</option>"
    buf = buf & "<option value='하나 146-910009-28804' " & ChkIIF(accountno="하나 146-910009-28804","selected","") & " >하나은행 146-910009-28804</option>"
    buf = buf & "<option value='기업 277-028182-01-046' " & ChkIIF(accountno="기업 277-028182-01-046","selected","") & " >기업은행 277-028182-01-046</option>"
    buf = buf & "<option value='농협 029-01-246118' " & ChkIIF(accountno="농협 029-01-246118","selected","") & " >농 협 029-01-246118</option>"
    buf = buf & "</select>"
    
    response.write buf
end Sub

'// 은행 목록 //
Sub DrawBankCombo(selectedname,selectedId)
    dim buf
	
	buf = "<select name='" & selectedname & "' id='bank'>"
	buf = buf + "<option value='' " & chkIIF(selectedId="","selected","") & " ></option>"
	buf = buf + "<option value='경남'" & chkIIF(selectedId="경남","selected","") & " >경남</option>"
	buf = buf + "<option value='광주'" & chkIIF(selectedId="광주","selected","") & " >광주</option>"
	buf = buf + "<option value='국민'" & chkIIF(selectedId="국민","selected","") & " >국민</option>"
	buf = buf + "<option value='기업'" & chkIIF(selectedId="기업","selected","") & " >기업</option>"
	buf = buf + "<option value='농협'" & chkIIF(selectedId="농협","selected","") & " >농협</option>"
	buf = buf + "<option value='단위농협'" & chkIIF(selectedId="단위농협","selected","") & " >단위농협</option>"
	buf = buf + "<option value='대구'" & chkIIF(selectedId="대구","selected","") & " >대구</option>"
	buf = buf + "<option value='도이치'" & chkIIF(selectedId="도이치","selected","") & " >도이치</option>"
	buf = buf + "<option value='부산'" & chkIIF(selectedId="부산","selected","") & " >부산</option>"
	buf = buf + "<option value='산업'" & chkIIF(selectedId="산업","selected","") & " >산업</option>"
	buf = buf + "<option value='새마을금고'" & chkIIF(selectedId="새마을금고","selected","") & " >새마을금고</option>"
	buf = buf + "<option value='수협'" & chkIIF(selectedId="수협","selected","") & " >수협</option>"
	buf = buf + "<option value='신한'" & chkIIF(selectedId="신한","selected","") & " >신한</option>"
	buf = buf + "<option value='외환'" & chkIIF(selectedId="외환","selected","") & " >외환</option>"
	buf = buf + "<option value='우리'" & chkIIF(selectedId="우리","selected","") & " >우리</option>"
	buf = buf + "<option value='우체국'" & chkIIF(selectedId="우체국","selected","") & " >우체국</option>"
	buf = buf + "<option value='전북'" & chkIIF(selectedId="전북","selected","") & " >전북</option>"
	buf = buf + "<option value='제일'" & chkIIF(selectedId="제일","selected","") & " >제일</option>"
	buf = buf + "<option value='조흥'" & chkIIF(selectedId="조흥","selected","") & " >조흥</option>"
	buf = buf + "<option value='평화'" & chkIIF(selectedId="평화","selected","") & " >평화</option>"
	buf = buf + "<option value='하나'" & chkIIF(selectedId="하나","selected","") & " >하나</option>"
	buf = buf + "<option value='시티'" & chkIIF(selectedId="시티","selected","") & " >시티</option>"
	buf = buf + "<option value='홍콩샹하이'" & chkIIF(selectedId="홍콩샹하이","selected","") & " >홍콩샹하이</option>"
	buf = buf + "<option value='ABN암로은행'" & chkIIF(selectedId="ABN암로은행","selected","") & " >ABN암로은행</option>"
	buf = buf + "<option value='UFJ은행'" & chkIIF(selectedId="UFJ은행","selected","") & " >UFJ은행</option>"
	buf = buf + "<option value='신협'" & chkIIF(selectedId="신협","selected","") & " >신협</option>"
	buf = buf + "</select>"
	
	response.write buf
end Sub

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

Function GetPolderName(pDept)
	On Error Resume Next
	Dim vScriptUrl		'/소스 경로저장 변수
	Dim vIndex2			'/ 2번째 슬래시 위치
	Dim vIndex3			'/ 3번째 슬래시 위치
	Dim vIndex4			'/ 4번째 슬래시 위치
	
	vScriptUrl = Request.ServerVariables("SCRIPT_NAME")
	vIndex2 = InStr(2, vScriptUrl, "/")

	Select Case pDept
		Case 2
			vIndex3 = InStr(vIndex2+1, vScriptUrl, "/")
			GetPolderName = Mid(vScriptUrl, vIndex2+1, vIndex3-vIndex2-1)
		Case 3
			vIndex3 = InStr(vIndex2+1, vScriptUrl, "/")
			vIndex4 = InStr(vIndex3+1, vScriptUrl, "/")
			GetPolderName = Mid(vScriptUrl, vIndex3+1, vIndex4-vIndex3-1)
		Case Else
			GetPolderName = Mid(vScriptUrl, 2, vIndex2-2)
	End Select
	On Error Goto 0
End Function

'' 2015/07/15 쿠키 검사 require MD5.asp
function TenOrderSerialHash(iorderserial)
    TenOrderSerialHash = LEFT(MD5(iorderserial&"ten"&iorderserial),20)
end Function

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

'// ston캐시 서버 썸네일 제작(퀄러티 함께)
function getStonReSizeImg(furl,wd,ht,qua)
    if (qua<>"100") and (qua<>"") then
        getStonReSizeImg = furl&"/10x10/resize/"&wd&"x"&ht&"/quality/"&qua&"/"
    else
        getStonReSizeImg = furl&"/10x10/resize/"&wd&"x"&ht&"/"
    end if
end function

'// ston캐시 서버 썸네일 제작(기존 포토서버 썸네일 변경) - 리스트 위주
function getStonThumbImgURL(furl,wd,ht,fit,ws)
    getStonThumbImgURL = furl&"/10x10/thumbnail/"&wd&"x"&ht&"/quality/80/"
end Function

Function fnBackPathURLChange(url)
	url = Replace(url,"/","%2F")
	url = Replace(url,".","%2E")
	url = Replace(url,"?","%3F")
	url = Replace(url,"=","%3D")
	url = Replace(url,"&","%26")
	fnBackPathURLChange = url
End Function

'작가/강사 체크
Function fnlecturerCheck(loginuserid)
	Dim i, torf, SQL
	torf = False
	i = 0

	SQL = "select count(*) as cnt"
	SQL = SQL & " from [db_academy].[dbo].tbl_corner_good "
	SQL = SQL & " where lecturer_id='" + CStr(loginuserid) + "'"

	rsget.CursorLocation = adUseClient
	rsget.Open SQL, dbget, adOpenForwardOnly, adLockReadOnly

	if NOT(rsget.EOF or rsget.BOF) then
		If rsget("cnt") > 0 Then
			torf = TRUE
		End If
	end if

	rsget.Close
	fnlecturerCheck = torf
End Function
%>
