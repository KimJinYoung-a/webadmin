<%
'+---------------------------------------------------------------------------------------------------------+
'|                                   문 자 열   공 통   함 수 선 언                                        |
'+------------------------------------------+--------------------------------------------------------------+
'|                함 수 명                  |                          기    능                            |
'+------------------------------------------+--------------------------------------------------------------+
'| Num2Str(inum,olen,cChr,oalign)           | 숫자를 지정한 길이의 문자열로 변환한다.                      |
'|                                          | 사용예 : Num2Str(425,4,"0","R") → 0425                       |
'+------------------------------------------+--------------------------------------------------------------+
'| SplitValue(orgStr,delim,pos)             | 문자열을 잘라 원하는 위치의 값을 반환                        |
'|                                          | 사용예 : SplitValue("A/B/C","/","2") → B                     |
'+------------------------------------------+--------------------------------------------------------------+
'| CurrFormat(byVal v)                      | 숫자열을 화폐형으로 변환                                     |
'|                                          | 사용예 : CurrFormat(1250) → 1,250                            |
'+------------------------------------------+--------------------------------------------------------------+
'| chrbyte(str,chrlen,dot)                  | 지정길이로 문자열 자르기                                     |
'|                                          | 사용예 : chrbyte("안녕하세요",3,"Y") → 안녕...               |
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
'| CheckCurse(txt)		                    | 문자열의 형식을 정규식으로 검사                              |
'|                                          | 사용예 :  CheckCurse(txt) -> 글자수 만큼 * 모양으로 치환	   |
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
'| getWeekName(yymmdd)                      | 입력된 날짜의 요일명 반환                                    |
'|                                          | 사용예 : weekName = getWeekName("2015-03-12") > "목"         |
'+------------------------------------------+--------------------------------------------------------------+
'| DrawOneDateBox(yyyy,mm,dd,tt)            | 날짜 선택 셀렉트박스 출력 (년원일시)                         |
'|                                          | 사용예 : call DrawOneDateBox("2006","08","10","15")          |
'+------------------------------------------+--------------------------------------------------------------+
'| DrawDateTimeBox(byval fyy,vyy,tsy,tey,fmm,vmm,fdd,vdd,fhh,vhh,fnn,vnn,fss,vss)                          |
'| 날짜시간 셀렉트박스 출력 (년원일시분초) 사용예 : call DrawDateTimeBox("yyyy",2006,1,1,"mm",8,"dd",10,"","","","","","") |
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
'| nl2br(v)                                 | 문자열내 CR/LF를 <br />태그로 치환                             |
'|                                          | 사용예 : Response.Write nl2br(Rs("contents"))                |
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
'| DrawTenBankAccount(accountnoName, accountno) |    무통장 입금 계좌                                      |
'|                                          | 사용예 : Call DrawTenBankAccount("accountno",accountno)      |
'+------------------------------------------+--------------------------------------------------------------+
'| DrawBankCombo(selectedname,selectedId)   | 은행 목록                                                    |
'|                                          | 사용예 : Call DrawBankCombo("bankname",bankname)             |
'+------------------------------------------+--------------------------------------------------------------+
'| getSpecialShopItemPrice(sellcash)        | 우수 회원샵 상품 할인 판매가                                 |
'|                                          | 사용예 : discountprice = getSpecialShopItemPrice(sellcah)    |
'+------------------------------------------+--------------------------------------------------------------+
'| getAjaxSiteResult(mtd,strURL,para)       | 지정한 URL의 실행결과값 가져오기                             |
'|                                          | 사용예 : Res = getAjaxSiteResult("GET","http://www.10x10.co.kr/index.asp","div=A") |
'+------------------------------------------+--------------------------------------------------------------+
'| checkFilePath(filePath)       | 파일이 해당경로에 있는지 체크한다.										|
'|                                          | 사용예 : bool = checkFIlePath(filePath)								 |
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
'| IsVIPUser()                         		| [회원등급]으로 VIP 인지 여부 return Boolean				   |
'|                                          | 사용예 : bool = IsVIPUser()                             	   |
'+------------------------------------------+--------------------------------------------------------------+
'| IsSpecialShopUser()                      | [회원등급]으로 우수회원샵 이용가능 여부 return Boolean	   |
'|                                          | 사용예 : bool = IsVIPUser()                             	   |
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
'| GetLoginCurrentTenCash()                 | 로그인 당시 예치금   - 예치금 변경시 세팅 필요               |
'|                                          | 사용예 : ret = GetLoginCurrentTenCash()                      |
'+------------------------------------------+--------------------------------------------------------------+
'| GetLoginCurrentTenGiftCard()             | 로그인 당시 기프트카드잔액   - 잔액 변경시 세팅 필요         |
'|                                          | 사용예 : ret = GetLoginCurrentTenGiftCard()                  |
'+------------------------------------------+--------------------------------------------------------------+
'| SetLoginCouponCount(couponcount)         | 로그인 당시 할인권 + 상품쿠푠 갯수 세팅                      |
'|                                          | 사용예 : call SetLoginCouponCount(couponcount)               |
'+------------------------------------------+--------------------------------------------------------------+
'| SetLoginCurrentMileage(currmileage)      | 로그인 당시 마일리지 세팅                                    |
'|                                          | 사용예 : call SetLoginCurrentMileage(currmileage)            |
'+------------------------------------------+--------------------------------------------------------------+
'| SetLoginCurrentTenCash(currmileage)      | 로그인 당시 예치금 세팅                                      |
'|                                          | 사용예 : call SetLoginCurrentTenCash(currtencash)            |
'+------------------------------------------+--------------------------------------------------------------+
'| SetLoginCurrentTenGiftCard(currmileage)  | 로그인 당시 기프트카드잔액 세팅                              |
'|                                          | 사용예 : call SetLoginCurrentTenGiftCard(currtengiftcardmileage)|
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
'|                                          | 사용예 : response.Cookies("tinfo")("shix") = HashTenID(userid)|
'+------------------------------------------+--------------------------------------------------------------+
'| getEncLoginUserID()          		    | 암호화된 아이디 검증 및 로그인된 아이디 가져옴 **md5.asp 필요|
'|                                          | 사용예 : userid = getEncLoginUserID()                        |
'+------------------------------------------+--------------------------------------------------------------+
'| chkSimplePwdComplex(uid,pwd)             | 패스워드 복잡성 검사 (웹사이트용)                            |
'|                                          | 사용예 : chkSimplePwdComplex('userid','password00')          |
'+------------------------------------------+--------------------------------------------------------------+
'| GetUserProfileImg(inum)                  | 유저 아이디별 프로필 이미지를 긁어옴                         |
'|                                          | 사용예 : Response.write GetUserProfileImg("1")               |
'+------------------------------------------+--------------------------------------------------------------+

'+---------------------------------------------------------------------------------------------------------+
'|                              2009/2010/2013 리뉴얼 추가 함수                                            |
'+------------------------------------------+--------------------------------------------------------------+
'| rw(), rwe()                              | response.write 축약, rwe는 dbget.close() : response.End 포함 |
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
'| fnPagingSSL()                            | 2009 SSL용 페이징 함수, 페이지값을 넘기는 파라미터명에 유의  |
'| 사용예 : <%=fnPagingSSL(페이지파라미터, 토탈레코드카운트, 현재페이지, 페이지사이즈, 블럭사이즈)%        |
'+------------------------------------------+--------------------------------------------------------------+
'| ArraySwapRows(ary,row1,row2)             | 배열의 행 치환                                               |
'|                                          | 사용예 : ArraySwapRows(배열명,바꿀열,대상열)                 |
'+------------------------------------------+--------------------------------------------------------------+
'| ArrayQuickSort(vec,loBound,hiBound,SortField) | 배열값을 내림차순으로 정렬                              |
'|                                          | 사용예 : ArrayQuickSort(배열명,최소수,배열최대수,기준열번호) |
'+------------------------------------------+--------------------------------------------------------------+
'| BinaryToText(BinaryData, CharSet)         | 바이너리 데이터 TEXT형태로 변환                             |
'|                                           | 사용예 : BinaryToText(objXML.ResponseBody, "euc-kr")        |
'+-------------------------------------------+-------------------------------------------------------------+
'| chkArrValue(aVal,cVal)                    | 콤마로 구분된 배열값에 지정된 값이 있는지 반환              |
'|                                           | 사용예 : chkArrValue("A,B,C", "B") → true                   |
'+-------------------------------------------+-------------------------------------------------------------+
'| chkArrValueLen(aVal,cVal,lmt)             | 콤마로 구분된 배열값에 지정된 값(지정길이)이 있는지 반환    |
'|                                           | 사용예 : chkArrValueLen("AA,BB,CC","B",1) → true            |
'+-------------------------------------------+-------------------------------------------------------------+
'| getThumbImgFromURL(furl,wd,ht,fit,ws)     | 포토서버 썸네일 제작(기존 파일명)                           |
'|                                           | 사용예 : getThumbImgFromURL(이미지파일경로,넓이,높이,"true","false") |
'+-------------------------------------------+-------------------------------------------------------------+

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
		Case "5"
			GetUserLevelStr = "ORANGE"
		Case "0"
			GetUserLevelStr = "YELLOW"
		Case "1"
			GetUserLevelStr = "GREEN"
		Case "2"
			GetUserLevelStr = "BLUE"
		Case "3"
			GetUserLevelStr = "VIP SILVER"
		Case "4"
			GetUserLevelStr = "VIP GOLD"
		Case "6"
			GetUserLevelStr = "VVIP"
		Case "7"
			GetUserLevelStr = "STAFF"
		Case "8"
			GetUserLevelStr = "FAMILY"
		Case "9"
			GetUserLevelStr = "BIZ"
		Case Else
			GetUserLevelStr = "ORANGE"
	end Select
end function

'// 사용자 등급의 해당명칭의 CSS 클래스를 반환 //
Function GetUserLevelCSSClass()
	Select Case CStr(request.Cookies("tinfo")("userlevel"))
		Case "5"	GetUserLevelCSSClass = "memORANGE" ''"orange"
		Case "0"	GetUserLevelCSSClass = "memYELLOW" ''"yellow"
		Case "1"	GetUserLevelCSSClass = "memGREEN"  ''"green"
		Case "2"	GetUserLevelCSSClass = "memBLUE"   ''"blue"
		Case "3"	GetUserLevelCSSClass = "memSILVER" ''"vipsilver"
		Case "4"	GetUserLevelCSSClass = "memGOLD" ''"vipgold"
		Case "6"	GetUserLevelCSSClass = "memVVIP" ''"vvip"
		Case "7"	GetUserLevelCSSClass = "memSTAFF" ''"staff"
		Case "8"	GetUserLevelCSSClass = "memSTAFF" ''"black"
		Case "9"	GetUserLevelCSSClass = "memSTAFF" ''"BIZ"
		Case Else	GetUserLevelCSSClass = "memORANGE" ''"orange"
	End Select
End Function


'// 로그인 레벨에 따른 색상 //
Function GetLoginUserColor()
    dim uselevel
    uselevel = request.Cookies("tinfo")("userlevel")

    Select Case Cstr(uselevel)
        Case "5"
            ''오렌지
            GetLoginUserColor = "#F2901B"
        Case "0"
            ''옐로우
            GetLoginUserColor = "#F2B705"
        Case "1"
            ''그린
            GetLoginUserColor = "#68A63C"
        Case "2"
            ''블루
            GetLoginUserColor = "#026DAF"
        Case "3"
            ''VIP
			GetLoginUserColor = "#999999"
        Case "4"
            ''VIP GOLD
            GetLoginUserColor = "#c9af20"
        Case "6"
            ''VVIP
            GetLoginUserColor = "#B70606"
        Case "7"
            ''Staff
            GetLoginUserColor = "#000000"
        Case "8"
            ''FAMILY
            GetLoginUserColor = "#000000"
        Case "9"
            ''BIZ
            GetLoginUserColor = "#000000"
        Case Else
			GetLoginUserColor = "#000000"
	End Select

End Function

'// 상품 이미지 경로를 계산하여 반환 //
function GetImageSubFolderByItemid(byval iitemid)
    if (iitemid <> "") then
	    GetImageSubFolderByItemid = Num2Str(CStr(Clng(iitemid) \ 10000),2,"0","R")
	else
	    GetImageSubFolderByItemid = ""
	end if
end function

'// HTML요소용 특수문자/기호 치환(웹표준)
function str2html(v)
	if Isnull(v) then Exit function

	v = replace(v, "&", "&amp;")
	v = replace(v, "<", "&lt;")
	v = replace(v, ">", "&gt;")
	v = replace(v, """", "&quot;")
	v = replace(v, "'", "&acute;")

	str2html = v
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
    v = Replace(v, "", "<br />")
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


'// 문자열내 CR/LF를 <br />태그로 치환 //
function nl2br(v)
	if IsNull(v) then
		nl2br = ""
		Exit function
	end if

    v = Replace(v, vbcrlf,"<br />")
    v = Replace(v, vbCr,"<br />")
    v = Replace(v, vbLf,"<br />")
    nl2br = v
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
		Case "0000.00.00-00:00"
			s = CStr(year(ddate)) & "." &_
				Num2Str(month(ddate),2,"0","R") & "." &_
				Num2Str(day(ddate),2,"0","R") & " " &_
				Num2Str(hour(ddate),2,"0","R") & ":" &_
				Num2Str(minute(ddate),2,"0","R")
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
		Case "00:00" '시분
			s = Num2Str(hour(ddate),2,"0","R") & ":" &_
				Num2Str(minute(ddate),2,"0","R")
		Case "M/D" '월일
			s = Num2Str(month(ddate),2,"0","R") & "월 " &_
				Num2Str(day(ddate),2,"0","R") & "일"
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

'// 지정날짜의 요일 반환 //
function getWeekName(yymmdd)
	dim ArrWkNm: ArrWkNm = Split("일,월,화,수,목,금,토",",")
	getWeekName = ArrWkNm(weekDay(CDate(yymmdd))-1)
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
		if (Format00(2,i)=Format00(2,mm)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"'>" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>월 "

    buf = buf + "<select name='dd' class='input_02'>"
    for i=1 to 31
		if (Format00(2,i)=Format00(2,dd)) then
	    buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
        buf = buf + "<option value='" + Format00(2,i) + "'>" + Format00(2,i) + "</option>"
		end if
    next
    buf = buf + "</select>일 "


    buf = buf & "<select name='tt' class='input_02'>"
    for i=9 to 18
		if (Format00(2,i)=Format00(2,tt)) then
        buf = buf & "<option value='" & CStr(i) & "' selected>" & CStr(i) & "~" & CStr(i + 2) & "</option>"
		else
        buf = buf & "<option value='" & CStr(i) & "'>" & CStr(i) & "~" & CStr(i + 2) & "</option>"
		end if
    next
    buf = buf & "</select>시 "

    response.write buf
end Sub


'// 날짜시간 선택상자 출력 - 공용 //
Sub DrawDateTimeBox(byval fyy,vyy,tsy,tey,fmm,vmm,fdd,vdd,fhh,vhh,fnn,vnn,fss,vss)
	dim buf,i
	buf = ""

	if fyy<>"" then
		buf = buf & "<select name='" & fyy & "' class='input_default'>"
	    for i=Year(date())-tsy to Year(date())+tey
			if (CStr(i)=CStr(vyy)) then
				buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
			else
	    		buf = buf + "<option value=" + CStr(i) + ">" + CStr(i) + "</option>"
			end if
		next
	    buf = buf + "</select>년 "
	end if

	if fmm<>"" then
		buf = buf + "<select name='" & fmm & "' class='input_default'>"
		for i=1 to 12
			if (Format00(2,i)=Format00(2,vmm)) then
				buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
			else
				buf = buf + "<option value='" + Format00(2,i) +"'>" + Format00(2,i) + "</option>"
			end if
		next
		buf = buf + "</select>월 "
	end if

	if fdd<>"" then
		buf = buf + "<select name='" & fdd & "' class='input_default'>"
		for i=1 to 31
			if (Format00(2,i)=Format00(2,vdd)) then
				buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
			else
				buf = buf + "<option value='" + Format00(2,i) +"'>" + Format00(2,i) + "</option>"
			end if
		next
		buf = buf + "</select>일 "
	end if

	if fhh<>"" then
		buf = buf + "<select name='" & fhh & "' class='input_default'>"
		for i=0 to 23
			if (Format00(2,i)=Format00(2,vhh)) then
				buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
			else
				buf = buf + "<option value='" + Format00(2,i) +"'>" + Format00(2,i) + "</option>"
			end if
		next
		buf = buf + "</select>시 "
	end if

	if fnn<>"" then
		buf = buf + "<select name='" & fnn & "' class='input_default'>"
		for i=0 to 59
			if (Format00(2,i)=Format00(2,vnn)) then
				buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
			else
				buf = buf + "<option value='" + Format00(2,i) +"'>" + Format00(2,i) + "</option>"
			end if
		next
		buf = buf + "</select>분 "
	end if

	if fss<>"" then
		buf = buf + "<select name='" & fss & "' class='input_default'>"
		for i=0 to 59
			if (Format00(2,i)=Format00(2,vss)) then
				buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
			else
				buf = buf + "<option value='" + Format00(2,i) +"'>" + Format00(2,i) + "</option>"
			end if
		next
		buf = buf + "</select>초 "
	end if

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


'// 숫자 형식을 000NNNN 형식으로 변환  //
function Format00(totallength,orgData)
    Format00 = ""

    if IsNULL(orgData) then Exit Function

    Format00 = Num2Str(orgData,totallength,"0","R")
end function


'// 제품번호를 문자열로 변환 //
function FormatCode(itemcode)
	FormatCode = Num2Str(itemcode,6,"0","R")
end function


'// 제품이미지 반환 //
function GetListImageUrl(byval itemid)
	GetListImageUrl = "/image/list/L" + Num2Str(itemid,9,"0","R") + ".jpg"
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


'// HTML코드 꺽은괄호로 치환 //
Function ReverseBracket(strng)
	strng = Replace(strng,"&lt;","<")
	strng = Replace(strng,"&gt;",">")
	ReverseBracket = strng
end Function


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
	if isNull(szSource) then  exit Function

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


'// VIP 인지 여부 //
Function IsVIPUser()
	If GetLoginUserLevel() = "3" OR GetLoginUserLevel() = "4" OR GetLoginUserLevel() = "6" Then
    	IsVIPUser = True
    Else
		IsVIPUser = False
	End If
End Function

'// VVIP 인지 여부 //
Function IsVVIPUser()
	If GetLoginUserLevel() = "6" Then
    	IsVVIPUser = True
    Else
		IsVVIPUser = False
	End If
End Function

'// 로그인 아이디 - 암호화 필요 //
Function GetLoginUserID()
    GetLoginUserID = request.cookies("partner")("userid")
End Function


'// 로그인 한 이름  //
Function GetLoginUserName()
    GetLoginUserName = response.Cookies("partner")("ssBctCname")
End Function


'// 로그인 이메일 //
Function GetLoginUserEmail()
    GetLoginUserEmail = response.Cookies("partner")("ssBctEmail")
End Function


'// 로그인 레벨 //
'Function GetLoginUserLevel()
'    dim uselevel
'    uselevel = request.Cookies("tinfo")("userlevel")
'    if (uselevel="") then
'		GetLoginUserLevel = "5"
'	else
'		GetLoginUserLevel = uselevel
'	end if
'End Function

'// 로그인 회원구분 //
Function GetLoginUserDiv()
    dim userDiv
    userDiv = response.Cookies("partner")("userdiv")
    if (userDiv="") then
		GetLoginUserDiv = "14"  ''확인.
	else
		GetLoginUserDiv = userDiv
	end if
End Function

'// 로그인 실명확인여부 //
'Function GetLoginRealNameCheck()
'    dim RealNameCheck
'    RealNameCheck = request.Cookies("tinfo")("realnamecheck")
'    if (RealNameCheck="") then
'		GetLoginRealNameCheck = "N"
'	else
'		GetLoginRealNameCheck = RealNameCheck
'	end if
'End Function

'// 로그인 레벨에 따른 색상 //
Function GetLoginUserColor()
    dim uselevel
    uselevel = request.Cookies("tinfo")("userlevel")

    Select Case Cstr(uselevel)
        Case "5"
            ''오렌지
            GetLoginUserColor = "#F2901B"
        Case "0"
            ''옐로우
            GetLoginUserColor = "#F2B705"
        Case "1"
            ''그린
            GetLoginUserColor = "#68A63C"
        Case "2"
            ''블루
            GetLoginUserColor = "#026DAF"
        Case "3"
            ''VIP SILVER
            GetLoginUserColor = "#999999"
        Case "4"
            ''VIP GOLD
            GetLoginUserColor = "#c9af20"
        Case "6"
            ''VVIP
            GetLoginUserColor = "#d50c0c"
        Case "7"
            ''Staff
            GetLoginUserColor = "#000000"
        Case "8"
            ''FAMILY
            GetLoginUserColor = "#000000"
        Case "9"
            ''BIZ
            GetLoginUserColor = "#000000"
        Case Else
			GetLoginUserColor = "#000000"
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


''// 장바구니 갯수세팅
Function SetCartCount(cartcount)
    dim tmp
    tmp = cartcount

    if (Not IsNumeric(tmp)) then Exit function
    if tmp<1 then tmp = 0

    response.Cookies("etc").domain = "10x10.co.kr"
    response.Cookies("etc")("cartCnt") = tmp
End Function


''// 로그인 당시 쿠폰 + 상품 쿠폰 갯수 - 쿠폰 받았을때 /사용했을때 세팅 필요 :
Function GetLoginCouponCount()
    dim tmp
    GetLoginCouponCount = 0

    tmp = request.cookies("etc")("couponCnt")

    if (Not IsNumeric(tmp)) then Exit function

    if tmp<1 then tmp = 0

    GetLoginCouponCount = tmp
End Function


''// 로그인 당시 마일리지 - 변경시 세팅 필요/ Display에만 사용 :
Function GetLoginCurrentMileage()
    dim tmp
    GetLoginCurrentMileage = 0

    tmp = request.cookies("etc")("currentmile")

    if (Not IsNumeric(tmp)) then Exit function

    if tmp<1 then tmp = 0

    GetLoginCurrentMileage = tmp
End Function

''// 로그인 당시 예치금 - 변경시 세팅 필요/ Display에만 사용 :
Function GetLoginCurrentTenCash()
    dim tmp
    GetLoginCurrentTenCash = 0

    tmp = request.cookies("etc")("currtencash")

    if (Not IsNumeric(tmp)) then Exit function

    if tmp<1 then tmp = 0

    GetLoginCurrentTenCash = tmp
End Function

''// 로그인 당시 기프트카드잔액 - 변경시 세팅 필요/ Display에만 사용 :
Function GetLoginCurrentTenGiftCard()
    dim tmp
    GetLoginCurrentTenGiftCard = 0

    tmp = request.cookies("etc")("currtengiftcard")

    if (Not IsNumeric(tmp)) then Exit function

    if tmp<1 then tmp = 0

    GetLoginCurrentTenGiftCard = tmp
End Function

''// 로그인 당시 쿠폰 + 상품쿠폰세팅
Function SetLoginCouponCount(couponcount)
    dim tmp
    tmp = couponcount

    if (Not IsNumeric(tmp)) then Exit function
    if tmp<1 then tmp = 0

    response.Cookies("etc").domain = "10x10.co.kr"
    response.Cookies("etc")("couponCnt") = tmp
End Function


''// 로그인 당시 마일리지 세팅
Function SetLoginCurrentMileage(currmileage)
    dim tmp
    tmp = currmileage

    if (Not IsNumeric(tmp)) then Exit function
    if tmp<1 then tmp = 0

    response.Cookies("etc").domain = "10x10.co.kr"
    response.Cookies("etc")("currentmile") = tmp
End Function

''// 로그인 당시 예치금 세팅
Function SetLoginCurrentTenCash(currtencash)
    dim tmp
    tmp = currtencash

    if (Not IsNumeric(tmp)) then Exit function
    if tmp<1 then tmp = 0

    response.Cookies("etc").domain = "10x10.co.kr"
    response.Cookies("etc")("currtencash") = tmp
End Function

''// 로그인 당시 기프트카드잔액 세팅
Function SetLoginCurrentTenGiftCard(currtengiftcard)
    dim tmp
    tmp = currtengiftcard

    if (Not IsNumeric(tmp)) then Exit function
    if tmp<1 then tmp = 0

    response.Cookies("etc").domain = "10x10.co.kr"
    response.Cookies("etc")("currtengiftcard") = tmp
End Function

'// 로그인 아이콘 //
Function GetLoginUserICon()
    GetLoginUserICon = request.cookies("etc")("usericon")
End Function



'// 주문번호 로그인  //
Function GetGuestLoginOrderserial()
    GetGuestLoginOrderserial = session("userorderserial") 'request.cookies("guestinfo")("orderserial")
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

''프로필 이미지 번호가 0일경우 사용.
function getDefaultProfileImgNo(vuid)
    dim t : t = LEFT(vuid,1)
    if (t="") then
        getDefaultProfileImgNo = "1"
        exit function
    end if
    getDefaultProfileImgNo = cStr((ASC(t) mod 30)+1)
end function

'// 유저 프로필 이미지
Function GetUserProfileImg(inum,vuid)
	Dim Rvinum
	Rvinum = getNumeric(inum)

	If Rvinum="0" or Rvinum="" Then
	    Rvinum = getDefaultProfileImgNo(vuid)
	end if

	GetUserProfileImg = "http://fiximage.10x10.co.kr/web2015/common/img_profile_"& Num2Str(Rvinum,2,"0","R") &".png"
End Function

'// 로그인 아이콘(번호로 변경) //
Function GetLoginUserICon()
    GetLoginUserICon = request.cookies("etc")("usericonNo")
End Function

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

'// 가격 할인율 계산 //
Function GetPricePercent(Sprice,Oprice,pt)
	if Sprice="" or Oprice="" or isNull(Sprice) or isNull(Oprice) then Exit Function
	if Sprice < Oprice then
		''GetPricePercent = FormatNumber(100-(Clng(Sprice)/Clng(Oprice)*100),pt) & "%"		''반올림 문제로 내림으로 변경(20150914; 허진원)
		GetPricePercent = FormatNumber(Clng((Oprice-Sprice)/Oprice*100),pt) & "%"
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
			"self.location.replace('" & targetURL & "');" & vbCrLf &_
			"</script>"
	Response.Write strTemp
End Sub

'// 무통장 입금 텐바이텐 계좌 //
Sub DrawTenBankAccount(accountnoName, accountno)
    dim buf
    buf = "<select name='" & accountnoName & "' class='input_02' style='width:200;height:18;font-size:11px;'>"
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

	buf = "<select name='" & selectedname & "' >"
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
	buf = buf + "<option value='시티'" & chkIIF(selectedId="시티","selected","") & " >시티</option>"
	buf = buf + "<option value='신한'" & chkIIF(selectedId="신한","selected","") & " >신한</option>"
	buf = buf + "<option value='신협'" & chkIIF(selectedId="신협","selected","") & " >신협</option>"
	buf = buf + "<option value='우리'" & chkIIF(selectedId="우리","selected","") & " >우리</option>"
	buf = buf + "<option value='우체국'" & chkIIF(selectedId="우체국","selected","") & " >우체국</option>"
	buf = buf + "<option value='전북'" & chkIIF(selectedId="전북","selected","") & " >전북</option>"
	buf = buf + "<option value='제일'" & chkIIF(selectedId="제일","selected","") & " >제일</option>"
	buf = buf + "<option value='제주'" & chkIIF(selectedId="제주","selected","") & " >제주</option>"
	buf = buf + "<option value='KEB하나'" & chkIIF(selectedId="KEB하나","selected","") & " >KEB하나</option>"
	buf = buf + "<option value='현대스위스상호저축은행'" & chkIIF(selectedId="현대스위스상호저축은행","selected","") & " >현대스위스상호저축은행</option>"
	buf = buf + "<option value='홍콩샹하이'" & chkIIF(selectedId="홍콩샹하이","selected","") & " >홍콩샹하이</option>"
	buf = buf + "<option value='ABN암로은행'" & chkIIF(selectedId="ABN암로은행","selected","") & " >ABN암로은행</option>"
	buf = buf + "<option value='UFJ은행'" & chkIIF(selectedId="UFJ은행","selected","") & " >UFJ은행</option>"
	buf = buf + "</select>"

	response.write buf
end Sub



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

Function IsSpecialShopUser()
	If GetLoginUserLevel() = "2" OR GetLoginUserLevel() = "3" OR GetLoginUserLevel() = "4" OR GetLoginUserLevel() = "6" OR GetLoginUserLevel() = "7" OR GetLoginUserLevel() = "8" Then
    	IsSpecialShopUser = True
    Else
		IsSpecialShopUser = False
	End If
End Function

Function getSpecialShopPercent()
	SELECT Case GetLoginUserLevel()
		case "2" : getSpecialShopPercent = "15"
		case "3" : getSpecialShopPercent = "20"
		case "4" : getSpecialShopPercent = "25"
		case "6" : getSpecialShopPercent = "25"
		case "7" : getSpecialShopPercent = "25"
		case "8" : getSpecialShopPercent = "20"
	End SELECT
End Function

''//우수회원샵 상품 가격
function getSpecialShopItemPrice(orgprice)
    dim rPrice, ulevel
    rPrice = orgprice
    ulevel = CStr(GetLoginUserLevel())

    Select Case ulevel
		Case "2"
			'' BLUE 15%
			rPrice = CLng(orgprice*0.85)
		Case "3"
			'' VIP silver 20%
			rPrice = CLng(orgprice*0.8)
		Case "4"
			'' VIP gold 25%
			rPrice = CLng(orgprice*0.75)
		Case "6"
			'' VVIP 25%
			rPrice = CLng(orgprice*0.75)
		Case "7"
			'' STAFF 30%->25%로변경 20150112
			rPrice = CLng(orgprice*0.75)
		Case "8"
			'' FAMILY 20%
			rPrice = CLng(orgprice*0.8)
		Case "9"
			'' BIZ 0%
			rPrice = CLng(orgprice*0.0)
	end Select

	getSpecialShopItemPrice = rPrice
end function

'// 지정URL 결과값 접수
function getAjaxSiteResult(mtd,strURL,para)
	dim oXML, strRes
	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open mtd, strURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.send para	'파라메터 전송/실행

	'// 한글 처리
	Dim responseStrm
	Set responseStrm = CreateObject("ADODB.Stream")

	responseStrm.Open
	responseStrm.Position = 0
	responseStrm.Type = 1
	responseStrm.Write oXML.responseBody
	responseStrm.Position = 0
	responseStrm.Type = 2
	responseStrm.Charset = "euc-kr"
	strRes = responseStrm.ReadText
	getAjaxSiteResult=strRes  			'결과 반환
	responseStrm.close
	Set responseStrm = Nothing

	Set oXML = Nothing	'컨퍼넌트 해제
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
		'VIP Silver 등급
		getCommonFreeBeasongLimit = 10000
	Case 4
		'VIP GOLD 등급 : 항상무료
		getCommonFreeBeasongLimit = 1
	Case 6
		'VVIP 등급 : 항상무료
		getCommonFreeBeasongLimit = 1
	Case 7
		'Staff 등급 : 항상무료
		getCommonFreeBeasongLimit = 1
	Case 8
		'Family 등급
		getCommonFreeBeasongLimit = 10000
	Case 9
		'BIZ 등급
		getCommonFreeBeasongLimit = 30000
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
	if Not(tmpIp=>"115094163043" and tmpIp<="115094163045") and Not(tmpIp=>"061252133001" and tmpIp<="061252133127") and Not(tmpIp=>"061252143070" and tmpIp<="061252143072") and Not(tmpIp=>"192168001001" and tmpIp<="192168001256") and tmpIp<>"211206236117" then
		If Response.Buffer Then
			Response.Clear
			Response.Expires = 0
		End If
		'// 2014년 리뉴얼 준비중
		Response.write "<html>"
		Response.write "<head><title>텐바이텐-서비스 점검중입니다</title></head>"
		Response.write "<body>"
		Response.write "<table width='100%' height='100%' cellpadding='0' cellspacing='0' border='0'>"
		Response.write "<tr>"
		''Response.write "	<td align='center' valign='middle'><img src='http://fiximage.10x10.co.kr/web2013/common/open_ready_2014.jpg' width='1140' height='910' border='0' alt='coming soon'></td>"
		Response.write "	<td align='center' valign='middle'><img src='http://fiximage.10x10.co.kr/web2015/common/2015_10x10_open_ready_PC.jpg' width='1104' border='0' ></td>"
		Response.write "</tr>"
		Response.write "</table>"
		Response.write "</body>"
		Response.write "</html>"
		response.End
	end if
End Sub

'/서버 주기적 업데이트 위한 공사중 처리 '2011.11.11 한용민 생성
'/리뉴얼시 이전해 주시고 지우지 말아 주세요
Sub serverupdate_underconstruction()
	dim isServerDown : isServerDown = false
		'isServerDown = true	' 서버다운
		isServerDown = false	' 서버활성화

	if Not(isServerDown) then exit Sub

	Response.write "<html>"
	Response.write "<head><title>서비스 점검중입니다</title></head>"
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
	''if GetLoginUserDiv()<>"01" then	'회원 구분이 일반회원이 아니라면 아이디 변환 안함(업체/직원 등 당첨자등 참고-2015.09.02;허진원 제거)
	if GetLoginUserLevel="7" then	'회원 등급이 STAFF라면 아이디 변환 안함(직원 당첨자 등 참고)
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
	CurrURLQ = "http://" & Request.ServerVariables("Server_name") & CurrURL()
	If Request.ServerVariables("REQUEST_METHOD") = "POST" then
		CurrURLQ = CurrURL & "?" & Request.Form
	 else
 		CurrURLQ = CurrURL & "?" & Request.QueryString
	End if
End Function

Function RefURLQ()
	dim strRef
	strRef = Split(Replace(request.ServerVariables("HTTP_REFERER"),"http://",""),"/")
	if ubound(strRef)>0 then
		RefURLQ = Replace(Replace(request.ServerVariables("HTTP_REFERER"),"http://",""),strRef(0),"")
	else
		RefURLQ = ""
	end if
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
	response.write str & "<br />"
End Sub

' response.write 개행 + dbget.close()	:	response.End
Sub rwe(ByVal str)
	response.write str & "<br />"
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

	Dim imgPrev01, imgNext01, imgPrev02, imgNext02
	imgPrev01	= "<img src=""http://fiximage.10x10.co.kr/web2009/common/btn_pageprev01.gif"" border=0 align=""absmiddle"" style='display:inline'>"
	imgNext01	= "<img src=""http://fiximage.10x10.co.kr/web2009/common/btn_pagenext01.gif"" border=0 align=""absmiddle"" style='display:inline'>"
	imgPrev02	= "<img src=""http://fiximage.10x10.co.kr/web2009/common/btn_pageprev02.gif"" border=0 align=""absmiddle"" style='display:inline'>"
	imgNext02	= "<img src=""http://fiximage.10x10.co.kr/web2009/common/btn_pagenext02.gif"" border=0 align=""absmiddle"" style='display:inline'>"

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

	fnPaging	= str

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
	imgPrev01	= "<img src=""/fiximage/web2009/common/btn_pageprev01.gif"" border=0 align=""absmiddle"" style='display:inline'>"
	imgNext01	= "<img src=""/fiximage/web2009/common/btn_pagenext01.gif"" border=0 align=""absmiddle"" style='display:inline'>"
	imgPrev02	= "<img src=""/fiximage/web2009/common/btn_pageprev02.gif"" border=0 align=""absmiddle"" style='display:inline'>"
	imgNext02	= "<img src=""/fiximage/web2009/common/btn_pagenext02.gif"" border=0 align=""absmiddle"" style='display:inline'>"

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
    dim eArr : eArr = Array("naver.com", "daum.net", "hanmail.net", "gmail.com", "nate.com", "empal.com")
	'' ,                  "netian.com","paran.com","hanmail.net","dreamwiz.com","nate.com" _
    ''             ,"empal.com","orgio.net","unitel.co.kr","chol.com","kornet.net","korea.com" _
    ''             ,"freechal.com","hanafos.com","hitel.net","hanmir.com","hotmail.com")
	emailVal = LCase(emailVal)

    RetVal = "<input name='"&txBoxName&"' type='text' title='이메일 직접 입력' class='txtInp' style='width:118px;display:none;' value=''/>"
    RetVal = RetVal & "<select name='"&cbBoxName&"' id='select3' title='이메일 서비스 선택' class='select offInput emailSelect' style='width:102px;' onChange=""jsShowMailBox('"&frmName&"','"&cbBoxName&"','"&txBoxName&"');""\>"
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

'// 배열 치환
Sub ArraySwapRows(ary,row1,row2)
  Dim x,tempvar
  For x = 0 to Ubound(ary,2)
    tempvar = ary(row1,x)
    ary(row1,x) = ary(row2,x)
    ary(row2,x) = tempvar
  Next
End Sub  'SwapRows

'// 배열 정렬
Sub ArrayQuickSort(vec,loBound,hiBound,SortField)
  Dim pivot(),loSwap,hiSwap,temp,counter
  Redim pivot (Ubound(vec,2))

  if hiBound - loBound = 1 then
    if vec(loBound,SortField) > vec(hiBound,SortField) then Call ArraySwapRows(vec,hiBound,loBound)
  End If

  For counter = 0 to Ubound(vec,2)
    pivot(counter) = vec(int((loBound + hiBound) / 2),counter)
    vec(int((loBound + hiBound) / 2),counter) = vec(loBound,counter)
    vec(loBound,counter) = pivot(counter)
  Next

  loSwap = loBound + 1
  hiSwap = hiBound

  do
    while loSwap < hiSwap and vec(loSwap,SortField) <= pivot(SortField)
      loSwap = loSwap + 1
    wend

    while vec(hiSwap,SortField) > pivot(SortField)
      hiSwap = hiSwap - 1
    wend

    if loSwap < hiSwap then Call ArraySwapRows(vec,loSwap,hiSwap)

  loop while loSwap < hiSwap

  For counter = 0 to Ubound(vec,2)
    vec(loBound,counter) = vec(hiSwap,counter)
    vec(hiSwap,counter) = pivot(counter)
  Next

  if loBound < (hiSwap - 1) then Call ArrayQuickSort(vec,loBound,hiSwap-1,SortField)
  if hiSwap + 1 < hibound then Call ArrayQuickSort(vec,hiSwap+1,hiBound,SortField)

End Sub

'// 정규식 패턴지정 문자열 처리/반환
Function RepWord(str,patrn,repval)
	Dim regEx

	SET regEx = New RegExp
	regEx.Pattern = patrn			' 패턴을 설정.
	regEx.IgnoreCase = True			' 대/소문자를 구분하지 않도록 .
	regEx.Global = True				' 전체 문자열을 검색하도록 설정.
	RepWord = regEx.Replace(str,repval)
End Function

'// 정규식 패턴지정된 결과 반환
Function RegExp(ByVal pattern, ByVal strText)
     Dim objRegExp
     Set objRegExp = new RegExp
     objRegExp.Pattern = pattern
     objRegExp.Global = True
     objRegExp.IgnoreCase = True

     Dim match, Matches
     Dim rst

     Set Matches = objRegExp.Execute(strText)
     If Matches.Count = 0 Then
          RegExp = ""
     Else
		For Each match In Matches
		   rst = match
		   exit for	'첫번째 찾은것만 반환후 종료
		Next
		RegExp = rst
     End If
     Set objRegExp = Nothing
End Function

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
    dim encedID : encedID = request.Cookies("tinfo")("shix")      ''암호화된쿠키값.
    getEncLoginUserID = ret

    if (planid="") then Exit function   '' 아이디 쿠키값없으면 로그인 안된것임

    if (UCASE(HashTenID(planid))=UCASE(encedID)) then   ''해쉬된 값과 현재 아이디가 같으면 정상 아이디 리턴 UCASE 추가 2012.12.02
        getEncLoginUserID = planid
        Exit function
    end if

    ''if (encedID="") then                '' 암호화된값이 없으면. 암호화전 운영인경우가 있으므로 일단 정상으로 판단. 차후 주석처리 2013/10/02 주석
    ''    getEncLoginUserID = planid
    ''    Exit function
    ''end if

    ''다른경우 조작된 경우임.
    ''관리자에게 메세지발송
    On Error Resume Next
    call InfoMsgMailSend("planid="&planid&"<br />"&"encedID="&encedID)
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
    strMsg = strMsg & "<li>서버:<br />"
	strMsg = strMsg & application("Svr_Info")
	strMsg = strMsg & "<br/ ><br /></li>"

	'// 접속자 브라우저 정보
	strMsg = strMsg & "<li>브라우저 종류:<br />"
	strMsg = strMsg & Server.HTMLEncode(Request.ServerVariables("HTTP_USER_AGENT"))
	strMsg = strMsg & "<br /><br /></li>"
	strMsg = strMsg & "<li>접속자 IP:<br />"
	strMsg = strMsg & Server.HTMLEncode(Request.ServerVariables("REMOTE_ADDR"))
	strMsg = strMsg & "<br /><br /></li>"
	strMsg = strMsg & "<li>경유페이지:<br />"
	strMsg = strMsg & request.ServerVariables("HTTP_REFERER")
	strMsg = strMsg & "<br /><br /></li>"
	'// 오류 페이지 정보
	strMsg = strMsg & "<li>페이지:<br />"
	strMethod = Request.ServerVariables("REQUEST_METHOD")
	strMsg = strMsg & "HOST : " & Request.ServerVariables("HTTP_HOST") & "<BR />"
	strMsg = strMsg & strMethod & " : "

	If strMethod = "POST" Then
		strMsg = strMsg & Request.TotalBytes & " bytes to "
	End If

	strMsg = strMsg & Request.ServerVariables("SCRIPT_NAME")
	strMsg = strMsg & "</li>"

	If strMethod = "POST" Then
		strMsg = strMsg & "<br /><li>POST Data:<br />"

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
		strMsg = strMsg & "<br /><li>GET Data:<br />"
		strMsg = strMsg & Request.QueryString
	End If
	strMsg = strMsg & "<br /><br /></li>"

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
	cdoMessage.SubJect 	= "["&date()&"] 10x10페이지 메세지 발생"
	cdoMessage.HTMLBody	= strMsg & "<br /><li>Message:<br />" & paramMsg &"</li>"

	cdoMessage.BodyPart.Charset="ks_c_5601-1987"         '/// 한글을 위해선 꼭 넣어 주어야 합니다.
    cdoMessage.HTMLBodyPart.Charset="ks_c_5601-1987"     '/// 한글을 위해선 꼭 넣어 주어야 합니다.

	cdoMessage.Send

	Set cdoMessage = nothing
	Set cdoConfig = nothing
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

Function fnKorChr(sChr)
	dim c, r
	c = Left(sChr,1)
	if c="" then Exit Function

	if c>="가" and c<"나" then r="ㄱ"
	if c>="나" and c<"다" then r="ㄴ"
	if c>="다" and c<"라" then r="ㄷ"
	if c>="라" and c<"마" then r="ㄹ"
	if c>="마" and c<"바" then r="ㅁ"
	if c>="바" and c<"사" then r="ㅂ"
	if c>="사" and c<"아" then r="ㅅ"
	if c>="아" and c<"자" then r="ㅇ"
	if c>="자" and c<"차" then r="ㅈ"
	if c>="차" and c<"카" then r="ㅊ"
	if c>="카" and c<"타" then r="ㅋ"
	if c>="타" and c<"파" then r="ㅌ"
	if c>="파" and c<"하" then r="ㅍ"
	if c>="하" and c<="힣" then r="ㅎ"
	fnKorChr = r
End Function

Function CheckCurse (ByVal txt) '욕 체크 2012-01-02 이종화
	Const Filename = "/curse.txt"
	Const ForReading = 1, ForWriting = 2, ForAppending = 3
	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

	Dim FilePath
	FilePath =  server.mappath("/chtml/curse/")

	Dim FSO
	set FSO = server.createObject("Scripting.FileSystemObject")

	if FSO.FileExists(Filepath&Filename) Then

		Dim TextStream
		Set TextStream = FSO.OpenTextFile(Filepath&Filename, ForReading, False, TristateUseDefault)

		Dim Contents '욕 -_-
		Contents = TextStream.ReadAll
		TextStream.Close
		Set TextStream = nothing

	Else

		Response.Write "<h3><i><font color=red> File " & Filename &_
					   " does not exist</font></i></h3>"

	End If
	Set FSO = Nothing

	If Contents <> "" Then

		Dim aFile
		aFile = Split(Contents,",")

		Dim Bftxt
		Dim i , ii
		Dim rpsStr ,costStr ,lenStr

		For i = 0 To UBound(aFile)
			If InStr(1,txt,aFile(i)) <> "0" Then
				lenStr = Len(aFile(i))
				'길이 만큼 * 모양
				For ii = 1 To lenStr
					costStr = "*"
					rpsStr = rpsStr & costStr
				Next
				'기존 텍스트 변환
				txt = Replace(txt,aFile(i),rpsStr)
				'별모양 초기화
				rpsStr = ""
			End If
			Bftxt = txt
		Next
	End if
	CheckCurse = Bftxt
End function

'2012-09-19 김진영 그누보드의 print_r2를 ASP화 변형형
Function r2(ByVal method)
	Dim i, pLength, msg, pValue, pJump, pLine, pTab
	pLine = "<br />"
	pTab = vbTab
	pTab = "    "
	msg = ""
	If IsArray(method) Then
		pLength = Ubound(method)
		For i=0 to pLength
			pValue = method(i)
			pJump = "&nbsp;&nbsp;" & pTab & "[" &  i & "]" & " => "
			If not IsArray(pValue) Then
				msg = msg & pJump & pValue & pLine
			Else
				msg = msg & pJump & r2(pValue)
			End If
		Next
		msg = "Array" & pLine & "(" & pLine & msg & ")" & pLine
		response.write "<span style='font-size:10pt; line-height:12px'>"&msg&"</span>"
	Else
		Dim key
		method = LCase(method)
		response.write "<table width=750 border=1 bordercolor='#cccccc' style='border-collapse:collapse;font:10pt'>" + vbcrlf
		response.write "<tr>" + vbcrlf
		response.write "	<td align='center' bgcolor='F1F1E5'>name</td>" + vbcrlf
		response.write "	<td align='center'>value</td>" + vbcrlf
		response.write "</tr>" + vbcrlf
		Select Case method
			Case "post" :
				For Each key in Request.Form
					response.write "<tr>" + vbcrlf
					response.write "<td bgcolor='#F1F1E5'>" & key & "</td>" + vbcrlf
					If IsArray(Request.Form(key)) Then
						response.write  "<td>" & r2(Request.Form(key)) & "</td>" + vbcrlf
					Else
						response.write  "<td>" & Request.Form(key) & "</td>" + vbcrlf
					End If
					response.write  "</tr>" + vbcrlf
				Next

			Case "get" :
				For Each key in Request.QueryString
					response.write  "<tr>" + vbcrlf
					response.write  "<td bgcolor='#F1F1E5'>" & key & "</td>" + vbcrlf
					If IsArray(Request.Form(key)) Then
						response.write  "<td>" & r2(Request.QueryString(key)) & "</td>" + vbcrlf
					Else
						response.write  "<td>" & Request.QueryString(key) & "</td>" + vbcrlf
					End If
					response.write  "</tr>" + vbcrlf
				Next
			Case "server" :
				For Each key in Request.ServerVariables
					response.write  "<tr>" + vbcrlf
					response.write  "<td bgcolor='#F1F1E5'>" & key & "</td>" + vbcrlf
					If IsArray(Request.Form(key)) Then
						response.write  "<td>" & r2(Request.ServerVariables(key)) & "</td>" + vbcrlf
					Else
						response.write  "<td>" & Request.ServerVariables(key) & "</td>" + vbcrlf
					End If
					response.write  "</tr>" + vbcrlf
				Next
				response.write  method
			Case Else : response.write method
		End Select
	End If
END function

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

'// 콤마로 구분된 배열값에 지정된 값이 있는지 반환(지정길이)
function chkArrValueLen(aVal,cVal,lmt)
	dim arrV, i
	chkArrValueLen = false
	arrV = split(aVal,",")
	for i=0 to ubound(arrV)
		if left(arrV(i),lmt)=left(cVal,lmt) then
			chkArrValueLen = true
			exit function
		end if
	next
end function

'// 콤마로 구분된 배열들에 지정된 값이 있는지 반환 (같은 수의 배열중 1에 있으면 2의 값을 반환)
function chkArrSelVal(aVal1,aVal2,cVal)
	dim arrV1, arrV2, i
	arrV1 = split(aVal1,",")
	arrV2 = split(aVal2,",")
	for i=0 to ubound(arrV1)
		if cStr(arrV1(i))=cStr(cVal) then
			chkArrSelVal = arrV2(i)
			exit function
		end if
	next
end function

'// ston캐시 서버 썸네일 제작(퀄러티 함께)
function getStonReSizeImg(furl,wd,ht,qua)
    getStonReSizeImg = furl&"/10x10/resize/"&wd&"x"&ht&"/quality/"&qua&"/"
end function

'// ston캐시 서버 썸네일 제작(기존 포토서버 썸네일 변경) - 리스트 위주
function getStonThumbImgURL(furl,wd,ht,fit,ws)
    getStonThumbImgURL = furl&"/10x10/thumbnail/"&wd&"x"&ht&"/quality/80/"
end function

'// 포토서버 썸네일 제작(기존 파일명)
function getThumbImgFromURL(furl,wd,ht,fit,ws)
	dim sCmd

	'도메인 치환
	if instr(furl,"imgstatic")>0 then
		furl = replace(furl,"imgstatic.10x10.co.kr/","thumbnail.10x10.co.kr/imgstatic/")
	elseif instr(furl,"webimage")>0 then
		furl = replace(furl,"webimage.10x10.co.kr/","thumbnail.10x10.co.kr/webimage/")
	end if

	'썸네일 커맨드
	sCmd = "?cmd=thumb"
	if wd<>"" then sCmd = sCmd & "&w=" & wd
	if ht<>"" then sCmd = sCmd & "&h=" & ht
	if fit<>"" then sCmd = sCmd & "&fit=" & fit
	if ws<>"" then sCmd = sCmd & "&ws=" & ws

	'변환주소 반환
	getThumbImgFromURL = furl & sCmd
end function

'// 포토서버 흑백아이콘 제작(기존 파일명)
function getGrayImgFromURL(furl)
	dim sCmd

	'도메인 치환
	if instr(furl,"imgstatic")>0 then
		furl = replace(furl,"imgstatic.10x10.co.kr/","thumbnail.10x10.co.kr/imgstatic/")
	elseif instr(furl,"webimage")>0 then
		furl = replace(furl,"webimage.10x10.co.kr/","thumbnail.10x10.co.kr/webimage/")
	end if

	'썸네일 커맨드
	sCmd = "?cmd=gray"

	'변환주소 반환
	getGrayImgFromURL = furl & sCmd
end function

Function MyBadge_IsExist_LoginDateCookie()
	dim loginDate, currDate

	MyBadge_IsExist_LoginDateCookie = False

	loginDate = request.cookies("mybadge")("logindate")
	if (Len(loginDate) <> 13) then
		exit Function
	end if

	currDate = Left(FormatDate(Now, "0000.00.00-00:00:00"), 13)

	if (Left(loginDate, 10) <> Left(currDate, 10)) then
		exit Function
	end if

	if (CInt(Right(loginDate,2)) < 4) and (CInt(Right(currDate,2)) >= 4) then
		exit Function
	end if

	MyBadge_IsExist_LoginDateCookie = True
end Function

Function MyBadge_IsExist_NewBadgeCookie()
	dim newBadge

	MyBadge_IsExist_NewBadgeCookie = False

	newBadge = request.cookies("mybadge")("newbadge")
	if (Len(newBadge) <> 1) then
		exit Function
	end if

	if (newBadge = "Y") then
		MyBadge_IsExist_NewBadgeCookie = True
	end if
end Function

Function MyAlarm_IsExist_CheckDateCookie()
	dim checkDate, currDate

	MyAlarm_IsExist_CheckDateCookie = False

	checkDate = request.cookies("myalarm")("checkdate")
	if (Len(checkDate) <> 13) then
		exit Function
	end if

	currDate = Left(FormatDate(Now, "0000.00.00-00:00:00"), 13)

	if (Left(checkDate, 10) <> Left(currDate, 10)) then
		exit Function
	end if

	if (CInt(Right(checkDate,2)) < 4) and (CInt(Right(currDate,2)) >= 4) then
		exit Function
	end if

	MyAlarm_IsExist_CheckDateCookie = True
end Function

Function MyAlarm_IsExist_NewMyAlarmCookie()
	dim newMyAlarm

	MyAlarm_IsExist_NewMyAlarmCookie = False

	newMyAlarm = request.cookies("myalarm")("newmyalarm")
	if (Len(newMyAlarm) <> 1) then
		exit Function
	end if

	if (newMyAlarm = "Y") then
		MyAlarm_IsExist_NewMyAlarmCookie = True
	end if
end Function

Function GetMyTenTenBgColor()
	dim my10x10_bgcolor : my10x10_bgcolor = request.cookies("my10x10")("bgcolor")
	dim my10x10_bgcolor_cookie : my10x10_bgcolor_cookie = request.cookies("cs_my10x10_bgcolor")

	if (my10x10_bgcolor_cookie <> "") then
		if (my10x10_bgcolor <> Left(my10x10_bgcolor_cookie, 10)) then
			my10x10_bgcolor = Left(my10x10_bgcolor_cookie, 10)
			response.cookies("my10x10")("bgcolor") = my10x10_bgcolor
		end if
	end if

	if (my10x10_bgcolor = "") then
		my10x10_bgcolor = "skinBlue"
		response.cookies("my10x10")("bgcolor") = my10x10_bgcolor
	end if

	GetMyTenTenBgColor = my10x10_bgcolor
End Function

'//날짜형식 2013-01-01 오후 03:00:00 형식을 2013-01-01 15:00:00로 변환		'/2013.04.22 한용민 생성
function dateconvert(dateval)
	dim tmpval
	if dateval = "" then exit function

	tmpval = year(dateval)
	tmpval = tmpval & "-" & Format00(2,month(dateval))
	tmpval = tmpval & "-" & Format00(2,day(dateval))
	tmpval = tmpval & " " & Format00(2,hour(dateval))
	tmpval = tmpval & ":" & Format00(2,minute(dateval))
	tmpval = tmpval & ":" & Format00(2,second(dateval))

	dateconvert = left(tmpval,19)
end Function


'// 해당경로에 파일이 있는지 체크한다.
Function checkFilePath(filePath)
	Dim fso, result
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.Fileexists(filePath) Then
		result = 1
	Else
		result = 0
	End If
	checkFilePath = result
	Set fso = nothing
End Function

'//카테고리 메인 GaParam
Function CateMain_GaParam(vDisp,Po,Num)
	Dim gParam , url , tDisp
	tDisp = Left(vDisp,3)
	url = getThisFullURL()

	Select Case CStr(tDisp)
	Case "101"
		gParam = "gaparam=stationery"
	Case "102"
		gParam = "gaparam=digital"
	Case "103"
		gParam = "gaparam=travel"
	Case "104"
		gParam = "gaparam=toy"
	Case "121"
		gParam = "gaparam=furniture"
	Case "122"
		gParam = "gaparam=deco"
	Case "120"
		gParam = "gaparam=fabric"
	Case "112"
		gParam = "gaparam=kitchen"
	Case "119"
		gParam = "gaparam=food"
	Case "117"
		gParam = "gaparam=fashion"
	Case "116"
		gParam = "gaparam=accessories"
	Case "118"
		gParam = "gaparam=beauty"
	Case "115"
		gParam = "gaparam=kids"
	Case "110"
		gParam = "gaparam=catdog"
	Case Else
		gParam = "gaparam="
	end Select

	If InStr(url,"?") > 0 Then
		CateMain_GaParam = "&" & gParam &"_"& Po &"_"& Num
	Else
		CateMain_GaParam = "?" & gParam &"_"& Po &"_"& Num
	End If

End function
%>
