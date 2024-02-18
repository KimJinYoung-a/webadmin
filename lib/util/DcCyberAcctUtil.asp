<%

function getLGD_FINANCECODE2Name(fCode)
    select Case fCode
        CASE "11" : getLGD_FINANCECODE2Name = "농협"
        CASE "06" : getLGD_FINANCECODE2Name = "국민"
        CASE "20" : getLGD_FINANCECODE2Name = "우리"
        CASE "26" : getLGD_FINANCECODE2Name = "신한"
        CASE "81" : getLGD_FINANCECODE2Name = "하나"
        CASE "03" : getLGD_FINANCECODE2Name = "기업"
        CASE "05" : getLGD_FINANCECODE2Name = "외환"
        CASE "39" : getLGD_FINANCECODE2Name = "경남"
        CASE "32" : getLGD_FINANCECODE2Name = "부산"
        CASE "71" : getLGD_FINANCECODE2Name = "우체국"
        CASE "07" : getLGD_FINANCECODE2Name = "수협"
        CASE "31" : getLGD_FINANCECODE2Name = "대구"
        CASE ELSE : getLGD_FINANCECODE2Name = ""
    end Select
end function

function CheckNChangeCyberAcct(iorderserial)
    dim sqlStr
    dim ipkumdiv, accountdiv, accountNo, cancelyn, subtotalPrice, OLDsubtotalPrice, OLDCancelyn, sumPaymentEtc
    ipkumdiv = 0
    OLDsubtotalPrice = 0
    OLDCancelyn      = ""

    CheckNChangeCyberAcct = false

    sqlStr = " select orderserial, ipkumdiv, accountdiv, accountNo, cancelyn, subtotalPrice, sumPaymentEtc"
    sqlStr = sqlStr & " from db_order.dbo.tbl_order_master"
    sqlStr = sqlStr & " where orderserial='" & iorderserial & "'"

    rsget.Open sqlStr,dbget,1
    if (Not rsget.Eof) then
        ipkumdiv    = rsget("ipkumdiv")
		accountdiv  = rsget("accountdiv")
		accountNo   = rsget("accountNo")
		cancelyn    = rsget("cancelyn")
		subtotalPrice = rsget("subtotalPrice")
        sumPaymentEtc = rsget("sumPaymentEtc")
    end if
	rsget.close

	if (ipkumdiv<>2) then Exit function
	if (accountdiv<>"7") then Exit function

	if (accountNo="국민 470301-01-014754") _
        or (accountNo="신한 100-016-523130") _
        or (accountNo="우리 092-275495-13-001") _
        or (accountNo="하나 146-910009-28804") _
        or (accountNo="기업 277-028182-01-046") _
        or (accountNo="농협 029-01-246118") then
            Exit function
    end if

    dim CLOSEDATE
    if (cancelyn<>"N") then
        CLOSEDATE = Replace(Left(CStr(now()),10),"-","") & "000000"
    else
        CLOSEDATE = Replace(Left(CStr(DateAdd("d",10,now())),10),"-","") & "235959"
    end if

    sqlStr = " select top 1 subtotalPrice, convert(varchar(19),CLOSEDATE,20) as CLOSEDATE "
    sqlStr = sqlStr & " from db_order.dbo.tbl_order_CyberAccountLog"
    sqlStr = sqlStr & " where orderserial='" & iorderserial & "'"
    sqlStr = sqlStr & " order by differencekey desc"
    rsget.Open sqlStr,dbget,1
    if (Not rsget.Eof) then
        OLDsubtotalPrice = rsget("subtotalPrice")
        OLDCancelyn      = rsget("CLOSEDATE")

        if (RIGHT(OLDCancelyn,8)="00:00:00") then
            OLDCancelyn="Y"
        else
            OLDCancelyn="N"
        end if
    end if
    rsget.close

    if (OLDsubtotalPrice<>subtotalPrice) or (OLDCancelyn<>Cancelyn) then
        '// 유플러스 전송시에는 보조결제 뺀금액 전송
        CheckNChangeCyberAcct = ChangeCyberAcct(iorderserial, subtotalPrice-sumPaymentEtc, CLOSEDATE)
    end if
end function

function CheckNAssignCyberAcct(asid, iorderserial, CyberAcctCode)
	'// CyberAcctCode = 입금계좌은행코드
    dim sqlStr
    dim ipkumdiv, accountdiv, accountNo, cancelyn, subtotalPrice, goodname, buyname, accountname, buyhp, buyemail, userid
    ipkumdiv = 0

    CheckNAssignCyberAcct = false

    sqlStr = " select orderserial, ipkumdiv, accountdiv, accountNo, cancelyn, subtotalPrice, buyname, accountname, buyhp, buyemail, userid "
    sqlStr = sqlStr & " from db_order.dbo.tbl_order_master"
    sqlStr = sqlStr & " where orderserial='" & iorderserial & "'"

    rsget.Open sqlStr,dbget,1
    if (Not rsget.Eof) then
        ipkumdiv    = rsget("ipkumdiv")
		accountdiv  = rsget("accountdiv")
		accountNo   = rsget("accountNo")
		cancelyn    = rsget("cancelyn")
		subtotalPrice = rsget("subtotalPrice")
		buyname    	= rsget("buyname")
		accountname	= rsget("accountname")
		buyhp    	= rsget("buyhp")
		buyemail    = rsget("buyemail")
		userid    	= rsget("userid")
    end if
	rsget.close

	if (ipkumdiv<>0) then Exit function
	if (accountdiv<>"7") then Exit function

    sqlStr = " select max(itemname) as itemname, count(*) as cnt "
    sqlStr = sqlStr & " from "
    sqlStr = sqlStr & " [db_order].[dbo].[tbl_order_detail] "
    sqlStr = sqlStr & " where orderserial = '" & iorderserial & "' and itemid <> 0 and cancelyn <> 'Y' "
    rsget.Open sqlStr,dbget,1
    if (Not rsget.Eof) then
		if rsget("cnt") > 0 then
			goodname = rsget("itemname")
			if rsget("cnt") > 1 then
				goodname = goodname & " 외 " & rsget("cnt")
			end if
		else
			goodname = "배송비"
		end if
    end if
	rsget.close

	CheckNAssignCyberAcct = AssignCyberAcct(asid, iorderserial, subtotalPrice, goodname, buyname, accountname, buyhp, buyemail, userid, CyberAcctCode)
end function

function ChangeCyberAcct(LGD_OID, LGD_AMOUNT, LGD_CLOSEDATE)
    '/*
    ' * [가상계좌 발급/변경요청 페이지]
    ' *
    ' * 가상계좌 발급 변경(CHANGE)은 금액과 마감일만 변경 할수 있습니다.
    ' */
    dim CST_PLATFORM : CST_PLATFORM         = ""         ' LG텔레콤 결제서비스 선택(test:테스트, service:서비스)
    IF application("Svr_Info")="Dev" THEN CST_PLATFORM = "test"
''CST_PLATFORM = ""

    dim CST_MID : CST_MID = "tenbyten01"                 ' LG텔레콤으로 부터 발급받으신 상점아이디를 입력하세요.

    dim LGD_MID                                                  ' 테스트 아이디는 't'를 제외하고 입력하세요.
    if CST_PLATFORM = "test" then                                ' 상점아이디(자동생성)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if

    dim LGD_METHOD : LGD_METHOD          = "CHANGE"                              ' ASSIGN:할당, CHANGE:변경

    'LGD_PRODUCTINFO   	 = trim(request("LGD_PRODUCTINFO"))  	 ' 상품정보
    'LGD_BUYER          	 = trim(request("LGD_BUYER"))         	 ' 구매자명
	'LGD_ACCOUNTOWNER     = trim(request("LGD_ACCOUNTOWNER"))  	 ' 입금자명
	'LGD_ACCOUNTPID       = trim(request("LGD_ACCOUNTPID"))       ' 입금자주민번호(옵션)
	'LGD_BUYERPHONE       = trim(request("LGD_BUYERPHONE"))       ' 구매자휴대폰번호
	'LGD_BUYEREMAIL       = trim(request("LGD_BUYEREMAIL"))       ' 구매자이메일(옵션)
	'LGD_BANKCODE         = trim(request("LGD_BANKCODE"))         ' 입금계좌은행코드
	'LGD_CASHRECEIPTUSE   = trim(request("LGD_CASHRECEIPTUSE"))   ' 현금영수증 발행구분('1':소득공제, '2':지출증빙)
	'LGD_CASHCARDNUM      = trim(request("LGD_CASHCARDNUM"))      ' 현금영수증 카드번호
	'LGD_TAXFREEAMOUNT    = trim(request("LGD_TAXFREEAMOUNT"))    ' 면세금액
	'LGD_CASNOTEURL       = "http://61.252.133.2:8888/admin/apps/DC_CA_noteurl.asp" ''"http://상점URL/cas_noteurl.asp"       ' 입금결과 처리를 위한 상점페이지를 반드시 설정해 주세요


    'configPath           = "C:/lgdacom"         				 ' LG텔레콤에서 제공한 환경파일("/conf/lgdacom.conf") 위치 지정.
    dim configPath : configPath				   = "C:/lgdacom" '''"C:/lgdacom/conf/" & CST_MID  ''conf 파일 통합 2013/02/15

    dim xpay
    Set xpay = server.CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM
    xpay.Init_TX(LGD_MID)

    xpay.Set "LGD_TXNAME", "CyberAccount"
    xpay.Set "LGD_METHOD", LGD_METHOD
    xpay.Set "LGD_OID", LGD_OID
    xpay.Set "LGD_AMOUNT", LGD_AMOUNT
    xpay.Set "LGD_CLOSEDATE", LGD_CLOSEDATE
    'xpay.Set "LGD_PRODUCTINFO", LGD_PRODUCTINFO
    'xpay.Set "LGD_BUYER", LGD_BUYER
    'xpay.Set "LGD_ACCOUNTOWNER", LGD_ACCOUNTOWNER
    'xpay.Set "LGD_ACCOUNTPID", LGD_ACCOUNTPID
    'xpay.Set "LGD_BUYERPHONE", LGD_BUYERPHONE
    'xpay.Set "LGD_BUYEREMAIL", LGD_BUYEREMAIL
    'xpay.Set "LGD_BANKCODE", LGD_BANKCODE
    'xpay.Set "LGD_CASHRECEIPTUSE", LGD_CASHRECEIPTUSE
    'xpay.Set "LGD_CASHCARDNUM", LGD_CASHCARDNUM

    'xpay.Set "LGD_TAXFREEAMOUNT", LGD_TAXFREEAMOUNT
    'xpay.Set "LGD_CASNOTEURL", LGD_CASNOTEURL


    '/*
    ' * 1. 가상계좌 발급/변경 요청 결과처리
    ' *
    ' * 결과 리턴 파라미터는 연동메뉴얼을 참고하시기 바랍니다.
    ' */
    Dim itemCount, itemName, resCount, i, j
    Dim sqlStr

    ChangeCyberAcct = false

    if (xpay.TX()) then
        if LGD_METHOD = "ASSIGN" then      '가상계좌 발급의 경우

'        	'1)가상계좌 발급결과 화면처리(성공,실패 결과 처리를 하시기 바랍니다.)
'        	Response.Write("가상계좌 발급 요청처리가 완료되었습니다. <br>")
'        	Response.Write("TX Response_code = " & xpay.resCode & "<br>")
'        	Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
'
'			Response.Write("결과코드 : " & xpay.Response("LGD_RESPCODE", 0) & "<br>")
'	    	Response.Write("거래번호 : " & xpay.Response("LGD_TID", 0) & "<p>")
'
'        	'아래는 결제요청 결과 파라미터를 모두 찍어 줍니다.
'
'        	itemCount = xpay.resNameCount
'        	resCount = xpay.resCount
'
'        	For i = 0 To itemCount - 1
'            	itemName = xpay.ResponseName(i)
'            	Response.Write(itemName & "&nbsp:&nbsp")
'            	For j = 0 To resCount - 1
'                	Response.Write(xpay.Response(itemName, j) & "<br>")
'            	Next
'        	Next

        else		'가상계좌 변경의 경우
        	'1)가상계좌 변경결과 화면처리(성공,실패 결과 처리를 하시기 바랍니다.)


        	ChangeCyberAcct = (Trim(xpay.resCode)="0000")

        	if (Trim(xpay.resCode)="0000") then
        	    sqlStr = " IF EXISTS (select orderserial from db_order.dbo.tbl_order_CyberAccountLog where orderserial='" & LGD_OID & "')" & VbCrlf
                sqlStr = sqlStr & " BEGIN" & VbCrlf
                sqlStr = sqlStr & "	Insert Into db_order.dbo.tbl_order_CyberAccountLog" & VbCrlf
                sqlStr = sqlStr & "	(orderserial, differencekey, userid, FINANCECODE,ACCOUNTNUM" & VbCrlf
                sqlStr = sqlStr & "	, subtotalPrice, CLOSEDATE"& VbCrlf
                sqlStr = sqlStr & "	,RefIP)" & VbCrlf
                sqlStr = sqlStr & "	select top 1 orderserial, (differencekey+1) as differencekey" & VbCrlf
                sqlStr = sqlStr & "	,userid, FINANCECODE, ACCOUNTNUM" & VbCrlf
                sqlStr = sqlStr & "	, " & LGD_AMOUNT & " as subtotalprice" & VbCrlf
                sqlStr = sqlStr & "	, '" & Left(LGD_CLOSEDATE,4) + "-" + Mid(LGD_CLOSEDATE,5,2) + "-" + Mid(LGD_CLOSEDATE,7,2) + " " + Mid(LGD_CLOSEDATE,9,2) + ":" + Mid(LGD_CLOSEDATE,11,2) + ":" + Mid(LGD_CLOSEDATE,13,2) & "' as CLOSEDATE" & VbCrlf
                sqlStr = sqlStr & "	, '" & Left(request.ServerVariables("REMOTE_ADDR"),32) & "' as refip" & VbCrlf
                sqlStr = sqlStr & "	from db_order.dbo.tbl_order_CyberAccountLog" & VbCrlf
                sqlStr = sqlStr & "	where orderserial='" & LGD_OID & "'" & VbCrlf
                sqlStr = sqlStr & "	order by differencekey desc" & VbCrlf
                sqlStr = sqlStr & " END"

                dbget.Execute sqlStr
            ELSE
            	Response.Write("결과코드 : " & xpay.Response("LGD_RESPCODE", 0) & "<br>")
                Response.Write("주문번호 : " & LGD_OID & "<br>")
                Response.Write("입금액 : " & LGD_AMOUNT & "<br>")
            	Response.Write("입금마감일 : " & LGD_CLOSEDATE & "<p>")


            	itemCount = xpay.resNameCount
            	resCount = xpay.resCount

            	For i = 0 To itemCount - 1
                	itemName = xpay.ResponseName(i)
                	Response.Write(itemName & "&nbsp:&nbsp")
                	For j = 0 To resCount - 1
                    	Response.Write(xpay.Response(itemName, j) & "<br>")
                	Next
            	Next
        	end if
        end if
    else
        '2)API 요청 실패 화면처리
        ''Response.Write("가상계좌 발급/변경 요청처리가 실패되었습니다. <br>")
        ''Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        ''Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
    end if

end function

function AssignCyberAcct(asid, iorderserial, subtotalPrice, goodname, buyname, accountname, buyhp, buyemail, userid, CyberAcctCode)
    '/*
    ' * [가상계좌 발급요청 페이지]
    ' */
	dim LGD_FINANCECODE, LGD_ACCOUNTNUM, Tid, accountno
	dim FINANCECODE, ACCOUNTNUM, CLOSEDATE, IsSuccess
	dim sqlStr, iresultmsg

    dim CST_PLATFORM : CST_PLATFORM         = ""         		' LG텔레콤 결제서비스 선택(test:테스트, service:서비스)
    IF application("Svr_Info")="Dev" THEN CST_PLATFORM = "test"

    dim CST_MID : CST_MID = "tenbyten01"                 		' LG텔레콤으로 부터 발급받으신 상점아이디를 입력하세요.

    dim LGD_MID                                                 ' 테스트 아이디는 't'를 제외하고 입력하세요.
    if CST_PLATFORM = "test" then                               ' 상점아이디(자동생성)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if

    dim LGD_METHOD       : LGD_METHOD        = "ASSIGN"             				' ASSIGN:할당, CHANGE:변경
    dim LGD_OID          : LGD_OID     		 = iorderserial    						' 주문번호(상점정의 유니크한 주문번호를 입력하세요)
    dim LGD_AMOUNT       : LGD_AMOUNT      	 = subtotalprice      					' 금액("," 를 제외한 금액을 입력하세요)
    dim LGD_PRODUCTINFO  : LGD_PRODUCTINFO   = trim(goodname)  	 					' 상품정보
    dim LGD_BUYER        : LGD_BUYER         = trim(buyname)         				' 구매자명
	dim LGD_ACCOUNTOWNER : LGD_ACCOUNTOWNER  = trim(accountname)  					' 입금자명
	dim LGD_ACCOUNTPID
	    LGD_ACCOUNTPID = Left(asid, 13)         									' 입금자주민번호(옵션)/아이디 MAX 13 ,금액체크

	dim LGD_BUYERPHONE   : LGD_BUYERPHONE       = trim(Replace(buyhp,"-",""))       ' 구매자휴대폰번호
	dim LGD_BUYEREMAIL   : LGD_BUYEREMAIL       = trim(buyemail)       				' 구매자이메일(옵션)
	dim LGD_BANKCODE     : LGD_BANKCODE         = trim(CyberAcctCode)         		' 입금계좌은행코드

	dim LGD_CASHRECEIPTUSE, LGD_CASHCARDNUM
''이니시스 현금영수증으로 사용
''	if (request.Form("cashreceiptreq")="Y") then
''	    LGD_CASHRECEIPTUSE   = trim(useopt+1)   ' 현금영수증 발행구분('1':소득공제, '2':지출증빙)
''	    LGD_CASHCARDNUM      = trim(request.Form("cashReceipt_ssn")) ''trim(request("LGD_CASHCARDNUM"))      ' 현금영수증 카드번호
''	else
''	    LGD_CASHRECEIPTUSE  =""
''	    LGD_CASHCARDNUM     =""
''    end if

	dim LGD_CLOSEDATE
		LGD_CLOSEDATE       = trim(Replace(Left(dateadd("d",10,now()),10),"-","") + "235959")        ' 입금 마감일 20100331 000000
	dim LGD_TAXFREEAMOUNT : LGD_TAXFREEAMOUNT   = "0 "    ' 면세금액
	dim LGD_CASNOTEURL    : LGD_CASNOTEURL      = "http://scm.10x10.co.kr/admin/apps/DC_CA_noteurl.asp"       ' 입금결과 처리를 위한 상점페이지를 반드시 설정해 주세요
IF application("Svr_Info")="Dev" THEN LGD_CASNOTEURL = "http://61.252.133.2:8888/admin/apps/DC_CA_noteurl.asp"

    dim configPath : configPath				   = "C:/lgdacom" '''/conf/" & CST_MID
    dim xpay

    On Error Resume Next
    Set xpay = server.CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM
    xpay.Init_TX(LGD_MID)

    IF (ERR) then
        response.write Err.Description
        response.write "<script language='javascript'>alert('결제는 이루어 지지 않았습니다. \n\n: 죄송합니다. 가상계좌 발급에 오류가 있습니다. \n\n잠시후 다시 시도해 주시기 바랍니다.');</script>"
        response.end
    End IF
    On Error Goto 0

    xpay.Set "LGD_TXNAME", "CyberAccount"
    xpay.Set "LGD_METHOD", LGD_METHOD
    xpay.Set "LGD_OID", LGD_OID
    xpay.Set "LGD_AMOUNT", LGD_AMOUNT
    xpay.Set "LGD_PRODUCTINFO", LGD_PRODUCTINFO
    xpay.Set "LGD_BUYER", LGD_BUYER
    xpay.Set "LGD_ACCOUNTOWNER", LGD_ACCOUNTOWNER
    xpay.Set "LGD_ACCOUNTPID", LGD_ACCOUNTPID
    xpay.Set "LGD_BUYERPHONE", LGD_BUYERPHONE
    xpay.Set "LGD_BUYEREMAIL", LGD_BUYEREMAIL
    xpay.Set "LGD_BANKCODE", LGD_BANKCODE
    xpay.Set "LGD_CASHRECEIPTUSE", LGD_CASHRECEIPTUSE
    xpay.Set "LGD_CASHCARDNUM", LGD_CASHCARDNUM
    xpay.Set "LGD_CLOSEDATE", LGD_CLOSEDATE
    xpay.Set "LGD_TAXFREEAMOUNT", LGD_TAXFREEAMOUNT
    xpay.Set "LGD_CASNOTEURL", LGD_CASNOTEURL

    xpay.Set "LGD_CUSTOM_CASSMSMSG", "[텐바이텐] [LGD_FINANCENAME] [LGD_SA] [LGD_COMPANYNAME] [LGD_AMOUNT]원 주문번호:"&iorderserial&" 감사합니다"  ''2015/07/22

    if xpay.TX() then
        if LGD_METHOD = "ASSIGN" then      '가상계좌 발급의 경우
            LGD_FINANCECODE = xpay.Response("LGD_FINANCECODE", 0)   ''은행
            LGD_ACCOUNTNUM = xpay.Response("LGD_ACCOUNTNUM", 0)   ''가상계좌
            Tid = xpay.Response("LGD_TID", 0)
        end if
    else
        response.write " [" + xpay.resCode + "] " & Replace(Left(xpay.resMsg,60),"'","")
		response.end
    end if

    IsSuccess = (xpay.resCode="0000")

    iresultmsg  = Left(xpay.resMsg,90)
    paygatetid = Tid

    if IsSuccess then
        FINANCECODE = LGD_FINANCECODE
        ACCOUNTNUM  = LGD_ACCOUNTNUM
        CLOSEDATE   = LGD_CLOSEDATE
        accountno = getLGD_FINANCECODE2Name(LGD_FINANCECODE) & " " & LGD_ACCOUNTNUM
        if (iresultmsg="") then
            iresultmsg =  "[가상계좌] " & accountno
        end if
    else
        iresultmsg = "[" & xpay.resCode & "]" & iresultmsg
    end if

    if Not IsSuccess then
        ''가상계좌도 실패건 있을 수 있도록 변경함.
		'// 주문내역 주문실패 처리
		sqlStr = " update [db_order].[dbo].tbl_order_master" + vbCrlf
		sqlStr = sqlStr + " set ipkumdiv='1' " + vbCrlf
		if (iresultmsg<>"") then
		    sqlStr = sqlStr + " ,resultmsg=convert(varchar(100),'" + iresultmsg + "')" + vbCrlf
		end if
		sqlStr = sqlStr + " where orderserial='" + CStr(iorderserial) + "'" + vbCrlf

		''response.write sqlStr & "<br>"
		dbget.Execute(sqlStr)
	else
		''' 주문 마스타 서머리 재저장
		sqlStr = " update [db_order].[dbo].tbl_order_master" + vbCrlf
		sqlStr = sqlStr + " set accountno='" + accountno + "' " + vbCrlf
		sqlStr = sqlStr + " ,ipkumdiv='2'" + vbCrlf

		if (paygatetid<>"") then
		    sqlStr = sqlStr + " ,paygatetid='" + paygatetid + "'" + vbCrlf
		end if

		if (iresultmsg<>"") then
		    sqlStr = sqlStr + " ,resultmsg=convert(varchar(100),'" + iresultmsg + "')" + vbCrlf
		end if

		sqlStr = sqlStr + " where orderserial='" + CStr(iorderserial) + "'" + vbCrlf

		''response.write sqlStr & "<br>"
		dbget.Execute(sqlStr)

        sqlStr = " insert into db_order.dbo.tbl_order_CyberAccountLog"
        sqlStr = sqlStr & " (orderserial, differencekey, userid, FINANCECODE, ACCOUNTNUM, subtotalPrice, CLOSEDATE, RefIP)"
        sqlStr = sqlStr & " values('" & iorderserial & "'"
        sqlStr = sqlStr & " ,0"
        sqlStr = sqlStr & " ,'" & userid & "'"
        sqlStr = sqlStr & " ,'" & FINANCECODE & "'"
        sqlStr = sqlStr & " ,'" & ACCOUNTNUM & "'"
        sqlStr = sqlStr & " ,'" & subtotalprice & "'"
        sqlStr = sqlStr & " ,'" & Left(CLOSEDATE,4) + "-" + Mid(CLOSEDATE,5,2) + "-" + Mid(CLOSEDATE,7,2) + " " + Mid(CLOSEDATE,9,2) + ":" + Mid(CLOSEDATE,11,2) + ":" + Mid(CLOSEDATE,13,2) & "'"
        sqlStr = sqlStr & " ,'" & Left(request.ServerVariables("REMOTE_ADDR"),32) & "'"
        sqlStr = sqlStr & " )"

        dbget.Execute sqlStr
    end if
    SET xpay = Nothing

	AssignCyberAcct = IsSuccess
end function

%>
