<%
    '/*
    ' * [가상계좌 발급/변경요청 페이지]
    ' *
    ' * 가상계좌 발급 변경(CHANGE)은 금액과 마감일만 변경 할수 있습니다. 
    ' */
    CST_PLATFORM         = trim(request("CST_PLATFORM"))         ' LG텔레콤 결제서비스 선택(test:테스트, service:서비스)
    CST_MID              = trim(request("CST_MID"))              ' LG텔레콤으로 부터 발급받으신 상점아이디를 입력하세요.
                                                                 ' 테스트 아이디는 't'를 제외하고 입력하세요.
    if CST_PLATFORM = "test" then                                ' 상점아이디(자동생성)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if
    LGD_METHOD           = trim(request("LGD_METHOD"))           ' ASSIGN:할당, CHANGE:변경
    LGD_OID     		 = trim(request("LGD_OID"))    			 ' 주문번호(상점정의 유니크한 주문번호를 입력하세요)
    LGD_AMOUNT      	 = trim(request("LGD_AMOUNT"))      	 ' 금액("," 를 제외한 금액을 입력하세요)
    LGD_PRODUCTINFO   	 = trim(request("LGD_PRODUCTINFO"))  	 ' 상품정보
    LGD_BUYER          	 = trim(request("LGD_BUYER"))         	 ' 구매자명
	LGD_ACCOUNTOWNER     = trim(request("LGD_ACCOUNTOWNER"))  	 ' 입금자명
	LGD_ACCOUNTPID       = trim(request("LGD_ACCOUNTPID"))       ' 입금자주민번호(옵션)
	LGD_BUYERPHONE       = trim(request("LGD_BUYERPHONE"))       ' 구매자휴대폰번호
	LGD_BUYEREMAIL       = trim(request("LGD_BUYEREMAIL"))       ' 구매자이메일(옵션)
	LGD_BANKCODE         = trim(request("LGD_BANKCODE"))         ' 입금계좌은행코드
	LGD_CASHRECEIPTUSE   = trim(request("LGD_CASHRECEIPTUSE"))   ' 현금영수증 발행구분('1':소득공제, '2':지출증빙)
	LGD_CASHCARDNUM      = trim(request("LGD_CASHCARDNUM"))      ' 현금영수증 카드번호
	LGD_CLOSEDATE        = trim(request("LGD_CLOSEDATE"))        ' 입금 마감일
	LGD_TAXFREEAMOUNT    = trim(request("LGD_TAXFREEAMOUNT"))    ' 면세금액
	LGD_CASNOTEURL       = "http://webadmin.10x10.co.kr/admin/apps/DC_CA_noteurl.asp" ''"http://상점URL/cas_noteurl.asp"       ' 입금결과 처리를 위한 상점페이지를 반드시 설정해 주세요
	

    'configPath           = "C:/lgdacom"         				 ' LG텔레콤에서 제공한 환경파일("/conf/lgdacom.conf") 위치 지정.
    configPath				   = "C:/lgdacom/conf/" & CST_MID

    Set xpay = server.CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM
    xpay.Init_TX(LGD_MID)

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
    

    '/*
    ' * 1. 가상계좌 발급/변경 요청 결과처리
    ' *
    ' * 결과 리턴 파라미터는 연동메뉴얼을 참고하시기 바랍니다.
    ' */
    if xpay.TX() then
        if LGD_METHOD = "ASSIGN" then      '가상계좌 발급의 경우
        
        	'1)가상계좌 발급결과 화면처리(성공,실패 결과 처리를 하시기 바랍니다.)
        	Response.Write("가상계좌 발급 요청처리가 완료되었습니다. <br>")
        	Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        	Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
			
			Response.Write("결과코드 : " & xpay.Response("LGD_RESPCODE", 0) & "<br>")
	    	Response.Write("거래번호 : " & xpay.Response("LGD_TID", 0) & "<p>")
        	
        	'아래는 결제요청 결과 파라미터를 모두 찍어 줍니다.
        	Dim itemCount
        	Dim resCount
        	itemCount = xpay.resNameCount
        	resCount = xpay.resCount

        	For i = 0 To itemCount - 1
            	itemName = xpay.ResponseName(i)
            	Response.Write(itemName & "&nbsp:&nbsp")
            	For j = 0 To resCount - 1
                	Response.Write(xpay.Response(itemName, j) & "<br>")
            	Next
        	Next
        
        else		'가상계좌 변경의 경우
        	'1)가상계좌 변경결과 화면처리(성공,실패 결과 처리를 하시기 바랍니다.)
        	Response.Write("가상계좌 변경 요청처리가 완료되었습니다. <br>")
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
        
        Response.Write("<p>")
            
    else
        '2)API 요청 실패 화면처리
        Response.Write("가상계좌 발급/변경 요청처리가 실패되었습니다. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
    end if
%>
