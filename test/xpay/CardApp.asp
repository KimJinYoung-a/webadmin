<%
	'공통정보
    CST_PLATFORM               = trim(request("CST_PLATFORM"))       'LG데이콤 결제 서비스 선택(test:테스트, service:서비스)
    CST_MID                    = trim(request("CST_MID"))            '상점아이디(LG데이콤으로 부터 발급받으신 상점아이디를 입력하세요)
                                                                     '테스트 아이디는 't'를 반드시 제외하고 입력하세요.
    if CST_PLATFORM = "test" then                                    '상점아이디(자동생성)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if
    LGD_OID                 = trim(request("LGD_OID"))            	'주문번호(상점정의 유니크한 주문번호를 입력하세요)
    LGD_OID                 = "2010010100003"
    LGD_BUYER          		= trim(request("LGD_BUYER"))     		'구매자명
    LGD_PRODUCTINFO         = trim(request("LGD_PRODUCTINFO"))     	'상품정보
    LGD_BUYEREMAIL          = trim(request("LGD_BUYEREMAIL"))     	'이메일주소(결제성공시 메일발송)
    LGD_AMOUNT              = trim(request("LGD_AMOUNT"))         	'결제금액("," 를 제외한 결제금액을 입력하세요)
    LGD_AUTHTYPE			= trim(request("LGD_AUTHTYPE"))		 	'인증유형(ISP인경우만  'ISP')
    LGD_CARDTYPE			= trim(request("LGD_CARDTYPE"))			'카드사코드
    
    '안심클릭 인증 또는 해외카드
    LGD_PAN                 = trim(request("LGD_PAN"))            	'카드번호    
    LGD_INSTALL             = trim(request("LGD_INSTALL"))        	'할부개월수(두자리숫자)
    LGD_NOINT				= trim(request("LGD_NOINT"))		    '무이자할부여부('1':상점부담무이자할부,'0':일반할부)
    LGD_EXPYEAR             = trim(request("LGD_EXPYEAR"))        	'유효기간년(YY)
   	LGD_EXPMON              = trim(request("LGD_EXPMON"))         	'유효기간월(MM)
    VBV_ECI             	= trim(request("VBV_ECI"))				'안심클릭ECI  
 	VBV_CAVV				= trim(request("VBV_CAVV"))			 	'안심클릭CAVV
 	VBV_XID				   	= trim(request("VBV_XID"))			 	'안심클릭XID    
    
    'ISP인증
    KVP_QUOTA				= trim(request("KVP_QUOTA"))			'할부개월수
    KVP_NOINT				= trim(request("KVP_NOINT"))			'무이자할부여부('1':상점부담무이자할부,'0':일반할부)
	KVP_CARDCODE			= trim(request("KVP_CARDCODE"))			'ISP카드코드
	KVP_SESSIONKEY			= trim(request("KVP_SESSIONKEY"))		'ISP세션키
	KVP_ENCDATA				= trim(request("KVP_ENCDATA"))		 	'ISP암호화데이터
	
    '' configPath				   = "C:/lgdacom"
    configPath				   = "C:/lgdacom/conf/" & CST_MID					 'LG데이콤에서 제공한 환경파일(/conf/lgdacom.conf, /conf/mall.conf)이 위치한 디렉토리 지정 
    
	Dim xpay
	Dim i, j
	Dim itemName
	
	Set xpay = server.CreateObject("XPayClientCOM.XPayClient")	
    xpay.Init configPath, CST_PLATFORM    
    xpay.Init_TX(LGD_MID)

    xpay.Set "LGD_TXNAME", "CardAuth"
    xpay.Set "LGD_OID", LGD_OID 
	xpay.Set "LGD_AMOUNT", LGD_AMOUNT
	xpay.Set "LGD_BUYER", LGD_BUYER
	xpay.Set "LGD_PRODUCTINFO", LGD_PRODUCTINFO
	xpay.Set "LGD_BUYEREMAIL", LGD_BUYEREMAIL
	xpay.Set "LGD_AUTHTYPE", LGD_AUTHTYPE
	xpay.Set "LGD_CARDTYPE", LGD_CARDTYPE
	xpay.Set "LGD_BUYERIP", Request.ServerVariables("REMOTE_ADDR")	'반드시 결제고객의 IP를 넘겨야 함
	
	if LGD_AUTHTYPE = "ISP" then
		xpay.Set "KVP_QUOTA", KVP_QUOTA
		xpay.Set "KVP_NOINT", KVP_NOINT
		xpay.Set "KVP_CARDCODE", KVP_CARDCODE
		xpay.Set "KVP_SESSIONKEY", KVP_SESSIONKEY
		xpay.Set "KVP_ENCDATA", KVP_ENCDATA 
	else
		xpay.Set "LGD_PAN", LGD_PAN
		xpay.Set "LGD_INSTALL", LGD_INSTALL
		xpay.Set "LGD_NOINT", LGD_NOINT
		xpay.Set "LGD_EXPYEAR", LGD_EXPYEAR
		xpay.Set "LGD_EXPMON", LGD_EXPMON
		xpay.Set "VBV_ECI", VBV_ECI
		xpay.Set "VBV_CAVV", VBV_CAVV
		xpay.Set "VBV_XID", VBV_XID
	end if 

    
    if  xpay.TX() then
        '1)결제결과 처리(성공,실패 결과 처리를 하시기 바랍니다.)
        Response.Write("결제요청이 완료되었습니다. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")

	    Response.Write("거래번호 : " & xpay.Response("LGD_TID", 0) & "<br>")
	    Response.Write("상점아이디 : " & xpay.Response("LGD_MID", 0) & "<br>")
	    Response.Write("상점주문번호 : " & xpay.Response("LGD_OID", 0) & "<br>")
	    Response.Write("결제금액 : " & xpay.Response("LGD_AMOUNT", 0) & "<br>")
	    Response.Write("결과코드 : " & xpay.Response("LGD_RESPCODE", 0) & "<br>")
	    Response.Write("결과메세지 : " & xpay.Response("LGD_RESPMSG", 0) & "<p>")

        Response.Write("[결제요청 결과 파라미터]<br>")

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
        
        Response.Write("<p>")
        
        if xpay.resCode = "0000" then
        	'최종결제요청 결과 성공 DB처리
        	Response.Write("최종결제요청 결과 성공 DB처리하시기 바랍니다." & "<br>")
        	            	            	
        	'최종결제요청 결과 성공 DB처리 실패시 Rollback 처리
        	isDBOK = true 'DB처리 실패시 false로 변경해 주세요.
        	
        	if isDBOK then
        	else
        		Response.Write("<p>")
        		xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
        		
                Response.Write("TX Rollback Response_code = " & xpay.resCode & "<br>")
                Response.Write("TX Rollback Response_msg = " & xpay.resMsg & "<p>")
        		
                if "0000" = xpay.resCode then
                	Response.Write("자동취소가 정상적으로 완료 되었습니다.<br>")
                else
                	Response.Write("자동취소가 정상적으로 처리되지 않았습니다.<br>")
                end if
        	end if            	
        else
        	'결제결제요청 결과 실패 DB처리
        	Response.Write("결제결제요청 결과 실패 DB처리하시기 바랍니다." & "<br>")
        end if
    else
        '2)API 요청실패 처리
        Response.Write("결제요청이 실패하였습니다. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
        
        '결제요청 결과 실패 상점 DB처리
        Response.Write("결제결제요청 결과 실패 DB처리하시기 바랍니다." & "<br>")
    end if 
%>
