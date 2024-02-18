<%
    '/*
    ' * [최종결제요청 페이지(STEP2-2)]
    ' *
    ' * LG데이콤으로 부터 내려받은 LGD_PAYKEY(인증Key)를 가지고 최종 결제요청.(파라미터 전달시 POST를 사용하세요)
    ' */

    'configPath = "C:/lgdacom/conf/" & "thefingers" & "/conf/" 'LG데이콤에서 제공한 환경파일("/conf/lgdacom.conf, /conf/mall.conf") 위치 지정.  
    ''configPath = "C:/lgdacom/conf/thefingers/"
    configPath = "C:/lgdacom/conf/youareagirl/"
    
    '/*
    ' *************************************************
    ' * 1.최종결제 요청 - BEGIN
    ' *  (단, 최종 금액체크를 원하시는 경우 금액체크 부분 주석을 제거 하시면 됩니다.)
    ' *************************************************
    ' */
response.write "LGD_OID=" + request("LGD_OID")
response.write "LGD_HASHDATA=" + request("LGD_HASHDATA")

    CST_PLATFORM               = trim(request("CST_PLATFORM"))
    CST_MID                    = trim(request("CST_MID"))
    if CST_PLATFORM = "test" then
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if
    LGD_PAYKEY                 = trim(request("LGD_PAYKEY"))

    Dim xpay            '결제요청 API 객체
    Dim amount_check    '금액비교 결과
    Dim i, j
    Dim itemName

	'해당 API를 사용하기 위해 setup.exe 를 설치해야 합니다.
    Set xpay = server.CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM

    xpay.Init_TX(LGD_MID)
    xpay.Set "LGD_TXNAME", "PaymentByKey"
    xpay.Set "LGD_PAYKEY", LGD_PAYKEY
    
response.write "LGD_PAYKEY="&LGD_PAYKEY
    
    '금액을 체크하시기 원하는 경우 아래 주석을 풀어서 이용하십시요.
	'DB_AMOUNT = "DB나 세션에서 가져온 금액" 	'반드시 위변조가 불가능한 곳(DB나 세션)에서 금액을 가져오십시요.
	''xpay.Set "LGD_AMOUNTCHECKYN", "Y"
	''xpay.Set "LGD_AMOUNT", 1000
	
	''주문번호 세팅 테스트 :: 안됨..
    ''xpay.Set "LGD_OID", "2010011100007" xpay.Set "LGD_OID", "2010011100007" 
    
    '/*
    ' *************************************************
    ' * 1.최종결제 요청(수정하지 마세요) - END
    ' *************************************************
    ' */

    '/*
    ' * 2. 최종결제 요청 결과처리
    ' *
    ' * 최종 결제요청 결과 리턴 파라미터는 연동메뉴얼을 참고하시기 바랍니다.
    ' */

    if  xpay.TX() then
        '1)결제결과 화면처리(성공,실패 결과 처리를 하시기 바랍니다.)
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
        '2)API 요청실패 화면처리
        Response.Write("결제요청이 실패하였습니다. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
            
        '결제요청 결과 실패 상점 DB처리
        Response.Write("결제결제요청 결과 실패 DB처리하시기 바랍니다." & "<br>")
        
        Response.Write("거래번호 : " & xpay.Response("LGD_TID", 0) & "<br>")
    end if
 %>
