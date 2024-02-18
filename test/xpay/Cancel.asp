<%
    '/*
    ' * [결제취소 요청 페이지]
    ' *
    ' * LG데이콤으로 부터 내려받은 거래번호(LGD_TID)를 가지고 취소 요청을 합니다.(파라미터 전달시 POST를 사용하세요)
    ' * (승인시 LG데이콤으로 부터 내려받은 PAYKEY와 혼동하지 마세요.)
    ' */
    CST_PLATFORM         = trim(request("CST_PLATFORM"))        ' LG데이콤 결제서비스 선택(test:테스트, service:서비스)
    CST_MID              = trim(request("CST_MID"))             ' LG데이콤으로 부터 발급받으신 상점아이디를 입력하세요.
                                                                ' 테스트 아이디는 't'를 제외하고 입력하세요.
    if CST_PLATFORM = "test" then                               ' 상점아이디(자동생성)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if
    LGD_TID              = trim(request("LGD_TID"))             ' LG데이콤으로 부터 내려받은 거래번호(LGD_TID)

    configPath           = "C:/lgdacom"         				' LG데이콤에서 제공한 환경파일("/conf/lgdacom.conf") 위치 지정.


    Set xpay = server.CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM
    xpay.Init_TX(LGD_MID)

    xpay.Set "LGD_TXNAME", "Cancel"
    xpay.Set "LGD_TID", LGD_TID
 

    '/*
    ' * 1. 결제취소 요청 결과처리
    ' *
    ' * 취소결과 리턴 파라미터는 연동메뉴얼을 참고하시기 바랍니다.
    ' */
    if xpay.TX() then
        '1)결제취소결과 화면처리(성공,실패 결과 처리를 하시기 바랍니다.)
        Response.Write("결제취소 요청이 완료되었습니다. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
    else
        '2)API 요청 실패 화면처리
        Response.Write("결제취소 요청이 실패하였습니다. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
    end if
%>
