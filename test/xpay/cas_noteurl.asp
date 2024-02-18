<!-- #include file="./lgdacom/md5.asp" -->
<%
    '/*
    ' * [상점 결제결과처리(DB) 페이지]
    ' *
    ' * 1) 위변조 방지를 위한 hashdata값 검증은 반드시 적용하셔야 합니다.
    ' *
    ' */
    LGD_RESPCODE            = trim(request("LGD_RESPCODE"))             '// 응답코드: 0000(성공) 그외 실패
    LGD_RESPMSG             = trim(request("LGD_RESPMSG"))              '// 응답메세지
    LGD_MID                 = trim(request("LGD_MID"))                  '// 상점아이디
    LGD_OID                 = trim(request("LGD_OID"))                  '// 주문번호
    LGD_AMOUNT              = trim(request("LGD_AMOUNT"))               '// 거래금액
    LGD_TID                 = trim(request("LGD_TID"))                  '// 데이콤이 부여한 거래번호
    LGD_PAYTYPE             = trim(request("LGD_PAYTYPE"))              '// 결제수단코드
    LGD_PAYDATE             = trim(request("LGD_PAYDATE"))              '// 거래일시(승인일시/이체일시)
    LGD_HASHDATA            = trim(request("LGD_HASHDATA"))             '// 해쉬값
    LGD_FINANCECODE         = trim(request("LGD_FINANCECODE"))          '// 결제기관코드(은행코드)
    LGD_FINANCENAME         = trim(request("LGD_FINANCENAME"))          '// 결제기관이름(은행이름)
    LGD_ESCROWYN            = trim(request("LGD_ESCROWYN"))             '// 에스크로 적용여부
    LGD_TIMESTAMP           = trim(request("LGD_TIMESTAMP"))            '// 타임스탬프
    LGD_ACCOUNTNUM          = trim(request("LGD_ACCOUNTNUM"))           '// 계좌번호(무통장입금)
    LGD_CASTAMOUNT          = trim(request("LGD_CASTAMOUNT"))           '// 입금총액(무통장입금)
    LGD_CASCAMOUNT          = trim(request("LGD_CASCAMOUNT"))           '// 현입금액(무통장입금)
    LGD_CASFLAG             = trim(request("LGD_CASFLAG"))              '// 무통장입금 플래그(무통장입금) - 'R':계좌할당, 'I':입금, 'C':입금취소
    LGD_CASSEQNO            = trim(request("LGD_CASSEQNO"))             '// 입금순서(무통장입금)
    LGD_CASHRECEIPTNUM      = trim(request("LGD_CASHRECEIPTNUM"))       '// 현금영수증 승인번호
    LGD_CASHRECEIPTSELFYN   = trim(request("LGD_CASHRECEIPTSELFYN"))    '// 현금영수증자진발급제유무 Y: 자진발급제 적용, 그외 : 미적용
    LGD_CASHRECEIPTKIND     = trim(request("LGD_CASHRECEIPTKIND"))      '// 현금영수증 종류 0: 소득공제용 , 1: 지출증빙용
	LGD_PAYER            	= trim(request("LGD_PAYER"))             	'// 입금자명
	
    '/*
    ' * 구매정보
    ' */
    LGD_BUYER               = trim(request("LGD_BUYER"))                '// 구매자
    LGD_PRODUCTINFO         = trim(request("LGD_PRODUCTINFO"))          '// 상품명
    LGD_BUYERID             = trim(request("LGD_BUYERID"))              '// 구매자 ID
    LGD_BUYERADDRESS        = trim(request("LGD_BUYERADDRESS"))         '// 구매자 주소
    LGD_BUYERPHONE          = trim(request("LGD_BUYERPHONE"))           '// 구매자 전화번호
    LGD_BUYEREMAIL          = trim(request("LGD_BUYEREMAIL"))           '// 구매자 이메일
    LGD_BUYERSSN            = trim(request("LGD_BUYERSSN"))             '// 구매자 주민번호
    LGD_PRODUCTCODE         = trim(request("LGD_PRODUCTCODE"))          '// 상품코드
    LGD_RECEIVER            = trim(request("LGD_RECEIVER"))             '// 수취인
    LGD_RECEIVERPHONE       = trim(request("LGD_RECEIVERPHONE"))        '// 수취인 전화번호
    LGD_DELIVERYINFO        = trim(request("LGD_DELIVERYINFO"))         '// 배송지
	
	
    '/*
    ' * hashdata 검증을 위한 mertkey는 상점관리자 -> 계약정보 -> 상점정보관리에서 확인하실수 있습니다.
    ' * LG데이콤에서 발급한 상점키로 반드시변경해 주시기 바랍니다.
    ' */
    LGD_MERTKEY = ""  '//mertkey
    LGD_HASHDATA2 = md5( LGD_MID & LGD_OID & LGD_AMOUNT & LGD_RESPCODE & LGD_TIMESTAMP & LGD_MERTKEY )

    '/*
    ' * 상점 처리결과 리턴메세지
    ' *
    ' * OK  : 상점 처리결과 성공
    ' * 그외 : 상점 처리결과 실패
    ' *
    ' * ※ 주의사항 : 성공시 'OK' 문자이외의 다른문자열이 포함되면 실패처리 되오니 주의하시기 바랍니다.
    ' */
    resultMSG = "결제결과 상점 DB처리(LGD_CASNOTEURL) 결과값을 입력해 주시기 바랍니다."

    if LGD_HASHDATA2 = LGD_HASHDATA then
        '//해쉬값 검증이 성공이면
        if LGD_RESPCODE = "0000" then
            '//결제가 성공이면
            if LGD_CASFLAG = "R" then
                '/*
                ' * 무통장 할당 성공 결과 상점 처리(DB) 부분
                ' * 상점 결과 처리가 정상이면 "OK"
                ' */
                resultMSG = "OK"
            elseif LGD_CASFLAG = "I" then
                '/*
                ' * 무통장 입금 성공 결과 상점 처리(DB) 부분
                ' * 상점 결과 처리가 정상이면 "OK"
                ' */
                resultMSG = "OK"
            elseif LGD_CASFLAG = "C" then
                '/*
                ' * 무통장 입금취소 성공 결과 상점 처리(DB) 부분
                ' * 상점 결과 처리가 정상이면 "OK"
                ' */
                resultMSG = "OK"
            end if
        else
            '//결제가 실패이면
            '/*
            ' * 거래실패 결과 상점 처리(DB) 부분
            ' * 상점결과 처리가 정상이면 "OK"
            ' */
            resultMSG = "OK"
        end if
    else
        '//해쉬값이 검증이 실패이면
        '/*
        ' * hashdata검증 실패 로그를 처리하시기 바랍니다.
        ' */
        resultMSG = "결제결과 상점 DB처리(LGD_CASNOTEURL) 해쉬값 검증이 실패하였습니다."
    end if

    Response.Write(resultMSG)
%>

