<%
Const C_HIGH_VERSION = "04"  '' 버전 1
Const C_LOW_VERSION  = "00"  '' 버전 2
Const C_CPCO_ID      = "RC0777"  '' 제휴사 코드

dim C_ERRCodeList , C_ErrCodeName
C_ERRCodeList = Array( "0000","1110" _
,"3110","3111","3112","3113","3114","3115","3116","3117","3118","3119","3120","3121" _
,"5001","5013","5014","5015","5016","5021","5022","5031","5032","5033","5034","5041" _
,"5051","5052","5061","5062","5071","5072","5081","5082","5091","5101","5102","5103" _
,"5201","5300","5301","5303","5304","5309","5311","5999" _
,"6660","6300","6630","6999","6810","6800","6900","6720" _
)

C_ErrCodeName = Array( "정상","전문오류 시스템실 연락바람" _
,"기프티콘 연동규격 버전이 일치하지 않습니다","쿠폰번호가 일치하지 않습니다","이 쿠폰은 본 매장에서 사용할 수 없습니다","쿠폰번호인증을 실패했습니다","쿠폰이 유효하지 않습니다","이미 사용된 쿠폰 입니다.","사용기간이 경과되어 사용할 수 없습니다","사용되지 않는 쿠폰입니다","이미 교환취소되었습니다","교환 당일이 아니면 반품이 불가능합니다","수신번호가 일치하지 않습니다","이미 사용한 쿠폰입니다" _
,"등록되지 않은 IP 입니다.","연동발송 ID값이 없습니다.","캠페인 ID값이 올바르지 않습니다.","캠페인 ID값이 존재하지 않습니다.","캠페인에 해당하는 연동 ID값이 올바르지 않습니다.","상품구분 값이 없습니다.","잘못된 상품구분 값입니다.","상품 ID값이 없습니다.","상품이 존재하지 않습니다.","상품 유효기간이 만료 되었습니다.","캠페인에 해당하는 상품이 아닙니다.","상품수량 값이 없습니다." _
,"MDN(수신자 번호)값이 없습니다.","MDN(수신자 번호)가 올바르지 않습니다.","SMS 메시지가 없습니다.","SMS 메시지 최대 값을 초과했습니다.","무선 안내페이지 값이 없습니다.","해당하는 무선 안내페이지가 존재하지 않습니다.","기프티콘 꾸미기 이미지 번호가 없습니다.","해당하는 기프티콘 꾸미기 이미지 값이 존재하지 않습니다.","기프티콘 꾸미기 메시지 최대 값을 초과했습니다.","고객식별ID 최대 값을 초과했습니다","회신번호 최대 값을 초과했습니다","TR_ID의 최대 값 50Byte를 초과했습니다." _
,"SMS_TYPE 지정 형식이 올바르지 않습니다.","파라미터 오류 입니다.","판매상품ID 매핑 데이터 없음.","꾸미기 번호 매핑 데이터 없음.","판매상품 할인 금액이 없습니다.","중복된 TR_ID 값 입니다.","초과된 상품 수량 요청입니다.","주문 연동 실패" _
,"취소 실패","DB 에러","기한 만료","네트워크 에러","매치된 데이터 없음","파라메터 에러","서버 에러","상태 정보 오류" _
)

''-------------------------------------------------------
function getErrCode2Name(iErrCode)
    Dim i
    if IsNULL(iErrCode) then Exit function

    for i=0 to UBound(C_ERRCodeList)
        if (C_ERRCodeList(i)=iErrCode) then
            getErrCode2Name = C_ErrCodeName(i)
            Exit For
        end if
    next

end function

function getNByteStr(orgBytes,stN,lenN)
    Dim i,s
    Dim byteLen
    If Not IsArray(orgBytes) then Exit function
    byteLen=Ubound(orgBytes)
    if (byteLen<stN+lenN-1) then Exit function

    For i=stN To stN+lenN-1
        s = s & Chr(AscB(MidB(orgBytes, i, 1)))
    Next
    getNByteStr = s ''Replace(s," ",".")
end function

function getNByteStrW(orgBytes,stN,lenN)
    Dim i,s
    Dim byteLen
    If Not IsArray(orgBytes) then Exit function
    byteLen=Ubound(orgBytes)
    if (byteLen<stN+lenN-1) then Exit function

    For i=stN To stN+lenN-1
        ''rw HEX(AscB(MidB(orgBytes, i , 1))) & "==" & CLNG("&H" & HEX(AscB(MidB(orgBytes, i , 1))) )
        If AscB(MidB(orgBytes, i , 1))>127 THEN
            'rw HEX(AscB(MidB(orgBytes, i , 1))) & HEX(AscB(MidB(orgBytes, i+1 , 1))) & "==" & CLNG("&H" & HEX(AscB(MidB(orgBytes, i , 1))) & HEX(AscB(MidB(orgBytes, i+1 , 1))) )
            ''s = s & Chr(AscB(MidB(orgBytes, i, 1)))
            s = s & Chr("&H" & HEX(AscB(MidB(orgBytes, i , 1))) & HEX(AscB(MidB(orgBytes, i+1 , 1))))
            i=i+1
        ELSE
            ''rw HEX(AscB(MidB(orgBytes, i , 1))) & "==" & CLNG("&H" & HEX(AscB(MidB(orgBytes, i , 1))) )
            ''s = s & Chr(AscB(MidB(orgBytes, i, 1)))
            s = s & Chr(AscB(MidB(orgBytes, i, 1))) ''Chr("&H" & HEX(AscB(MidB(orgBytes, i , 1))) )
        END IF
    Next
    getNByteStrW = s ''Replace(s," ",".")
end function

function getNByteLng(orgBytes,stN,lenN)
    Dim i,s
    Dim byteLen
    If Not IsArray(orgBytes) then Exit function
    byteLen=Ubound(orgBytes)
    if (byteLen<stN+lenN-1) then Exit function

    For i=stN To stN+lenN-1
        s = s + 16^((stN+lenN-i-1)*2)*AscB(MidB(orgBytes,i,1))
    Next
    getNByteLng = s
end function

function getByteLength(oStr)
    dim i,ln
    ln =0
    for i=0 to Len(oStr)-1
        if (ASC(MID(oStr,i+1,1))<0) Then
            ln = ln+2
        else
            ln = ln + 1
        end if
    next
    getByteLength = ln
end function

Function MakeRightBalnkChar(orgData,MaxLen)
    Dim i, Ret
    Dim orgLen
    orgLen = getByteLength(orgData)
    ''response.write  "orgLen="&orgLen
    IF (orgLen>MaxLen) then
        Ret = LeftB(orgData,MaxLen)   ''수정요망
    Else
        Ret = orgData
        for i=0 to MaxLen-orgLen-1
            Ret = Ret & " "
        next
    End IF

    MakeRightBalnkChar = Ret
End Function


function DecTo4ByteChar(idecimal)
    DecTo4ByteChar = Hex2ByteArray(Dec2Hex(idecimal,4))
end function

function Hex2ByteArray(iHexa)
    dim i, ret
    for i=0 to Len(iHexa)-1
        ret = ret & CHR("&H"&Mid(iHexa,i+1,2))
        i=i+1
    next
    Hex2ByteArray = ret
end function

function Dec2Hex(decVal,nbyte)
    dim iHexa, i, buf
    iHexa = HEX(decVal)

    ''Fill Zero
    for i=0 to nbyte*2-1
        buf = buf & "0"
    next

    Dec2Hex = Left(buf,Len(buf)-Len(CStr(iHexa)))&iHexa
end function

sub dPByteArrayDEcimal(byteArray)
    Dim d, i
    d = ""
    For i = 1 To LenB(byteArray)
        d = d & CStr(AscB(MidB(byteArray, i, 1))) & ","
    Next
    Response.Write "<p>" & d & "</p>"
end sub

Function MakeDefaultParam(svcCode,iCouponNo,iTraceNum)
    dim iheader, Param1, Param2
    Dim NowYYYYMMDD : NowYYYYMMDD = Replace(Left(now(),10),"-","")
    Dim NowHHNNSS   : NowHHNNSS   = Replace(FormatDateTime(time,4),":","") + Right(FormatDateTime(time,3),2)

    set iheader = New CGiftiConCommonHeader
    iheader.FSERVICE_CODE = svcCode
    iheader.FTRACE_NUMBER = iTraceNum
    iheader.FTRANS_DATE   = NowYYYYMMDD
    iheader.FTRANS_TIME   = NowHHNNSS
    IF svcCode="P100" THEN
        iheader.FBODY_LENGTH=DecTo4ByteChar(255)  ''조회시 255 Byte
        ''iheader.FBODY_LENGTH= Chr("&H00")&Chr("&H00")&Chr("&H00")&Chr("&HFF")
    ELSEIF svcCode="P110" THEN
        iheader.FBODY_LENGTH=DecTo4ByteChar(343)  ''승인시 343 Byte
    ELSEIF svcCode="P120" THEN
        iheader.FBODY_LENGTH=DecTo4ByteChar(343)  ''취소시 343 Byte
    ENd IF
    Param1 = iheader.MakeParamString
    set iheader = Nothing

    set iheader = New CGiftiConCommonBody
    iheader.FCOUPON_NUMBER    = iCouponNo
    iheader.FPOS_REQUEST_DATE = NowYYYYMMDD
    iheader.FPOS_REQUEST_TIME = NowHHNNSS

    Param2 = iheader.MakeParamString
    set iheader = Nothing

    MakeDefaultParam = Param1 & Param2
End Function

''공통해더
CLASS CGiftiConCommonHeader
    public FSERVICE_CODE    '' Char(4)  '' 서비스 전문번호 P100:조회/P101:조회응답, P110:승인/P111:승인응답, P120:승인취소/P121:승인취소응, P130:망상취소/P131:망상취소응답
    public FHIGH_VERSION    '' Char(2)  '' 버전 1
    public FLOW_VERSION     '' Char(2)  '' 버전 2
    public FORG_CODE        '' Char(4)  '' 기관코드
    public FTRANS_DATE      '' Char(8)  '' 전송일자 YYYYMMDD
    public FTRANS_TIME      '' Char(6)  '' 전송시간 HHNNSS
    public FTRACE_NUMBER    '' Char(10) '' 추적번호
    public FBODY_LENGTH     '' 4Byte Int ''바디길이 (전송시 255) ::확인.
    public FERROR_CDOE_1    '' Char(2)
    public FERROR_CDOE_2    '' Char(2)
    public FHD_FILER        '' Char(10) '' 예비

    public Function MakeParamString
        Dim Ret
        Ret = FSERVICE_CODE
        Ret = Ret & FHIGH_VERSION
        Ret = Ret & FLOW_VERSION
        Ret = Ret & FORG_CODE
        Ret = Ret & FTRANS_DATE
        Ret = Ret & FTRANS_TIME
        Ret = Ret & MakeRightBalnkChar(FTRACE_NUMBER,10)
        Ret = Ret & FBODY_LENGTH
        Ret = Ret & FERROR_CDOE_1
        Ret = Ret & FERROR_CDOE_2
        Ret = Ret & FHD_FILER

        MakeParamString = ret
    End function

    Private Sub Class_Initialize()
        FHIGH_VERSION = C_HIGH_VERSION
        FLOW_VERSION  = C_LOW_VERSION
        FORG_CODE     = "    "
        FERROR_CDOE_1 = "  "
        FERROR_CDOE_2 = "  "
        FHD_FILER     = "          "
        FBODY_LENGTH  = Chr("&H00")&Chr("&H00")&Chr("&H00")&Chr("&HFF") ''"    " ''
	End Sub

	Private Sub Class_Terminate()

	End Sub
ENd CLASS

''공통Body
CLASS CGiftiConCommonBody
    public FCPCO_ID             '' Char(6)  '' 제휴사코드
    public FFRANCHISE_ID        '' Char(10) '' 가맹점코드 (자체)
    public FFRANCHISE_NAME      '' Char(80) '' 가맹점명
    public FPOS_ID              '' Char(16) '' 포스번호 (자체)
    public FPOS_REQUEST_DATE    '' Char(8)  '' POS상조회일자 YYYYMMDD
    public FPOS_REQUEST_TIME    '' Char(6)  '' POS상조회시각 HHNNSS
    public FCOUPON_NUMBER       '' Char(12)  '' 쿠폰번호
    public FBARCODE_SCAN        '' Char(1)  ‘0’: barcode scan, ‘1’: key in
    public FSECURE_MOD          '' Char(1)  ‘0’: 사용, ‘1’: 미사용
    public FRECEIVER_MDN        '' Char(11) '' 수신자번호 FSECURE_MOD ‘1’인 경우 space

    public Function MakeParamString
        Dim Ret
        Ret = FCPCO_ID
        Ret = Ret & FFRANCHISE_ID
        Ret = Ret & FFRANCHISE_NAME
        Ret = Ret & FPOS_ID
        Ret = Ret & FPOS_REQUEST_DATE
        Ret = Ret & FPOS_REQUEST_TIME
        Ret = Ret & MakeRightBalnkChar(FCOUPON_NUMBER,12)
        Ret = Ret & FBARCODE_SCAN
        Ret = Ret & FSECURE_MOD
        Ret = Ret & FRECEIVER_MDN

        MakeParamString = ret
    End function

    Private Sub Class_Initialize()
        FCPCO_ID = C_CPCO_ID
        FFRANCHISE_ID = "TENONLINE "
        FFRANCHISE_NAME = "TENBYTEN 온라인                                                                 "
        FPOS_ID  = "10000           "
        FBARCODE_SCAN = "1"
        FSECURE_MOD   = "1"
        FRECEIVER_MDN = "           "
	End Sub

	Private Sub Class_Terminate()

	End Sub
ENd CLASS

CLASS CGiftiConResult
    private FRectReceivedBites
    public FSERVICE_CODE
    public FTRANS_DATE
    public FTRANS_TIME
    public FTRACE_NUMBER
    public FBODY_LENGTH
    public FERROR_CDOE_1
    public FERROR_CDOE_2
    public FCOUPON_NUMBER
    public FMESSAGE
    public FEXCHANGE_COUNT

    ''' 여러개 가능.
    public FSubItemCode
    public FSubItemBarCode
    public FSubItemEa
    public FSubSupplyID
    public FSubSupplyPrice
    public FSubPartnerCharge
    public FSubSupplyerCharge
    public FSubSubItemType
    public FSubLimitPrice
    public FSubDiscountPrice
    public FSubNotice
    public FSubFiller

    public FApprovNO        ''승인번호
    public FExchangePrice   ''상품교환가

    public function getResultCode
        getResultCode = FERROR_CDOE_1 & FERROR_CDOE_2
    end function

    public function getResultStr
        getResultStr = getErrCode2Name(FERROR_CDOE_1 & FERROR_CDOE_2)
    end function

    ''상품가격 FSubItemEa 확인.
    public function getItemPrice
        If IsNumeric(FSubSupplyPrice) and IsNumeric(FSubPartnerCharge) and IsNumeric(FSubSupplyerCharge) THEN
            getItemPrice = CLNG(FSubSupplyPrice)+CLNG(FSubPartnerCharge)+CLNG(FSubSupplyerCharge)
        else
            getItemPrice = 0
        End IF
    end function

    public function parseResult(irecbytes)
        FRectReceivedBites = irecbytes
        parseResult = False

        FSERVICE_CODE = getNByteStr(FRectReceivedBites,1,4)
        FTRANS_DATE   = getNByteStr(FRectReceivedBites,13,8)
        FTRANS_TIME   = getNByteStr(FRectReceivedBites,21,6)
        FTRACE_NUMBER = getNByteStr(FRectReceivedBites,27,10)
        FBODY_LENGTH  = getNByteLng(FRectReceivedBites,37,4)
        FERROR_CDOE_1 = getNByteStr(FRectReceivedBites,41,2)
        FERROR_CDOE_2 = getNByteStr(FRectReceivedBites,43,2)

        FCOUPON_NUMBER  = getNByteStr(FRectReceivedBites,181,12)

        IF (FSERVICE_CODE="P101") THEN ''조회응답
            FMESSAGE        = getNByteStrW(FRectReceivedBites,310,64)
            FEXCHANGE_COUNT = getNByteLng(FRectReceivedBites,374,4)

            '''교환 상품이 여러개 일 수 있음.. // 정책적으로 결정.. ==> 단일 상품만 사용.
            IF (FEXCHANGE_COUNT>0) Then
                ''For i=0 to FEXCHANGE_COUNT-1
                FSubItemCode        = getNByteStr(FRectReceivedBites,378,8)
                FSubItemBarCode     = getNByteStr(FRectReceivedBites,386,13)
                FSubItemEa          = getNByteLng(FRectReceivedBites,399,4)
                FSubSupplyID        = getNByteStr(FRectReceivedBites,403,6)
                FSubSupplyPrice     = getNByteLng(FRectReceivedBites,409,4)
                FSubPartnerCharge   = getNByteLng(FRectReceivedBites,413,4)
                FSubSupplyerCharge  = getNByteLng(FRectReceivedBites,417,4)
                FSubSubItemType     = getNByteStr(FRectReceivedBites,421,2)     '01: 일반상품 /02: 상품권(Gifticon을 종이상품권으로)/ 03: 정액 할인/04: 특정 상품 할인
                FSubLimitPrice      = getNByteLng(FRectReceivedBites,423,4)     ''제약사항	N	4	할인권의 경우 최소구매금액
                FSubDiscountPrice   = getNByteLng(FRectReceivedBites,427,4)     ''할인금액
                FSubNotice          = getNByteStrW(FRectReceivedBites,431,100)  ''AN	100	교환주의사항
                FSubFiller          = getNByteStr(FRectReceivedBites,531,50)
                ''Next
            End IF
        ELSEIF (FSERVICE_CODE="P111") or (FSERVICE_CODE="P121") THEN ''승인/취소응답
            FApprovNO       = getNByteStr(FRectReceivedBites,206,20)
            FExchangePrice  = getNByteLng(FRectReceivedBites,226,4)
            FMESSAGE        = getNByteStrW(FRectReceivedBites,230,64)
        ENd IF

        parseResult= true
    end function


    Private Sub Class_Initialize()
        FEXCHANGE_COUNT = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

CLASS CGiftiCon
    private FConIP
    private FConPort
    private FConSSL
    private FmaxWaitMillisec

    public FConResult
    public FLASTERROR

    private function connectSocket(byVal params, byRef receivedBites, byRef RetERROR )
        Dim iSocket, ret1, receiveLen

        set iSocket = Server.CreateObject("Chilkat.Socket")   '''컴포넌트 설치 해야함.
        ret1 = iSocket.UnlockComponent("10X10CSocket_AwvVPpd2JD6l") '''("Anything for 30-day trial")  ''컴포넌트 상용키
        If (ret1 <> 1) Then
            connectSocket = False
            RetERROR = "Failed to unlock component"
            set iSocket=Nothing
            Exit Function
        End If

        ret1 = iSocket.Connect(FConIP,FConPort,FConSSL,FmaxWaitMillisec)
        If (ret1 <> 1) Then
            connectSocket = False
            RetERROR = iSocket.LastErrorText
            set iSocket=Nothing
            Exit Function
        End If

        '  Set maximum timeouts for reading an writing (in millisec)
        iSocket.MaxReadIdleMs = 10000
        iSocket.MaxSendIdleMs = 10000

        iSocket.StringCharset = "euc-kr"   '' "euc-kr"             '''중요
        ret1 = iSocket.SendString(Params)
        If (ret1 <> 1) Then
            connectSocket = False
            RetERROR = iSocket.LastErrorText
            set iSocket=Nothing
            Exit Function
        End If

        receivedBites= iSocket.ReceiveBytes
        receiveLen = LenB(receivedBites)
        ''response.write "receiveLen["&receiveLen&"]"

        set iSocket=Nothing

        if (receiveLen<1) then
            connectSocket = False
            RetERROR = "소켓 연결오류 - [ERR001]"
            Exit Function
        end if

        connectSocket = true
    end function

    ''쿠폰 조회
    public function reqCouponState(byVal iCouponNo, byVal iTraceNum)
        Dim Params
        Dim receivedBites, RetERROR

        Params = MakeDefaultParam("P100",iCouponNo, iTraceNum)
        Params = Params  & MakeRightBalnkChar("",50)

        if (Not connectSocket(params, receivedBites, RetERROR )) then
            reqCouponState = FALSE
            FLASTERROR = RetERROR
            Exit Function
        end if

        set FConResult = new CGiftiConResult
        If (Not FConResult.parseResult(receivedBites)) then
            reqCouponState = False
            FLASTERROR = "parsing Error"
            Exit Function
        end if

        reqCouponState = true
    end function

    ''쿠폰 승인
    public function reqCouponApproval(byVal iCouponNo, byVal iTraceNum, byVal exChangePrice)
        Dim Params
        Dim receivedBites, RetERROR

        Params = MakeDefaultParam("P110",iCouponNo, iTraceNum)
        Params = Params  & MakeRightBalnkChar("",20)        ''승인번호
        Params = Params  & DecTo4ByteChar(exChangePrice)    ''가격 4Byte   ''' 상품 교환가격 ExchangePrice 으로 응답
        Params = Params  & MakeRightBalnkChar("",64)        ''응답메세지
        Params = Params  & MakeRightBalnkChar("",50)        ''Filler

        if (Not connectSocket(params, receivedBites, RetERROR )) then
            reqCouponApproval = FALSE
            FLASTERROR = RetERROR
            Exit Function
        end if

        set FConResult = new CGiftiConResult
        If (Not FConResult.parseResult(receivedBites)) then
            reqCouponApproval = False
            FLASTERROR = "parsing Error"
            Exit Function
        end if

        reqCouponApproval = true
    end function

    public function reqCouponCancel(byVal iCouponNo, byVal iTraceNum, byVal exChangePrice)
        Dim Params
        Dim receivedBites, RetERROR

        Params = MakeDefaultParam("P120",iCouponNo, iTraceNum)
        Params = Params  & MakeRightBalnkChar("",20)        ''승인번호
        Params = Params  & DecTo4ByteChar(exChangePrice)    ''가격 4Byte   ''' 상품 교환가격 ExchangePrice 으로 응답
        Params = Params  & MakeRightBalnkChar("",64)        ''응답메세지
        Params = Params  & MakeRightBalnkChar("",50)        ''Filler

        if (Not connectSocket(params, receivedBites, RetERROR )) then
            reqCouponCancel = FALSE
            FLASTERROR = RetERROR
            Exit Function
        end if

        set FConResult = new CGiftiConResult
        If (Not FConResult.parseResult(receivedBites)) then
            reqCouponCancel = False
            FLASTERROR = "parsing Error"
            Exit Function
        end if

        reqCouponCancel = true
    end function


    Private Sub Class_Initialize()
        ''테스트 서버	113.217.246.45	9091
        ''상용 서버	    172.28.94.240	9091

        IF (application("Svr_Info")="Dev") THEN
            ''FConIP = "113.217.246.45"
            
            FConIP = "nstgauth.gifticon.com"
        ELSE
            ''FConIP = "172.28.94.240"
            
            FConIP = "auth.gifticon.com"
        END IF
        FConPort = 9091
        FConSSL = 0
        FmaxWaitMillisec = 20000
	End Sub

	Private Sub Class_Terminate()

	End Sub
END CLASS
%>