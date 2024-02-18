<%
Const C_HIGH_VERSION = "04"  '' ���� 1
Const C_LOW_VERSION  = "00"  '' ���� 2
Const C_CPCO_ID      = "RC0777"  '' ���޻� �ڵ�

dim C_ERRCodeList , C_ErrCodeName
C_ERRCodeList = Array( "0000","1110" _
,"3110","3111","3112","3113","3114","3115","3116","3117","3118","3119","3120","3121" _
,"5001","5013","5014","5015","5016","5021","5022","5031","5032","5033","5034","5041" _
,"5051","5052","5061","5062","5071","5072","5081","5082","5091","5101","5102","5103" _
,"5201","5300","5301","5303","5304","5309","5311","5999" _
,"6660","6300","6630","6999","6810","6800","6900","6720" _
)

C_ErrCodeName = Array( "����","�������� �ý��۽� �����ٶ�" _
,"����Ƽ�� �����԰� ������ ��ġ���� �ʽ��ϴ�","������ȣ�� ��ġ���� �ʽ��ϴ�","�� ������ �� ���忡�� ����� �� �����ϴ�","������ȣ������ �����߽��ϴ�","������ ��ȿ���� �ʽ��ϴ�","�̹� ���� ���� �Դϴ�.","���Ⱓ�� ����Ǿ� ����� �� �����ϴ�","������ �ʴ� �����Դϴ�","�̹� ��ȯ��ҵǾ����ϴ�","��ȯ ������ �ƴϸ� ��ǰ�� �Ұ����մϴ�","���Ź�ȣ�� ��ġ���� �ʽ��ϴ�","�̹� ����� �����Դϴ�" _
,"��ϵ��� ���� IP �Դϴ�.","�����߼� ID���� �����ϴ�.","ķ���� ID���� �ùٸ��� �ʽ��ϴ�.","ķ���� ID���� �������� �ʽ��ϴ�.","ķ���ο� �ش��ϴ� ���� ID���� �ùٸ��� �ʽ��ϴ�.","��ǰ���� ���� �����ϴ�.","�߸��� ��ǰ���� ���Դϴ�.","��ǰ ID���� �����ϴ�.","��ǰ�� �������� �ʽ��ϴ�.","��ǰ ��ȿ�Ⱓ�� ���� �Ǿ����ϴ�.","ķ���ο� �ش��ϴ� ��ǰ�� �ƴմϴ�.","��ǰ���� ���� �����ϴ�." _
,"MDN(������ ��ȣ)���� �����ϴ�.","MDN(������ ��ȣ)�� �ùٸ��� �ʽ��ϴ�.","SMS �޽����� �����ϴ�.","SMS �޽��� �ִ� ���� �ʰ��߽��ϴ�.","���� �ȳ������� ���� �����ϴ�.","�ش��ϴ� ���� �ȳ��������� �������� �ʽ��ϴ�.","����Ƽ�� �ٹ̱� �̹��� ��ȣ�� �����ϴ�.","�ش��ϴ� ����Ƽ�� �ٹ̱� �̹��� ���� �������� �ʽ��ϴ�.","����Ƽ�� �ٹ̱� �޽��� �ִ� ���� �ʰ��߽��ϴ�.","���ĺ�ID �ִ� ���� �ʰ��߽��ϴ�","ȸ�Ź�ȣ �ִ� ���� �ʰ��߽��ϴ�","TR_ID�� �ִ� �� 50Byte�� �ʰ��߽��ϴ�." _
,"SMS_TYPE ���� ������ �ùٸ��� �ʽ��ϴ�.","�Ķ���� ���� �Դϴ�.","�ǸŻ�ǰID ���� ������ ����.","�ٹ̱� ��ȣ ���� ������ ����.","�ǸŻ�ǰ ���� �ݾ��� �����ϴ�.","�ߺ��� TR_ID �� �Դϴ�.","�ʰ��� ��ǰ ���� ��û�Դϴ�.","�ֹ� ���� ����" _
,"��� ����","DB ����","���� ����","��Ʈ��ũ ����","��ġ�� ������ ����","�Ķ���� ����","���� ����","���� ���� ����" _
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
        Ret = LeftB(orgData,MaxLen)   ''�������
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
        iheader.FBODY_LENGTH=DecTo4ByteChar(255)  ''��ȸ�� 255 Byte
        ''iheader.FBODY_LENGTH= Chr("&H00")&Chr("&H00")&Chr("&H00")&Chr("&HFF")
    ELSEIF svcCode="P110" THEN
        iheader.FBODY_LENGTH=DecTo4ByteChar(343)  ''���ν� 343 Byte
    ELSEIF svcCode="P120" THEN
        iheader.FBODY_LENGTH=DecTo4ByteChar(343)  ''��ҽ� 343 Byte
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

''�����ش�
CLASS CGiftiConCommonHeader
    public FSERVICE_CODE    '' Char(4)  '' ���� ������ȣ P100:��ȸ/P101:��ȸ����, P110:����/P111:��������, P120:�������/P121:���������, P130:�������/P131:�����������
    public FHIGH_VERSION    '' Char(2)  '' ���� 1
    public FLOW_VERSION     '' Char(2)  '' ���� 2
    public FORG_CODE        '' Char(4)  '' ����ڵ�
    public FTRANS_DATE      '' Char(8)  '' �������� YYYYMMDD
    public FTRANS_TIME      '' Char(6)  '' ���۽ð� HHNNSS
    public FTRACE_NUMBER    '' Char(10) '' ������ȣ
    public FBODY_LENGTH     '' 4Byte Int ''�ٵ���� (���۽� 255) ::Ȯ��.
    public FERROR_CDOE_1    '' Char(2)
    public FERROR_CDOE_2    '' Char(2)
    public FHD_FILER        '' Char(10) '' ����

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

''����Body
CLASS CGiftiConCommonBody
    public FCPCO_ID             '' Char(6)  '' ���޻��ڵ�
    public FFRANCHISE_ID        '' Char(10) '' �������ڵ� (��ü)
    public FFRANCHISE_NAME      '' Char(80) '' ��������
    public FPOS_ID              '' Char(16) '' ������ȣ (��ü)
    public FPOS_REQUEST_DATE    '' Char(8)  '' POS����ȸ���� YYYYMMDD
    public FPOS_REQUEST_TIME    '' Char(6)  '' POS����ȸ�ð� HHNNSS
    public FCOUPON_NUMBER       '' Char(12)  '' ������ȣ
    public FBARCODE_SCAN        '' Char(1)  ��0��: barcode scan, ��1��: key in
    public FSECURE_MOD          '' Char(1)  ��0��: ���, ��1��: �̻��
    public FRECEIVER_MDN        '' Char(11) '' �����ڹ�ȣ FSECURE_MOD ��1���� ��� space

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
        FFRANCHISE_NAME = "TENBYTEN �¶���                                                                 "
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

    ''' ������ ����.
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

    public FApprovNO        ''���ι�ȣ
    public FExchangePrice   ''��ǰ��ȯ��

    public function getResultCode
        getResultCode = FERROR_CDOE_1 & FERROR_CDOE_2
    end function

    public function getResultStr
        getResultStr = getErrCode2Name(FERROR_CDOE_1 & FERROR_CDOE_2)
    end function

    ''��ǰ���� FSubItemEa Ȯ��.
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

        IF (FSERVICE_CODE="P101") THEN ''��ȸ����
            FMESSAGE        = getNByteStrW(FRectReceivedBites,310,64)
            FEXCHANGE_COUNT = getNByteLng(FRectReceivedBites,374,4)

            '''��ȯ ��ǰ�� ������ �� �� ����.. // ��å������ ����.. ==> ���� ��ǰ�� ���.
            IF (FEXCHANGE_COUNT>0) Then
                ''For i=0 to FEXCHANGE_COUNT-1
                FSubItemCode        = getNByteStr(FRectReceivedBites,378,8)
                FSubItemBarCode     = getNByteStr(FRectReceivedBites,386,13)
                FSubItemEa          = getNByteLng(FRectReceivedBites,399,4)
                FSubSupplyID        = getNByteStr(FRectReceivedBites,403,6)
                FSubSupplyPrice     = getNByteLng(FRectReceivedBites,409,4)
                FSubPartnerCharge   = getNByteLng(FRectReceivedBites,413,4)
                FSubSupplyerCharge  = getNByteLng(FRectReceivedBites,417,4)
                FSubSubItemType     = getNByteStr(FRectReceivedBites,421,2)     '01: �Ϲݻ�ǰ /02: ��ǰ��(Gifticon�� ���̻�ǰ������)/ 03: ���� ����/04: Ư�� ��ǰ ����
                FSubLimitPrice      = getNByteLng(FRectReceivedBites,423,4)     ''�������	N	4	���α��� ��� �ּұ��űݾ�
                FSubDiscountPrice   = getNByteLng(FRectReceivedBites,427,4)     ''���αݾ�
                FSubNotice          = getNByteStrW(FRectReceivedBites,431,100)  ''AN	100	��ȯ���ǻ���
                FSubFiller          = getNByteStr(FRectReceivedBites,531,50)
                ''Next
            End IF
        ELSEIF (FSERVICE_CODE="P111") or (FSERVICE_CODE="P121") THEN ''����/�������
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

        set iSocket = Server.CreateObject("Chilkat.Socket")   '''������Ʈ ��ġ �ؾ���.
        ret1 = iSocket.UnlockComponent("10X10CSocket_AwvVPpd2JD6l") '''("Anything for 30-day trial")  ''������Ʈ ���Ű
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

        iSocket.StringCharset = "euc-kr"   '' "euc-kr"             '''�߿�
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
            RetERROR = "���� ������� - [ERR001]"
            Exit Function
        end if

        connectSocket = true
    end function

    ''���� ��ȸ
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

    ''���� ����
    public function reqCouponApproval(byVal iCouponNo, byVal iTraceNum, byVal exChangePrice)
        Dim Params
        Dim receivedBites, RetERROR

        Params = MakeDefaultParam("P110",iCouponNo, iTraceNum)
        Params = Params  & MakeRightBalnkChar("",20)        ''���ι�ȣ
        Params = Params  & DecTo4ByteChar(exChangePrice)    ''���� 4Byte   ''' ��ǰ ��ȯ���� ExchangePrice ���� ����
        Params = Params  & MakeRightBalnkChar("",64)        ''����޼���
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
        Params = Params  & MakeRightBalnkChar("",20)        ''���ι�ȣ
        Params = Params  & DecTo4ByteChar(exChangePrice)    ''���� 4Byte   ''' ��ǰ ��ȯ���� ExchangePrice ���� ����
        Params = Params  & MakeRightBalnkChar("",64)        ''����޼���
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
        ''�׽�Ʈ ����	113.217.246.45	9091
        ''��� ����	    172.28.94.240	9091

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