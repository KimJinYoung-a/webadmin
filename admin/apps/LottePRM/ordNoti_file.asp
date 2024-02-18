<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LotteiMall/incLotteiMallFunction.asp"-->
<!-- #include virtual="/admin/etc/LotteiMall/lotteiMallcls.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<%
''제휴 주문입력 //아침9시 점심1시 저녁5시쯤 이렇게 3번하고 잇어용
''평일 07:50 ~ 17:50 에만 수신 하도록 변경. (취소 주문등 고려)

dim nowTime : nowTime = now()
dim valid1 : valid1 = CDate(Left(CStr(now()),10)+" 07:50:00")
dim valid2 : valid2 = CDate(Left(CStr(now()),10)+" 17:50:00")

''휴일인지 검색
Dim iwd : iwd = CStr(weekDay(now()))

if (nowTime<valid1) or (nowTime>valid2) or (iwd="1") or (iwd="7") then
    response.write "ERR"
    response.end
end if


''공휴일 수신 안함
Dim sqlStr, isholiday : isholiday = FALSE
sqlStr = " select top 1 * from db_cs.dbo.tbl_Holiday"
sqlStr = sqlStr & " where holiday='"&Left(CStr(nowTime),10)&"'"

rsget.Open sqlStr,dbget,1
IF  not rsget.EOF  then
    isholiday = TRUE
ENd IF
rsget.close

if (isholiday) then
    response.write "ERR"
    response.end
end if


function responseErrXML(iErrMSG)
    dim retXML
    retXML = "<?xml version='1.0' encoding='utf-8'?>"&vbcrlf
    retXML = retXML &"<OrderInfoResult>"&vbcrlf
    retXML = retXML &"<MessageHeader>"&vbcrlf
    retXML = retXML &"	<SENDER>TENBYTEN</SENDER>"&vbcrlf
    retXML = retXML &"	<RECEIVER>LotteH</RECEIVER>"&vbcrlf
    retXML = retXML &"	<DATETIME>"&replace(Left(FormatDateTime(now,0),10),"-","")&" "&Left(FormatDateTime(now,4),5)&Right(FormatDateTime(now,0),3)&"</DATETIME>"&vbcrlf
    retXML = retXML &"	<DOCUMENTID>ORDERINFO</DOCUMENTID>"&vbcrlf
    retXML = retXML &"	<ERROROCCUR>Y</ERROROCCUR>"&vbcrlf
    retXML = retXML &"	<ERRORMESSAGE>"&iErrMSG&"</ERRORMESSAGE>"&vbcrlf
    retXML = retXML &"</MessageHeader>"&vbcrlf
    retXML = retXML &"<MessageBody>"&vbcrlf
    retXML = retXML &"<OrderMasterResult>"&vbcrlf
    retXML = retXML &"	<SHORT_ORDER_NO></SHORT_ORDER_NO>"&vbcrlf
    retXML = retXML &"	<ORDER_DT></ORDER_DT>"&vbcrlf
    retXML = retXML &"<OrderDetailResult>"&vbcrlf
    retXML = retXML &"	<ORDER_NO></ORDER_NO>"&vbcrlf
    retXML = retXML &"	<ORDER_SEQ></ORDER_SEQ>"&vbcrlf
    retXML = retXML &"	<RESULT>Y</RESULT>"&vbcrlf
    retXML = retXML &"	<RESULT_MSG></RESULT_MSG>"&vbcrlf
    retXML = retXML &"</OrderDetailResult>"&vbcrlf
    retXML = retXML &"</OrderMasterResult>"&vbcrlf
    retXML = retXML &"</MessageBody>"&vbcrlf
    retXML = retXML &"</OrderInfoResult>"&vbcrlf

    response.write retXML
end function

Function SimpleBinaryToString(Binary)
  'SimpleBinaryToString converts binary data (VT_UI1 | VT_ARRAY Or MultiByte string)
  'to a string (BSTR) using MultiByte VBS functions
  Dim I, S
  For I = 1 To LenB(Binary)
    S = S & Chr(AscB(MidB(Binary, I, 1)))
  Next
  SimpleBinaryToString = S
End Function

Function RSBinaryToString(xBinary)
  'Antonin Foller, http://www.motobit.com
  'RSBinaryToString converts binary data (VT_UI1 | VT_ARRAY Or MultiByte string)
  'to a string (BSTR) using ADO recordset

  Dim Binary
  'MultiByte data must be converted To VT_UI1 | VT_ARRAY first.
  If vartype(xBinary)=8 Then Binary = MultiByteToBinary(xBinary) Else Binary = xBinary

  Dim RS, LBinary
  Const adLongVarChar = 201
  Set RS = CreateObject("ADODB.Recordset")
  LBinary = LenB(Binary)

  If LBinary>0 Then
    RS.Fields.Append "mBinary", adLongVarChar, LBinary
    RS.Open
    RS.AddNew
      RS("mBinary").AppendChunk Binary
    RS.Update
    RSBinaryToString = RS("mBinary")
  Else
    RSBinaryToString = ""
  End If
End Function

Function MultiByteToBinary(MultiByte)
  ' 2000 Antonin Foller, http://www.motobit.com
  ' MultiByteToBinary converts multibyte string To real binary data (VT_UI1 | VT_ARRAY)
  ' Using recordset
  Dim RS, LMultiByte, Binary
  Const adLongVarBinary = 205
  Set RS = CreateObject("ADODB.Recordset")
  LMultiByte = LenB(MultiByte)
  If LMultiByte>0 Then
    RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
    RS.Open
    RS.AddNew
      RS("mBinary").AppendChunk MultiByte & ChrB(0)
    RS.Update
    Binary = RS("mBinary").GetChunk(LMultiByte)
  End If
  Set RS = Nothing
  MultiByteToBinary = Binary
End Function

'Dim buf1
'buf1 = BinaryToText(Request.BinaryRead(Request.TotalBytes), "utf-8")  ''한글 깨짐.
''''buf1 = SimpleBinaryToString(Request.BinaryRead(Request.TotalBytes))
''buf1 = RSBinaryToString(Request.BinaryRead(Request.TotalBytes))  ''한글 일부 깨짐.
'buf1 = replace(buf1,"&","")
'buf1 = ReplaceText(buf1,"(배송전연락주시고부재시경비실에맡겨주세요)[\s\S]*(부탁해요~)","배송전연락주시고부재시경비실에맡겨주세요")
'CALL XMLFileSave(buf1,"IMSI",1)

Dim buf1
dim fs,objFile, ipath
ipath = "C:\home\cube1010\admin2009scm\admin\etc\LotteiMall\xmlFiles\2013-07-01\ORD_2013-07-01_55812.14_0.xml"



Set fs = Server.CreateObject("Scripting.FileSystemObject")
Set objFile = fs.OpenTextFile(ipath)
buf1 = objFile.readAll()
Set objFile =Nothing
Set fs =Nothing


Dim reqBufALL
Dim xmlDOM, listItem, bufText, retXML, retERRXML, OrderAlreadyInputed
set xmlDOM = CreateObject("MSXML2.DomDocument.3.0")
xmlDOM.async = False
xmlDOM.LoadXML buf1
''xmlDOM.Load Request ''바이너리 데이터 사용시

''xmlDOM.LoadXML getORDNotiSampleXML '' xmlDOM.LoadXML getORDNotiSampleXML
''CALL XMLFileSave(xmlDOM.text,"ORD",1)
''아래 구문 추가 & 가 있는경우 CData로 안 묶어서 오류남..
''reqBufALL = replace(xmlDOM.text,"&","")
''xmlDOM.LoadXML reqBufALL

Dim refip : refip = request.ServerVariables("REMOTE_ADDR")
Dim ErrMSG


dim buf
buf = xmlDOM.xml ''xmlDOM.text


'' 임시저장.
CALL XMLFileSave(buf,"ORD",0)


Dim i
Dim ORDER_NO_ARR, ERR_ORDER_NO_ARR
Dim ORDER_NO,ORDER_SEQ,ORDER_DT,PAY_DT,GOODS_ID,GOODS_NAME,ENTP_DT_CODE,GOODSDT_INFO
Dim O_NAME,O_TEL,O_HTEL,O_EMAIL,S_NAME,S_TEL,S_HTEL,S_POST,S_ADDR,CS_MSG,QTY,SALE_PRICE
Dim DELY_TYPE,DELY_COST
Dim OrderEntryLine,SubNodes
Dim DATETIME,DOCUMENTID
On Error resume Next
DOCUMENTID  = xmlDOM.getElementsByTagName("DOCUMENTID").item(0).text
If (ERR) then
    CALL responseErrXML("XML parsing ERROR : NO DOCUMENTID")
    CALL SendMail("webserver@10x10.co.kr", "kjy8517@10x10.co.kr", "[롯데iMall]주문수신오류"&now(),Err.description)
    response.end
End IF
DATETIME    = xmlDOM.getElementsByTagName("DATETIME").item(0).text
If (ERR) then
    CALL responseErrXML("XML parsing ERROR : NO DATETIME")
    response.end
End IF
buf = DATETIME
DATETIME    = Left(buf,4)&"-"&MID(buf,5,2)&"-"&MID(buf,7,2)&" "&Right(buf,8)
Set OrderEntryLine = xmlDOM.getElementsByTagName("OrderEntryLineItem")

If (ERR) then
    CALL responseErrXML("XML parsing ERROR : NO OrderEntryLineItem")
    response.end
End IF
On Error  Goto 0

'''무조건 1건씩 날라옴..
Dim enti : enti=0
for each SubNodes in OrderEntryLine
    enti=enti+1
    ORDER_NO    = Trim(SubNodes.getElementsByTagName("ORDER_NO").item(0).text)                  ''여러주문건이 한번에 오는지 검토 필요..
    ORDER_SEQ   = Trim(SubNodes.getElementsByTagName("ORDER_SEQ").item(0).text)
    ORDER_DT    = Trim(SubNodes.getElementsByTagName("ORDER_DT").item(0).text)
''    PAY_DT      = Trim(SubNodes.getElementsByTagName("PAY_DT").item(0).text)
    GOODS_ID    = Trim(SubNodes.getElementsByTagName("GOODS_ID").item(0).text)
    GOODS_NAME  = Trim(SubNodes.getElementsByTagName("GOODS_NAME").item(0).text)
    ENTP_DT_CODE= Trim(SubNodes.getElementsByTagName("ENTP_DT_CODE").item(0).text)
    GOODSDT_INFO= Trim(SubNodes.getElementsByTagName("GOODSDT_INFO").item(0).text)
    O_NAME      = Trim(SubNodes.getElementsByTagName("O_NAME").item(0).text)
    O_TEL       = Trim(SubNodes.getElementsByTagName("O_TEL").item(0).text)
    O_HTEL      = Trim(SubNodes.getElementsByTagName("O_HTEL").item(0).text)
    O_EMAIL     = Trim(SubNodes.getElementsByTagName("O_EMAIL").item(0).text)
    S_NAME      = Trim(SubNodes.getElementsByTagName("S_NAME").item(0).text)
    S_TEL       = Trim(SubNodes.getElementsByTagName("S_TEL").item(0).text)
    S_HTEL      = Trim(SubNodes.getElementsByTagName("S_HTEL").item(0).text)
    S_POST      = Trim(SubNodes.getElementsByTagName("S_POST").item(0).text)
    S_ADDR      = Trim(SubNodes.getElementsByTagName("S_ADDR").item(0).text)
    CS_MSG      = Trim(SubNodes.getElementsByTagName("CS_MSG").item(0).text)
    QTY         = Trim(SubNodes.getElementsByTagName("QTY").item(0).text)
    SALE_PRICE  = Trim(SubNodes.getElementsByTagName("SALE_PRICE").item(0).text)
    DELY_TYPE   = Trim(SubNodes.getElementsByTagName("DELY_TYPE").item(0).text)
    DELY_COST   = Trim(SubNodes.getElementsByTagName("DELY_COST").item(0).text)

    if InStr(ORDER_NO_ARR,ORDER_NO)<1 then
        ORDER_NO_ARR = ORDER_NO_ARR + ORDER_NO + ","
    end if

''    buf = ORDER_DT
''    ORDER_DT    = Left(buf,4)&"-"&MID(buf,5,2)&"-"&MID(buf,7,2)&" "&Right(buf,8)      '''- 포함 들어옴.
''    IF (PAY_DT<>"") then
''        buf = PAY_DT
''        PAY_DT    = Left(buf,4)&"-"&MID(buf,5,2)&"-"&MID(buf,7,2)&" "&Right(buf,8)
''    END IF

    ''중복CHECK
    OrderAlreadyInputed =false
    sqlStr = " select top 1 ORDER_DT from db_temp.dbo.tbl_LTiMall_OrdNoti"
    sqlStr = sqlStr& " where ORDER_NO='"&ORDER_NO&"'"
    sqlStr = sqlStr& " and ORDER_SEQ='"&ORDER_SEQ&"'"
    rsget.Open sqlStr,dbget,1
    IF  not rsget.EOF  then
        OrderAlreadyInputed = true
        ERR_ORDER_NO_ARR = ERR_ORDER_NO_ARR + ORDER_NO + ORDER_SEQ + ","
        retERRXML = retERRXML &"<OrderMasterResult>"&vbcrlf
        retERRXML = retERRXML &"	<SHORT_ORDER_NO>"&ORDER_NO&"</SHORT_ORDER_NO>"&vbcrlf
        retERRXML = retERRXML &"	<ORDER_DT>"&ORDER_DT&"</ORDER_DT>"&vbcrlf
        retERRXML = retERRXML &"<OrderDetailResult>"&vbcrlf
        retERRXML = retERRXML &"	<ORDER_NO>"&ORDER_NO&"-"&ORDER_SEQ&"</ORDER_NO>"&vbcrlf
        retERRXML = retERRXML &"	<ORDER_SEQ>"&ORDER_SEQ&"</ORDER_SEQ>"&vbcrlf
        retERRXML = retERRXML &"	<RESULT>F</RESULT>"&vbcrlf
        retERRXML = retERRXML &"	<RESULT_MSG>Already Sended Order</RESULT_MSG>"&vbcrlf
        retERRXML = retERRXML &"</OrderDetailResult>"&vbcrlf
        retERRXML = retERRXML &"</OrderMasterResult>"&vbcrlf
    end IF
    rsget.Close

    IF (Not OrderAlreadyInputed) and (ORDER_NO<>"") and (ORDER_SEQ<>"") then
        sqlStr = " Insert into db_temp.dbo.tbl_LTiMall_OrdNoti"
        sqlStr = sqlStr& " (ORDER_NO,ORDER_SEQ,ORDER_DT"
        IF (PAY_DT<>"") then
        sqlStr = sqlStr& " ,PAY_DT"
        ENd IF
        sqlStr = sqlStr& " ,GOODS_ID,GOODS_NAME,ENTP_DT_CODE,GOODSDT_INFO"
        sqlStr = sqlStr& " ,O_NAME,O_TEL,O_HTEL,O_EMAIL,S_NAME,S_TEL,S_HTEL,S_POST,S_ADDR,CS_MSG,QTY,SALE_PRICE"
        sqlStr = sqlStr& " ,DELY_TYPE,DELY_COST,DATETIME,DOCUMENTID,refip)"
        sqlStr = sqlStr& " values('"&ORDER_NO&"'"
        sqlStr = sqlStr& " ,'"&ORDER_SEQ&"'"
        sqlStr = sqlStr& " ,'"&ORDER_DT&"'"
        IF (PAY_DT<>"") then
        sqlStr = sqlStr& " ,'"&PAY_DT&"'"
        ENd IF
        sqlStr = sqlStr& " ,'"&GOODS_ID&"'"
        sqlStr = sqlStr& " ,'"&html2DB(GOODS_NAME)&"'"
        sqlStr = sqlStr& " ,'"&ENTP_DT_CODE&"'"
        sqlStr = sqlStr& " ,'"&html2DB(GOODSDT_INFO)&"'"
        sqlStr = sqlStr& " ,'"&html2DB(O_NAME)&"'"
        sqlStr = sqlStr& " ,'"&html2DB(O_TEL)&"'"
        sqlStr = sqlStr& " ,'"&html2DB(O_HTEL)&"'"
        sqlStr = sqlStr& " ,'"&html2DB(O_EMAIL)&"'"
        sqlStr = sqlStr& " ,'"&html2DB(S_NAME)&"'"
        sqlStr = sqlStr& " ,'"&html2DB(S_TEL)&"'"
        sqlStr = sqlStr& " ,'"&html2DB(S_HTEL)&"'"
        sqlStr = sqlStr& " ,'"&S_POST&"'"
        sqlStr = sqlStr& " ,'"&html2DB(S_ADDR)&"'"
        sqlStr = sqlStr& " ,'"&html2DB(CS_MSG)&"'"
        sqlStr = sqlStr& " ,"&QTY&""
        sqlStr = sqlStr& " ,"&SALE_PRICE&""
        sqlStr = sqlStr& " ,'"&DELY_TYPE&"'"
        sqlStr = sqlStr& " ,"&DELY_COST&""
        sqlStr = sqlStr& " ,'"&DATETIME&"'"
        sqlStr = sqlStr& " ,'"&DOCUMENTID&"'"
        sqlStr = sqlStr& " ,'"&refip&"'"
        sqlStr = sqlStr& " )"
    ''rw sqlStr
        dbget.Execute sqlStr
    ENd IF

    IF (ORDER_NO="") then
        CALL responseErrXML("XML parsing ERROR : InValid param ORDER_NO")
        response.end
    end if

    IF (ORDER_SEQ="") then
        CALL responseErrXML("XML parsing ERROR : InValid param ORDER_SEQ")
        response.end
    end if

next

SET xmlDOM=Nothing

IF (enti<1) then
    CALL responseErrXML("XML parsing ERROR : NO OrderEntryLineItem, NoLine")
    response.end
end if

Dim p_ORDER_NO, p_ORDER_DT

if Right(ORDER_NO_ARR,1)="," then
    ORDER_NO_ARR = Left(ORDER_NO_ARR,Len(ORDER_NO_ARR)-1)
end if
ORDER_NO_ARR = "'"&replace(ORDER_NO_ARR,",","','")&"'"

if Right(ERR_ORDER_NO_ARR,1)="," then
    ERR_ORDER_NO_ARR = Left(ERR_ORDER_NO_ARR,Len(ERR_ORDER_NO_ARR)-1)
end if
ERR_ORDER_NO_ARR = "'"&replace(ERR_ORDER_NO_ARR,",","','")&"'"

retXML = "<?xml version='1.0' encoding='utf-8'?>"&vbcrlf
retXML = retXML &"<OrderInfoResult>"&vbcrlf
retXML = retXML &"<MessageHeader>"&vbcrlf
retXML = retXML &"	<SENDER>TENBYTEN</SENDER>"&vbcrlf
retXML = retXML &"	<RECEIVER>LotteH</RECEIVER>"&vbcrlf
retXML = retXML &"	<DATETIME>"&replace(Left(FormatDateTime(now,0),10),"-","")&" "&Left(FormatDateTime(now,4),5)&Right(FormatDateTime(now,0),3)&"</DATETIME>"&vbcrlf
retXML = retXML &"	<DOCUMENTID>ORDERINFO</DOCUMENTID>"&vbcrlf
retXML = retXML &"	<ERROROCCUR>"&CHKIIF(retERRXML<>"","Y","N")&"</ERROROCCUR>"&vbcrlf
retXML = retXML &"	<ERRORMESSAGE></ERRORMESSAGE>"&vbcrlf
retXML = retXML &"</MessageHeader>"&vbcrlf
retXML = retXML &"<MessageBody>"&vbcrlf

IF (retERRXML<>"") then
    retXML = retXML & retERRXML&vbcrlf
ELSE
    sqlStr = "select ORDER_NO, convert(varchar(19),ORDER_DT,21) as ORDER_DT, ORDER_SEQ"
    sqlStr = sqlStr& " from db_temp.dbo.tbl_LTiMall_OrdNoti"
    sqlStr = sqlStr& " where ORDER_NO='"&ORDER_NO&"'"               ''sqlStr = sqlStr& " where ORDER_NO in ("&ORDER_NO_ARR&")"
    sqlStr = sqlStr& " and ORDER_SEQ='"&ORDER_SEQ&"'"
    ''sqlStr = sqlStr& " and DATETIME='"&REplace(DATETIME,"-","")&"'"
    sqlStr = sqlStr& " and DOCUMENTID='"&DOCUMENTID&"'"
    ''if (ERR_ORDER_NO_ARR<>"") then
    ''    sqlStr = sqlStr& " and (ORDER_NO+ORDER_SEQ) not in ("&ERR_ORDER_NO_ARR&")"
    ''end if
    sqlStr = sqlStr& " order by ORDER_NO, ORDER_SEQ"
    rsget.Open sqlStr,dbget,1

    if  not rsget.EOF  then
        retXML = retXML &"<OrderMasterResult>"&vbcrlf
        retXML = retXML &"	<SHORT_ORDER_NO>"&rsget("ORDER_NO")&"</SHORT_ORDER_NO>"&vbcrlf
        retXML = retXML &"	<ORDER_DT>"&rsget("ORDER_DT")&"</ORDER_DT>"&vbcrlf
        retXML = retXML &"<OrderDetailResult>"&vbcrlf
        retXML = retXML &"	<ORDER_NO>"&rsget("ORDER_NO")&"-"&rsget("ORDER_SEQ")&"</ORDER_NO>"&vbcrlf
        retXML = retXML &"	<ORDER_SEQ>"&rsget("ORDER_SEQ")&"</ORDER_SEQ>"&vbcrlf
        retXML = retXML &"	<RESULT>S</RESULT>"&vbcrlf
        retXML = retXML &"	<RESULT_MSG></RESULT_MSG>"&vbcrlf
        retXML = retXML &"</OrderDetailResult>"&vbcrlf
        retXML = retXML &"</OrderMasterResult>"&vbcrlf
    end if
    rsget.Close
END IF
retXML = retXML &"</MessageBody>"&vbcrlf
retXML = retXML &"</OrderInfoResult>"&vbcrlf

'' 임시저장.
CALL XMLFileSave(retXML ,"ORD_RET",0)
response.write retXML

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->