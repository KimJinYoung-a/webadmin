<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LotteiMall/incLotteiMallFunction.asp"-->
<!-- #include virtual="/admin/etc/LotteiMall/lotteiMallcls.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<%
function responseErrXML(iErrMSG)
    dim retXML
    retXML = "<?xml version='1.0' encoding='utf-8'?>"&vbcrlf 
    retXML = retXML &"<OrderInfoResult>"&vbcrlf 
    retXML = retXML &"<MessageHeader>"&vbcrlf 
    retXML = retXML &"	<SENDER>TENBYTEN</SENDER>"&vbcrlf 
    retXML = retXML &"	<RECEIVER>LotteH</RECEIVER>"&vbcrlf 
    retXML = retXML &"	<DATETIME>"&replace(Left(FormatDateTime(now,0),10),"-","")&" "&Left(FormatDateTime(now,4),5)&Right(FormatDateTime(now,0),3)&"</DATETIME>"&vbcrlf 
    retXML = retXML &"	<DOCUMENTID>ORDERCNCLINFO</DOCUMENTID>"&vbcrlf 
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

Dim reqBufALL 
Dim xmlDOM, listItem, bufText, retXML, retERRXML, OrderAlreadyInputed
set xmlDOM = CreateObject("MSXML2.DomDocument.3.0")
xmlDOM.async = False 
xmlDOM.Load Request '''xmlDOM.LoadXML Request  :: Load 메서드가 맞는듯 옵션등에 

Dim refip : refip = request.ServerVariables("REMOTE_ADDR")
Dim ErrMSG


dim buf 
buf = xmlDOM.xml ''xmlDOM.text 


'' 임시저장.
CALL XMLFileSave(buf,"CS",0)


Dim sqlStr, i
Dim ORDER_NO_ARR, ERR_ORDER_NO_ARR
Dim ORDER_NO,ORDER_SEQ,ORDER_DT,PAY_DT,CNCL_CNT,CNCL_AMT,ENTP_DT_CODE,ORDER_STTS
Dim CNCL_DT
Dim DELY_TYPE,DELY_COST
Dim OrderCnclEntry,SubNodes
Dim DATETIME,DOCUMENTID
On Error resume Next
DOCUMENTID  = xmlDOM.getElementsByTagName("DOCUMENTID").item(0).text
If (ERR) then
    CALL responseErrXML("XML parsing ERROR : NO DOCUMENTID")
    CALL SendMail("webserver@10x10.co.kr", "kjy8517@10x10.co.kr", "[롯데iMall]CS주문수신오류"&now(),Err.description)
    response.end
End IF
DATETIME    = xmlDOM.getElementsByTagName("DATETIME").item(0).text
If (ERR) then
    CALL responseErrXML("XML parsing ERROR : NO DATETIME")
    response.end
End IF
buf = DATETIME
DATETIME    = Left(buf,4)&"-"&MID(buf,5,2)&"-"&MID(buf,7,2)&" "&Right(buf,8)
Set OrderCnclEntry = xmlDOM.getElementsByTagName("OrderCnclEntry")

If (ERR) then
    CALL responseErrXML("XML parsing ERROR : NO OrderCnclEntry")
    response.end
End IF
On Error  Goto 0



'''무조건 1건씩 날라옴..
Dim enti : enti=0
for each SubNodes in OrderCnclEntry
    

    enti=enti+1
    ORDER_NO    = Trim(SubNodes.getElementsByTagName("ORDER_NO").item(0).text)                  ''여러주문건이 한번에 오는지 검토 필요..(하나씩 옴)
   

    ORDER_SEQ   = Trim(SubNodes.getElementsByTagName("ORDER_SEQ").item(0).text)
    ORDER_DT    = Trim(SubNodes.getElementsByTagName("ORDER_DT").item(0).text)
    CNCL_CNT    = Trim(SubNodes.getElementsByTagName("CNCL_CNT").item(0).text)
    CNCL_AMT    = Trim(SubNodes.getElementsByTagName("CNCL_AMT").item(0).text)
    ''ENTP_DT_CODE= Trim(SubNodes.getElementsByTagName("ENTP_DT_CODE").item(0).text)
    ORDER_STTS  = Trim(SubNodes.getElementsByTagName("ORDER_STTS").item(0).text)                  ''주문상태	VC2	10		01:취소, 02:반품, 03:교환
    
    
    CNCL_DT     = Trim(SubNodes.getElementsByTagName("CNCL_DT").item(0).text)
    
    

    if InStr(ORDER_NO_ARR,ORDER_NO)<1 then
        ORDER_NO_ARR = ORDER_NO_ARR + ORDER_NO + ","    
    end if
    

    ''중복CHECK
    OrderAlreadyInputed =false
    sqlStr = " select top 1 ORDER_DT from db_temp.dbo.tbl_LTiMall_CSNoti"
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
        sqlStr = " Insert into db_temp.dbo.tbl_LTiMall_CSNoti"
        sqlStr = sqlStr& " (ORDER_NO,ORDER_SEQ,mallid,ORDER_DT"
        sqlStr = sqlStr& " ,CNCL_CNT,CNCL_AMT,ORDER_STTS"
        sqlStr = sqlStr& " ,CNCL_DT,refip)"
        sqlStr = sqlStr& " values('"&ORDER_NO&"'"
        sqlStr = sqlStr& " ,'"&ORDER_SEQ&"'"
        sqlStr = sqlStr& " ,'"&CMALLNAME&"'"
        sqlStr = sqlStr& " ,'"&ORDER_DT&"'"
        sqlStr = sqlStr& " ,"&CNCL_CNT&""
        sqlStr = sqlStr& " ,"&CNCL_AMT&""
        sqlStr = sqlStr& " ,'"&ORDER_STTS&"'"
        sqlStr = sqlStr& " ,'"&CNCL_DT&"'"
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
    CALL responseErrXML("XML parsing ERROR : NO OrderCnclEntryLineItem, NoLine")
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
retXML = retXML &"	<DOCUMENTID>ORDERCNCLINFO</DOCUMENTID>"&vbcrlf 
retXML = retXML &"	<ERROROCCUR>"&CHKIIF(retERRXML<>"","Y","N")&"</ERROROCCUR>"&vbcrlf 
retXML = retXML &"	<ERRORMESSAGE></ERRORMESSAGE>"&vbcrlf 
retXML = retXML &"</MessageHeader>"&vbcrlf 
retXML = retXML &"<MessageBody>"&vbcrlf 

IF (retERRXML<>"") then
    retXML = retXML & retERRXML&vbcrlf 
ELSE
    sqlStr = "select ORDER_NO, convert(varchar(19),ORDER_DT,21) as ORDER_DT, ORDER_SEQ"
    sqlStr = sqlStr& " from db_temp.dbo.tbl_LTiMall_CSNoti"
    sqlStr = sqlStr& " where ORDER_NO='"&ORDER_NO&"'"               ''sqlStr = sqlStr& " where ORDER_NO in ("&ORDER_NO_ARR&")"
    sqlStr = sqlStr& " and ORDER_SEQ='"&ORDER_SEQ&"'"
    sqlStr = sqlStr& " and mallid='"&CMALLNAME&"'"
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
CALL XMLFileSave(retXML ,"CS_RET",0)
response.write retXML

%> 
<!-- #include virtual="/lib/db/dbclose.asp" -->