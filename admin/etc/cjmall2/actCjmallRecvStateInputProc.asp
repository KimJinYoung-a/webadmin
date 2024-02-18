<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/cjmall/incCJmallFunction.asp"-->
<!-- #include virtual="/admin/etc/incOutMallRecvCheckFunction.asp"-->

<%

function getCjMallRecvXMLStr(mallordernoAll, delicompCd, wbNo, vendorOrdId, recvNm)
    Dim strRst
    Dim mallorderno,mallOrderSeq1,mallOrderSeq2,mallOrderSeq3

    mallorderno     = splitValue(mallordernoAll,"-",0)
    mallOrderSeq1   = splitValue(mallordernoAll,"-",1)
    mallOrderSeq2   = splitValue(mallordernoAll,"-",2)
    mallOrderSeq3   = splitValue(mallordernoAll,"-",3)

    strRst = ""
    strRst = strRst &"<?xml version=""1.0"" encoding=""EUC-KR""?>"
    strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_04_05"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_04_05.xsd"">"
    strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
    strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
	strRst = strRst &"<tns:receiveComplete>"
    strRst = strRst &"<tns:ordNo>"&mallorderno&"</tns:ordNo>" ''주문정보 - 주문번호
    strRst = strRst &"<tns:ordGSeq>"&mallOrderSeq1&"</tns:ordGSeq>" ''주문정보 - 주문상품순번
    strRst = strRst &"<tns:ordDSeq>"&mallOrderSeq2&"</tns:ordDSeq>" ''주문정보 - 주문상세순번
    strRst = strRst &"<tns:ordWSeq>"&mallOrderSeq3&"</tns:ordWSeq>" ''주문정보 - 주문처리순번
    strRst = strRst &"<tns:recvNm>"&recvNm&"</tns:recvNm>" ''받는사람
    strRst = strRst &"</tns:receiveComplete>"
    strRst = strRst &"</tns:ifRequest>"

    getCjMallRecvXMLStr = strRst
end function


	'// 변수선언
	dim LotteGoodNo, LotteStatCd
	dim strSql, actCnt, lp
    dim AssignedCNT

	actCnt = 0			'실갱신건수

Dim proc_gubun : proc_gubun = "sfin"

Dim ten_ord_no : ten_ord_no     = request("ten_ord_no")
Dim ord_no     : ord_no     = request("ord_no")
Dim ord_dtl_sn : ord_dtl_sn = request("ord_dtl_sn")
Dim hdc_cd     : hdc_cd     = request("hdc_cd")
Dim inv_no     : inv_no     = request("inv_no")
Dim rcv_nm     : rcv_nm     = request("rcv_nm")
Dim paramAdd
Dim tenOrderSerial,orgsubtotalprice ,minusOrderSerial ,minusSubtotalprice
dim recvCheckResult
dim errCode, errMSG

inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")

Dim ORG_ord_no : ORG_ord_no = ord_no


recvCheckResult = FnCheckNSaveRecvState(hdc_cd, inv_no, errCode, errMSG)

if (Not recvCheckResult) then
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set recvSendReqCnt=IsNull(recvSendReqCnt, 0) + 1"
    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
    strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
    strSql = strSql & "	and matchstate in ('O','C','Q','A')"

	dbget.Execute strSql

	response.write "<script>alert('배송완료 이전입니다.'); location.href='" + CStr(FnGetRecvCheckURL(hdc_cd, inv_no)) + "';</script>"
	dbget.close()
	response.end
end if

'' response.write "TEST" & hdc_cd
'' response.end



Dim xmlStr : xmlStr = getCjMallRecvXMLStr(ord_dtl_sn, hdc_cd, inv_no, ten_ord_no, rcv_nm)


Dim retDoc, sURL
Dim successYn, errorMsg
sURL = cjMallAPIURL

SET retDoc = xmlSend(sURL, xmlStr)

'On Error Resume next
successYn = retDoc.getElementsByTagName("ns1:successYn").item(0).text
errorMsg = retDoc.getElementsByTagName("ns1:errorMsg").item(0).text
'On Error Goto 0
SET retDoc = Nothing

'rw successYn  (true, false)
'rw errorMsg
rw successYn
rw errorMsg
Dim IsSuccss : IsSuccss=(successYn="true")

if (IsSuccss) then
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set recvSendState = 100"
	strSql = strSql & "	, recvSendReqCnt=IsNull(recvSendReqCnt, 0) + 1 "
    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
    strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
    strSql = strSql & "	and matchstate in ('O')"

	dbget.Execute strSql,AssignedCNT

    IF (AssignedCNT>0) then
	    if (IsAutoScript) then
	        rw "OK|"&ord_no&" "&ord_dtl_sn
	    ELSE
    	    response.write "OK"
    	ENd IF
    ENd IF
else
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set recvSendReqCnt=IsNull(recvSendReqCnt, 0) + 1 "
    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"&VBCRLF
    strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"&VBCRLF
    strSql = strSql & "	and matchstate in ('O','C','Q','A')"

	dbget.Execute strSql

    rw "<font color=red>"&errorMsg&"</font>"

    rw ten_ord_no
    rw ord_no
    rw ord_dtl_sn
    rw hdc_cd
    rw inv_no
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
