<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/cjmall/incCJmallFunction.asp"-->
<%
function getCjMallSongjangXMLStr(mallordernoAll,delicompCd,wbNo,vendorOrdId)
    Dim strRst
    Dim mallorderno,mallOrderSeq1,mallOrderSeq2,mallOrderSeq3

    mallorderno     = splitValue(mallordernoAll,"-",0)
    mallOrderSeq1   = splitValue(mallordernoAll,"-",1)
    mallOrderSeq2   = splitValue(mallordernoAll,"-",2)
    mallOrderSeq3   = splitValue(mallordernoAll,"-",3)

    strRst = ""
    strRst = strRst &"<?xml version=""1.0"" encoding=""EUC-KR""?>"
    strRst = strRst &"<tns:ifRequest xmlns:tns=""http://www.example.org/ifpa"" tns:ifId=""IF_04_04"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:schemaLocation=""http://www.example.org/ifpa ../IF_04_04.xsd"">"
    strRst = strRst &"<tns:vendorId>411378</tns:vendorId>"
    strRst = strRst &"<tns:vendorCertKey>CJ03074113780</tns:vendorCertKey>"
    strRst = strRst &"<tns:takeout>"
    strRst = strRst &"<tns:ordNo>"&mallorderno&"</tns:ordNo>" ''�ֹ����� - �ֹ���ȣ
    strRst = strRst &"<tns:ordGSeq>"&mallOrderSeq1&"</tns:ordGSeq>" ''�ֹ����� - �ֹ���ǰ����
    strRst = strRst &"<tns:ordDSeq>"&mallOrderSeq2&"</tns:ordDSeq>" ''�ֹ����� - �ֹ��󼼼���
    strRst = strRst &"<tns:ordWSeq>"&mallOrderSeq3&"</tns:ordWSeq>" ''�ֹ����� - �ֹ�ó������
    strRst = strRst &"<tns:delicompCd>"&delicompCd&"</tns:delicompCd>" ''�ù��
    strRst = strRst &"<tns:wbNo>"&wbNo&"</tns:wbNo>" ''������ȣ
    strRst = strRst &"<tns:vendorOrdId>"&vendorOrdId&"</tns:vendorOrdId>" ''���»��ֹ���ȣ
    strRst = strRst &"</tns:takeout>"
    strRst = strRst &"</tns:ifRequest>"

    getCjMallSongjangXMLStr = strRst
end function

dim mode      : mode=request("mode")
If mode = "updateSendState" Then
	Dim sqlStr, AssignedRow
	sqlStr = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	sqlStr = sqlStr & "	Set sendState='"&request("updateSendState")&"'"
	sqlStr = sqlStr & "	,sendReqCnt=sendReqCnt+1"

	if (request("updateSendState") = "952") then
		'// ����ֹ��� �μ����۵� skip
		sqlStr = sqlStr & " , recvSendState = 100 "
		sqlStr = sqlStr & " , recvSendReqCnt=IsNull(recvSendReqCnt, 0) + 1 "
	end if

	sqlStr = sqlStr & "	where OutMallOrderSerial='"&request("ord_no")&"'"
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ord_dtl_sn")&"'"
	sqlStr = sqlStr & "	and sellsite='cjmall'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"�� �Ϸ� ó��.');window.close()</script>"
	response.end
End If

'// ��������
dim LotteGoodNo, LotteStatCd
dim strSql, actCnt, lp
dim AssignedCNT

actCnt = 0			'�ǰ��ŰǼ�

Dim proc_gubun : proc_gubun = "sfin"

Dim ten_ord_no : ten_ord_no     = request("ten_ord_no")
Dim ord_no     : ord_no     = request("ord_no")
Dim ord_dtl_sn : ord_dtl_sn = request("ord_dtl_sn")
Dim hdc_cd     : hdc_cd     = request("hdc_cd")
Dim inv_no     : inv_no     = Left(request("inv_no"), 15)					'// 15�� ������ ����
Dim paramAdd
Dim tenOrderSerial,orgsubtotalprice ,minusOrderSerial ,minusSubtotalprice

inv_no = replace(inv_no,"-","")
inv_no = replace(inv_no," ","")

Dim ORG_ord_no : ORG_ord_no = ord_no

'ord_no = replace(ord_no,":01","")  '''2011-11-30-4100282:01
'ord_no = replace(ord_no,"_1","")  ''2012-03-11-8159343_1
'ord_no = replace(ord_no,"_2","")
'ord_no = replace(ord_no,"_3","")
'ord_no = replace(ord_no,"_4","")

'if (inv_no="������ ���� ������ �����е�") then inv_no="��Ÿ"
'paramAdd = "&ord_no="+Replace(ord_no,"-","")
'paramAdd = paramAdd + "&ord_dtl_sn="+ord_dtl_sn
'paramAdd = paramAdd + "&proc_gubun="+proc_gubun
'paramAdd = paramAdd + "&hdc_cd="+hdc_cd
'paramAdd = paramAdd + "&inv_no="+server.UrlEncode(replace(inv_no,"-",""))

'ord_no = "20130523069391"
'ord_dtl_sn = "20130523069391-001-001-001"

'20130523069250-001-001-001
'20130523069304-001-001-001
'20130523069391-001-001-001

Dim xmlStr : xmlStr = getCjMallSongjangXMLStr(ord_dtl_sn,hdc_cd,inv_no,ten_ord_no)

'response.write xmlStr
'response.end

Dim retDoc, sURL
Dim successYn, errorMsg
sURL = cjMallAPIURL

SET retDoc = xmlSend(sURL, xmlStr)

''response.end

'    If (isCJ_DebugMode) Then
'        CALL XMLFileSave(retDoc.XML, "RET_SONGJANG", ord_dtl_sn)
'    End If

'On Error Resume next
successYn = retDoc.getElementsByTagName("ns1:successYn").item(0).text
errorMsg = retDoc.getElementsByTagName("ns1:errorMsg").item(0).text
'On Error Goto 0
SET retDoc = Nothing

'rw successYn  (true, false)
'rw errorMsg
'rw successYn
'rw errorMsg
Dim IsSuccss : IsSuccss=(successYn="true")

if (IsSuccss) then
    strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	strSql = strSql & "	Set sendState=1"
	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
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
	strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
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

	'���� ����Ƚ���� 3ȸ�� ������ ����ó�� ����
	'updateSendState = 951		������ ����
	'updateSendState = 952		����ֹ�
	Dim errCount : errCount = 0
	strSql = ""
	strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
	strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
	strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
	strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
	rsget.Open strSql,dbget,1
	If Not rsget.Eof Then
		errCount = rsget("cnt")
	End If
	rsget.Close

	If errCount > 0 Then
		response.write  "<select name='updateSendState' id=""updateSendState"">" &_
						"	<option value=''>����</option>" &_
						"	<option value='951'>������ ����</option>" &_
						"	<option value='952'>����ֹ�</option>" &_
						"</select>&nbsp;&nbsp;"
		response.write "<input type='button' value='�Ϸ�ó��' onClick=""fnSetSendState('"&ORG_ord_no&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)"">"
		response.write "<script language='javascript'>"&VbCRLF
		response.write "function fnSetSendState(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
		response.write "    if(selectValue == ''){"&VbCRLF
		response.write "    	alert('�������ּ���');"&VbCRLF
		response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
		response.write "    	return;"&VbCRLF
		response.write "    }"&VbCRLF
		response.write "    var uri = 'actCjmallSongjangInputProc.asp?mode=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If

end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
