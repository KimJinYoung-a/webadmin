<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<%

dim entrId      : entrId="10X10"                    ''������ : ���޾�ü_ID
dim ordclmNo    : ordclmNo=request("ordclmNo")      ''������ũ �ֹ���ȣ
dim ordSeq      : ordSeq=request("ordSeq")          ''������ũ �ֹ�����
dim delvDt      : delvDt=request("delvDt")          ''YYYYMMDD ���Ϸ�����
dim delvEntrNo  : delvEntrNo=request("delvEntrNo")  ''�ù���ڵ�
dim invoNo      : invoNo=request("invoNo")          ''������ȣ ���ڸ� ������.
dim optPrdTp    : optPrdTp=request("optPrdTp")      '' �ɼǻ�ǰ����	01 (�Ϲݴ�ǰ��ǰ)  02 (�߰��ɼ������ԵȺθ��ǰ)
dim optOrdSeqList  : optOrdSeqList=request("optOrdSeqList") ''�ֹ���������Ʈ


'2013/02/28 �����߰�
dim mode      : mode=request("mode")
If mode = "updateSendState" Then
	Dim sqlStr, AssignedRow
	sqlStr = "Update db_temp.dbo.tbl_xSite_TMPOrder "
	sqlStr = sqlStr & "	Set sendState='"&request("updateSendState")&"'"
	sqlStr = sqlStr & "	,sendReqCnt=sendReqCnt+1"
	sqlStr = sqlStr & "	where OutMallOrderSerial='"&request("ordclmNo")&"'"
	sqlStr = sqlStr & "	and OrgDetailKey='"&request("ordSeq")&"'"
	sqlStr = sqlStr & "	and sellsite='interpark'"
	dbget.Execute sqlStr,AssignedRow
	response.write "<script>alert('"&AssignedRow&"�� �Ϸ� ó��.');window.close()</script>"
	response.end
End If
'2013/02/28 �����߰� ��



'''01 �Ϲݴܺջ�ǰ�� ��� �ش� �ֹ�����  02 �߰��ɼ������ԵȺθ��ǰ�� ��� �θ��ǰ�� �ֹ����� �� �߰��ɼǻ�ǰ�� �ֹ�����
''* �ֹ������� �ϳ� �̻��� ��� ������(��|��)�� �����Ѵ�.
''ex) sc.optOrdSeqList=1|2|3|4


invoNo= trim(replace(invoNo,"-",""))

	'// ��������
	dim strSql, actCnt, lp
    dim AssignedCNT, paramAdd

	actCnt = 0			'�ǰ��ŰǼ�

'Dim ordclmNo : ordclmNo = ord_no
'ord_no = replace(ord_no,":01","")  '''2011-11-30-4100282:01
'ord_no = replace(ord_no,"_1","")  ''2012-03-11-8159343_1
'ord_no = replace(ord_no,"_2","")
'ord_no = replace(ord_no,"_3","")
'ord_no = replace(ord_no,"_4","")



paramAdd = "&sc.entrId="+entrId
if (Right(ordclmNo,2)="_1") then
    paramAdd = paramAdd + "&sc.ordclmNo="+replace(ordclmNo,"_1","")
else
    paramAdd = paramAdd + "&sc.ordclmNo="+ordclmNo
end if
paramAdd = paramAdd + "&sc.ordSeq="+ordSeq
paramAdd = paramAdd + "&sc.delvDt="+delvDt
paramAdd = paramAdd + "&sc.delvEntrNo="+delvEntrNo
paramAdd = paramAdd + "&sc.invoNo="+invoNo
paramAdd = paramAdd + "&sc.optPrdTp="+optPrdTp
paramAdd = paramAdd + "&sc.optOrdSeqList="+optOrdSeqList


Dim iResult, iMessage
Dim iParkURL, iParams, replyXML, ErrMsg

''iParkURL = "http://www.interpark.com/order/OrderClmAPI.do"
iParkURL = "https://joinapi.interpark.com/order/OrderClmAPI.do"   '''�Ǽ����� https://joinapi.interpark.com
iParams  = "_method=delvCompForComm" & paramAdd


'rw "delvEntrNo="&delvEntrNo
''rw iParkURL &"?"& iParams
''response.end


replyXML = SendReqGet(iParkURL, iParams)

    Select Case left(replyXML,5)
    	Case "[401]","[404]","[500]","[err]"
    		ErrMsg = replyXML
    	Case Else
    		ErrMsg = ""
    end Select

    if (ErrMsg<>"") then
        if (IsAutoScript) then
            rw "������ũ �����Է���  ������ �߻��߽��ϴ�. "&ordclmNo&" "&ordclmNo&"_"&ordSeq
        else
    	    Response.Write "<script language=javascript>alert('������ũ �����Է���  ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
    	ENd IF
    	dbget.Close: Response.End

    else

    	'XML�� ���� DOM ��ü ����
    	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
    	xmlDOM.async = False
    	'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
''rw "objXML.ResponseBody="&BinaryToText(objXML.ResponseBody, "euc-kr")
 ''response.end

        ''����� �� (20120702) ���� ����.. :: ������?
        IF (Trim(replyXML)="") THEN
            rw "��� ��"
            iResult="1"
        ELSE
        	xmlDOM.LoadXML replyXML

        	if Err<>0 then
        		Response.Write "<script language=javascript>alert('�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
        		dbget.Close: Response.End
        	end if
            On Error Resume next
        	iResult		= Trim(xmlDOM.getElementsByTagName("CODE").item(0).text)			'���
            iMessage    = xmlDOM.getElementsByTagName("MESSAGE").Item(0).Text
        	On Error Goto 0
        END IF
    '	'// ����
        IF (iResult="000") THEN
        	strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
        	strSql = strSql & "	Set sendState=1"
        	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
            strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
            strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
            strSql = strSql & "	and sellsite='interpark'"
            strSql = strSql & "	and matchstate in ('O')"
       ''rw strSql
        	dbget.Execute strSql,AssignedCNT
        	actCnt = actCnt+AssignedCNT

        	IF (AssignedCNT>0) then
        	    if (IsAutoScript) then
        	        rw "OK|"&ordclmNo&"_"&ordSeq
        	    ELSE
            	    response.write "OK"
            	ENd IF
            ENd IF
        ELSE
            ''�̹� ������ �ԷµȰ��.// ������ N�� ��õ��� ����.
            ''���� �ʰ� �Է� : �ֹ��Ϸκ��� 30���� ����Ͽ� �߼ۿϷ� �Ұ��մϴ�. �����Ϳ� �����Ͻñ� �ٶ��ϴ�. ����ó) 02-3289-3236
            ''��ҵȰ�� : �߼ۿϷ� ó���� �ֹ��� �������� �ʽ��ϴ�.
            ''Maybe ������ : ���� �ֹ��������°� �߼ۿϷḦ ó���� �� �����ϴ�
            IF LEft(iMessage,Len("�̹� ���ó���� �Ǿ����ϴ�."))="�̹� ���ó���� �Ǿ����ϴ�." THEN
                strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
            	strSql = strSql & "	Set sendState=9"
            	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
                strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
                strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
                strSql = strSql & "	and sellsite='interpark'"
                strSql = strSql & "	and matchstate in ('O','C','Q','A')"  ''A �߰�
'rw  strSql
            	dbget.Execute strSql,AssignedCNT

            	if (IsAutoScript) then
            	    rw "SKIP ó�� "&ordclmNo&" "&ordSeq
            	ELSE
            	    rw "SKIP ó��"
            	ENd IF

            ELSE
                '' �õ� ȸ�� �߰� sendReqCnt
                strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
            	strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
                strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
                strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
                strSql = strSql & "	and sellsite='interpark'"
                strSql = strSql & "	and matchstate in ('O','C','Q')" '','A'
'rw strSql
            	dbget.Execute strSql

                iMessage = "<font color=red>"&iMessage&"</font>"
            END IF

            if (IsAutoScript) then
                rw "iMessage="&iMessage&":"&ordclmNo&" "&ordclmNo&"_"&ordSeq
            else
                rw "iMessage="&iMessage
            ENd IF

'2013/02/28 ������ �߰�
'���� ����Ƚ���� 3ȸ�� �����鼭 minusOrderSerial�� ������ �� �ش�
'updateSendState = 901		�߼�ó������ �����ϰ�
'updateSendState = 902		����� ��������
'updateSendState = 903		��ǰó����
			Dim errCount
			strSql = ""
			strSql = strSql & " SELECT Count(*) as cnt FROM db_temp.dbo.tbl_xSite_TMPOrder " & VBCRLF
			strSql = strSql & "	where OutMallOrderSerial='"&ordclmNo&"'"
			strSql = strSql & "	and OrgDetailKey='"&ordSeq&"'"
			strSql = strSql & " and sendReqCnt >= 3" & VBCRLF
			rsget.Open strSql,dbget,1
			If Not rsget.Eof Then
				errCount = rsget("cnt")
			End If
			rsget.Close

			If errCount > 0 Then
				response.write  "<select name='updateSendState' id=""updateSendState"">" &_
								"	<option value=''>����</option>" &_
								"	<option value='901'>�߼�ó������ �����ϰ�</option>" &_
								"	<option value='902'>����� ��������</option>" &_
								"	<option value='903'>��ǰó����</option>" &_
								"</select>&nbsp;&nbsp;"
				response.write "<input type='button' value='�Ϸ�ó��' onClick=""finCancelOrd2('"&ordclmNo&"','"&ordSeq&"',document.getElementById('updateSendState').value)"">"
				response.write "<script language='javascript'>"&VbCRLF
				response.write "function finCancelOrd2(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
				response.write "    if(selectValue == ''){"&VbCRLF
				response.write "    	alert('�������ּ���');"&VbCRLF
				response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
				response.write "    	return;"&VbCRLF
				response.write "    }"&VbCRLF
				response.write "    var uri = 'actInterparkSongjangInputProc.asp?mode=updateSendState&ordclmNo='+ORG_ord_no+'&ordSeq='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
				response.write "    location.replace(uri);"&VbCRLF
				response.write "}"&VbCRLF
				response.write "</script>"&VbCRLF
			End If
'2013/02/28 ������ �߰� ��

        END IF
    	Set xmlDOM = Nothing
    end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->