<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<%
	'// ��������
	dim LotteGoodNo, LotteStatCd
	dim strSql, actCnt, lp
    dim AssignedCNT
    
	actCnt = 0			'�ǰ��ŰǼ�

Dim proc_gubun : proc_gubun = "sfin"
Dim ord_no     : ord_no     = request("ord_no")
Dim ord_dtl_sn : ord_dtl_sn = request("ord_dtl_sn")
Dim hdc_cd     : hdc_cd     = request("hdc_cd")
Dim inv_no     : inv_no     = request("inv_no")
Dim paramAdd 
Dim tenOrderSerial,orgsubtotalprice ,minusOrderSerial ,minusSubtotalprice


Dim ORG_ord_no : ORG_ord_no = ord_no

ord_no = replace(ord_no,":01","")  '''2011-11-30-4100282:01
ord_no = replace(ord_no,"_1","")  ''2012-03-11-8159343_1
ord_no = replace(ord_no,"_2","")
ord_no = replace(ord_no,"_3","")
ord_no = replace(ord_no,"_4","")

if (inv_no="������ ���� ������ �����е�") then inv_no="��Ÿ"
paramAdd = "&ord_no="+Replace(ord_no,"-","")
paramAdd = paramAdd + "&ord_dtl_sn="+ord_dtl_sn
paramAdd = paramAdd + "&proc_gubun="+proc_gubun
paramAdd = paramAdd + "&hdc_cd="+hdc_cd
paramAdd = paramAdd + "&inv_no="+server.UrlEncode(replace(inv_no,"-",""))

''rw lotteAPIURL & "/openapi/registDeliver.lotte?subscriptionId=" & lotteAuthNo & paramAdd
'response.end

Dim iResult, iMessage

    Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
    objXML.Open "GET", lotteAPIURL & "/openapi/registDeliver.lotte?subscriptionId=" & lotteAuthNo & paramAdd, false
    objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objXML.Send()
    If objXML.Status = "200" Then
    
    	'//���޹��� ���� Ȯ��
    	'Response.contentType = "text/xml; charset=euc-kr"
    	'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
    	'response.End
    
    	'XML�� ���� DOM ��ü ����
    	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
    	xmlDOM.async = False
    	'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
''rw "objXML.ResponseBody="&BinaryToText(objXML.ResponseBody, "euc-kr")
 ''response.end
        
        ''����� �� (20120702) ���� ����.. :: ������? 
        IF (Trim(objXML.ResponseBody)="") THEN
            rw "��� ��"
            iResult="1"
        ELSE
        	xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
        	
        	if Err<>0 then
        		Response.Write "<script language=javascript>alert('�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
        		dbget.Close: Response.End
        	end if
            On Error Resume next
        	iResult		= Trim(xmlDOM.getElementsByTagName("Result").item(0).text)			'���
            	If ERR THEN
            	    iMessage = xmlDOM.getElementsByTagName("Message").Item(0).Text
            	    iMessage2 = xmlDOM.getElementsByTagName("Message").Item(0).Text
            	ENd IF
        	On Error Goto 0    
        END IF
    '	'// ����

        IF (iResult="1") THEN
        	strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
        	strSql = strSql & "	Set sendState=1"
        	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
            strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
            strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
            strSql = strSql & "	and matchstate in ('O')" 
       ''rw strSql     
        	dbget.Execute strSql,AssignedCNT
        	actCnt = actCnt+AssignedCNT
        	
        	IF (AssignedCNT>0) then
        	    if (IsAutoScript) then
        	        rw "OK|"&ord_no&" "&ord_dtl_sn
        	    ELSE
            	    response.write "OK"
            	ENd IF
            ENd IF
        ELSE
            ''�̹� ������ �ԷµȰ��.// ������ N�� ��õ��� ����.
            ''���� �ʰ� �Է� : �ֹ��Ϸκ��� 30���� ����Ͽ� �߼ۿϷ� �Ұ��մϴ�. �����Ϳ� �����Ͻñ� �ٶ��ϴ�. ����ó) 02-3289-3236
            ''��ҵȰ�� : �߼ۿϷ� ó���� �ֹ��� �������� �ʽ��ϴ�. 
            ''Maybe ������ : ���� �ֹ��������°� �߼ۿϷḦ ó���� �� �����ϴ�
            IF LEft(iMessage,Len("���� �ֹ��������°� �߼ۿϷḦ ó���� �� �����ϴ�"))="���� �ֹ��������°� �߼ۿϷḦ ó���� �� �����ϴ�" THEN
                strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
            	strSql = strSql & "	Set sendState=9"
            	strSql = strSql & "	,sendReqCnt=sendReqCnt+1"
                strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
                strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
                strSql = strSql & "	and matchstate in ('O','C','Q','A')" 
''rw  strSql
            	dbget.Execute strSql,AssignedCNT
            	
            	if (IsAutoScript) then
            	    rw "SKIP ó�� "&ord_no&" "&ord_dtl_sn
            	ELSE
            	    rw "SKIP ó��"
            	ENd IF
            	
            ELSE
                '' �õ� ȸ�� �߰� sendReqCnt
                strSql = "Update db_temp.dbo.tbl_xSite_TMPOrder "
            	strSql = strSql & "	Set sendReqCnt=sendReqCnt+1"
                strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
                strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
                strSql = strSql & "	and matchstate in ('O','C','Q','A')" 

            	dbget.Execute strSql
            
                iMessage = "<font color=red>"&iMessage&"</font>"
            END IF
            
            if (IsAutoScript) then
                rw "iMessage="&iMessage&":"&ord_no&" "&ord_dtl_sn
            else
                rw "iMessage="&iMessage
            ENd IF

        END IF
    	Set xmlDOM = Nothing



    	IF (iResult<>"1") and (Not IsAutoScript) THEN
    	    ''
    	    strSql = "select * from db_temp.dbo.tbl_xSite_TMPOrder "
    	    strSql = strSql & "	where OutMallOrderSerial='"&ORG_ord_no&"'"
            strSql = strSql & "	and OrgDetailKey='"&ord_dtl_sn&"'"
            strSql = strSql & "	and sendState=0"
            
            rsget.Open strSql,dbget,1
            if Not rsget.Eof then
                tenOrderSerial = rsget("Orderserial")
            end if
            rsget.Close
            
            if (tenOrderSerial<>"") then
                strSql = "select top 10 m1.orderserial,m1.subtotalprice,m2.orderserial as minusOrderSerial,m2.subtotalprice as minusSubtotalprice"
                strSql = strSql & "	from db_order.dbo.tbl_order_master m1"
                strSql = strSql & "		left join db_order.dbo.tbl_order_master m2"
                strSql = strSql & "		on m2.linkorderserial='"&tenOrderSerial&"'"
                strSql = strSql & "		and m2.jumundiv='9'"
                strSql = strSql & "		and m2.cancelyn='N'"
                strSql = strSql & "	where m1.orderserial='"&tenOrderSerial&"'"
                
                rsget.Open strSql,dbget,1
                if Not rsget.Eof then
                    orgsubtotalprice    = rsget("subtotalprice")
                    minusOrderSerial    = rsget("minusOrderSerial")
                    minusSubtotalprice  = rsget("minusSubtotalprice")
                end if
                rsget.Close
            end if

            if (minusOrderSerial<>"") then
                
                rw "���� �ݾ� : " & FormatNumber(orgsubtotalprice,0) & " / " & FormatNumber(minusSubtotalprice,0)
                rw "<br>�ֹ���ȣ :" & tenOrderSerial & " / " & minusOrderSerial
                rw "<input type='button' value='�Ϸ�ó��' onClick=""finCancelOrd('"&ORG_ord_no&"','"&ord_dtl_sn&"')"">"
                response.write VbCRLF
                response.write "<script language='javascript'>"&VbCRLF
                response.write "function finCancelOrd(ORG_ord_no,ord_dtl_sn){"&VbCRLF
                response.write "    var uri = 'actRegLotteItem.asp?mode=etcSongjangFin&ORG_ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn;"&VbCRLF
                response.write "    var popwin = window.open(uri,'finCancelOrd','width=200,height=200');"&VbCRLF
                response.write "    popwin.focus()"&VbCRLF
                response.write "}"&VbCRLF
                response.write "</script>"&VbCRLF
'2013/02/28 ������ Else�κ� �߰�
'���� ����Ƚ���� 3ȸ�� �����鼭 minusOrderSerial�� ������ �� �ش�
'updateSendState = 901		�߼�ó������ �����ϰ� 
'updateSendState = 902		����� ��������
'updateSendState = 903		��ǰó����
			Else
				Dim errCount
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
									"	<option value='901'>�߼�ó������ �����ϰ�</option>" &_
									"	<option value='902'>����� ��������</option>" &_
									"	<option value='903'>��ǰó����</option>" &_
									"</select>&nbsp;&nbsp;"
					response.write "<input type='button' value='�Ϸ�ó��' onClick=""finCancelOrd2('"&ORG_ord_no&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)"">"
					response.write "<script language='javascript'>"&VbCRLF
					response.write "function finCancelOrd2(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
					response.write "    if(selectValue == ''){"&VbCRLF
					response.write "    	alert('�������ּ���');"&VbCRLF
					response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
					response.write "    	return;"&VbCRLF
					response.write "    }"&VbCRLF
					response.write "    var uri = 'actRegLotteItem.asp?mode=updateSendState&ORG_ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
	                response.write "    var popwin = window.open(uri,'finCancelOrd2','width=200,height=200');"&VbCRLF
	                response.write "    popwin.focus()"&VbCRLF
					response.write "}"&VbCRLF
					response.write "</script>"&VbCRLF
				End If
'2013/02/28 ������ else�κ� �߰� ��
            end if
    	end if
    else
        if (IsAutoScript) then
            rw "�Ե����İ� ����߿� ������ �߻��߽��ϴ�. "&ord_no&" "&ord_dtl_sn
        else    
    	    Response.Write "<script language=javascript>alert('�Ե����İ� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
    	ENd IF
    	dbget.Close: Response.End
    end if
    Set objXML = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->