<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/inc_dailyAuthCheck.asp"-->
<%
'// ��������
Dim LotteGoodNo, LotteStatCd
Dim strSql, actCnt, lp, waitCnt, rjtCnt
Dim AssignedCNT, GoodsCount

actCnt = 0			'�ǰ��ŰǼ�
waitCnt = 0
on Error Resume Next

strSql = ""
strSql = strSql & " SELECT TOP 100 itemid, LtiMallTmpGoodNo FROM db_item.dbo.tbl_ltiMall_regItem "
strSql = strSql & " WHERE LtiMallStatCd in ('10','20','51','52') "
strSql = strSql & " and dateDiff(hh,IsNULL(lastConfirmdate,'2001-01-01'),getdate()) > 5 "
strSql = strSql & " and LtiMallTmpGoodNo is Not NULL"
strSql = strSql & " ORDER BY IsNULL(lastConfirmdate,'2001-01-01') ASC, regdate ASC"
rsget.Open strSql,dbget,1
If Not(rsget.EOF or rsget.BOF) Then
	'// �Ե����̸� ���û�ǰ��ȣ ��������
	Do Until rsget.EOF
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.Open "GET", ltiMallAPIURL & "/openapi/getRdToPrGoodsNoApi.lotte?subscriptionId=" & ltiMallAuthNo & "&goods_req_no=" & rsget("LtiMallTmpGoodNo"), false
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objXML.Send()
			If objXML.Status = "200" Then
				Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
					xmlDOM.async = False
					xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

					If Err <> 0 then
						Response.Write "<script language=javascript>alert('�Ե����̸� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����.1');</script>"
						Set xmlDOM = Nothing
						Set objXML = Nothing
						dbget.Close: Response.End
					End If

					GoodsCount 		= Trim(xmlDOM.getElementsByTagName("GoodsCount").item(0).text)			'�˻���
					LotteGoodNo		= Trim(xmlDOM.getElementsByTagName("GoodsNo").item(0).text)			'���û�ǰ��ȣ
					LotteStatCd		= Trim(xmlDOM.getElementsByTagName("ConfStatCd").item(0).text)		'���������ڵ�

					strSql =""
					strSql = strSql & " UPDATE db_item.dbo.tbl_ltiMall_regItem "
					strSql = strSql & "	SET lastConfirmdate = getdate() "
					If (LotteStatCd <> "") then
						If LotteStatCd = "30" Then
							LotteStatCd = "7"
						End If
						strSql = strSql & "	,LtiMallStatCd='" & LotteStatCd & "' "
					End If
		
					If (LotteGoodNo > "0") and (LotteGoodNo <> "") Then
						strSql = strSql & " ,LtiMallGoodNo='" & LotteGoodNo & "' "
					End If
	
					strSql = strSql & " WHERE itemid='" & rsget("itemid") & "'"
					dbget.Execute strSql, AssignedCNT
					If (LotteStatCd = "30") Then
					    actCnt = actCnt + AssignedCNT
					ElseIf (LotteStatCd = "20") Then
					    waitCnt = waitCnt + AssignedCNT
					ElseIf (LotteStatCd = "40") Then
					    rjtCnt = rjtCnt + AssignedCNT
					End If

				Set xmlDOM = Nothing
			Else
				Response.Write "<script language=javascript>alert('�Ե����̸��� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����.2');</script>"
				dbget.Close: Response.End
			End If
		Set objXML = Nothing
		rsget.MoveNext
	Loop
End If

rsget.Close

'##### DB ���� ó�� #####
If Err.Number = 0 Then
	If actCnt > 0 or waitCnt > 0 or rjtCnt > 0 Then
	    If (IsAutoScript) then
	        rw  "OK|"&actCnt & "���� ����." & waitCnt& "���� ���δ��." & rjtCnt & "�� �ݷ�"
	    Else
	        If (session("ssBctID") = "icommang" or session("ssBctID") = "kjy8517") Then
	            rw actCnt & "���� ����." & waitCnt& "���� ���δ��." & rjtCnt & "�� �ݷ�"
	        Else
    		    Response.Write "<script language=javascript>alert('" & actCnt & "���� ���������� ���ŵǾ����ϴ�.');parent.history.go(0);</script>"
    	    End If
    	End if
	Else
	    If (IsAutoScript) Then
	        rw  "OK|"&actCnt & "���� ����." & waitCnt& "���� ���δ��."  & rjtCnt & "�� �ݷ�"
	    Else
    		Response.Write "<script language=javascript>alert('������ �ӽõ�� ��ǰ�� �����ϴ�.');parent.history.go(0);</script>"
    	End If
	End If
Else
    If (IsAutoScript) Then
        rw "S_ERR|ó�� �߿� ������ �߻��߽��ϴ�"
    Else
        Response.Write "<script language=javascript>alert('ó�� �߿� ������ �߻��߽��ϴ�.\n�����ڿ��� �������ּ���.');</script>"
    End If
End If

on Error Goto 0
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->