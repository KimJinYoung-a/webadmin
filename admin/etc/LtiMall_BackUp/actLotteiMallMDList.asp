<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/inc_dailyAuthCheck.asp"-->
<%
Dim MDCode, MDName, SellFeeType, NormalSellFee, EventSellFee
Dim strSql, actCnt
actCnt = 0		'�ǰ��ŰǼ�

'// �Ե����̸� ���MD ��ȸ
Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", ltiMallAPIURL & "/openapi/searchMDListOpenApi.lotte?subscriptionId=" & ltiMallAuthNo, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()
If objXML.Status = "200" Then
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
	On Error Resume Next
		MDCnt = xmlDOM.getElementsByTagName("MDCount").item(0).text		'���MD��
		If Err <> 0 Then
			Response.Write "<script language=javascript>alert('�Ե����̸� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
			Response.End
		End If

		If MDCnt > 0 Then
			'// Ʈ������ ����
			dbget.beginTrans
			'��� MD��뿩�� ����
			strSql = "UPDATE db_temp.dbo.tbl_lotteiMall_MDInfo SET isUsing = 'N', lastupdate = getdate() WHERE isUsing = 'Y' "
			dbget.Execute(strSql)

			'// MDInfo Loop
			Set MDInfo = xmlDOM.getElementsByTagName("MDInfo")
			For each SubNodes in MDInfo
				MDCode			= Trim(SubNodes.getElementsByTagName("MDCode").item(0).text)		'���MD�ڵ�
				MDName			= Trim(SubNodes.getElementsByTagName("MDName").item(0).text)		'���MD��
				SellFeeType		= Trim(SubNodes.getElementsByTagName("SellFeeType").item(0).text)	'��������
				NormalSellFee	= Trim(SubNodes.getElementsByTagName("NormalSellFee").item(0).text)	'���������
				EventSellFee	= Trim(SubNodes.getElementsByTagName("EventSellFee").item(0).text)	'��������

				'MD���翩�� Ȯ��
				strSql = "Select count(MDCode) From db_temp.dbo.tbl_lotteiMall_MDInfo Where MDCode='" & MDCode & "'"
				rsget.Open strSql,dbget,1

				If rsget(0) > 0 Then
					'// ���� -> �����
					strSql = "update db_temp.dbo.tbl_lotteiMall_MDInfo Set isUsing = 'Y' Where MDCode = '" & MDCode & "'"
					dbget.Execute(strSql)
				Else
					'// ���� -> �űԵ��
					strSql = "Insert into db_temp.dbo.tbl_lotteiMall_MDInfo (MDCode, MDName, SellFeeType, NormalSellFee, EventSellFee) VALUES " &_
							" ('" & MDCode & "'" &_
							", '" & html2db(MDName) & "'" &_
							", '" & SellFeeType & "'" &_
							", '" & NormalSellFee & "'" &_
							", '" & EventSellFee & "')"
					dbget.Execute(strSql)
					actCnt = actCnt+1
				End If

				rsget.Close
			Next
			Set MDInfo = Nothing

			'##### DB ���� ó�� #####
		    If Err.Number = 0 Then
		    	dbget.CommitTrans				'Ŀ��(����)
		    	Response.Write "<script language=javascript>alert('" & actCnt & "���� ���������� ���ŵǾ����ϴ�.');parent.history.go(0);</script>"
		    Else
		        dbget.RollBackTrans				'�ѹ�(�����߻���)
		        Response.Write "<script language=javascript>alert('�ڷ� ���� �߿� ������ �߻��߽��ϴ�.\n�����ڿ��� �������ּ���.');</script>"
		    End If
		Else
			Response.Write "<script language=javascript>alert('�Ե����Ŀ� �����Ǿ��ִ� ���MD�� �����ϴ�.\n�Ե����� ����ڿ��� �������ּ���.');</script>"
			Response.End
		End If
	On Error Goto 0

	Set xmlDOM = Nothing
Else
	Response.Write "<script language=javascript>alert('�Ե����İ� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
	Response.End
End If
Set objXML = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->