<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<%
	'// ��������
	dim MDCode, MDName, SellFeeType, NormalSellFee, EventSellFee
	dim strSql, actCnt

	actCnt = 0		'�ǰ��ŰǼ�

	'// �Ե����� ���MD ��ȸ
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", lotteAPIURL & "/openapi/searchMDListOpenApi.lotte?subscriptionId=" & lotteAuthNo, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()
	If objXML.Status = "200" Then

		'//���޹��� ���� Ȯ��
		'Response.contentType = "text/xml; charset=euc-kr"
		'response.write BinaryToText(objXML.ResponseBody, "euc-kr")

		'XML�� ���� DOM ��ü ����
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		
		on Error Resume Next
			MDCnt = xmlDOM.getElementsByTagName("MDCount").item(0).text		'���MD��
			if Err<>0 then
				Response.Write "<script language=javascript>alert('�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
				Response.End
			end if

			if MDCnt>0 then
				'// Ʈ������ ����
				dbget.beginTrans

				'��� MD��뿩�� ����
				strSql = "update db_temp.dbo.tbl_lotte_MDInfo Set isUsing='N', lastupdate=getdate() Where isUsing='Y' "
				dbget.Execute(strSql)

				'// MDInfo Loop
				Set MDInfo = xmlDOM.getElementsByTagName("MDInfo")
				for each SubNodes in MDInfo
					MDCode			= Trim(SubNodes.getElementsByTagName("MDCode").item(0).text)		'���MD�ڵ�
					MDName			= Trim(SubNodes.getElementsByTagName("MDName").item(0).text)		'���MD��
					SellFeeType		= Trim(SubNodes.getElementsByTagName("SellFeeType").item(0).text)	'��������
					NormalSellFee	= Trim(SubNodes.getElementsByTagName("NormalSellFee").item(0).text)	'���������
					EventSellFee	= Trim(SubNodes.getElementsByTagName("EventSellFee").item(0).text)	'��������

					'MD���翩�� Ȯ��
					strSql = "Select count(MDCode) From db_temp.dbo.tbl_lotte_MDInfo Where MDCode='" & MDCode & "'"
					rsget.Open strSql,dbget,1

					if rsget(0)>0 then
						'// ���� -> �����
						strSql = "update db_temp.dbo.tbl_lotte_MDInfo Set isUsing='Y' Where MDCode='" & MDCode & "'"
						dbget.Execute(strSql)
					else
						'// ���� -> �űԵ��
						strSql = "Insert into db_temp.dbo.tbl_lotte_MDInfo (MDCode, MDName, SellFeeType, NormalSellFee, EventSellFee) values " &_
								" ('" & MDCode & "'" &_
								", '" & html2db(MDName) & "'" &_
								", '" & SellFeeType & "'" &_
								", '" & NormalSellFee & "'" &_
								", '" & EventSellFee & "')"
						dbget.Execute(strSql)
						actCnt = actCnt+1
					end if

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
			else
				Response.Write "<script language=javascript>alert('�Ե����Ŀ� �����Ǿ��ִ� ���MD�� �����ϴ�.\n�Ե����� ����ڿ��� �������ּ���.');</script>"
				Response.End
			end if
		on Error Goto 0

		Set xmlDOM = Nothing
	else
		Response.Write "<script language=javascript>alert('�Ե����İ� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
		Response.End
	end if
	Set objXML = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->