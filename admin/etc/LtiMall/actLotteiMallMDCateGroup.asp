<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/inc_dailyAuthCheck.asp"-->
<%
Dim SGrpCnt, GrpCnt, lp, SGrpInfo, GrpInfo
Dim MDCode, groupCode, SuperGroupName, GroupName
Dim strSql, actCnt
actCnt = 0		'�ǰ��ŰǼ�
MDCode = Request("mdcd")

'// �Ե����̸� MD��ǰ�� ��ȸ
Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", ltiMallAPIURL & "/openapi/searchMDGsgrListOpenApi.lotte?subscriptionId=" & ltiMallAuthNo & "&md_id=" & MDCode, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()

If objXML.Status = "200" Then
	'//���޹��� ���� Ȯ��
'	Response.contentType = "text/xml; charset=euc-kr"
'	response.write BinaryToText(objXML.ResponseBody, "euc-kr")
'	response.End

	'XML�� ���� DOM ��ü ����
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
	'on Error Resume Next
		SGrpCnt = xmlDOM.getElementsByTagName("SuperGroupCount").item(0).text		'������ǰ�� ī��Ʈ
		If Err <> 0 then
			Response.Write "<script language=javascript>alert('�Ե����̸� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
			Response.End
		end if

		If SGrpCnt > 0 Then
			dbget.beginTrans
			strSql = "UPDATE db_temp.dbo.tbl_lotteiMall_MDCateGrp SET isUsing = 'N', lastupdate = getdate() WHERE isUsing = 'Y' and MDCode = '"&MDCode&"'"
			dbget.Execute(strSql)

			Set SGrpInfo = xmlDOM.getElementsByTagName("SuperGroupInfo")
			For each SGNodes in SGrpInfo
				SuperGroupName	= Trim(SGNodes.getElementsByTagName("SuperGroupName").item(0).text)		'������ǰ����
				GrpCnt			= SGNodes.getElementsByTagName("SubGroupCount").item(0).text			'������ǰ����
				If GrpCnt > 0 Then
					Set GrpInfo = SGNodes.getElementsByTagName("SubGroupInfo")
					For each SubNodes in GrpInfo
						groupCode	= Trim(SubNodes.getElementsByTagName("GroupCode").item(0).text)		'�׷��ڵ�
						GroupName	= Trim(SubNodes.getElementsByTagName("GroupName").item(0).text)		'�׷��
						strSql = "SELECT count(*) FROM db_temp.dbo.tbl_lotteiMall_MDCateGrp WHERE groupCode = '" & groupCode & "' and MDCode='" & MDCode & "'"
						rsget.Open strSql,dbget,1
						If rsget(0) > 0 Then
							strSql = "UPDATE db_temp.dbo.tbl_lotteiMall_MDCateGrp SET isUsing='Y' WHERE groupCode = '" & groupCode & "' and MDCode='" & MDCode & "'"
							dbget.Execute(strSql)
							actCnt = actCnt+1
						Else
							strSql = "INSERT INTO db_temp.dbo.tbl_lotteiMall_MDCateGrp (groupCode, MDCode, SuperGroupName, GroupName) VALUES " &_
									" ('" & groupCode & "'" &_
									", '" & MDCode & "'" &_
									", '" & html2db(SuperGroupName) & "'" &_
									", '" & html2db(GroupName) & "')"
							dbget.Execute(strSql)
							actCnt = actCnt+1
						End If
						rsget.Close
					Next
				End If
			Next
			Set SGrpInfo = Nothing

			'##### DB ���� ó�� #####
		    If Err.Number = 0 Then
		    	dbget.CommitTrans				'Ŀ��(����)
		    	Response.Write "<script language=javascript>alert('" & actCnt & "���� ���������� ���ŵǾ����ϴ�.');parent.history.go(0);</script>"
		    Else
		        dbget.RollBackTrans				'�ѹ�(�����߻���)
		        Response.Write "<script language=javascript>alert('�ڷ� ���� �߿� ������ �߻��߽��ϴ�.\n�����ڿ��� �������ּ���.');</script>"
		    End If
		Else
			Response.Write "<script language=javascript>alert('" & MDCode & " MD�� �����Ǿ��ִ� MD��ǰ���� �����ϴ�.\n�Ե����� ����ڿ��� �������ּ���.');</script>"
			Response.End
		End If
	'on Error Goto 0
	Set xmlDOM = Nothing
Else
	Response.Write "<script language=javascript>alert('�Ե����̸��� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
	Response.End
End If
Set objXML = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->