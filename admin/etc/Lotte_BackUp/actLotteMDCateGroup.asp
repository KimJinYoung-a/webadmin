<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<%
	'// ��������
	dim SGrpCnt, GrpCnt, lp, SGrpInfo, GrpInfo
	dim MDCode, groupCode, SuperGroupName, GroupName
	dim strSql, actCnt

	actCnt = 0		'�ǰ��ŰǼ�

	MDCode = Request("mdcd")

	'// �Ե����� MD��ǰ�� ��ȸ
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", lotteAPIURL & "/openapi/searchMDGsgrListOpenApi.lotte?subscriptionId=" & lotteAuthNo & "&md_id=" & MDCode, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()
rw lotteAPIURL & "/openapi/searchMDGsgrListOpenApi.lotte?subscriptionId=" & lotteAuthNo & "&md_id=" & MDCode
'response.end

	If objXML.Status = "200" Then

		'//���޹��� ���� Ȯ��
		'Response.contentType = "text/xml; charset=euc-kr"
		'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		'response.End

		'XML�� ���� DOM ��ü ����
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		
		'on Error Resume Next
			SGrpCnt = xmlDOM.getElementsByTagName("SuperGroupCount").item(0).text		'������ǰ�� ī��Ʈ
			if Err<>0 then
				Response.Write "<script language=javascript>alert('�Ե����� ��� �м� �߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
				Response.End
			end if
rw "SGrpCnt="&SGrpCnt
			if SGrpCnt>0 then
				'// Ʈ������ ����
				dbget.beginTrans

				'��� ��ǰ����뿩�� ����
				strSql = "update db_temp.dbo.tbl_lotte_MDCateGrp Set isUsing='N', lastupdate=getdate() Where isUsing='Y' and MDCode='"&MDCode&"'"
				''rw strSql
				dbget.Execute(strSql)

				'// SGrpInfo Loop
				Set SGrpInfo = xmlDOM.getElementsByTagName("SuperGroupInfo")
				for each SGNodes in SGrpInfo
					SuperGroupName	= Trim(SGNodes.getElementsByTagName("SuperGroupName").item(0).text)		'������ǰ����
					GrpCnt			= SGNodes.getElementsByTagName("SubGroupCount").item(0).text			'������ǰ����

					if GrpCnt>0 then
						Set GrpInfo = SGNodes.getElementsByTagName("SubGroupInfo")
						for each SubNodes in GrpInfo
							groupCode	= Trim(SubNodes.getElementsByTagName("GroupCode").item(0).text)		'�׷��ڵ�
							GroupName	= Trim(SubNodes.getElementsByTagName("GroupName").item(0).text)		'�׷��
	
							'��ǰ�����翩�� Ȯ��
							strSql = "Select count(*) From db_temp.dbo.tbl_lotte_MDCateGrp Where groupCode='" & groupCode & "' and MDCode='" & MDCode & "'"
							rsget.Open strSql,dbget,1
		
							if rsget(0)>0 then
								'// ���� -> �����
								strSql = "update db_temp.dbo.tbl_lotte_MDCateGrp Set isUsing='Y' Where groupCode='" & groupCode & "' and MDCode='" & MDCode & "'"
								dbget.Execute(strSql)
								actCnt = actCnt+1
							else
								'// ���� -> �űԵ��
								strSql = "Insert into db_temp.dbo.tbl_lotte_MDCateGrp (groupCode, MDCode, SuperGroupName, GroupName) values " &_
										" ('" & groupCode & "'" &_
										", '" & MDCode & "'" &_
										", '" & html2db(SuperGroupName) & "'" &_
										", '" & html2db(GroupName) & "')"
								dbget.Execute(strSql)
								actCnt = actCnt+1
							end if
		
							rsget.Close
						Next
					end if
				next
				Set SGrpInfo = Nothing

				'##### DB ���� ó�� #####
			    If Err.Number = 0 Then
			    	dbget.CommitTrans				'Ŀ��(����)
			    	Response.Write "<script language=javascript>alert('" & actCnt & "���� ���������� ���ŵǾ����ϴ�.');parent.history.go(0);</script>"
			    Else
			        dbget.RollBackTrans				'�ѹ�(�����߻���)
			        Response.Write "<script language=javascript>alert('�ڷ� ���� �߿� ������ �߻��߽��ϴ�.\n�����ڿ��� �������ּ���.');</script>"
			    End If
			else
				Response.Write "<script language=javascript>alert('" & MDCode & " MD�� �����Ǿ��ִ� MD��ǰ���� �����ϴ�.\n�Ե����� ����ڿ��� �������ּ���.');</script>"
				Response.End
			end if
		'on Error Goto 0

		Set xmlDOM = Nothing
	else
		Response.Write "<script language=javascript>alert('�Ե����İ� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');</script>"
		Response.End
	end if
	Set objXML = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->