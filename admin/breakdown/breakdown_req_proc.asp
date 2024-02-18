<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/breakdown/breakdownCls.asp"-->
<!-- #include virtual="/lib/email/smslib.asp"-->

<%
	Dim vGubun, strSql, vQuery, vReqIdx, vReqDIdx, vWorkType, vWorkTarget, vReqEquipment, vReqComment, vReqCapimage1, vWorkState, vWorkComment, vWorkPartSn
	vGubun 			= requestCheckVar(Request("gb"),1)
	vReqDIdx 		= requestCheckVar(Request("reqdidx"),10)
	vWorkPartSn		= requestCheckVar(Request("work_part_sn"),20)
	vWorkType		= requestCheckVar(Request("work_type"),2)
	vWorkTarget		= requestCheckVar(Request("work_target"),100)
	vReqEquipment	= requestCheckVar(Request("req_equipment"),2)
	vReqComment		= html2db(Request("req_comment"))
	vReqCapimage1	= requestCheckVar(Request("req_capimage1"),100)
	vWorkState		= requestCheckVar(Request("work_state"),2)
	vWorkComment	= html2db(Request("work_comment"))

	if checkNotValidHTML(vReqComment) or checkNotValidHTML(vWorkComment) then
'		Alert_move "���뿡 ��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���.","about:blank"
'		dbget.close()	:	response.End
		Response.Write "<script>alert('���뿡 HTML�±״� ���� �� �����ϴ�..');location.href='/admin/breakdown/?menupos="&request("menupos")&"';</script>"
		dbget.close()
		Response.End
	end if

	vReqComment = stripHTML(vReqComment)
	vWorkComment = stripHTML(vWorkComment)

	If vGubun = "I" Then	'### �����Է�
		vQuery = "INSERT INTO [db_temp].[dbo].[tbl_breakdown_request](req_userid, req_part_sn) VALUES('" & session("ssBctId") & "', '" & session("ssAdminPsn") & "')"
		dbget.execute vQuery

		strSql = " SELECT SCOPE_IDENTITY() "
		rsget.Open strSql,dbget
 		IF Not rsget.EOF THEN
 			vReqIdx = rsget(0)
 		ELSE
			Call sbAlertMsg ("������ ó���� ������ �߻��Ͽ����ϴ�.[1]", "back", "")
 		END IF
 		rsget.close

		vQuery = "INSERT INTO [db_temp].[dbo].[tbl_breakdown_request_detail](req_idx, work_type, work_target, req_equipment, req_comment, req_captimage, work_state, work_part_sn) " & _
					"VALUES('" & vReqIdx & "', '" & vWorkType & "', '" & vWorkTarget & "', '" & vReqEquipment & "', '" & vReqComment & "', '" & vReqCapimage1 & "', '1', '" & vWorkPartSn & "')"
		dbget.execute vQuery


		If (session("ssAdminPsn") = "30") Then
			'### ��������
			Call SendNormalSMS("010-4782-3272","",session("ssBctCname") & "���� �۾���û-"&fnWorkType(vWorkType)&"("&fnWorkTargetName(vWorkTarget)&") http://scm.10x10.co.kr/admin/breakdown/m.asp?a="&vReqIdx&"")
			Call SendNormalSMS("010-2618-7652","",session("ssBctCname") & "���� �۾���û-"&fnWorkType(vWorkType)&"("&fnWorkTargetName(vWorkTarget)&") http://scm.10x10.co.kr/admin/breakdown/m.asp?a="&vReqIdx&"")
		End If

	ElseIf vGubun = "U" Then	'### �����Է�
		vQuery = "UPDATE [db_temp].[dbo].[tbl_breakdown_request_detail] SET " & _
				 " work_type = '" & vWorkType & "', work_target = '" & vWorkTarget & "', req_equipment = '" & vReqEquipment & "', req_comment = '" & vReqComment & "', work_part_sn = '" & vWorkPartSn & "' " & _
				 " , req_captimage = '" & vReqCapimage1 & "' WHERE idx = '" & vReqDIdx & "' "
		dbget.execute vQuery
	ElseIf vGubun = "C" Then	'### ����Ʈ�� �۾��� �ڸ�Ʈ�� �Է�
		vQuery = "UPDATE [db_temp].[dbo].[tbl_breakdown_request_detail] SET " & _
				 " work_comment = '" & vWorkComment & "', work_lastupdate = getdate() WHERE idx = '" & vReqDIdx & "' "
		dbget.execute vQuery
	ElseIf vGubun = "D" Then
		vQuery = "UPDATE [db_temp].[dbo].[tbl_breakdown_request_detail] SET " & _
				 " isusing = 'N', work_lastupdate = getdate() WHERE idx = '" & vReqDIdx & "' "
		dbget.execute vQuery
		''response.Write vQuery
	ElseIf vGubun = "S" Then	'### ����Ʈ�� �۾��ڰ� �����Ҷ� �Է�
		vQuery = "UPDATE [db_temp].[dbo].[tbl_breakdown_request_detail] SET " & _
				 " work_state = '" & vWorkState & "', work_comment = '" & vWorkComment & "', now_worker = '" & session("ssBctId") & "', work_lastupdate = getdate() WHERE idx = '" & vReqDIdx & "' "
		dbget.execute vQuery

		If vWorkState = "3" Then
			vQuery = "UPDATE [db_temp].[dbo].[tbl_breakdown_request_detail] SET " & _
					 " work_startdate = getdate() WHERE idx = '" & vReqDIdx & "' "
			dbget.execute vQuery
		End If

		If (session("ssAdminPsn") = "30") Then
			'### ��������
			If session("ssBctId") = "kei0329" Then
				Call SendNormalSMS("010-2618-7652","",request("smsmessage") & "-����� �۾���")
			ElseIf session("ssBctId") = "skygo1222" Then
				Call SendNormalSMS("010-4782-3272","",request("smsmessage") & "-�̼��� �۾���")
			End IF
		End If

	End IF

	Response.Write "<script>alert('����Ǿ����ϴ�.');location.href='/admin/breakdown/?menupos="&request("menupos")&"';</script>"
	dbget.close()
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->