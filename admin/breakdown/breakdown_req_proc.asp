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
'		Alert_move "내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요.","about:blank"
'		dbget.close()	:	response.End
		Response.Write "<script>alert('내용에 HTML태그는 넣을 수 없습니다..');location.href='/admin/breakdown/?menupos="&request("menupos")&"';</script>"
		dbget.close()
		Response.End
	end if

	vReqComment = stripHTML(vReqComment)
	vWorkComment = stripHTML(vWorkComment)

	If vGubun = "I" Then	'### 새로입력
		vQuery = "INSERT INTO [db_temp].[dbo].[tbl_breakdown_request](req_userid, req_part_sn) VALUES('" & session("ssBctId") & "', '" & session("ssAdminPsn") & "')"
		dbget.execute vQuery

		strSql = " SELECT SCOPE_IDENTITY() "
		rsget.Open strSql,dbget
 		IF Not rsget.EOF THEN
 			vReqIdx = rsget(0)
 		ELSE
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
 		END IF
 		rsget.close

		vQuery = "INSERT INTO [db_temp].[dbo].[tbl_breakdown_request_detail](req_idx, work_type, work_target, req_equipment, req_comment, req_captimage, work_state, work_part_sn) " & _
					"VALUES('" & vReqIdx & "', '" & vWorkType & "', '" & vWorkTarget & "', '" & vReqEquipment & "', '" & vReqComment & "', '" & vReqCapimage1 & "', '1', '" & vWorkPartSn & "')"
		dbget.execute vQuery


		If (session("ssAdminPsn") = "30") Then
			'### 문자전송
			Call SendNormalSMS("010-4782-3272","",session("ssBctCname") & "님의 작업신청-"&fnWorkType(vWorkType)&"("&fnWorkTargetName(vWorkTarget)&") http://scm.10x10.co.kr/admin/breakdown/m.asp?a="&vReqIdx&"")
			Call SendNormalSMS("010-2618-7652","",session("ssBctCname") & "님의 작업신청-"&fnWorkType(vWorkType)&"("&fnWorkTargetName(vWorkTarget)&") http://scm.10x10.co.kr/admin/breakdown/m.asp?a="&vReqIdx&"")
		End If

	ElseIf vGubun = "U" Then	'### 수정입력
		vQuery = "UPDATE [db_temp].[dbo].[tbl_breakdown_request_detail] SET " & _
				 " work_type = '" & vWorkType & "', work_target = '" & vWorkTarget & "', req_equipment = '" & vReqEquipment & "', req_comment = '" & vReqComment & "', work_part_sn = '" & vWorkPartSn & "' " & _
				 " , req_captimage = '" & vReqCapimage1 & "' WHERE idx = '" & vReqDIdx & "' "
		dbget.execute vQuery
	ElseIf vGubun = "C" Then	'### 리스트에 작업자 코멘트만 입력
		vQuery = "UPDATE [db_temp].[dbo].[tbl_breakdown_request_detail] SET " & _
				 " work_comment = '" & vWorkComment & "', work_lastupdate = getdate() WHERE idx = '" & vReqDIdx & "' "
		dbget.execute vQuery
	ElseIf vGubun = "D" Then
		vQuery = "UPDATE [db_temp].[dbo].[tbl_breakdown_request_detail] SET " & _
				 " isusing = 'N', work_lastupdate = getdate() WHERE idx = '" & vReqDIdx & "' "
		dbget.execute vQuery
		''response.Write vQuery
	ElseIf vGubun = "S" Then	'### 리스트에 작업자가 컨텍할때 입력
		vQuery = "UPDATE [db_temp].[dbo].[tbl_breakdown_request_detail] SET " & _
				 " work_state = '" & vWorkState & "', work_comment = '" & vWorkComment & "', now_worker = '" & session("ssBctId") & "', work_lastupdate = getdate() WHERE idx = '" & vReqDIdx & "' "
		dbget.execute vQuery

		If vWorkState = "3" Then
			vQuery = "UPDATE [db_temp].[dbo].[tbl_breakdown_request_detail] SET " & _
					 " work_startdate = getdate() WHERE idx = '" & vReqDIdx & "' "
			dbget.execute vQuery
		End If

		If (session("ssAdminPsn") = "30") Then
			'### 문자전송
			If session("ssBctId") = "kei0329" Then
				Call SendNormalSMS("010-2618-7652","",request("smsmessage") & "-유재규 작업중")
			ElseIf session("ssBctId") = "skygo1222" Then
				Call SendNormalSMS("010-4782-3272","",request("smsmessage") & "-이성모 작업중")
			End IF
		End If

	End IF

	Response.Write "<script>alert('저장되었습니다.');location.href='/admin/breakdown/?menupos="&request("menupos")&"';</script>"
	dbget.close()
	Response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->