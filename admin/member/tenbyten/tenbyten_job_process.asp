<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim job_sn, job_name, job_isDel, mode
	Dim SQL, strMsg

	mode 		= requestCheckvar(request("mode"),32)
	job_sn 		= requestCheckvar(request("job_sn"),32)
	job_name 	= requestCheckvar(request("job_name"),32)
	job_isDel 	= requestCheckvar(request("job_isDel"),32)

	'트랜젝션 시작
	dbget.beginTrans

	'// 처리 분기 //
	Select Case mode
		Case "add"
			strMsg = "직책정보가 등록되었습니다."
			SQL =	"Insert into db_partner.dbo.tbl_jobInfo " &_
					" (job_name, job_isDel) values " &_
					" ('" & job_name & "'" &_
					" ,'N')"
			dbget.Execute(SQL)
		Case "modi"
			strMsg = "직책정보가 수정되었습니다."
			SQL =	"Update db_partner.dbo.tbl_jobInfo Set " &_
					"	job_name = '" & job_name & "' " &_
					"Where job_sn=" & job_sn
			dbget.Execute(SQL)
		Case "del"
			strMsg = "처리가 완료되었습니다."
			SQL =	"Update db_partner.dbo.tbl_jobInfo Set " &_
					"	job_isDel = '" + job_isDel + "' " &_
					"Where job_sn=" & job_sn
			dbget.Execute(SQL)
	End Select

	'오류검사 및 실행
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)

		response.write	"<script language='javascript'>" &_
						"	alert('" & strMsg & "');" &_
						"	opener.history.go(0);" &_
						"	self.close();" &_
						"</script>"
	Else
	    dbget.RollBackTrans				'롤백(에러발생시)

		response.write	"<script language='javascript'>" &_
						"	alert('처리중 에러가 발생했습니다.');" &_
						"	history.back();" &_
						"</script>"

	End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->