<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim posit_sn, posit_name, posit_isDel, mode
	Dim SQL, strMsg

	posit_sn = Request("posit_sn")
	posit_name = Request("posit_name")
	posit_isDel = Request("posit_isDel")
	mode = Request("mode")

	'트랜젝션 시작
	dbget.beginTrans

	'// 처리 분기 //
	Select Case mode
		Case "add"
			strMsg = "직급정보가 등록되었습니다."
			SQL =	"Insert into db_partner.dbo.tbl_positInfo " &_
					" (posit_name, posit_isDel) values " &_
					" ('" & posit_name & "'" &_
					" ,'N')"
			dbget.Execute(SQL)
		Case "modi"
			strMsg = "직급정보가 수정되었습니다."
			SQL =	"Update db_partner.dbo.tbl_positInfo Set " &_
					"	posit_name = '" & posit_name & "' " &_
					"Where posit_sn=" & posit_sn
			dbget.Execute(SQL)
		Case "del"
			strMsg = "처리가 완료되었습니다."
			SQL =	"Update db_partner.dbo.tbl_positInfo Set " &_
					"	posit_isDel = '" + posit_isDel + "' " &_
					"Where posit_sn=" & posit_sn
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