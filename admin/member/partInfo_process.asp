<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim part_sn, part_name, part_sort, part_isDel, mode
	Dim SQL, strMsg

	part_sn = Request("part_sn")
	part_name = Request("part_name")
	part_sort = Request("part_sort")
	part_isDel = Request("part_isDel")
	mode = Request("mode")

	'트랜젝션 시작
	dbget.beginTrans

	'// 처리 분기 //
	Select Case mode
		Case "add"
			strMsg = "부서정보가 등록되었습니다."
			SQL =	"Insert into db_partner.dbo.tbl_partInfo " &_
					" (part_name, part_sort, part_isDel) values " &_
					" ('" & part_name & "'" &_
					" ," & part_sort &_
					" ,'N')"
			dbget.Execute(SQL)
		Case "modi"
			strMsg = "부서정보가 수정되었습니다."
			SQL =	"Update db_partner.dbo.tbl_partInfo Set " &_
					"	part_name = '" & part_name & "' " &_
					"	,part_sort = " & part_sort & " " &_
					"Where part_sn=" & part_sn
			dbget.Execute(SQL)
		Case "del"
			strMsg = "처리가 완료되었습니다."
			SQL =	"Update db_partner.dbo.tbl_partInfo Set " &_
					"	part_isDel = '" + part_isDel + "' " &_
					"Where part_sn=" & part_sn
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