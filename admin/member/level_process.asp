<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim level_sn, level_no, level_name, level_isDel, strLevel, mode
	Dim SQL, strMsg, strCnt

	level_sn = Request("level_sn")
	level_no = Request("level_no")
	level_name = Request("level_name")
	level_isDel = Request("level_isDel")
	strLevel = Request("strLevel")
	mode = Request("mode")

	'트랜젝션 시작
	dbget.beginTrans

	'// 처리 분기 //
	Select Case mode
		Case "add"
			strMsg = "등급정보가 등록되었습니다."
			SQL =	"Insert into db_partner.dbo.tbl_level " &_
					" (level_no, level_name, level_isDel) values " &_
					" (" & level_no &_
					" ,'" & level_name & "'" &_
					" ,'N')"
			dbget.Execute(SQL)
		Case "modi"
			strMsg = "등급정보가 수정되었습니다."
			SQL =	"Update db_partner.dbo.tbl_level Set " &_
					"	level_no = " & level_no &_
					"	, level_name = '" & level_name & "' " &_
					"Where level_sn=" & level_sn
			dbget.Execute(SQL)
		Case "del"
			strMsg = "처리가 완료되었습니다."
			SQL =	"Update db_partner.dbo.tbl_level Set " &_
					"	level_isDel = '" + level_isDel + "' " &_
					"Where level_sn=" & level_sn
			dbget.Execute(SQL)
		Case "dp_chk"
			SQL =	"Select count(*) From db_partner.dbo.tbl_level " &_
					"Where level_no=" & strLevel
			rsget.Open SQL,dbget,1
				strCnt = rsget(0)
			rsget.Close
			if strCnt>0 then
				response.write	"<script language='javascript'>" &_
								"	alert('이미 사용중인 등급번호입니다.\n다른 등급번호를 선택해주십시요.');" &_
								"</script>"
					response.End
			else
				response.write	"<script language='javascript'>" &_
								"	alert('사용가능한 등급번호입니다.');" &_
								"	parent.document.frm_level.level_no.value='" & strLevel & "';" &_
								"	parent.document.frm_level.level_name.focus();" &_
								"</script>"
					response.End
			end if
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