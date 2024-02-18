<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	'// 변수 선언 //
	dim commCd, msg, SQL

	commCd = RequestCheckvar(Request("commCd"),10)

	SQL = "Select count(commCd) as cnt From db_academy.dbo.tbl_CommCd where commCd='" & commCd & "'"
	rsACADEMYget.Open sql, dbACADEMYget, 1
		if rsACADEMYget("cnt")>0 then
			msg = "중복된 코드입니다."
		else
			msg = "사용 가능한 코드입니다."
		end if
	rsACADEMYget.close

	'//결과 메시지 출력
	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"</script>"
	dbget.close()	:	response.End
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->