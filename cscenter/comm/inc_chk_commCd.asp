<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	'// ���� ���� //
	dim comm_cd, msg, SQL

	comm_cd = Request("comm_cd")

	SQL = "Select count(comm_cd) as cnt From db_cs.dbo.tbl_cs_comm_code where comm_cd='" & comm_cd & "'"
	rsget.Open sql, dbget, 1
		if rsget("cnt")>0 then
			msg = "�ߺ��� �ڵ��Դϴ�."
		else
			msg = "��� ������ �ڵ��Դϴ�."
		end if
	rsget.close

	'//��� �޽��� ���
	response.write	"<script language='javascript'>" &_
					"	alert('" & msg & "');" &_
					"</script>"
	dbget.close()	:	response.End
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
