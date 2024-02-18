<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
Dim idx, mode, cName, mallid, strSql
idx			= request("idx")
mallid		= request("mallid")
mode		= request("mode")
cName		= db2html(request("cName"))

If mode = "chgName" Then
	strSql = ""
	strSql = strSql & " UPDATE db_etcmall.[dbo].[tbl_Outmall_option_Manager] SET "
	strSql = strSql & " itemnameChange = '"&cName&"' "
	strSql = strSql & " WHERE idx = '"&idx&"' "
	dbget.Execute strSql
	response.write 	"<script language='javascript'>alert('저장되었습니다');</script>"
End If
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->