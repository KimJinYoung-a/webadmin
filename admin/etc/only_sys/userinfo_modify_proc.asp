<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/etc/only_sys/check_auth.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/etc/only_sys/only_sys_cls.asp"-->

<%
	Dim vQuery, vUserID, vUserName, vRealChk
	vUserID = requestCheckVar(Request("userid"),50)
	vUserName = Trim(requestCheckVar(Request("username"),100))
	vRealChk = requestCheckVar(Request("realnamecheck"),2)
	
	If vUserID = "" Then
		dbget.close()
		Response.Write "<script>alert('잘못된접근');location.href='/admin/etc/only_sys/userinfo_modify.asp';</script>"
		Response.End
	End If
	
	vQuery = "UPDATE [db_user].[dbo].[tbl_user_n] SET username = '" & vUserName & "'"
	If vRealChk <> "" Then
		IF vRealChk <> "Y" AND vRealChk <> "N" Then
			dbget.close()
			Response.Write "<script>alert('realnamecheck가 Y or N 아님.');location.href='/admin/etc/only_sys/userinfo_modify.asp';</script>"
			Response.End
		End If
		vQuery = vQuery & ", realnamecheck = '" & vRealChk & "'"
	End IF
	vQuery = vQuery & "WHERE userid = '" & vUserID & "'"
	dbget.Execute vQuery
	
%>

<script language="javascript">
document.location.href = "/admin/etc/only_sys/userinfo_modify.asp?userid=<%=vUserID%>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->