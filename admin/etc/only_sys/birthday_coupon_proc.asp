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
	Dim vQuery, vUserID
	vUserID = requestCheckVar(Request("userid"),50)
	
	If vUserID = "" Then
		dbget.close()
		Response.Write "<script>alert('잘못된접근');location.href='/admin/etc/only_sys/birthday_coupon.asp';</script>"
		Response.End
	End If
	
	vQuery = "INSERT INTO [db_user].dbo.tbl_user_coupon" & vbCrLf
	vQuery = vQuery & "(masteridx,userid,coupontype,couponvalue,couponname,minbuyprice" & vbCrLf
	vQuery = vQuery & ",startdate,expiredate,targetitemlist,couponmeaipprice,reguserid)" & vbCrLf
	vQuery = vQuery & "select 126,userid,'2','5000','[생일쿠폰] 생일을 축하드려요','30000'" & vbCrLf
	vQuery = vQuery & "	,convert(varchar(10),getdate(),21) + ' 00:00:00',convert(varchar(10),dateadd(d,14,getdate()),21) + ' 23:59:59','',0,'system'" & vbCrLf
	vQuery = vQuery & "from db_user.dbo.tbl_user_n" & vbCrLf
	vQuery = vQuery & "WHERE userid = '" & vUserID & "'"
	dbget.Execute vQuery
	
%>

<script language="javascript">
document.location.href = "/admin/etc/only_sys/birthday_coupon.asp?userid=<%=vUserID%>";
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->