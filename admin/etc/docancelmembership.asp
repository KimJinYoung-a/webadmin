<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/skmembershippointcls.asp"-->
<%
dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim idx,password,mode

idx	= request.form("idx")
password = request.form("password")
mode = request.form("mode")

if C_ADMIN_AUTH<>true then
	if (password<>"ehdvkf") then
		response.write "not valid..."
		response.write "<script>alert('패스워드가 올바르지 않습니다.');</script>"
		response.write "<script>history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

dim osktmem
dim resultcode
if mode="cancel" then
	set osktmem = new CSkMembershipJunmun
	osktmem.CancelPreSavedJunmun idx

	 resultcode = osktmem.GetResultMsg
	response.write resultcode
	set osktmem = Nothing
end if

%>
<script language='javascript'>
<% if (resultcode="[00]정상") then %>
alert('취소되었습니다.');
location.replace('<%= refer %>');
<% end if %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->