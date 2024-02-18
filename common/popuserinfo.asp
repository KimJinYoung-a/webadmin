<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
dim uid
uid = requestCheckVar(request("uid"),32)

dim username,userphone,userhp,useremail
dim resultcount
dim sqlStr

resultcount = 0

sqlStr = "select top 1 * from [db_user].[dbo].tbl_user_n"
sqlStr = sqlStr + " where userid='" + uid + "'"
rsget.open sqlStr,dbget,1
resultcount = rsget.RecordCount
if not rsget.Eof then
	username	= db2html(rsget("username"))
	userphone	= rsget("userphone")
	userhp		= rsget("usercell")
	useremail	= db2html(rsget("usermail"))
end if
rsget.close
%>
<% if resultcount>0 then %>
<script language='javascript'>
opener.ReActUser('<%= username %>','<%= userphone %>','<%= userhp %>','<%= useremail %>');
window.close();
</script>
<% else %>
<script language='javascript'>
alert('해당 아이디가 존재하지 않습니다.');
window.colse();
</script>
<% end if %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->