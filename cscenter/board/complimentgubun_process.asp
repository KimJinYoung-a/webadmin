<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode,code,masterid
dim cname

masterid = request("masterid")
code = request("code")
cname = request("cname")
mode = request("mode")

dim sqlStr


if (mode = "add") then

	sqlStr = " insert into [db_cs].[dbo].tbl_qna_compliment_gubun(masterid,code,cname)"
	sqlStr = sqlStr + " values('" + Cstr(masterid) + "','"  + Cstr(code) + "','" + Cstr(cname) + "')"
	rsget.Open sqlStr, dbget, 1

else

	sqlStr = "update [db_cs].[dbo].tbl_qna_compliment_gubun"
	sqlStr = sqlStr + " set cname = '" + Cstr(cname) + "'"
	sqlStr = sqlStr + " where code = '" + Cstr(code) + "'"
	sqlStr = sqlStr + " and masterid = '" + Cstr(masterid) + "'"
	rsget.Open sqlStr, dbget, 1

end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('처리 되었습니다.');
location.replace('/cscenter/board/cs_qna_compliment_gubun.asp');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->