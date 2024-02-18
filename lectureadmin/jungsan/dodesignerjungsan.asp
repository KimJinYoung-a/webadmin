<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<%
dim idx, mode, rd_state
idx = RequestCheckvar(request("idx"),10)
mode  = RequestCheckvar(request("mode"),16)
rd_state = RequestCheckvar(request("rd_state"),6)

dim sqlStr
if mode="statechange" then
	sqlStr = "update [db_jungsan].[dbo].tbl_designer_jungsan_master"
	sqlStr = sqlStr + " set finishflag='" + rd_state + "'"
	sqlStr = sqlStr + " where id=" + CStr(idx)

	rsget.Open sqlStr,dbget,1
end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>

<script language="javascript">
alert('저장 되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->