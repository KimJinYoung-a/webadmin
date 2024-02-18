<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode,idx,masterid
dim gubun,contents

idx = request("idx")
masterid = request("masterid")
gubun = request("gubun")
contents = request("contents")
mode = request("mode")

dim sqlStr


if (mode = "add") then

	sqlStr = " insert into [db_cs].[10x10].tbl_qna_preface(masterid,gubun,contents)"
	sqlStr = sqlStr + " values('" + Cstr(masterid) + "','" + Cstr(gubun) + "','" + Cstr(contents) + "')"
	rsget.Open sqlStr, dbget, 1

else

	sqlStr = "update [db_cs].[10x10].tbl_qna_preface"
	sqlStr = sqlStr + " set gubun='" + Cstr(gubun) + "'"
	sqlStr = sqlStr + ",contents = '" + Cstr(contents) + "'"
	sqlStr = sqlStr + " where idx = " + Cstr(idx)
	rsget.Open sqlStr, dbget, 1

end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('처리 되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->