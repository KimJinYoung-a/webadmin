<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode,cd1,itemid
dim viewidx,disptitle
dim i,masterid

cd1 = request("cd1")
masterid = request("masterid")
itemid = request("itemid")

itemid = replace(itemid,",","','")

itemid = Left(itemid,Len(itemid)-2)

dim sqlStr

	sqlStr = "delete from [db_cs].[10x10].tbl_qna_preface_gubun"
	sqlStr = sqlStr + " where code in ('" + itemid + ")"
	sqlStr = sqlStr + " and masterid = '" + Cstr(masterid) + "'"

	rsget.Open sqlStr,dbget,1

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('사용 안함 처리 되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->