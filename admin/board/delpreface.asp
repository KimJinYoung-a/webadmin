<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode,cd1,itemid
dim viewidx,disptitle
dim i

cd1 = request("cd1")
itemid = request("itemid")

itemid = Left(itemid,Len(itemid)-1)

dim sqlStr

	sqlStr = "update [db_cs].[10x10].tbl_qna_preface"
	sqlStr = sqlStr + " set isusing='N'"
	sqlStr = sqlStr + " where idx in (" + itemid + ")"

	rsget.Open sqlStr,dbget,1

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('사용 안함 처리 되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->