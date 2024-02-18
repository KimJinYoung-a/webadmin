<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode,cd1,itemid
dim viewidx,disptitle
dim i

mode = request("mode")
cd1 = request("cd1")
itemid = request("itemid")

itemid = Left(itemid,Len(itemid)-1)

dim sqlStr

if (mode = "del") then
	sqlStr = "update [db_cs].[dbo].tbl_qna_preface"
	sqlStr = sqlStr + " set isusing='N'"
	sqlStr = sqlStr + " where idx in (" + itemid + ")"
	rsget.Open sqlStr,dbget,1
	response.write "<script>alert('사용 안함 처리 되었습니다.')</script>"
elseif (mode = "re") then
	sqlStr = "update [db_cs].[dbo].tbl_qna_preface"
	sqlStr = sqlStr + " set isusing='Y'"
	sqlStr = sqlStr + " where idx in (" + itemid + ")"
	response.write "<script>alert('사용 전환 처리 되었습니다.')</script>"
	rsget.Open sqlStr,dbget,1
end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")

%>
<script language="javascript">
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
