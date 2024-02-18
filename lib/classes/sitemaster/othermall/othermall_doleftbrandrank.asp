<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode,makerid,idx,cdl
dim i

mode = request("mode")
cdl = request("cdl")
makerid = request("makerid")

idx = request("itemid")
If idx <> "" then
idx = Left(idx,Len(idx)-1)
End if

dim sqlStr
if mode="del" then
	sqlStr = "delete from [db_contents].[dbo].tbl_category_left_brand_rank"
	sqlStr = sqlStr + " where idx in (" + idx + ")"
	response.write sqlStr
	rsget.Open sqlStr,dbget,1
elseif mode="add" then
	sqlStr = "insert into [db_contents].[dbo].tbl_category_left_brand_rank(cdl,makerid)"
	sqlStr = sqlStr + " values('" + Cstr(cdl) + "','" + makerid + "')"
	rsget.Open sqlStr,dbget,1
end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('적용 되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->