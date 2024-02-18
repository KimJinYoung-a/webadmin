<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim itemid, mode, sailyn, dissellprice, disbuyprice
itemid = request("itemid")
mode = request("mode")
sailyn = request("sailyn")
dissellprice = request("dissellprice")
disbuyprice = request("disbuyprice")

dim sqlstr, i

	sqlstr = "update [db_item].[dbo].tbl_item" + VbCrlf
	sqlstr = sqlstr + " set sellcash=orgprice" + VbCrlf
	sqlstr = sqlstr + " ,buycash=orgsuplycash" + VbCrlf
	sqlstr = sqlstr + " ,sailyn='N'" + VbCrlf
	sqlstr = sqlstr + " where itemid=" + itemid

	rsget.Open sqlStr ,dbget,1


dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>

<script language="javascript">
alert('수정되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
