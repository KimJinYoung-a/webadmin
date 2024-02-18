<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim idx, gubun
idx = request("idx")

dim sqlStr, orderserial

sqlStr = " select top 1 T.orderserial "
sqlStr = sqlStr + " from "
sqlStr = sqlStr + " 	[db_order].[dbo].[tbl_order_temp] T "
sqlStr = sqlStr + " 	join [db_order].[dbo].[tbl_onlineApp_log] a on a.PGkey = T.P_TID "
sqlStr = sqlStr + " where "
sqlStr = sqlStr + " 	1 = 1 "
sqlStr = sqlStr + " 	and a.idx = " & idx

rsget.CursorLocation = adUseClient
rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
if not rsget.Eof Then
    orderserial = rsget("orderserial")
end if
rsget.close

%>
<script language="javascript">

</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="logidx" value="<%= idx %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td>주문번호</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" height="30">
    	<td>
			<%= orderserial %>
		</td>
    </tr>
	</form>
</table>

<br>

<div align="center">
<input type="button" class="button" value=" 닫 기 " onClick="opener.focus(); window.close();">
</div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
