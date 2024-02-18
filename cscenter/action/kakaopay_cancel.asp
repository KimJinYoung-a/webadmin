<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<script language='javascript'>
function ActCancel(frm){
    if (confirm('승인 취소 하시겠습니까?')){
        frm.action="kakaopay_cancel_process.asp";
        frm.submit();
    }
}

function fnSumit(frm){
    frm.submit();
}
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmCanncel" method="post" action="kakaopay_cancel_process.asp">
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">PG사 ID</td>
    <td bgcolor="#FFFFFF">
    	<input type="text" class="text_ro" name="paygateTid" size="60">
    </td>
</tr>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">환불액</td>
    <td bgcolor="#FFFFFF">
    	<input type="text" class="text_ro" name="refundrequire" size="60">
    </td>
</tr>
<tr height="25">
    <td colspan="2" align="center" bgcolor="#FFFFFF">
    <input type="button" class="button" value=" 승인 취소 " onClick="ActCancel(this.form)">
    </td>
</tr>
</form>
</table>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSearch" method="post">
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">ID</td>
    <td bgcolor="#FFFFFF">
    	<input type="text" class="text_ro" name="userid" size="60"> <input type="button" class="button" value=" 검색 " onClick="fnSumit(this.form)">
    </td>
</tr>
</form>
</table>
<%
dim userid
userid = requestCheckVar(request("userid"),16)

if (userid<>"") then
%>
<table width="300" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>">가격</td>
    <td bgcolor="#FFFFFF">
    	TID
    </td>
</tr>
<%
dim sqlStr
    sqlStr = " select Top 100 price, P_TID"
    sqlStr = sqlStr + " from db_order.dbo.tbl_order_temp"
    sqlStr = sqlStr + " where userid='" + userid + "'"
    sqlStr = sqlStr + " and pggubun='KK'"
    sqlStr = sqlStr + " order by temp_idx desc"
    rsget.Open sqlStr, dbget, 1
    if not rsget.EOF then
        do until rsget.eof
%>
<tr height="25">
    <td bgcolor="<%= adminColor("topbar") %>"><%=rsget("price")%></td>
    <td bgcolor="#FFFFFF"><%=rsget("P_TID")%></td>
</tr>
<%
        rsget.MoveNext
        loop
    end if
    rsget.close
%>
</table>
<%
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/poptail.asp"-->