<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim vitemid, vmallid, madeDate, useDate, mode, strSQL
mode	= request("mode")
vitemid = request("itemid")
vmallid = request("mallid")
madeDate = request("madeDate")
useDate = request("useDate")

If mode = "" Then
	strSQL = ""
	strSQL = strSQL & " SELECT TOP 1 * FROM db_item.dbo.tbl_OutMall_etcLink WHERE mallid = '"&vmallid&"' and linkgbn = 'auctionDate' and valtype = 4 and itemid = '"&vitemid&"' "
	rsget.Open strSQL, dbget, 1
	IF not rsget.EOF THEN
		madeDate = rsget("madeDate")
		useDate = rsget("useDate")
	END IF
	rsget.Close
ElseIf mode = "I" Then
	strSQL = ""
	strSQL = strSQL & " IF EXISTS (SELECT TOP 1 * FROM db_item.dbo.tbl_OutMall_etcLink WHERE mallid = '"&vmallid&"' and linkgbn = 'auctionDate' and valtype = 4 and itemid = '"&vitemid&"') "
	strSQL = strSQL & " 	UPDATE db_item.dbo.tbl_OutMall_etcLink SET "
	strSQL = strSQL & " 	madeDate = '"&madeDate&"' "
	strSQL = strSQL & " 	,useDate = '"&useDate&"' "
	strSQL = strSQL & "		WHERE itemid = '"&vitemid&"' "
	strSQL = strSQL & " 	and mallid = '"&vmallid&"' and linkgbn = 'auctionDate' and valtype = 4  "
	strSQL = strSQL & " ELSE "
	strSQL = strSQL & " 	INSERT INTO db_item.dbo.tbl_OutMall_etcLink(itemid, mallid, linkgbn, linkyn, valtype, intval, shortVal, textVal, stDt, edDt, regdate, madeDate, useDate) VALUES "
	strSQL = strSQL & " 	('"&vitemid&"', '"&vmallid&"', 'auctionDate', 'Y', '4', '', '', '', getdate(), '9999-12-31 00:00:00.000', getdate(), '"&madeDate&"', '"&useDate&"') "
	dbget.execute strSql
	response.write "<script>alert('저장 되었습니다');top.window.close();</script>"
End If
%>
<script language="javascript">
function fnSaveForm() {
	var frm = document.frmAct;
	if(confirm("저장 하시겠습니까?")) {
		frm.action="popAuctionDate.asp";
		frm.submit();
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmAct" target="xLink">
<input type="hidden" value="I" name="mode">
<input type="hidden" value="<%=vitemid%>" name="itemid">
<input type="hidden" value="<%=vmallid%>" name="mallid">
<tr bgcolor="#FFFFFF">
	<td>상품코드</td>
	<td><%= vitemid %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>제조일</td>
	<td>
		<input type="text" name="madeDate" size="10" maxlength=10 readonly value="<%=Left(madeDate,10)%>">
		<a href="javascript:calendarOpen(frmAct.madeDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>유효일</td>
	<td>
		<input type="text" name="useDate" size="10" maxlength=10 readonly value="<%=Left(useDate,10)%>">
		<a href="javascript:calendarOpen(frmAct.useDate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center"><input type="button" class="button" onclick="fnSaveForm();" value="저장"></td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="1" width="0" height="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
