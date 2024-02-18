<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/que/queItemCls.asp"-->
<%
Dim queGroup, arrRows, i
SET queGroup = new COutmall
	arrRows = queGroup.fnQueGroupCntList
SET queGroup = nothing
%>
<script>
function goifrm(v){
	document.ifrm.location.href = "/admin/etc/que/queActionList.asp?mallid="+v;
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td width="20%" valign="top">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr valign="top">
			<td>
				<table width="100%" align="center" cellpadding="5" cellspacing="0" class="a">
				<tr>
					<td>
						<table width="100%" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
						<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
							<td width="50%" align="center">제휴몰</td>
							<td width="50%" align="center">카운트</td>
						</tr>
				<%
					If isArray(arrRows) Then
						For i =0 To UBound(arrRows,2)
				%>
						<tr height="25" bgcolor="FFFFFF" style="cursor:pointer;" onclick="goifrm('<%= arrRows(0, i) %>');"  onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
							<td><%= arrRows(0, i) %></td>
							<td><%= arrRows(1, i) %></td>
						</tr>
				<%
						Next
					Else
				%>
						<tr height="25" bgcolor="FFFFFF">
							<td colspan="2" height="50" align="center">데이터가 없습니다.</td>
						</tr>
				<%
					End If
				%>
						</table>
					</td>
				</tr>
				</table>
				<table><tr height="7"><td></td></tr></table>
				<br>
				<p>&nbsp;</p>
			</td>
		</tr>
		</table>
	</td>
	<td width="10"></td>
    <td valign="top">
        <iframe id="ifrm" name="ifrm" src="/admin/etc/que/queActionList.asp" name="board" width="100%" height="200" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->