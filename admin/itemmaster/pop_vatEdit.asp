<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
	Dim vItemID, vVat, vAction
	vAction	= requestCheckvar(request("action"),6)
	vItemID = requestCheckvar(request("itemid"),10)
	vVat	= requestCheckvar(request("vat"),2)
	
	If vAction = "update" Then
		Call UpdateVat()
	End If
%>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm1" action="pop_vatEdit.asp" method="post">
<input type="hidden" name="action" value="update">
<input type="hidden" name="itemid" value="<%=vItemID%>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>">
	<td width="50" bgcolor="<%= adminColor("gray") %>">상품코드</td>
	<td bgcolor="white"><%=vItemID%></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>">
	<td width="50" bgcolor="<%= adminColor("gray") %>">과세면세</td>
	<td bgcolor="white">
		<input type="radio" name="vat" value="Y" <%=CHKIIF(vVat="Y","checked","")%>>과세&nbsp;&nbsp;&nbsp;
		<input type="radio" name="vat" value="N" <%=CHKIIF(vVat="N","checked","")%>>면세
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("topbar") %>">
	<td colspan="2" align="right" bgcolor="white"><input type="submit" class="button" value="저장">&nbsp;</td>
</tr>
</form>
</table>

<%
	Function UpdateVat()
		Dim vQuery, vItemID, vVat, vAction
		vAction	= requestCheckvar(request("action"),6)
		vItemID = requestCheckvar(request("itemid"),10)
		vVat	= requestCheckvar(request("vat"),2)

		If vAction = "update" Then
			vQuery = "UPDATE [db_item].[dbo].[tbl_item] SET vatinclude = '" & vVat & "', lastupdate = getdate() WHERE itemid = '" & vItemID & "'"
			dbget.Execute vQuery
			
			Response.Write "<script language='javascript'>alert('저장되었습니다.');opener.document.location.reload();window.close();</script>"
			dbget.close()
			Response.End
		End IF
	End Function
%>

<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->