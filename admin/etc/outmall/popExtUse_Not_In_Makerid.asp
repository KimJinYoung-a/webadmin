<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/etc/outmallConfirmCls.asp"-->
<!-- #include virtual="/admin/etc/incOutMallCommonFunction.asp"-->
<%
Dim mallgubun, makerid, oOutmall, page, i
mallgubun	= request("mallgubun")
makerid		= request("makerid")
page		= request("page")

If page = "" Then page = 1
If makerid <> "" Then
	response.redirect "popJaeHyu_Not_In_Makerid.asp?mallgubun="&mallgubun&"&makerid="&makerid&"&menupos="&menupos
End If
SET oOutMall = new cOutmall
	oOutMall.FCurrPage			= page
	oOutMall.FPageSize			= 20
	oOutMall.FRectMakerid		= makerid
	oOutMall.FRectMallgubun		= mallgubun
	oOutMall.getExtUseList
%>
<script language='javascript'>
function goPage(pg){
	frm.page.value = pg;
	frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="<%=CurrURL()%>" style="margin:0px;">
<input type="hidden" name="page">
<input type="hidden" name="mallgubun" value="<%=mallgubun%>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
		    <td width="90%">Mall ���� : <%=mallgubun%></td>
		    <td rowspan="4" width="10%"><input type="submit" value="�� ��" style="width:50px;height:50px;"></td>
		</tr>
		<tr>
			<td >
			�귣��ID : <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20"> <input type="button" class="button" value="ID�˻�" onclick="jsSearchBrandID(this.form.name,'makerid');" >
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<br>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" class="a">
		<tr>
			<td width="80%">

			</td>
			<td width="20%" align="right">�귣��� : <b><%=oOutMall.FTotalCount%></b></td>
		</tr>
		</table>
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
    <td width="30%">������</td>
	<td width="30%">�귣��ID</td>
	<td width="20%">�����</td>
	<td width="15%">�����</td>
	<td width="5%">����</td>
</tr>
<%
	If oOutMall.FResultCount > 0 Then
		For i =0 To oOutMall.FResultCount - 1
%>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td><%=oOutMall.FItemList(i).FMallID%></td>
	<td><%=oOutMall.FItemList(i).FMakerid%></td>
	<td><%=oOutMall.FItemList(i).FRegdate%></td>
	<td><%=oOutMall.FItemList(i).FRegUserid%></td>
	<td><a href="popJaeHyu_Not_In_Makerid.asp?mallgubun=<%=mallgubun%>&makerid=<%=oOutMall.FItemList(i).FMakerid%>&menupos=<%= menupos %>">[����]</a></td>
</tr>
<%		Next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9" align="center">
	<% If oOutMall.HasPreScroll Then %>
		<a href="javascript:goPage('<%= oOutMall.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + oOutMall.StartScrollPage To oOutMall.FScrollCount + oOutMall.StartScrollPage - 1 %>
		<% If i>oOutMall.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If oOutMall.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
<%	Else %>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="5" align="center" class="page_link">[�����Ͱ� �����ϴ�.]</td>
	</tr>
<%	End If %>
</table>
</form>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->