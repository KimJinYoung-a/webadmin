<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/wemakepriceEp/epShopCls.asp"-->
<%
Dim mode, NOTmakerid, sqlStr, makerid, orderby
Dim nMaker, page, itemidarr, isusingarr, isusing
page		= requestCheckvar(request("page"), 20)
mode		= request("mode")
NOTmakerid	= requestCheckvar(request("NOTmakerid"), 32)
itemidarr	= request("itemidarr")
isusingarr	= request("isusingarr")
isusing		= requestCheckvar(request("isusing"),1)
makerid		= requestCheckvar(request("makerid"), 32)
orderby		= requestCheckvar(request("orderby"), 10)

If page = "" Then page = 1

SET nMaker = new epShop
	nMaker.FCurrPage					= page
	nMaker.FPageSize					= 20
	nMaker.FMakerId						= makerid
	nMaker.FRectIsusing					= isusing
	nMaker.FRectOrderby					= orderby
    nMaker.EpshopnotinmakeridList
%>
<script language='javascript'>
function goPage(pg){
    var frm = document.frmsearch;
    frm.page.value=pg;
	frm.submit();
}
</script>
<!-- #include virtual="/admin/etc/wemakepriceEp/inc_wemakepriceHead.asp" -->
<!-- �˻� ���� -->
<form name="frmsearch" method="post" action="notinmakerid.asp" >
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="50">
		<table width="100%" class="a">
		<tr>
		    <td width="90%">Mall ���� : ������EP</td>
		    <td rowspan="4" width="10%"><input type="submit" value="�� ��" style="width:50px;height:50px;"></td>
		</tr>
		<tr>
			<td >
				�귣��ID : <input type="text" class="text" name="makerid" value="<%=makerid%>" size="20"> <input type="button" class="button" value="ID�˻�" onclick="jsSearchBrandID(this.form.name,'makerid');" >&nbsp;
				�Ǹſ��� : 
				<select name="isusing" class="select">
					<option value="">-Choice-</option>
					<option value="Y" <%= Chkiif(isusing = "Y", "selected", "") %> >�Ǹ�</option>
					<option value="N" <%= Chkiif(isusing = "N", "selected", "") %> >�Ǹž���</option>
				</select>
				&nbsp;
				���ı��� : 
				<select name="orderby" class="select">
					<option value="">-Choice-</option>
					<option value="lastupdate" <%= Chkiif(orderby = "lastupdate", "selected", "") %> >����������</option>
					<option value="best" <%= Chkiif(orderby = "best", "selected", "") %> >����Ʈ�귣��</option>
				</select>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fitem" method="post" style="margin:0px;">
<input type="hidden" name="sortarr" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
    <td>������</td>
	<td>�귣��ID</td>
	<td>�����</td>
	<td>�����</td>
	<td>����������</td>
	<td>����������</td>
	<td>�Ǹſ���</td>
</tr>
<% If nMaker.FResultCount > 0 Then %>
<% For i = 0 To nMaker.FResultCount - 1 %>
<tr bgcolor="#FFFFFF" height="30" align="center" height="25">
	<td><%=nMaker.FItemList(i).FMallgubun%></td>
	<td><%=nMaker.FItemList(i).FMakerid%></td>
	<td><%=nMaker.FItemList(i).FRegdate%></td>
	<td><%=nMaker.FItemList(i).FRegid%></td>
	<td><%=nMaker.FItemList(i).FLastupdate%></td>
	<td><%=nMaker.FItemList(i).FUpdateid%></td>
	<td>
		<%
			If nMaker.FItemList(i).FIsusing = "Y" Then
				response.write "�Ǹ���"
			Else
				response.write "�Ǹž���"
			End If
		%>
	</td>
</tr>
<% Next %>
<tr height="30">
	<td colspan="16" align="center" bgcolor="#FFFFFF">
	<% If nMaker.HasPreScroll Then %>
		<a href="javascript:goPage('<%= nMaker.StartScrollPage-1 %>');">[pre]</a>
	<% Else %>
		[pre]
	<% End If %>
	<% For i=0 + nMaker.StartScrollPage To nMaker.FScrollCount + nMaker.StartScrollPage - 1 %>
		<% If i>nMaker.FTotalpage Then Exit For %>
		<% If CStr(page)=CStr(i) Then %>
		<font color="red">[<%= i %>]</font>
		<% Else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% End If %>
	<% Next %>
	<% If nMaker.HasNextScroll Then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% Else %>
	[next]
	<% End If %>
	</td>
</tr>
<% Else %>
<tr height="50">
	<td colspan="16" align="center" bgcolor="#FFFFFF">
		��ϵ� �귣�尡 �����ϴ�
	</td>
</tr>
<% End If %>
</form>
</table>
<% SET nMaker = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->