<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<%
'###############################################
' Discription : ����ī�װ�-��ǰ�Ӽ� ��� Ajax
' History : 2013.08.06 ������ : �ű� ����
'###############################################
Response.CharSet = "euc-kr"

'// ���� ����
Dim dispCate
Dim oAttrib, lp
Dim page

'// �Ķ���� ����
dispCate = request("dispcate")
page = request("page")
if page="" then page="1"

'// ���������� ���
	set oAttrib = new CAttrib
	oAttrib.FPageSize = 40
	oAttrib.FCurrPage = page
	oAttrib.FRectDispCate = dispCate
    oAttrib.GetDispCateAttribList
%>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="2">
		�˻���� : <b><%=oAttrib.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oAttrib.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>ī�װ�</td>
    <td>�Ӽ�����</td>
</tr>
<%	for lp=0 to oAttrib.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><a href="javascript:viewDispCateAttrib('<%=oAttrib.FItemList(lp).Fcatecode%>')"><%="[" & oAttrib.FItemList(lp).Fcatecode & "] " & Replace(oAttrib.FItemList(lp).Fcatename,"^^"," > ") %></a></td>
	<td><%="[" & oAttrib.FItemList(lp).FattribDiv & "] " & oAttrib.FItemList(lp).FattribDivName%></td>
</tr>
<%	Next %>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center">
    <% if oAttrib.HasPreScroll then %>
		<a href="javascript:goPage('<%=dispCate%>','<%= oAttrib.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for lp=0 + oAttrib.StartScrollPage to oAttrib.FScrollCount + oAttrib.StartScrollPage - 1 %>
		<% if lp>oAttrib.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(lp) then %>
		<font color="red">[<%= lp %>]</font>
		<% else %>
		<a href="javascript:goPage('<%=dispCate%>','<%= lp %>');">[<%= lp %>]</a>
		<% end if %>
	<% next %>

	<% if oAttrib.HasNextScroll then %>
		<a href="javascript:goPage('<%=dispCate%>','<%= lp %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
<%
	set oAttrib = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->