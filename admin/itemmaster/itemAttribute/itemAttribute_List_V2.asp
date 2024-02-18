<%@ language=vbscript %>
<% option explicit %>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<%

Dim page, i
dim masteridx

page = requestCheckVar(request("page"), 32)
masteridx = requestCheckVar(request("masteridx"), 32)

if page="" then page="1"

dim oAttrib
set oAttrib = new CAttrib
oAttrib.FPageSize = 20
oAttrib.FCurrPage = page
oAttrib.FRectMasterIDX = masteridx

oAttrib.GetAttribList_V2

%>
<!-- ��� �˻��� ���� -->
<form name="frm" method="get" action="" style="margin:0;">
<input type="hidden" name="research" value="on" />
<input type="hidden" name="page" value="" />
<input type="hidden" name="menupos" value="<%= request("menupos") %>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
	    �Ӽ�����:
	    <%= drawSelectAttributeMaster("masteridx", masteridx, "") %>
	</td>
	<td width="80" rowspan="2" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" value="�˻�" />
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 5px 0;">
<tr>
    <td align="left">
    	<input type="button" value="��������" class="button" onClick="saveList()" title="�켱���� �� ��뿩�θ� �ϰ������մϴ�.">
    </td>
    <td align="right">
    	<input type="button" value="�űԼӼ� ���" class="button" onClick="popAttribute('');">
    </td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ��� ���� -->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="attrArr">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="table-layout: fixed;">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		�˻���� : <b><%=oAttrib.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oAttrib.FtotalPage%></b>
	</td>
</tr>
<colgroup>
	<col width="40" />
    <col width="50" />
    <col width="80" />
    <col width="*" />
    <col width="80" />
    <col width="80" />
    <col width="*" />
    <col width="80" />
	<col width="160" />
</colgroup>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><span class="ui-icon ui-icon-arrowthick-2-n-s"></span></td>
    <td><input type="checkbox" name="allChk" onclick="chkAllItem()"></td>
    <td>IDX</td>
    <td>�Ӽ�����</td>
    <td>ǥ�ü���</td>
    <td>IDX</td>
    <td>�Ӽ���</td>
    <td>ǥ�ü���</td>
	<td><span class="ui-icon ui-icon-wrench"></span></td>
</tr>
<tbody id="attrList">
<%	for i = 0 to oAttrib.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><span class="rowHaddle ui-icon ui-icon-grip-solid-horizontal" style="cursor:grab;" title="���ļ����� �����մϴ�."></span></td>
    <td><input type="checkbox" name="chkCd" value="<%= oAttrib.FItemList(i).Fidx %>" /></td>
    <td><%= oAttrib.FItemList(i).Fidx %></td>
    <td><%= oAttrib.FItemList(i).FattMasterName %></td>
    <td><%= oAttrib.FItemList(i).Fdispno %></td>
    <td><%= oAttrib.FItemList(i).Fdetailidx %></td>
    <td><%= oAttrib.FItemList(i).FattDetailName %></td>
    <td><%= oAttrib.FItemList(i).Fdetaildispno %></td>
	<td>
	</td>
</tr>
<%	Next %>
</tbody>
<tr bgcolor="#FFFFFF">
    <td colspan="9" align="center">
    <% if oAttrib.HasPreScroll then %>
		<a href="javascript:goPage('<%= oAttrib.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i = 0 + oAttrib.StartScrollPage to oAttrib.FScrollCount + oAttrib.StartScrollPage - 1 %>
		<% if i>oAttrib.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oAttrib.HasNextScroll then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
</form>
<!-- ��� �� -->

<%
	set oAttrib = Nothing
%>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
