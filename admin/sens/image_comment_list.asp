<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.04.15 �ѿ�� ����
'	Description : ��������
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/sens/image_commentcls.asp"-->
<%
dim page
page = request("page")
if page="" then page=1

dim oitem
set oitem = new CItemImage
oitem.FCurrPage=page
oitem.FPageSize=20
oitem.GetItemImageList

dim i
%>
<script language="javascript">
// �űԵ��
function fnNew(){
	document.location.href="image_comment_edit.asp?mode=add&menupos=<%= menupos %>";
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">	
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="�űԵ��" onclick="fnNew();">
		</td>
		<td align="right">
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oitem.fresultcount >0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= oitem.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= oitem.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>
			��ȣ
		</td>
		<td>
			������
		</td>
		<td>
			����ǥ����
		</td>
		<td>
			�����
		</td>
		<td>
			��뿩��
		</td>
    </tr>
	<% for i=0 to oitem.FResultcount -1 %>
	    <tr align="center" bgcolor="#FFFFFF">
			<td>
				<%= oitem.FItemList(i).Fidx %>
			</td>
			<td>
				<a href="image_comment_edit.asp?mode=edit&reviewid=<%= oitem.FItemList(i).Fidx %>&menupos=<%= menupos %>"><img src="<%= oitem.FItemList(i).FIconUrl %>" width =40 height=40 border="0"></a>
			</td>
			<td>
				<%= FormatDateTime(oitem.FItemList(i).Fviewdate,2) %>
			</td>
			<td>
				<%= FormatDateTime(oitem.FItemList(i).FRegDate,2) %>
			</td>
			<td>
				<% if oitem.FItemList(i).FIsusing = "Y" then %>Y<% else %><font color="red">N</font><% end if %>
			</td>
		</tr>
	<% next %>

	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
	
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if oitem.HasPreScroll then %>
				<a href="?page=<%= oitem.StarScrollPage-1 %>&menupos=<%= menupos %>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + oitem.StarScrollPage to oitem.FScrollCount + oitem.StarScrollPage - 1 %>
				<% if i>oitem.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if oitem.HasNextScroll then %>
				<a href="?page=<%= i %>&menupos=<%= menupos %>">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
set oitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->