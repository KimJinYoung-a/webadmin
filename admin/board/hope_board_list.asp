<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/hope_boardcls.asp" -->
<%

dim i, ix
dim page

page = request("page")
if (page = "") then
        page = "1"
end if

'==============================================================================
'��������
dim ohope
set ohope = New CHopeBoard

ohope.FPageSize = 10
ohope.FCurrPage = CInt(page)
ohope.list

%>

<script language="JavaScript">
<!--

function NextPage(ipage){
	document.noticeform.page.value= ipage;
	document.noticeform.submit();
}

//-->
</script>
<table width="780" cellspacing="1"  class="a" bgcolor=#3d3d3d>
<form method=post name="noticeform">
<input type="hidden" name="page" value="1">
<tr bgcolor="#FFFFFF">
	<td colspan="4" height="25" align="right">�˻���� : �� <font color="red"><% = ohope.FTotalCount %></font>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="100" align="center">��ȣ</td>
	<td align="center">����</td>
	<td width="100" align="center">�۾���</td>
	<td width="100" align="center">��¥</td>
</tr>
<% for i = 0 to (ohope.FResultCount - 1) %>
  <tr bgcolor="#FFFFFF">
    <td width="50"><%= ohope.FItemList(i).Fidx %></td>
	<td width="300">
	<a href="hope_board_read.asp?idx=<%= ohope.FItemList(i).Fidx %>"><%= ohope.FItemList(i).Ftitle %></a>
	<% if datediff("d",ohope.FItemList(i).Fregdate,now())<8 then %>
	&nbsp;&nbsp;&nbsp;[<font color=red><b>new</b></font>]
	<% end if %>
	</td>
    <td width="100"><%= ohope.FItemList(i).Fusername %></td>
    <td width="100"><%= FormatDateTime(ohope.FItemList(i).Fregdate,2) %></td>
  </tr>
<% next %>
  <tr bgcolor="#FFFFFF">
    <td align="center" colspan="4">
		<% if ohope.HasPreScroll then %>
			<a href="javascript:NextPage('<%= ohope.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + ohope.StarScrollPage to ohope.FScrollCount + ohope.StarScrollPage - 1 %>
			<% if ix>ohope.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if ohope.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
  </tr>
</form>
</table>
<table class="a" width="780">
<tr>
	<td align="right"><a href="offshop_notice_write.asp"><font color="red">�۾���</font></a></td>
</tr>
</table>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->