<%@ language=vbscript %>
<% option explicit %>
<%

response.write "��������޴�"
	response.End

%>


<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- include virtual="/lib/classes/board/offshop_noticecls.asp" -->
<%



dim i, ix
dim page

page = request("page")
if (page = "") then
        page = "1"
end if

'==============================================================================
'��������
dim boardnotice
set boardnotice = New CNotice

boardnotice.FPageSize = 10
boardnotice.FCurrPage = CInt(page)
boardnotice.FRectListAll = "on"
boardnotice.list

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
	<td colspan="4" height="25" align="right">�˻���� : �� <font color="red"><% = boardnotice.FTotalCount %></font>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="50" align="center">��ȣ</td>
	<td align="center">����</td>
	<td width="100" align="center">�۾���</td>
	<td width="100" align="center">��¥</td>
</tr>
<% for i = 0 to (boardnotice.FResultCount - 1) %>
  <tr bgcolor="#FFFFFF">
    <td width="50" align="center"><%= boardnotice.FItemList(i).Fidx %></td>
	<td>&nbsp;
	<a href="offshop_notice_modify.asp?idx=<%= boardnotice.FItemList(i).Fidx %>"><%= boardnotice.FItemList(i).Ftitle %></a>
	<% if datediff("d",boardnotice.FItemList(i).Fregdate,now())<8 then %>
	&nbsp;&nbsp;&nbsp;[<font color=red><b>new</b></font>]
	<% end if %>
	</td>
    <td width="100" align="center"><%= boardnotice.FItemList(i).Fusername %></td>
    <td width="100" align="center"><%= FormatDateTime(boardnotice.FItemList(i).Fregdate,2) %></td>
  </tr>
<% next %>
  <tr bgcolor="#FFFFFF">
    <td align="center" colspan="4">
		<% if boardnotice.HasPreScroll then %>
			<a href="javascript:NextPage('<%= boardnotice.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + boardnotice.StartScrollPage to boardnotice.FScrollCount + boardnotice.StartScrollPage - 1 %>
			<% if ix>boardnotice.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if boardnotice.HasNextScroll then %>
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