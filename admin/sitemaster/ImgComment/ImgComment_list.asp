<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/SitemasterClass/ImgCommentCls.asp"-->
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
// 신규등록
function fnNew(){
	document.location.href="ImgComment_edit.asp?mode=add&menupos=<%= menupos %>";
}
</script>
<table border="1" cellpadding="0" cellspacing="0" class="a" bordercolordark="White" bordercolorlight="#B2B2B2">
	<tr>
		<td colspan="9" align="right" height="30"><input type="button" class="button" value="신규등록" onclick="fnNew();"></td>
	</tr>
	<tr align="center">
		<td width="50">
			번호
		</td>
		<td width="100">
			아이콘
		</td>
		<td width="70">
			메인표시일
		</td>
		<td width="70">
			등록일
		</td>
		<td width="60">
			사용여부
		</td>
	</tr>
	<% for i=0 to oitem.FResultcount -1 %>
	<tr align="center">
		<td>
			<%= oitem.FItemList(i).Fidx %>
		</td>
		<td>
			<a href="ImgComment_edit.asp?mode=edit&reviewid=<%= oitem.FItemList(i).Fidx %>&menupos=<%= menupos %>"><img src="<%= oitem.FItemList(i).FIconUrl %>" border="0"></a>
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
	<tr align="center">
		<td colspan="9" height="1" bgcolor="#AAAAAA"></td>
	</tr>
	<tr align="center">
		<td colspan="14" align="center">
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
