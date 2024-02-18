<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionoffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/lib/classes/board/offshop_galleryCls.asp" -->


<%

dim i, j, page, shopid, isusing, research

page        = request("page")
shopid      = request("shopid")
isusing     = request("isusing")
research    = request("research")

if page="" then page=1
if (research="") and (isusing="") then isusing="Y"

dim offnews
set offnews = New COffshopGallery
offnews.FRectShopid = shopid
offnews.FPageSize = 20
offnews.FCurrPage = page
offnews.FScrollCount = 10
offnews.GetOffshopGalleryList

%>

<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 9pt; }
    INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #CACACA; color: #000000; }
-->
</STYLE>

<script language='javascript'>
function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}
</script>

<table width="100%" border="1" bordercolordark="White" bordercolorlight="black" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF">
  <tr bgcolor="#DDDDFF" height="25">
    <td width="50" align="center">번호</td>
    <td width="100" align="center">Shop</td>
    <td align="center">Image</td>
    <td width="100" align="center">사용여부</td>
    <td width="100" align="center">작성일</td>
    <td width="50" align="center">수정</td>
  </tr>
<% for i = 0 to (offnews.FResultCount - 1) %>
  <tr height="20">
    <td align="center">&nbsp;<%= offnews.FItemList(i).FIdx %></td>
    <td align="center"><%= offnews.FItemList(i).FShopID %></td>
    <td align="center">
    	<img src="<%= offnews.FItemList(i).FImageURL %>" width="100" height="100">
    </td>
	<td align="center"><%= offnews.FItemList(i).FUseYN%></td>
    <td align="center"><%= FormatDate(offnews.FItemList(i).FRegdate, "0000.00.00") %></td>
    <td align="center"><input type="button" value="수정" onClick="location.href='offshop_gallery_write.asp?idx=<%= offnews.FItemList(i).FIdx %>'"></td>
  </tr>
<% next %>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td align="center" height="30">
		<% if offnews.HasPreScroll then %>
			<a href="javascript:NextPage('<%= offnews.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + offnews.StartScrollPage to offnews.FScrollCount + offnews.StartScrollPage - 1 %>
			<% if i>offnews.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if offnews.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<tr>
	<td align="right"><a href="offshop_gallery_write.asp"><font color="red">Gallery 등록</font></a>&nbsp;&nbsp;&nbsp;</td>
</tr>
</table>

<form name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="<%=page%>">
</form>

<% set offnews = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->