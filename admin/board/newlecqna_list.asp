<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/item_qnacls.asp" -->
<%
dim notupbea, imageon, mifinish, makerid, research, page
notupbea = request("notupbea")
mifinish = request("mifinish")
makerid = request("makerid")
research = request("research")
page = request("page")

if page="" then page=1
if research="" and mifinish="" then mifinish="on"

dim itemqna
set itemqna = new CItemQna
itemqna.FPageSize = 20
itemqna.FCurrpage = page
itemqna.FReckMiFinish = mifinish
itemqna.FRectMakerid = makerid
itemqna.FRectOnlyTenBeasong = notupbea
itemqna.CItemDiv=90
itemqna.ItemQnaList

dim i
%>
<script language='javascript'>
function NextPage(page){
	frm.page.value=page;
	frm.submit();
}
</script>
<table width="98%" border="0" cellpadding="3" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" >
	<tr>
		<td class="a" >
		브랜드:<% drawSelectBoxDesignerwithName "makerid",makerid  %>
		<input type=checkbox name=notupbea <% if notupbea="on" then response.write "checked" %> >업체배송검색안함
		<input type=checkbox name=mifinish <% if mifinish="on" then response.write "checked" %> >미처리만검색
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>
<table width="98%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#CCCCCC">
  <tr bgcolor="#DDDDFF" height="25">
    <td width="130" align="center">고객명(아이디)</td>
    <td align="center" width="120">분류</td>
    <td align="center">내용</td>
    <td width="50" align="center">상품ID</td>
    <td width="80" align="center">브랜드</td>
    <td width="36" align="center">배송</td>
    <td width="80" align="center">작성일</td>
    <td width="80" align="center">답변자</td>
    <td width="80" align="center">답변일</td>
  </tr>
<% for i = 0 to (itemqna.FResultCount - 1) %>
  <tr height="20" bgcolor="#FFFFFF" >
    <td>&nbsp;<%= itemqna.FItemList(i).Fusername %>(<%= itemqna.FItemList(i).Fuserid %>)</td>
    <td><%= itemqna.FItemList(i).GetQaName %></td>
    <td >&nbsp;<a href="newitemqna_view.asp?id=<%= itemqna.FItemList(i).Fid %>&menupos=<%= menupos %>"><%= db2html(itemqna.FItemList(i).Ftitle) %></a></td>
    <td align="center"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= itemqna.FItemList(i).FItemID %>" target=_blank><%= itemqna.FItemList(i).FItemID %></a></td>
    <td align="center"><%= itemqna.FItemList(i).FBrandName %></td>
    <td align="center"><font color="<%= itemqna.FItemList(i).GetDeliveryTypeColor %>"><%= itemqna.FItemList(i).GetDeliveryTypeName %></font></td>
    <td align="center"><%= FormatDate(itemqna.FItemList(i).Fregdate, "0000-00-00") %></td>
    <td align="center"><%= itemqna.FItemList(i).Freplyuser %></td>
    <td align="center">
    <% if Not IsNULL(itemqna.FItemList(i).FReplydate) then %>
    <%= FormatDate(itemqna.FItemList(i).FReplydate, "0000-00-00") %>
    <% end if %>
    </td>
  </tr>
<% next %>
</table>
<table width="800" border="0" cellpadding="0" cellspacing="1" class="a" height=30>
	<tr bgcolor="#FFFFFF" >
		<td class="link_black" align="center" colspan=9>
			<% if itemqna.HasPreScroll then %>
				<a href="javascript:NextPage('<%= CStr(itemqna.StartScrollPage - 1) %>')">[prev]</a>
			<% else %>
				[prev]
			<% end if %>
			<% for i = itemqna.StartScrollPage to (itemqna.StartScrollPage + itemqna.FScrollCount - 1) %>
			  <% if (i > itemqna.FTotalPage) then Exit For %>
			  <% if CStr(i) = CStr(itemqna.FCurrPage) then %>
				 [<%= i %>]
			  <% else %>
				 <a href="javascript:NextPage('<%= i %>')" class="id_link">[<%= i %>]</a>
			  <% end if %>
			<% next %>
			<% if itemqna.HasNextScroll then %>
				<a href="javascript:NextPage('<%= CStr(itemqna.StartScrollPage + itemqna.FScrollCount) %>')">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>
<%
set itemqna = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp" -->