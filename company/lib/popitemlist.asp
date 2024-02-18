<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/itemlistcls.asp"-->
<!-- #include virtual="/company/lib/defaulthead.asp"-->
<%
dim page
dim designerid, itemid, itemname, dispyn, sellyn
dim obuyprice

dim frmname,bxitemid, bxitemname, bxitemimage
page = request("page")
designerid = session("ssBctId")
itemid  = request("itemid")
itemname= request("itemname")
dispyn  = request("dispyn")
sellyn  = request("sellyn")

if dispyn="" then dispyn="Y"
if sellyn="" then sellyn="Y"

if designerid="nanishow" then
	designerid="clayplay"
end if

frmname  = request("frmname")
bxitemid = request("bxitemid")
bxitemname = request("bxitemname")
bxitemimage = request("bxitemimage")

if (page="") then page=1

set obuyprice = new CItemList
obuyprice.FCurrPage = page
obuyprice.FPageSize = 15
obuyprice.FSearchItemName = itemname
obuyprice.FSearchDesigner = designerid
obuyprice.FSearchItemid = itemid
obuyprice.FSearchDispYn = dispyn
obuyprice.FSearchSellYn = sellyn
obuyprice.getItemList

dim i
%>
<script language='javascript'>
function selitem(frm){
	var parentfrm;
	parentfrm = opener.document.<%= frmname %>;
	parentfrm.<%= bxitemid %>.value = frm.iitemid.value;
	parentfrm.<%= bxitemname %>.value = frm.iitemname.value;

	parentfrm.<%= bxitemimage %>.src = frm.iimglist.value;

	close();
}
</script>
<table width="560" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<tr>
		<td class="a" >
		상품ID :
		<input type="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9">
		상품명 :
		<input type="text" name="itemname" value="<%= itemname %>" size="10" maxlength="32">
		전시여부 :
		<select name="dispyn">
     	<option value='' selected>선택</option>
     	<option value='Y' <% if dispyn="Y" then response.write "selected" %> >Y</option>
     	<option value='N' <% if dispyn="N" then response.write "selected" %> >N</option>
     	</select>
		판매여부 :
		<select name="sellyn">
     	<option value='' selected>선택</option>
     	<option value='Y' <% if sellyn="Y" then response.write "selected" %> >Y</option>
     	<option value='N' <% if sellyn="N" then response.write "selected" %> >N</option>
     	</select>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>
<table width="560" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td width="80">아이디</td>
	<td width="50">이미지</td>
	<td width="120">아이템명</td>
	<td width="80">판매가</td>
	<td width="80">전시여부</td>
	<td width="80">판매여부</td>
	<td width="80">선택</td>
</tr>
<% for i=0 to obuyprice.FresultCount-1 %>
<form name="sel_<%= obuyprice.FItemList(i).FItemID %>">
<input type="hidden" name="iitemid" value="<%= obuyprice.FItemList(i).FItemID %>">
<input type="hidden" name="iimglist" value="<%= obuyprice.FItemList(i).FImageList %>">
<input type="hidden" name="iimgsmall" value="<%= obuyprice.FItemList(i).FImageSmall %>">
<input type="hidden" name="iitemname" value="<%= obuyprice.FItemList(i).FItemName %>">
<tr>
	<td ><%= obuyprice.FItemList(i).FItemID %></td>
	<td ><img src="<%= obuyprice.FItemList(i).FImageSmall %>" width="50" height="50"></td>
	<td ><%= obuyprice.FItemList(i).FItemName %></td>
	<td ><%= obuyprice.FItemList(i).FSellPrice %></td>
	<td ><%= obuyprice.FItemList(i).FDisplayYn %></td>
	<td ><%= obuyprice.FItemList(i).FSellYn %></td>
	<td ><input type="button" value="선택" onClick="selitem(sel_<%= obuyprice.FItemList(i).FItemID %>)"></td>
</tr>
</form>
<% next %>
<tr>
	<td colspan="14" align="center">
	<% if obuyprice.HasPreScroll then %>
		<a href="?page=<%= obuyprice.StarScrollPage-1 %>&itemid=<%= itemid %>&itemname=<%= itemname %>&designerid=<%= designerid %>&dispyn=<%= dispyn %>&sellyn=<%= sellyn %>&frmname=<%= frmname %>&bxitemid=<%= bxitemid %>&bxitemname=<%= bxitemname %>&bxitemimage=<%= bxitemimage %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + obuyprice.StarScrollPage to obuyprice.FScrollCount + obuyprice.StarScrollPage - 1 %>
		<% if i>obuyprice.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&itemid=<%= itemid %>&itemname=<%= itemname %>&designerid=<%= designerid %>&dispyn=<%= dispyn %>&sellyn=<%= sellyn %>&frmname=<%= frmname %>&bxitemid=<%= bxitemid %>&bxitemname=<%= bxitemname %>&bxitemimage=<%= bxitemimage %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if obuyprice.HasNextScroll then %>
		<a href="?page=<%= i %>&itemid=<%= itemid %>&itemname=<%= itemname %>&designerid=<%= designerid %>&dispyn=<%= dispyn %>&sellyn=<%= sellyn %>&frmname=<%= frmname %>&bxitemid=<%= bxitemid %>&bxitemname=<%= bxitemname %>&bxitemimage=<%= bxitemimage %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set obuyprice = nothing
%>
<!-- #include virtual="/company/lib/defaulttail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->