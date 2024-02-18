<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/rackipgocls.asp"-->

<%
dim makerid, mwdiv, sellyn, isusing, diffrackcode, research
dim upbaeZeroStock
makerid = request("makerid")
mwdiv = request("mwdiv")
sellyn = request("sellyn")
isusing = request("isusing")
diffrackcode = request("diffrackcode")
research = request("research")
upbaeZeroStock = request("upbaeZeroStock")

dim i
if mwdiv="" then mwdiv="MW"
if (research="") and (diffrackcode="") then diffrackcode="on"
if (research="") and (isusing="") then isusing="Y"

if (upbaeZeroStock = "on") then
	mwdiv = "U"
end if

dim orackcode_branditem
set orackcode_branditem = new CRackIpgo
orackcode_branditem.FRectMakerid = makerid
orackcode_branditem.FRectMwdiv = mwdiv
orackcode_branditem.FRectSellYN = sellyn
orackcode_branditem.FRectIsUsingYN = isusing
orackcode_branditem.FRectdiffrackcode = diffrackcode
''orackcode_branditem.FRectUpbaeZeroStock = upbaeZeroStock
orackcode_branditem.GetRackBrandItemList

%>

<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
}

function PopBrandInfo(v){
    PopBrandInfoEdit(v);

	//var popwin = window.open("/admin/lib/popbrandinfoonly.asp?designer=" + v,"popbrandinfoonly","width=640 height=580 scrollbars=yes resizable=yes");
	//popwin.focus();
}

</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			판매:<% drawSelectBoxSellYN "sellyn", sellyn %>
	     	&nbsp;
			사용:<% drawSelectBoxUsingYN "isusing", isusing %>
			&nbsp;
	     	거래구분:<% drawSelectBoxMWU "mwdiv", mwdiv %>
	     	&nbsp;
        	<input type=checkbox name="diffrackcode" value="on" <% if diffrackcode="on" then response.write "checked" %>>브랜드랙코드와 상이한 상품만
			&nbsp;
        	<input type=checkbox name="upbaeZeroStock" value="on" <% if upbaeZeroStock="on" then response.write "checked" %>>랙코드등록된 재고없는 상품만(업배)
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="120">브랜드ID</td>
    	<td width="40">상품ID</td>
    	<td width="50">브랜드<br>랙코드</td>
    	<td width="50">상품<br>랙코드</td>
    	<td width="50">이미지</td>
    	<td>상품명</td>

    	<td width="30">거래<br>구분</td>

		<td width="30">판매<br>여부</td>
		<td width="30">사용<br>여부</td>
		<td width="30">한정<br>여부</td>

		<td>비고</td>
    </tr>
<% for i=0 to orackcode_branditem.FResultCount - 1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><a href="javascript:PopBrandInfo('<%= orackcode_branditem.FItemList(i).Fmakerid %>')"><%= orackcode_branditem.FItemList(i).Fmakerid %></a></td>
		<td><a href="javascript:PopItemSellEdit('<%= orackcode_branditem.FItemList(i).Fitemid %>');"><%= orackcode_branditem.FItemList(i).Fitemid %></a></td>
		<td><%= orackcode_branditem.FItemList(i).Frackcode %></td>
		<td>
			<% if (orackcode_branditem.FItemList(i).Fitemrackcode <> orackcode_branditem.FItemList(i).Frackcode) then %>
			<b><font color="red"><%= orackcode_branditem.FItemList(i).Fitemrackcode %></font></b>
			<% else %>
			<%= orackcode_branditem.FItemList(i).Fitemrackcode %>
			<% end if %>
		</td>
		<td><img src="<%= orackcode_branditem.FItemList(i).Fimgsmall %>" width=50 height=50></td>
		<td align="left"><%= orackcode_branditem.FItemList(i).Fitemname %></td>

		<td><font color="<%= mwdivColor(orackcode_branditem.FItemList(i).Fmwdiv) %>"><%= mwdivName(orackcode_branditem.FItemList(i).Fmwdiv) %></font></td>


		<td><font color="<%= yncolor(orackcode_branditem.FItemList(i).Fsellyn) %>"><%= orackcode_branditem.FItemList(i).Fsellyn %></font></td>
		<td><font color="<%= yncolor(orackcode_branditem.FItemList(i).FIsusing) %>"><%= orackcode_branditem.FItemList(i).FIsusing %></font></td>
		<td><font color="<%= yncolor(orackcode_branditem.FItemList(i).Flimityn) %>"><%= orackcode_branditem.FItemList(i).Flimityn %></font></td>
		<td></td>
	</tr>
<% next %>
</table>




<%
set orackcode_branditem = Nothing
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
