<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 정기구독 상품검색
' History : 2016.06.16 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim itemid, i, menupos, frmname, itemgubunfrm, itemidfrm, itemoptionfrm, itemnamefrm
	itemid = getNumeric(requestcheckvar(request("itemid"),10))
	menupos = requestcheckvar(request("menupos"),10)
	frmname = request("frmname")
	itemgubunfrm = request("itemgubunfrm")
	itemidfrm = request("itemidfrm")
	itemoptionfrm = request("itemoptionfrm")
	itemnamefrm = request("itemnamefrm")

dim oitem
set oitem = new CItem
	oitem.FRectItemID = itemid
	oitem.frectitemdivexists = "01"

	if itemid<>"" then
		oitem.GetOneItem
	end if

dim oitemoption
set oitemoption = new CItemOption
	oitemoption.FRectItemID = itemid
	oitemoption.frectitemdivexists = "01"

	if itemid<>"" then
		oitemoption.GetItem_Option
	end if

%>
<script type="text/javascript">

	function popselected(itemgubun,itemid,itemoption,itemname){
		var frmname; frmname="<%= frmname %>";
		var itemgubunfrm; itemgubunfrm="<%= itemgubunfrm %>";
		var itemidfrm; itemidfrm="<%= itemidfrm %>";
		var itemoptionfrm; itemoptionfrm="<%= itemoptionfrm %>";
		var itemnamefrm; itemnamefrm="<%= itemnamefrm %>";

		eval("opener."+ frmname + "." + itemgubunfrm).value=itemgubun;
		eval("opener."+ frmname + "." + itemidfrm).value=itemid;
		eval("opener."+ frmname + "." + itemoptionfrm).value=itemoption;
		eval("opener."+ frmname + "." + itemnamefrm).value=itemname;
		self.close();
	}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		※ 상품구분 중에 일반상품만 검색 됩니다. 마일리지상품이나 기타 상품들은 검색되지 않습니다.
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="frmname" value="<%= frmname %>">
<input type="hidden" name="itemgubunfrm" value="<%= itemgubunfrm %>">
<input type="hidden" name="itemidfrm" value="<%= itemidfrm %>">
<input type="hidden" name="itemoptionfrm" value="<%= itemoptionfrm %>">
<input type="hidden" name="itemnamefrm" value="<%= itemnamefrm %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="3">
		<img src="/images/icon_star.gif" border="0" align="absbottom">
		<b>상품 검색</b>
	</td>
</tr>
<% if (oitem.FResultCount<1) then %>
<tr height="25" bgcolor="FFFFFF">
	<td width="120" bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
	<td>
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="6" maxlength="9">
		<input type="button" class="button" value="검색" onClick="document.frm.submit();">
	</td>
</tr>
<tr bgcolor="FFFFFF">
    <td colspan="3" align="center">[검색 결과가 없습니다.]</td>
</tr>
<% else %>
<tr height="25" bgcolor="FFFFFF">
	<td width="120" bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
	<td>
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="6" maxlength="9">
		<input type="button" class="button" value="검색" onClick="document.frm.submit();">
	</td>
	<td rowspan="4" width="100" align="right">
	    <img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100">
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">상품명</td>
	<td><%= oitem.FOneItem.FItemName %></td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">브랜드ID</td>
	<td><%= oitem.FOneItem.FMakerid %></td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">소비자가/매입가</td>
	<td>
	    <% if (oitem.FOneItem.Fsailyn="Y") then %>
			<%= FormatNumber(oitem.FOneItem.FOrgPrice,0) %> / <%= FormatNumber(oitem.FOneItem.Forgsuplycash,0) %>
			&nbsp;
			<%= fnPercent(oitem.FOneItem.Forgsuplycash,oitem.FOneItem.FOrgPrice,1) %>
			&nbsp;&nbsp;
			<%= fnColor(oitem.FOneItem.FMwDiv,"mw") %>

			<br>

			<font color=#F08050>(할)<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %></font>
			&nbsp;
			<%= fnPercent(oitem.FOneItem.FBuycash,oitem.FOneItem.FSellcash,1) %>
			&nbsp;&nbsp;
			<%= fnColor(oitem.FOneItem.FMwDiv,"mw") %>

			<% if (oitem.FOneItem.IsCouponItem) then %>
			<br><font color=#10F050>(쿠) <%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %></font>
			<% end if %>
		<% else %>
			<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
			&nbsp;
			<%= fnPercent(oitem.FOneItem.FBuycash,oitem.FOneItem.FSellcash,1) %>
			&nbsp;&nbsp;
			<%= fnColor(oitem.FOneItem.FMwDiv,"mw") %>

			<% if (oitem.FOneItem.IsCouponItem) then %>
			<br><font color=#10F050>(쿠) <%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %> <!-- / <%= FormatNumber(oitem.FOneItem.Fcouponbuyprice) %> --> &nbsp;<%= oitem.FOneItem.GetCouponDiscountStr %> 할인 </font>
			<% end if %>
		<% end if %>
	</td>
</tr>
<% end if %>
</form>
</table>

<% if oitem.FResultCount>0 then %>
	<br>
	<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			검색결과 : <b><%= oitemoption.FtotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td>옵션코드</td>
	    <td>옵션명</td>
	    <td>옵션<br>사용여부</td>
		<td>비고</td>
	</tr>
	
	<% if oitemoption.FtotalCount>0 then %>
		<%
		'/단일 옵션
		if not(oitemoption.IsMultipleOption) then
		%>
			<% for i=0 to oitemoption.FResultCount - 1 %>
			<tr bgcolor="<%=chkIIF(oitemoption.FItemList(i).Foptisusing="Y","#FFFFFF","#DDDDDD")%>" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='<%=chkIIF(oitemoption.FItemList(i).Foptisusing="Y","#FFFFFF","#DDDDDD")%>';>
			    <td align="center">
			    	<%= oitemoption.FItemList(i).fitemoption %>
			    </td>
			    <td>
			    	<%= oitemoption.FItemList(i).foptionname %>
			    </td>
			    <td align="center">
			    	<%= oitemoption.FItemList(i).Foptisusing %>
			    </td>
			    <td align="center">
			    	<input type="button" onclick="popselected('10','<%= itemid %>','<%= oitemoption.FItemList(i).fitemoption %>','<%= replace(oitem.FOneItem.FItemName,"""","'") %>');" value="선택" class="button">
			    </td>
			</tr>
			<% Next %>
		<% else %>
			<tr bgcolor="#FFFFFF">
				<td colspan="20" align="center">검색결과가 없습니다.</td>
			</tr>
		<% end if %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="20" align="center">검색결과가 없습니다.</td>
		</tr>
	<% end if %>
	</table>
<% end if %>

<%
set oitem = Nothing
set oitemoption = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
