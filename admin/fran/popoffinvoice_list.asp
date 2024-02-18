<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->

<%

dim page, shopid, openerfrm, mode
dim research, i

page = RequestCheckVar(request("page"), 32)
shopid = RequestCheckVar(request("shopid"), 32)
research = RequestCheckVar(request("research"), 32)
openerfrm = RequestCheckVar(request("frm"), 32)
mode = RequestCheckVar(request("mode"), 32)

if (page = "") then
	page = 1
end if



'================================================================================
dim ocoffinvoice

set ocoffinvoice = new COffInvoice

ocoffinvoice.FRectShopid = shopid

ocoffinvoice.FCurrPage = page
ocoffinvoice.Fpagesize = 25

ocoffinvoice.GetMasterList

%>
<script language='javascript'>

function InsertInvoiceInfo(frm) {
	var openerfrm = eval(opener.<%= openerfrm %>);
	var mode = "<%= mode %>";

	if (mode == "INVOICE") {
		openerfrm.exporteraddr.value = frm.exporteraddr.value;
		openerfrm.riskmesseraddr.value = frm.riskmesseraddr.value;
		openerfrm.notifyaddr.value = frm.notifyaddr.value;
		openerfrm.portname.value = frm.portname.value;
		openerfrm.destinationname.value = frm.destinationname.value;
		openerfrm.carriername.value = frm.carriername.value;
		openerfrm.lccomment.value = frm.lccomment.value;
		openerfrm.lcbank.value = frm.lcbank.value;
		openerfrm.comment.value = frm.comment.value;
		// openerfrm.goodscomment1.value = frm.goodscomment1.value;
		// openerfrm.goodscomment2.value = frm.goodscomment2.value;
	} else {
		openerfrm.delivermethod.value = frm.delivermethod.value;
		openerfrm.exportmethod.value = frm.exportmethod.value;
		openerfrm.jungsantype.value = frm.jungsantype.value;
		openerfrm.priceunit.value = frm.priceunit.value;
		openerfrm.exchangerate.value = frm.exchangerate.value;
	}

	opener.focus();
	window.close();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="frm" value="<%= openerfrm %>">
	<input type="hidden" name="mode" value="<%= mode %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			ShopID : 
			<% 'drawSelectBoxOffShop "shopid",shopid %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="13">
			검색결과 : <b><%= ocoffinvoice.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= ocoffinvoice.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="40">IDX</td>
		<td>샵아이디</td>
		<td>인보이스<br>DATE</td>
		<td>운송방법</td>
		<td>운임부담</td>
		<td>정산시기</td>
		<td>작성화폐</td>
		<td>박스<br>수량</td>
		<td>총상품금액</td>
		<td>총운임</td>
		<td width="60">작성자</td>
		<td width="80">등록일</td>
		<td>비고</td>
	</tr>
	<% if ocoffinvoice.FResultCount >0 then %>
	<% for i=0 to ocoffinvoice.FResultcount-1 %>
	<form name="frmMaster<%= i %>" method="post" action="">
	<input type="hidden" name="exporteraddr" value="<%= ocoffinvoice.FItemList(i).Fexporteraddr %>">
	<input type="hidden" name="riskmesseraddr" value="<%= ocoffinvoice.FItemList(i).Friskmesseraddr %>">
	<input type="hidden" name="notifyaddr" value="<%= ocoffinvoice.FItemList(i).Fnotifyaddr %>">
	<input type="hidden" name="portname" value="<%= ocoffinvoice.FItemList(i).Fportname %>">
	<input type="hidden" name="destinationname" value="<%= ocoffinvoice.FItemList(i).Fdestinationname %>">
	<input type="hidden" name="carriername" value="<%= ocoffinvoice.FItemList(i).Fcarriername %>">
	<input type="hidden" name="lccomment" value="<%= ocoffinvoice.FItemList(i).Flccomment %>">
	<input type="hidden" name="lcbank" value="<%= ocoffinvoice.FItemList(i).Flcbank %>">
	<input type="hidden" name="comment" value="<%= ocoffinvoice.FItemList(i).Fcomment %>">
	<input type="hidden" name="goodscomment1" value="<%= ocoffinvoice.FItemList(i).Fgoodscomment1 %>">
	<input type="hidden" name="goodscomment2" value="<%= ocoffinvoice.FItemList(i).Fgoodscomment2 %>">

	<input type="hidden" name="delivermethod" value="<%= ocoffinvoice.FItemList(i).Fdelivermethod %>">
	<input type="hidden" name="exportmethod" value="<%= ocoffinvoice.FItemList(i).Fexportmethod %>">
	<input type="hidden" name="jungsantype" value="<%= ocoffinvoice.FItemList(i).Fjungsantype %>">
	<input type="hidden" name="priceunit" value="<%= ocoffinvoice.FItemList(i).Fpriceunit %>">
	<input type="hidden" name="exchangerate" value="<%= ocoffinvoice.FItemList(i).Fexchangerate %>">

	<tr bgcolor="#FFFFFF">
		<td align="center"><%= ocoffinvoice.FItemList(i).Fidx %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).Fshopid %><br><%= ocoffinvoice.FItemList(i).Fshopname %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).Finvoicedate %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).GetDeliverMethodName %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).GetExportMethodName %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).GetJungsanTypeName %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).Fpriceunit %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).Ftotalboxno %></td>
		<td align="center"><%= FormatNumber(ocoffinvoice.FItemList(i).Ftotalgoodsprice, 2) %></td>
		<td align="center"><%= FormatNumber(ocoffinvoice.FItemList(i).Ftotalboxprice, 2) %></td>
		<td align="center"><%= ocoffinvoice.FItemList(i).Freguserid %></td>
		<td align="center"><%= Left(ocoffinvoice.FItemList(i).Fregdate, 10) %></td>
		<td align="center">
			<input type="button" class="button" value=" 선택 " onClick="InsertInvoiceInfo(frmMaster<%= i %>)">
		</td>
	</tr>
	</form>
	<% next %>
	<% else %>
<tr bgcolor="#FFFFFF">
		<td colspan=13 align=center>[ 검색결과가 없습니다. ]</td>
	</tr>
	<% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="13" align="center">
			<%
			dim strparam
			strparam = "&shopid=" + CStr(shopid)

			%>
			<% if ocoffinvoice.HasPreScroll then %>
				<a href="?page=<%= ocoffinvoice.StartScrollPage-1 %>&research=on<%= strparam %>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + ocoffinvoice.StartScrollPage to ocoffinvoice.FScrollCount + ocoffinvoice.StartScrollPage - 1 %>
				<% if i>ocoffinvoice.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&research=on<%= strparam %>">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if ocoffinvoice.HasNextScroll then %>
				<a href="?page=<%= i %>&research=on<%= strparam %>">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
set ocoffinvoice = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->