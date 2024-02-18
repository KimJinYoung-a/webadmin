<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/base64.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/newitemcouponcls.asp"-->
<%
	'변수선언
	dim itemcouponidx, itemid
	dim oitemcouponmaster, ocouponitemlist

	itemcouponidx	= request("icpidx")
	itemid			= request("iid")

	'타겟쿠폰 확인
	set oitemcouponmaster = new CItemCouponMaster
	oitemcouponmaster.FRectItemCouponIdx = itemcouponidx
	oitemcouponmaster.GetOneItemCouponMaster

	if oitemcouponmaster.FResultCount<1 then
		Call Alert_Close("잘못된 쿠폰입니다.")
		response.End
	end if

	'쿠폰상품 확인
	set ocouponitemlist = new CItemCouponMaster
	ocouponitemlist.FPageSize=1
	ocouponitemlist.FCurrPage=1
	ocouponitemlist.FRectItemCouponIdx = itemcouponidx
	ocouponitemlist.FRectsRectItemidArr = itemid
	ocouponitemlist.GetItemCouponItemList

	if ocouponitemlist.FResultCount<1 then
		Call Alert_Close("없거나 잘못된 상품입니다.")
		response.End
	end if
%>
<script type="text/javascript">
// 클립보드로 복사
function fnCBCopy(iid,dvc) {
	var doc, dmn
	switch(dvc) {
		case "w":
			dmn = "http://www.10x10.co.kr/shopping/category_prd.asp";
			break;
		case "m":
			dmn = "http://m.10x10.co.kr/category/category_itemprd.asp";
			break;
		case "a":
			dmn = "http://m.10x10.co.kr/apps/appcom/wish/web2014/category/category_itemprd.asp";
			break;
	}
	doc = dmn + "?itemid=" + iid + "&ldv=<%=server.URLencode(Base64encode(oitemcouponmaster.FOneItem.FItemCouponIdx))%>";
	clipboardData.setData("Text", doc);
	alert('링크가 복사되었습니다. 사용하실 곳에 Ctrl+V 하시면됩니다.');
}
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<tr bgcolor="#FFFFFF">
    <td colspan="4"><b>타켓쿠폰 링크 확인/복사</b></td>
</tr>
<tr bgcolor="#E8E8EE">
	<td width="80">쿠폰명</td>
	<td colspan="3" bgcolor="#FFFFFF"><%= oitemcouponmaster.FOneItem.Fitemcouponname %></td>
</tr>
<tr bgcolor="#E8E8EE">
	<td>할인율</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<%= oitemcouponmaster.FOneItem.GetDiscountStr %>
	</td>
</tr>
<tr bgcolor="#E8E8EE">
	<td>적용기간</td>
	<td colspan="3" bgcolor="#FFFFFF">
	<%= oitemcouponmaster.FOneItem.Fitemcouponstartdate %> ~ <%= oitemcouponmaster.FOneItem.Fitemcouponexpiredate %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>상품번호</td>
	<td colspan="3" bgcolor="#FFFFFF"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td>상품명</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<%= ocouponitemlist.FitemList(0).FItemName %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="80">쿠폰판매가</td>
	<td bgcolor="#FFFFFF">
		<%= FormatNumber(ocouponitemlist.FitemList(0).GetCouponSellcash,0) %>원
		<% if ocouponitemlist.FitemList(0).Fitemcoupontype="3" then %><font color="red">(무료배송)</font><% end if %>
	</td>
	<td width="80">현재판매가</td>
	<td bgcolor="#FFFFFF">
		<%= FormatNumber(ocouponitemlist.FitemList(0).FSellcash,0) %>원
	</td>
</tr>
</table><br>
※ 아래 링크를 통해서 접속하면 할인가격과 쿠폰다운로드가 표시됩니다.<br>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<tr bgcolor="#DDDDFF" align="center">
	<td>사용처</td>
	<td>링크</td>
	<td>복사</td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>PC웹</td>
	<td><input type="text" value="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=itemid%>&ldv=<%=server.URLencode(Base64encode(oitemcouponmaster.FOneItem.FItemCouponIdx))%>" class="text" readonly style="width:100%" onfocus="this.select();"></td>
	<td><input type="button" onclick="fnCBCopy(<%=itemid%>,'w')" value="복사"></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>모바일웹</td>
	<td><input type="text" value="http://m.10x10.co.kr/category/category_itemprd.asp?itemid=<%=itemid%>&ldv=<%=server.URLencode(Base64encode(oitemcouponmaster.FOneItem.FItemCouponIdx))%>" class="text" readonly style="width:100%" onfocus="this.select();"></td>
	<td><input type="button" onclick="fnCBCopy(<%=itemid%>,'m')" value="복사"></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td>wishApp</td>
	<td><input type="text" value="http://m.10x10.co.kr/apps/appcom/wish/web2014/category/category_itemprd.asp?itemid=<%=itemid%>&ldv=<%=server.URLencode(Base64encode(oitemcouponmaster.FOneItem.FItemCouponIdx))%>" class="text" readonly style="width:100%" onfocus="this.select();"></td>
	<td><input type="button" onclick="fnCBCopy(<%=itemid%>,'a')" value="복사"></td>
</tr>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="3"><input type="button" value="창닫기" onclick="window.close()"></td>
</tr>
</table>
<%
	set oitemcouponmaster = Nothing
	set ocouponitemlist = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->