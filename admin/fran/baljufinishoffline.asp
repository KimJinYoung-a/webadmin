<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/baljucls_off.asp"-->
<%
' 사용안하는듯
response.end

dim baljudate, baljunum, baljuid, searchtype

baljudate = request("baljudate")
baljunum = request("baljunum")
baljuid = request("baljuid")
searchtype = request("searchtype")

if (searchtype = "") then
        searchtype = "M"
end if

if (baljudate = "") then
        baljudate = Left(now, 10)
end if


dim baljuitemoff
set baljuitemoff = New COfflineBalju
baljuitemoff.FRectBaljuNum = baljunum
baljuitemoff.FRectBaljuId = baljuid
'baljuitemoff.FRectBaljuDate = baljudate
baljuitemoff.GetOfflineBaljuItemListForFinish

dim TotalBaljucount, TotalUpchecount, TotalTenBaljucount
dim TotalNoPackCount, TotalPackCount, TotalDeliverCount, TotalEtcCount

dim i, minboxno, maxboxno

minboxno = -1

%>
<script>
function WriteBarcode(itemgubun, itemid, itemoption) {
        if (1*itemid>=1000000){
            var tmp = "" + (100000000 + (1 * itemid));
        }else{
            var tmp = "" + (1000000 + (1 * itemid));
        }
        document.frm.barcode.value = itemgubun + tmp.substring(1) + itemoption;
        barcodechulgo();
}

function FinishBalju() {
        var f = document.frmarr;
        var u = document.uparr;

        u.itemgubun.value = "";
        u.itemid.value = "";
        u.itemoption.value = "";
        u.comment.value = "";

        for (var i = 0; i < f.elements.length; i++) {
                if ((f[i].name == "comment") && (f[i].selectedIndex != 0)) {
                        u.itemgubun.value = u.itemgubun.value + "|" + f[i-3].value;
                        u.itemid.value = u.itemid.value + "|" + f[i-2].value;
                        u.itemoption.value = u.itemoption.value + "|" + f[i-1].value;

                        u.comment.value = u.comment.value + "|" + f[i][f[i].selectedIndex].value;
                }
        }

        if (confirm("해당 샆발주건에 대한 발주를 출고완료합니다\n발주는 출고완료로 전환되며, 5일출고로 표시한 상품은 자동으로 재주문 작성됩니다.") == true) {
                u.submit();
        }
}
</script>

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<form name="frm" onsubmit="return false;">
	<input type=hidden name="baljunum" value="<%= baljunum %>">
	<tr height="10">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>상품출고</strong></font>
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr height="25">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td bgcolor="#F3F3FF">
			<br>
			&nbsp;
			발주코드 : <%= baljunum %>
			<!--발주일자 : <%= Left(baljudate,10) %>-->
			<!--<input type="hidden" name="baljudate" value="<%= Left(baljudate,10) %>">-->
			&nbsp;
			가맹점 : <% drawSelectBoxOffShop "baljuid",baljuid %>
			<input type=button value=" 검색 " onclick="document.frm.submit();">
			<input type=button value=" 목록으로 " onclick="location.href='baljulistoffline.asp'">
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr height="10" valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
	</form>
</table>

<p>

<table width="100%" height="200" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
  <form name="researchFrm" method=get>
  <input type=hidden name="baljuid" value="<%= baljuid %>">
  <input type=hidden name="baljunum" value="<%= baljunum %>">
  <input type=hidden name="baljudate" value="<%= baljudate %>">
  <tr height="10" valign="bottom">
    <td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
    <td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
    <td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
  </tr>
  <tr height="25" valign="top">
    <td background="/images/tbl_blue_round_04.gif"></td>
    <td background="/images/tbl_blue_round_06.gif">
    	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    	<tr>
    	  <td>
            <img src="/images/icon_star.gif" align="absbottom">
            <font color="red"><strong>미출고내역</strong></font> &nbsp
          </td>
    	  <td align="right"><input type=button value=" 출고처리 " onclick="FinishBalju()">
    	  </td>
    	</tr>
    	</table>
    </td>
    <td background="/images/tbl_blue_round_05.gif"></td>
  </tr>
  </form>
  <tr valign="top">
    <td background="/images/tbl_blue_round_04.gif"></td>
    <td bgcolor="#FFFFFF">
		<table width="100%" border="0" cellspacing="2" cellpadding="0" class="a">
		<tr>
			<td>브랜드ID</td>
			<td align="center" width=30>배송<br>방식</td>
			<td width=25>구분</td>
			<td align="right" width=35>상품<br>코드</td>
			<td align="center" width=35>옵션</td>

			<td >(현)상품명</td>
			<td >(현)옵션명</td>
			<td width=20></td>
			<td align=center width=45>발주<br>수량</td>
			<td align=center width=50>미출고<br>수량</td>
			<td align=center width=50>출고준비<br>(Box in)</td>
			<td align=center width=45>패킹<br>완료</td>
			<td align=center width=80>비고</td>
		</tr>
		<form name="frmarr">
		<% for i=0 to baljuitemoff.FResultCount -1 %>
		<%
                        if ((searchtype = "A") or ((searchtype = "M") and (baljuitemoff.FItemList(i).Ftotalnopackno > 0)) or ((searchtype = "P") and (baljuitemoff.FItemList(i).Ftotalpackno > 0)) or ((searchtype = "C") and (baljuitemoff.FItemList(i).Ftotaldeliverno > 0))) then
                                if ((minboxno = -1) or ((minboxno > baljuitemoff.FItemList(i).FRealBoxNo) and (baljuitemoff.FItemList(i).FBoxSongjangNo = "0"))) then
                                        minboxno = baljuitemoff.FItemList(i).FRealBoxNo
                                end if

                                if (maxboxno < baljuitemoff.FItemList(i).FRealBoxNo) then
                                        maxboxno = baljuitemoff.FItemList(i).FRealBoxNo
                                end if

                                TotalBaljucount      = TotalBaljucount + baljuitemoff.FItemList(i).Ftotalbaljuno
                                TotalUpchecount      = TotalUpchecount +  baljuitemoff.FItemList(i).Ftotalupcheno
                                TotalTenBaljucount   = TotalTenBaljucount +  baljuitemoff.FItemList(i).Ftotaltenbaeno

                                TotalNoPackCount     = TotalNoPackCount + baljuitemoff.FItemList(i).Ftotalnopackno
                                TotalPackCount       = TotalPackCount + baljuitemoff.FItemList(i).Ftotalpackno
                                TotalDeliverCount    = TotalDeliverCount + baljuitemoff.FItemList(i).Ftotaldeliverno
                                TotalEtcCount        = TotalEtcCount + baljuitemoff.FItemList(i).Ftotaletcno

		%>
		<tr>
			<td height="1" colspan="13" bgcolor="#CCCCCC"></td>
		</tr>
		<tr>
			<!--
			<td align="center"><%= Format00(4,baljuitemoff.FItemList(i).Fprtidx) %></td>
			-->
			<td ><%= baljuitemoff.FItemList(i).FMakerid %></td>
			<td align="center">
                        <% if (baljuitemoff.FItemList(i).Fmaeipdiv = "U") then %>
                          업배
                        <% elseif (baljuitemoff.FItemList(i).Fmaeipdiv = "9") then %>
                          오프
                        <% else %>
                          <!--텐배-->
                        <% end if %>
		        </td>
		        <input type=hidden name=itemgubun value="<%= baljuitemoff.FItemList(i).FItemGubun %>">
		        <input type=hidden name=itemid value="<%= baljuitemoff.FItemList(i).FItemID %>">
		        <input type=hidden name=itemoption value="<%= baljuitemoff.FItemList(i).FItemOption %>">
			<td align="center"><%= baljuitemoff.FItemList(i).FItemGubun %></td>
			<td align="right"><%= baljuitemoff.FItemList(i).FItemID %></td>
			<td align="center"><%= baljuitemoff.FItemList(i).FItemOption %></td>
			<td ><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= baljuitemoff.FItemList(i).FItemID %>" target="_blank"><%= baljuitemoff.FItemList(i).FItemName %></a></td>
			<td ><%= baljuitemoff.FItemList(i).FItemOptionName %></td>
			<td>
			<% if ((searchtype <> "P") and (searchtype <> "C")) then %>
			  <a href="javascript:WriteBarcode('<%= baljuitemoff.FItemList(i).FItemGubun %>','<%= baljuitemoff.FItemList(i).FItemID %>','<%= baljuitemoff.FItemList(i).FItemOption %>')">=&gt;</a>
			<% end if %>
			</td>
			<td align=center><%= baljuitemoff.FItemList(i).Ftotalbaljuno %></td>
			<td align=center>
                        <% if (baljuitemoff.FItemList(i).Ftotalnopackno <> 0) then %>
			  <font color="blue"><%= baljuitemoff.FItemList(i).Ftotalnopackno %></font>
                        <% else %>
                          <%= baljuitemoff.FItemList(i).Ftotalnopackno %>
                        <% end if %>
		        </td>
			<td align=center>
                        <% if (baljuitemoff.FItemList(i).Ftotalpackno <> baljuitemoff.FItemList(i).Ftotalbaljuno) then %>
			  <font color="blue"><%= baljuitemoff.FItemList(i).Ftotalpackno %></font>
                        <% else %>
                          <%= baljuitemoff.FItemList(i).Ftotalpackno %>
                        <% end if %>
			</td>
			<td align=center>
                        <% if (baljuitemoff.FItemList(i).Ftotaldeliverno <> baljuitemoff.FItemList(i).Ftotalbaljuno) then %>
			  <b><font color="red"><%= baljuitemoff.FItemList(i).Ftotaldeliverno %></font></b>
                        <% else %>
                          <b><%= baljuitemoff.FItemList(i).Ftotaldeliverno %></b>
                        <% end if %>
		        </td>
		        <td align=center><% DrawMiChulgoDiv "comment", "" %></td>
		</tr>
		        <% end if %>
		<% next %>
		</form>
		<tr>
			<td height="1" colspan="13" bgcolor="#CCCCCC"></td>
		</tr>
		<tr>
			<td align=center>Total</td>
			<td ></td>
			<td ></td>
			<td ></td>
			<td ></td>
			<td ></td>
			<td ></td>
			<td ></td>
			<td align=center><%= TotalBaljucount %></td>
			<td align=center><%= TotalNoPackCount %></td>
			<td align=center><%= TotalPackCount %></td>
			<td align=center><b><%= TotalDeliverCount %></b></td>
			<td ></td>
		</tr>
		</table>
    </td>
    <td background="/images/tbl_blue_round_05.gif"></td>
  </tr>

  <tr height="10" valign="top">
    <td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
  </tr>
</table>
<form name="uparr" method="post" action="baljufinishoffline_process.asp">
<input type=hidden name=mode value="chulgoproc">
<input type=hidden name=baljunum value="<%= baljunum %>">
<input type=hidden name=baljuid value="<%= baljuid %>">
<input type=hidden name=itemgubun value="">
<input type=hidden name=itemid value="">
<input type=hidden name=itemoption value="">
<input type=hidden name=comment value="">
</form>
<%

set baljuitemoff = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->