<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : OFFLINE 삽별상품별 판매목록
' History : 2017.04.10 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<%

dim ipchulflag,designer,itemgubun,itemid,itemoption,shopid, showOrder
ipchulflag  = RequestCheckVar(request("ipchulflag"),9)
designer    = RequestCheckVar(request("designer"),32)
itemgubun   = RequestCheckVar(request("itemgubun"),2)
itemid      = RequestCheckVar(request("itemid"),9)
itemoption  = RequestCheckVar(request("itemoption"),4)
shopid      = RequestCheckVar(request("shopid"),32)
showOrder      = RequestCheckVar(request("showOrder"),32)

if itemgubun="" then itemgubun="10"
if itemoption="" then itemoption="0000"

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim fromdate,todate

fromdate = requestCheckVar(request("fromdate"),10)
todate = requestCheckVar(request("todate"),10)

if fromdate<>"" then
	yyyy1 = Left(fromdate,4)
	mm1 = Mid(fromdate,6,2)
	dd1 = Mid(fromdate,9,2)
else
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
end if

if todate<>"" then
	yyyy2 = Left(todate,4)
	mm2 = Mid(todate,6,2)
	dd2 = Mid(todate,9,2)
else
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
end if

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromdate = CStr(DateSerial(yyyy1, mm1, dd1))
todate = CStr(DateSerial(yyyy2, mm2, dd2+1))

dim osell

set osell = new COffShopSellReport
	osell.FRectStartday = fromDate
	osell.FRectEndday   = toDate
	osell.FRectDesigner = designer
	osell.FRectItemGubun= itemgubun
	osell.FRectItemID   = itemid
	osell.FRectItemOption= itemoption
	osell.FRectShopid   = shopid
    osell.FRectShowOrder   = showOrder
	osell.GetDaylySellItemListByShopByItem

dim i, totitemno

totitemno=0
%>


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr valign="bottom">
		<td width="10" height="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td height="10" valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" height="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top">
		<td height="20" background="/images/tbl_blue_round_04.gif"></td>
		<td height="20" background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>OFFLINE 삽별상품별 판매목록</strong></font></td>
		<td height="20" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			샆에서 발생한 판매에 대한 일별 정보입니다.(최근 3개월에 대한 정보만 표시됩니다.)
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td height="10"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td height="10" background="/images/tbl_blue_round_08.gif"></td>
		<td height="10"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>


<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<input type="hidden" name="ipchulflag" value="S">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<form name="frm" method="get" action="">
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	상품코드 : <select class="select" name="itemgubun">
	        		<option value="10" <%= chkIIF(itemgubun="10","selected","") %> >10</option>
	        		<option value="70" <%= chkIIF(itemgubun="70","selected","") %> >70</option>
	        		<option value="80" <%= chkIIF(itemgubun="80","selected","") %> >80</option>
	        		<option value="90" <%= chkIIF(itemgubun="90","selected","") %> >90</option>
	        	</select>
	        	<input type="text" class="text" name="itemid" value="<%= itemid %>" Maxlength="9" size="9">&nbsp;&nbsp;
	        	<input type="text" class="text_ro" name="itemoption" value="<%= itemoption %>" size=4 maxlength=4 readonly>
	        	매장 : <% drawSelectBoxOffShop "shopid",shopid %>
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	검색기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>&nbsp;&nbsp;
	        	브랜드 : <% drawSelectBoxDesignerwithName "designer",designer  %>&nbsp;&nbsp;
                <input type="checkbox" name="showOrder" value="Y" <%= CHKIIF(showOrder<>"", "checked", "") %>> 주문번호 표시
	        </td>
	        <td align="right" bgcolor="F4F4F4">(최대 1,000건)</td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
      <% if showOrder<>"" then %>
      <td>주문번호</td>
      <% end if %>
      <td width="60">판매일</td>
      <td width="100">브랜드</td>
      <td width="100">판매처</td>
      <td width="25">구분</td>
      <td width="50">상품코드</td>
      <td>아이템명</td>
      <td>옵션</td>
      <td>정산구분</td>
      <td width="50">소비자가</td>
      <td width="50">판매가</td>
      <td width="30">수량</td>
    </tr>
    <% for i=0 to osell.FResultCount-1 %>
    <%
    totitemno = totitemno + osell.FItemList(i).FItemNo
    %>
    <tr align="center" bgcolor="#FFFFFF">
      <% if showOrder<>"" then %>
      <td><%= osell.FItemList(i).ForderNo %></td>
      <% end if %>
      <td><%= osell.FItemList(i).Fshopregdate %></td>
      <td><%= osell.FItemList(i).FMakerID %></td>
      <td><%= osell.FItemList(i).Fshopid %></td>
      <td><%= osell.FItemList(i).FItemgubun %></td>
      <td><%= osell.FItemList(i).FItemID %></td>
      <td><%= osell.FItemList(i).FItemName %></td>
      <td><%= osell.FItemList(i).FItemOptionName %></td>
      <td><%= osell.FItemList(i).Fjcomm_cd %>(<%= osell.FItemList(i).Fcomm_name %>)</td>
      <td align="right"><%= FormatNumber(osell.FItemList(i).Fsellprice,0) %></td>
      <td align="right"><%= FormatNumber(osell.FItemList(i).Frealsellprice,0) %></td>
      <td align="center"><%= -1 * osell.FItemList(i).FItemNo %></td>
    </tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
	  <td align=center>Total</td>
	  <td colspan=<%= CHKIIF(showOrder<>"", "10", "9") %>></td>
	  <td align=center><%= FormatNumber(-1 * totitemno,0) %></td>
	</tr>
</table>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right" bgcolor="F4F4F4">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<%
set osell = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
