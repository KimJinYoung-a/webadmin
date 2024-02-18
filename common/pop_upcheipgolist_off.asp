<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%

dim ipchulflag,chargeid, designer,itemgubun,itemid,itemoption,shopid, reSeach
ipchulflag  = RequestCheckVar(request("ipchulflag"),9)
chargeid    = RequestCheckVar(request("chargeid"),32)
designer    = RequestCheckVar(request("designer"),32)
itemgubun   = RequestCheckVar(request("itemgubun"),2)
itemid      = RequestCheckVar(request("itemid"),9)
itemoption  = RequestCheckVar(request("itemoption"),4)
shopid      = RequestCheckVar(request("shopid"),32)
reSeach     = RequestCheckVar(request("reSeach"),2)

if itemgubun="" and reSeach="" then itemgubun="10"
if itemoption="" and reSeach="" then itemoption="0000"

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim fromdate,todate

fromdate = request("fromdate")
todate = request("todate")

if fromdate<>"" then
	yyyy1 = Left(fromdate,4)
	mm1 = Mid(fromdate,6,2)
	dd1 = Mid(fromdate,9,2)
else
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
end if

if todate<>"" then
	yyyy2 = Left(todate,4)
	mm2 = Mid(todate,6,2)
	dd2 = Mid(todate,9,2)
else
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
end if


if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromdate = CStr(DateSerial(yyyy1, mm1, dd1))
todate = CStr(DateSerial(yyyy2, mm2, dd2+1))



dim oupcheipchul

set oupcheipchul = new CShopIpChul

oupcheipchul.FRectStartday = fromDate
oupcheipchul.FRectEndday   = toDate
oupcheipchul.FRectChargeId = chargeid
oupcheipchul.FRectMakerId = designer
oupcheipchul.FRectItemgubun = itemgubun
oupcheipchul.FRectItemID = itemid
oupcheipchul.FRectItemOption = itemoption
oupcheipchul.FRectShopid = shopid

oupcheipchul.GetIpChulDetailByShopByItem

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
		<font color="red"><strong>OFFLINE 삽별상품별 업체입출고목록</strong></font></td>
		<td height="20" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			샆에 대한 업체입출고에 대한 정보입니다.
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
	<input type="hidden" name="reSeach" value="ON">
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	상품코드 :
                <% drawSelectBoxItemGubun "itemgubun", itemgubun %>
	        	<input type="text" class="text" name="itemid" value="<%= itemid %>" Maxlength="9" size="9">&nbsp;&nbsp;
	        	<input type="text" class="text" name="itemoption" value="<%= itemoption %>" size=4 maxlength=4 >
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
	        	출고처 : <% drawSelectBoxDesignerwithName "chargeid",chargeid  %>&nbsp;&nbsp;
				브랜드 : <% drawSelectBoxDesignerwithName "designer",designer  %>&nbsp;&nbsp;
	        </td>
	        <td align="right" bgcolor="F4F4F4">(최대 1,000건)</td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
      <td width="60">입출코드</td>
      <td width="60">입출고일</td>
      <td width="100">출고처</td>
      <td width="100">출고구분</td>
      <td width="100">입고처ID</td>
      <td width="100">브랜드ID</td>
	  <td width="25">구분</td>
      <td width="50">상품코드</td>
	  <td width="40">옵션코드</td>
      <td width="70">바코드</td>
	  <td>아이템명</td>
      <td>옵션</td>
	  <td width="50">소비자가</td>
      <td width="50">공급가</td>
      <td width="30">수량</td>
    </tr>
    <% for i=0 to oupcheipchul.FResultCount-1 %>
    <%
    totitemno = totitemno + oupcheipchul.FItemList(i).FItemNo
    %>
    <tr align="center" bgcolor="#FFFFFF">
      <td><a target="_blank" href="/common/offshop/shop_ipchuldetail.asp?idx=<%= oupcheipchul.FItemList(i).Fidx %>&menupos=196"><%= oupcheipchul.FItemList(i).Fidx %></a></td>
      <td><%= oupcheipchul.FItemList(i).Fexecdt %></td>
      <td><%= oupcheipchul.FItemList(i).Fchargeid %></td>
      <td><%= oupcheipchul.FItemList(i).Fcomm_cd %></td>
      <td><%= oupcheipchul.FItemList(i).Fshopid %></td>
      <td><%= oupcheipchul.FItemList(i).Fdesignerid %></td>
	  <td><%= oupcheipchul.FItemList(i).FItemgubun %></td>
      <td><%= oupcheipchul.FItemList(i).FItemID %></td>
	  <td><%= oupcheipchul.FItemList(i).Fitemoption %></td>
	  <td><%= oupcheipchul.FItemList(i).GetBarCode() %></td>
      <td><%= oupcheipchul.FItemList(i).FItemName %></td>
      <td><%= oupcheipchul.FItemList(i).FItemOptionName %></td>
      <td align="right"><%= FormatNumber(oupcheipchul.FItemList(i).FSellCash,0) %></td>
      <td align="right"><%= FormatNumber(oupcheipchul.FItemList(i).FsuplyCash,0) %></td>
      <td align="center"><%= oupcheipchul.FItemList(i).FItemNo %></td>
    </tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
	  <td align=center colspan="2">Total</td>
	  <td colspan="12"></td>
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
set oupcheipchul = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
