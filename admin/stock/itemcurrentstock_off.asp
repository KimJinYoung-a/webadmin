<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/realjaegocls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<%

const C_STOCK_DAY=7

dim itemgubun, itemid, itemoption
itemgubun = request("itemgubun")
itemid = request("itemid")
itemoption = request("itemoption")

if itemgubun="" then itemgubun="10"
if itemoption="" then itemoption="0000"

dim ojaegoitem
set ojaegoitem = new CRealJaeGo
ojaegoitem.FRectItemID = itemid
if itemid<>"" then
	ojaegoitem.GetItemDefaultData
end if

dim oitemoption
set oitemoption = new CItemOptionInfo
oitemoption.FRectItemID =  itemid
if itemid<>"" then
	oitemoption.getOptionList
end if

if (oitemoption.FResultCount<1) then
	itemoption = "0000"
end if



dim BasicMonth


BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)


dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectStartDate = BasicMonth + "-01"
osummarystock.FRectItemGubun = itemgubun
osummarystock.FRectItemID =  itemid
osummarystock.FRectItemOption =  itemoption
if itemid<>"" then
	osummarystock.GetCurrentItemStock
	osummarystock.GetDaily_Logisstock_Summary
end if

dim offstock
set offstock = new COffShopDailyStock
offstock.FRectItemGubun = itemgubun
offstock.FRectItemid = itemid
offstock.FRectItemoption = itemoption
if itemid<>"" then
	if ojaegoitem.FResultCount>0 then
		offstock.FRectMakerid = ojaegoitem.FItemList(0).Fmakerid
	end if

	offstock.GetCurrentAllShopItemStock
end if

dim i
dim sum_ipgono,sum_reipgono,sum_sellno,sum_resellno

dim sum_offchulgono, sum_offrechulgono, sum_etcchulgono, sum_etcrechulgono
dim sum_totsysstock, sum_availsysstock, sum_realstock
dim sum_errbaditemno, sum_errrealcheckno
dim sum_offsell

dim sysstock, sysavailstock, realstock, maystock
dim offstockno
%>

<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
}

function RefreshRecentStock(yyyymmdd,itemgubun,itemid,itemoption){
	if (confirm('최근 2달 내역을 새로고침 하시겠습니까?')){
		frmrefresh.mode.value="itemrecentipchulrefresh";
		frmrefresh.submit();
	}
}

function RefreshTodayStock(itemgubun,itemid,itemoption){
	if (confirm('금일 내역을 새로고침 하시겠습니까?')){
		frmrefresh.mode.value="itemtodayipchulrefresh";
		frmrefresh.submit();
	}
}


function RefreshALLStock(yyyymmdd,itemgubun,itemid,itemoption){
	if (confirm('전체 내역을 새로고침 하시겠습니까?')){
		frmrefresh.mode.value="itemallipchulrefresh";
		frmrefresh.submit();
	}
}

function PopStockBaditem(fromdate,todate,itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poperritemlist.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popbaditemlist','width=800,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function popRealErrList(fromdate,todate,itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poperritemlist.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'poperritemlist','width=800,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function popRealErrInput(itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poprealerrinput.asp?itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&BasicMonth=<%= BasicMonth %>','poprealerrinput','width=900,height=460,scrollbar=yes,resizable=yes')
	popwin.focus();
}
</script>


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr valign="bottom">
		<td width="10" height="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td height="10" valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" height="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top">
		<td height="20" background="/images/tbl_blue_round_04.gif"></td>
		<td height="20" background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>OFFLINE상품별재고현황</strong></font></td>
		<td height="20" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			오프라인 샵별 실시간 상품재고 정보입니다..
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method=get>
	<input type=hidden name=menupos value="<%= menupos %>">
    <tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
    </tr>
    <tr height="25" valign="bottom">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">
        	상품코드: <input type=text name=itemid value="<%= itemid %>" size=9 maxlength=9>
        	&nbsp;
			<% if oitemoption.FResultCount>0 then %>
			옵션선택 :
			<select name="itemoption">
			<option value="0000">----
			<% for i=0 to oitemoption.FResultCount-1 %>
			<option value="<%= oitemoption.FItemList(i).FItemOption %>" <% if itemoption=oitemoption.FItemList(i).FItemOption then response.write "selected" %> ><%= oitemoption.FItemList(i).FItemOptionName %>
			<% next %>
			</select>
			<% end if %>
			&nbsp;
        	<input type=button value="검색" onclick="document.frm.submit();">
        </td>
        <td valign="top" align="right">
        <% if itemid<>"" then %>
        	최종업데이트시간 : <b><%= osummarystock.FOneItem.Flastupdate %></b>
        <% end if %>

        <% if C_ADMIN_AUTH=true then %>
        <!-- <input type="button" value="전체내역새로고침" onclick="RefreshALLStock();"> -->
        <input type="button" value="2개월 새로고침" onclick="RefreshRecentStock();" disabled >
        <% end if %>
        <input type="button" value="금일 새로고침" onclick="RefreshTodayStock();" disabled >
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- 표 상단바 끝-->

<% if ojaegoitem.FResultCount>0 then %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#EEEEEE">
		<td colspan="6">&nbsp;<b> *Center 재고 </b></td>
	</tr>
	<tr bgcolor="#FFFFFF">
    	<td rowspan=<%= 5 + ojaegoitem.FResultCount -1 %> width="110" valign=top align=center><img src="<%= ojaegoitem.FItemList(0).FImageList %>" width="100" height="100"></td>
      	<td width="60"><b>*상품정보</b></td>
      	<td width="300">
      	<input type="button" value="수정" onclick="PopItemSellEdit('<%= itemid %>');">
      	</td>
      	<td width="60">배송구분 :</td>
      	<td colspan=2><%= ojaegoitem.FItemList(0).GetDeliveryName %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>상품코드 :</td>
      	<td>10 <b><%= CHKIIF(ojaegoitem.FItemList(0).FItemID>=1000000,Format00(8,ojaegoitem.FItemList(0).FItemID),Format00(6,ojaegoitem.FItemList(0).FItemID)) %></b> <%= itemoption %></td>
      	<td>전시여부 : </td>
      	<td colspan=2></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>브랜드ID :</td>
      	<td><%= ojaegoitem.FItemList(0).FMakerid %></td>
      	<td>판매여부 : </td>
      	<td colspan=2><%= ojaegoitem.FItemList(0).FSellyn %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td>상품명 :</td>
      	<td><%= ojaegoitem.FItemList(0).FItemName %></td>
      	<td>사용여부 : </td>
      	<td colspan=2><%= ojaegoitem.FItemList(0).FIsUsing %></td>
    </tr>
    <% for i=0 to ojaegoitem.FResultCount -1 %>
	    <% if ojaegoitem.FItemList(i).Foptionusing<>"Y" then %>
	    <tr bgcolor="#FFFFFF">
	      	<td><font color="#AAAAAA">옵션명 :</font></td>
	      	<td><font color="#AAAAAA"><%= ojaegoitem.FItemList(i).FItemOptionName %></font></td>
	      	<td><font color="#AAAAAA">한정여부 : </font></td>
	      	<td><font color="#AAAAAA"><%= ojaegoitem.FItemList(i).FLimitYn %> (<%= ojaegoitem.FItemList(i).GetLimitStr %>)</font></td>
	      	<td>
	      		<%= ojaegoitem.FItemList(i).Foldstockcurrno %> : (OLD)
	      		<%= ojaegoitem.FItemList(i).GetCheckStockNo %> : (NEW)
	      	</td>
	    </tr>
	    <% else %>

	    <% if ojaegoitem.FItemList(i).FItemOption=itemoption then %>
	    <tr bgcolor="#EEEEEE">
	    <% else %>
	    <tr bgcolor="#FFFFFF">
	    <% end if %>
	      	<td>옵션명 :</td>
	      	<td><%= ojaegoitem.FItemList(i).FItemOptionName %></td>
	      	<td>한정여부 : </td>
	      	<td><%= ojaegoitem.FItemList(i).FLimitYn %> (<%= ojaegoitem.FItemList(i).GetLimitStr %>)</td>
	      	<td>
	      		<%= ojaegoitem.FItemList(i).Foldstockcurrno %> : (OLD)
	      		<%= ojaegoitem.FItemList(i).GetCheckStockNo %> : (NEW)
	      	</td>
	    </tr>
	    <% end if %>
    <% next %>
</table>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="20" valign="bottom">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<br>시스템 총재고 = 입고/반품합 + 업체입고/반품합 - 총OFF판매합 + 기타출고/반품합
		<br><br>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->



<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="20" valign="bottom">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td><b>*샵별 재고</b>(기준시간 : 금일 새벽 1시)</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->
<%

dim colcount, offtotalstock
dim offtotipno, offtotreno, offtotupcheipno, offtotupchereno, offtotsellno, offtotcurrno
colcount = offstock.FResultCount
dim fromdate, todate

fromdate = "2001-10-10"
todate = Left(now(), 10)

%>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="100">shopID</td>
    	<td width="60">거래조건</td>
    	<td width="60">입고<br>(텐바이텐)</td>
    	<td width="60">반품<br>(텐바이텐)</td>
    	<td width="60">입고<br>(업체)</td>
    	<td width="60">반품<br>(업체)</td>
    	<td width="60">총판매</td>
    	<td width="60" bgcolor="F4F4F4">시스템재고</td>
    	<td width="60">샘플</td>
    	<td width="60">불량</td>
    	<td width="60" bgcolor="F4F4F4">유효재고</td>
    	<td width="60">오차</td>
    	<td width="60" bgcolor="F4F4F4">예상재고</td>
    	<td>비고</td>
    </tr>
    <% for i=0 to offstock.FResultcount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= offstock.FItemList(i).FShopid %></td>
    	<td><acronym title="매입마진 : <%= offstock.FItemList(i).Fdefaultmargin %>&#13공급마진 : <%= offstock.FItemList(i).Fdefaultsuplymargin %>"><font color="<%= offstock.FItemList(i).getChargeDivColor %>"><%= offstock.FItemList(i).GetChargedivName %></font></acronym></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= fromdate %>','<%= todate %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= offstock.FItemList(i).FShopid %>');"><%= offstock.FItemList(i).Fipno %></a></td>
    	<td><a href="javascript:PopItemIpChulListOffLine('<%= fromdate %>','<%= todate %>','<%= itemgubun %>','<%= itemid %>','<%= Itemoption %>','S', '<%= offstock.FItemList(i).FShopid %>');"><%= -1 * offstock.FItemList(i).Freno %></a></td>
    	<td><%= offstock.FItemList(i).Fupcheipno %></td>
    	<td><%= -1 * offstock.FItemList(i).Fupchereno %></td>
    	<td><%= -1 * offstock.FItemList(i).Fsellno %></td>
    	<td bgcolor="F4F4F4"><b><%= offstock.FItemList(i).Fcurrno %></b></td>
    	<td></td>
    	<td></td>
    	<td bgcolor="F4F4F4"></td>
    	<td></td>
    	<td bgcolor="F4F4F4"></td>
    	<td></td>
    </tr>
    <%
    if not IsNULL(offstock.FItemList(i).Fipno) then 	offtotipno = offtotipno + offstock.FItemList(i).Fipno
    if not IsNULL(offstock.FItemList(i).Freno) then 	offtotreno = offtotreno + offstock.FItemList(i).Freno
    if not IsNULL(offstock.FItemList(i).Fupcheipno) then 	offtotupcheipno = offtotupcheipno + offstock.FItemList(i).Fupcheipno
    if not IsNULL(offstock.FItemList(i).Fupchereno) then 	offtotupchereno = offtotupchereno + offstock.FItemList(i).Fupchereno
    if not IsNULL(offstock.FItemList(i).Fsellno) then 	offtotsellno = offtotsellno + offstock.FItemList(i).Fsellno
    if not IsNULL(offstock.FItemList(i).Fcurrno) then 	offtotalstock = offtotalstock + offstock.FItemList(i).Fcurrno
    %>
    <% next %>
    <tr align="center" bgcolor="#EEEEEE">
    	<td></td>
    	<td></td>
    	<td><%= offtotipno %></td>
    	<td><%= -1 * offtotreno %></td>
    	<td><%= offtotupcheipno %></td>
    	<td><%= -1 * offtotupchereno %></td>
    	<td><%= offtotsellno %></td>
    	<td bgcolor="F4F4F4"><b><%= offtotalstock %></b></td>
    	<td></td>
    	<td></td>
    	<td bgcolor="F4F4F4"></td>
    	<td></td>
    	<td bgcolor="F4F4F4"></td>
    	<td></td>
</table>



<% else %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#000000">
    <tr align="center" bgcolor="#DDDDFF">
    <td align=center bgcolor="#FFFFFF">검색 결과가 없습니다.</td>
    </tr>
</table>
<% end if %>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->


<% if (oitemoption.FResultCount>0) and (itemoption="0000") then %>
<script language='javascript'>
alert('옵션 선택 후 재 검색하세요.');
</script>
<% elseif (oitemoption.FResultCount<1) and (itemoption<>"0000") then %>
<script language='javascript'>
alert('재 검색하세요.');
</script>
<% end if %>
<%
set oitemoption = Nothing
set ojaegoitem = Nothing
set osummarystock = Nothing
set offstock = Nothing
%>
<form name=frmrefresh method=post action="dostockrefresh.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="refreshstartdate" value="<%= BasicMonth + "-01" %>">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="<%= itemoption %>">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->