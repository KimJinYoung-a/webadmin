<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%

dim page, research, i
dim yyyy1, mm1, yyyy2, mm2, stplace, targetGbn, itemgubun
dim ipgoMWdiv, itemMWdiv, itemid, itemoption
dim startYYYYMMDD, endYYYYMMDD
dim addInfoType
dim lastmwdiv, lastmakerid
dim tmpDate


page       	= requestCheckvar(request("page"),10)
research	= requestCheckvar(request("research"),10)
yyyy1       = requestCheckvar(request("yyyy1"),10)
mm1         = requestCheckvar(request("mm1"),10)
yyyy2       = requestCheckvar(request("yyyy2"),10)
mm2         = requestCheckvar(request("mm2"),10)
stplace     = requestCheckvar(request("stplace"),10)
itemgubun   = requestCheckvar(request("itemgubun"),10)
itemid   	= requestCheckvar(request("itemid"),10)
itemoption  = requestCheckvar(request("itemoption"),10)
lastmwdiv	= requestCheckvar(request("lastmwdiv"),10)
lastmakerid	= requestCheckvar(request("lastmakerid"),32)


page = 1
if (yyyy1="") then
	tmpDate = Left(DateAdd("m", -1, Now()), 7)
	yyyy1 = Left(tmpDate, 4)
	mm1 = Right(tmpDate, 2)
	yyyy2 = yyyy1
	mm2 = mm1

	yyyy1 = "2014"
	mm1 = "01"
end if

if (itemgubun = "") then
	itemgubun = "10"
end if


'// ============================================================================
dim ojaego
set ojaego = new CMonthlyStock

ojaego.FPageSize = 100
ojaego.FCurrPage = page
ojaego.FRectStartYYYYMM = yyyy1 + "-" + mm1
ojaego.FRectEndYYYYMM = yyyy2 + "-" + mm2
ojaego.FRectPlaceGubun = stplace

ojaego.FRectItemGubun = itemgubun
ojaego.FRectItemid = itemid
ojaego.FRectItemOption = itemoption

ojaego.FRectMwDiv = lastmwdiv

if (itemid <> "") then
	ojaego.GetMonthlyAvgPriceLogics
end if
''startYYYYMMDD = yyyy1 + "-" + mm1 + "-01"
''endYYYYMMDD = Left(DateAdd("d", -1, DateSerial(yyyy1, mm1 + 1, 1)), 10)

%>

<script language='javascript'>

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
			&nbsp;
			<font color="#CC3333">년/월 :</font> <% DrawYMYMBox yyyy1,mm1,yyyy2,mm2 %> 월 평균매입가
			&nbsp;
			<font color="#CC3333">입고처:</font>
		    <select name="stplace" class="select">
        		<option value="L" <%= CHKIIF(stplace="L","selected" ,"") %> >물류
        	</select>
			&nbsp;
	    	<font color="#CC3333">매입구분(재고자산):</font>
	        <select name="lastmwdiv" class="select">
				<option value="" <%= CHKIIF(lastmwdiv="","selected" ,"") %> >전체</option>
				<option value="M" <%= CHKIIF(lastmwdiv="M","selected" ,"") %> >매입</option>
				<option value="W" <%= CHKIIF(lastmwdiv="W","selected" ,"") %> >위탁</option>
				<option value="X" <%= CHKIIF(lastmwdiv="X","selected" ,"") %> >기타</option>
        	</select>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			&nbsp;
	    	<font color="#CC3333">상품구분:</font>
        	<select name="itemgubun" class="select">
				<option value="" <%= CHKIIF(itemgubun="","selected" ,"") %> >전체
				<option value="10" <%= CHKIIF(itemgubun="10","selected" ,"") %> >일반(10)
				<option value="70" <%= CHKIIF(itemgubun="70","selected" ,"") %> >소포품(70)
				<option value="85" <%= CHKIIF(itemgubun="85","selected" ,"") %> >사은품(85)
				<option value="80" <%= CHKIIF(itemgubun="80","selected" ,"") %> >사은품(80)
				<option value="90" <%= CHKIIF(itemgubun="90","selected" ,"") %> >오프전용(90)
        	</select>
			&nbsp;
			<font color="#CC3333">상품코드:</font>
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="8">
			&nbsp;
			<font color="#CC3333">옵션:</font>
			<input type="text" class="text" name="itemoption" value="<%= itemoption %>" size="4">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->
<p>

	<h5>작업중...</h5>
	* 물류입고 상품은 물류 및 매장의 시스템재고를 합산하여 계산합니다.(매입구분이 동일한 경우)<br>
	* 매입구분이 다르거나 물류입고 상품이 아닌 경우 매장별로 평균매입가가 계산됩니다.<br>

<p>

<% if (itemid = "") then %>
	<h5>먼저 상품코드를 선택하세요.</h5>
<% else %>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="60" rowspan="2">연월</td>
		<td width="40" rowspan="2">입고처</td>
		<td width="100" rowspan="2">매장</td>
		<td width="30" rowspan="2">구분</td>
		<td width=70 rowspan="2">상품코드</td>
		<td width=40 rowspan="2">옵션</td>

		<td colspan="2">전월재고(물류)</td>

		<td colspan="2">전월재고(매장)</td>

		<td colspan="2">당월입고(물류)</td>

		<td colspan="2">평균매입가(물류)</td>

		<td width=60 rowspan="2">매입구분<br>(물류)</td>
		<td width=120 rowspan="2">브랜드</td>

		<td rowspan="2">비고</td>
	</tr>

    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width=40>수량</td>
		<td width=60>금액</td>
		<td width=40>수량</td>
		<td width=60>금액</td>
		<td width=40>수량</td>
		<td width=60>금액</td>
		<td width=60>전월</td>
		<td width=60>당월</td>
	</tr>

	<% if ojaego.FResultCount >0 then %>
	<% for i=0 to ojaego.FResultcount-1 %>
	<tr bgcolor="#FFFFFF" height=25>
		<td align=center><%= ojaego.FItemList(i).Fyyyymm %></td>
		<td align=center><%= ojaego.FItemList(i).GetStockPlaceName %></td>
		<td align=center><%= ojaego.FItemList(i).Fshopid %></td>
		<td align=center><%= ojaego.FItemList(i).Fitemgubun %></td>
		<td align=center><%= ojaego.FItemList(i).Fitemid %></td>
		<td align=center><%= ojaego.FItemList(i).Fitemoption %></td>

		<td align=right><%= FormatNumber(ojaego.FItemList(i).FtotsysstockPrev, 0) %></td>
		<td align=right>
			<% if Not IsNull(ojaego.FItemList(i).FavgipgoPriceSumPrev) then %>
			<%= FormatNumber(ojaego.FItemList(i).FavgipgoPriceSumPrev, 0) %>
			<% end if %>
		</td>

		<td align=right><%= FormatNumber(ojaego.FItemList(i).FtotsysstockShopPrev, 0) %></td>
		<td align=right><%= FormatNumber(ojaego.FItemList(i).FtotsysstockBuySumShopPrev, 0) %></td>

		<td align=right><%= FormatNumber(ojaego.FItemList(i).FtotItemNo, 0) %></td>
		<td align=right><%= FormatNumber(ojaego.FItemList(i).FtotBuyCash, 0) %></td>

		<td align=right>
			<% if Not IsNull(ojaego.FItemList(i).FavgipgoPricePrev) then %>
			<%= FormatNumber(ojaego.FItemList(i).FavgipgoPricePrev, 0) %>
			<% end if %>
		</td>
		<td align=right><%= FormatNumber(ojaego.FItemList(i).FavgipgoPrice, 0) %></td>

		<td align=center>
			<% if Not IsNull(ojaego.FItemList(i).FlastmwdivPrev) then %>
				<% if (ojaego.FItemList(i).FlastmwdivPrev <> ojaego.FItemList(i).Flastmwdiv) then %>
					<%= ojaego.FItemList(i).FlastmwdivPrev %> -&gt;
				<% end if %>
			<% end if %>
			<%= ojaego.FItemList(i).Flastmwdiv %>
		</td>
		<td align=center>
			<% if Not IsNull(ojaego.FItemList(i).FmakeridPrev) then %>
				<% if (ojaego.FItemList(i).FmakeridPrev <> ojaego.FItemList(i).Fmakerid) then %>
					<%= ojaego.FItemList(i).FmakeridPrev %> -&gt;
				<% end if %>
			<% end if %>
			<%= ojaego.FItemList(i).Fmakerid %>
		</td>

		<td>

	    </td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan=17 align=center>[ 검색결과가 없습니다. ]</td>
	</tr>
<% end if %>

</table>
<% end if %>
<%
set ojaego = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
