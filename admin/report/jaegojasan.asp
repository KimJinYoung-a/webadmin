<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jaegocls.asp"-->
<H3>사용안함 - 수정중</H3>
<%
dim yyyy1,mm1,designer,mwgubun,isusing
designer = request("designer")
yyyy1 = request("yyyy1")
mm1 = request("mm1")
isusing = request("isusing")
mwgubun = request("mwgubun")

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

dim ojaego, yyyymm, enddate, pre3month
yyyymm = yyyy1 + "-" + mm1
enddate = dateserial(yyyy1,mm1+1,1)
pre3month = dateserial(yyyy1,mm1-2,1)

set ojaego = new CJaegoEval
ojaego.FRectYYYYMM   = yyyymm
ojaego.FRectIsusing = isusing
'ojaego.FRectDesigner = designer
ojaego.GetMonthJeagoSum

dim ojaegomaker
set ojaegomaker = new CJaegoEval

if mwgubun<>"" then
	ojaegomaker.FRectYYYYMM = yyyymm
	ojaegomaker.FRectStartDate = yyyymm + "-01"
	ojaegomaker.FRectEndDate = CStr(enddate)
	ojaegomaker.FRect3MonthStartDate = CStr(pre3month)
	ojaegomaker.FRectMwDiv = mwgubun
	''ojaegomaker.GetMonthJeagoSumByMaker
end if

dim totno, totbuy, totsell,i
dim totonlinemeaip, totofflinemeaip, totoffchulgobuycash
dim totoffchulgosuplycash, totFMonthMeachulSum, tot3MonthMeachulSum, totoff3monthchulgosuplycash
%>
<script language='javascript'>
function popStockJasan(mwdiv,yyyy1,mm1,designer){
	var popwin = window.open("jaegojasandetail.asp?mwdiv=" + mwdiv + "&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&designer=" + designer,"stockdetail","width=1000,height=620,scrollbars=yes, resizable=yes");
	popwin.focus();
}
</script>


<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>월말재고자산 및 회전율</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br>물류센터의 월말 재고자산 및 브랜드별 회전율 정보입니다.
			<br>
			<br>사용함 상품을 기본설정으로....
			<br>사용안하는 브랜드 먹표시
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	검색 : <% DrawYMBox yyyy1,mm1 %> 말일자 재고자산
	        	&nbsp;&nbsp;&nbsp;
	        	사용구분:
	        	<input type="radio" name="isusing" value="">전체
	        	<input type="radio" name="isusing" value="Y">사용함
	        	<input type="radio" name="isusing" value="N">사용안함
	        </td>
	        <td valign="top" align="right">
	        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="100">매입구분</td>
    	<td width="100">총재고수량</td>
    	<td width="100">소비자가</td>
    	<td width="100">평균마진</td>
    	<td width="100">매입가</td>
    	<td>비고</td>
    </tr>
    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    totno   = totno + ojaego.FItemList(i).FTotCount
    totbuy  = totbuy + ojaego.FItemList(i).FTotBuySum
    totsell = totsell + ojaego.FItemList(i).FTotSellSum
    %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><a href="?menupos=<%= menupos %>&mwgubun=<%= ojaego.FItemList(i).FMaeIpGubun %>&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>" target="_blank"><%= ojaego.FItemList(i).getMaeipGubunName %></a></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotSellSum,0) %></td>
    	<td><%= clng((1-(ojaego.FItemList(i).FTotBuySum)/(ojaego.FItemList(i).FTotSellSum))*100)/100 %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>
    	<td></td>
    </tr>
    <% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>총계</td>
    	<td align="right" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<td></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>

    	<td></td>
    </tr>
</table>

<% if mwgubun="M" then %>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
            <br>
            * 매입 내역 - 브랜드별<br>
            * 오프 출고액 - 매입 특정 구분없이 총 출고액
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<% elseif mwgubun="W" then %>

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr height="5" valign="top">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
            <br>
            * 특정 내역 - 브랜드별
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<% end if %>

<%
totno = 0
totbuy = 0
totsell = 0
%>

<% if mwgubun<>"" then %>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td rowspan=2>브랜드</td>
    	<td rowspan=2>총재고</td>
        <td rowspan=2>재고총액<br>(매입가)</td>
    <!--	<td rowspan=2>재고액<br>(소비자가)</td>     -->
    	<td colspan=2><%= mm1 %>월 총매입액</td>
    	<td rowspan=2><%= mm1 %>월 오프라인<br>매입액</td>

    	<td colspan=3><%= mm1 %>월 회전율

    	<td colspan=3>3개월 회전율
    </tr>
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="100"><%= mm1 %>월 매입액</td>
    	<td width="100"><%= mm1 %>월 off별도매입</td>

    	<td width="100"><%= mm1 %>월 온라인<br>매출액</td>
    	<td width="100"><%= mm1 %>월 오프라인<br>매출(출고)액</td>
    	<td width="100"><%= mm1 %>월회전율</td>

    	<td width="100">3개월 온라인<br>매출액</td>
    	<td width="100">3개월 오프라인<br>매출(출고)액</td>
    	<td width="100">3개월회전율</td>
    </tr>
    <% for i=0 to ojaegomaker.FResultCount -1 %>
    <%
        totno   = totno + ojaegomaker.FItemList(i).FTotCount
        totbuy  = totbuy + ojaegomaker.FItemList(i).FTotBuySum
        totsell = totsell + ojaegomaker.FItemList(i).FTotSellSum

        totonlinemeaip = totonlinemeaip + ojaegomaker.FItemList(i).Fonlinemeaip
        totofflinemeaip = totofflinemeaip + ojaegomaker.FItemList(i).Fofflinemeaip
        totoffchulgobuycash = totoffchulgobuycash + ojaegomaker.FItemList(i).Foffchulgobuycash
        totoffchulgosuplycash = totoffchulgosuplycash + ojaegomaker.FItemList(i).Foffchulgosuplycash
        totoff3monthchulgosuplycash = totoff3monthchulgosuplycash + ojaegomaker.FItemList(i).Foff3monthchulgosuplycash
        totFMonthMeachulSum = totFMonthMeachulSum + ojaegomaker.FItemList(i).FMonthMeachulSum
        tot3MonthMeachulSum = tot3MonthMeachulSum + ojaegomaker.FItemList(i).F3MonthMeachulSum

    %>
    <tr bgcolor="#FFFFFF">
    	<td><a href="javascript:popStockJasan('<%= mwgubun %>','<%= yyyy1 %>','<%= mm1 %>','<%= ojaegomaker.FItemList(i).Fmakerid %>');"><%= ojaegomaker.FItemList(i).Fmakerid %></a></td>
    	<td align="center"><%= FormatNumber(ojaegomaker.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).FTotBuySum,0) %></td>
    <!--	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).FTotSellSum,0) %></td>  -->
    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).Fonlinemeaip,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).Fofflinemeaip,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).Foffchulgobuycash,0) %></td>

    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).FMonthMeachulSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).Foffchulgosuplycash,0) %></td>
    	<td align="center">
    	<% if ojaegomaker.FItemList(i).FTotBuySum<>0 then %>
    		<%= CLng((ojaegomaker.FItemList(i).Foffchulgosuplycash+ojaegomaker.FItemList(i).FMonthMeachulSum)/ojaegomaker.FItemList(i).FTotBuySum*100)/100 %>
    	<% end if %>
    	</td>

    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).F3MonthMeachulSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegomaker.FItemList(i).Foff3monthchulgosuplycash,0) %></td>
    	<td align="center">
    	<% if ojaegomaker.FItemList(i).FTotBuySum<>0 then %>
    		<%= CLng((ojaegomaker.FItemList(i).Foff3monthchulgosuplycash+ojaegomaker.FItemList(i).F3MonthMeachulSum)/ojaegomaker.FItemList(i).FTotBuySum*100)/100 %>
    	<% end if %>
    	</td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
    	<td></td>
    	<td align="center"><%= FormatNumber(totno,0) %></td>
    	<td align="right"><%= FormatNumber(totbuy,0) %></td>
    <!--	<td align="right"><%= FormatNumber(totsell,0) %></td>   -->
    	<td align="right"><%= FormatNumber(totonlinemeaip,0) %></td>
    	<td align="right"><%= FormatNumber(totofflinemeaip,0) %></td>
    	<td align="right"><%= FormatNumber(totoffchulgobuycash,0) %></td>

    	<td align="right"><%= FormatNumber(totFMonthMeachulSum,0) %></td>
    	<td align="right"><%= FormatNumber(totoffchulgosuplycash,0) %></td>
    	<td align="center">
    	<% if totbuy<>0 then %>
    		<%= CLng((totoffchulgosuplycash+totFMonthMeachulSum)/totbuy*100)/100 %>
    	<% end if %>
    	</td>
    	<td align="right"><%= FormatNumber(tot3MonthMeachulSum,0) %></td>
    	<td align="right"><%= FormatNumber(totoff3monthchulgosuplycash,0) %></td>
    	<td align="center">
    	<% if totbuy<>0 then %>
    		<%= CLng((totoff3monthchulgosuplycash+tot3MonthMeachulSum)/totbuy*100)/100 %>
    	<% end if %>
    	</td>
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

<%
set ojaegomaker = Nothing
set ojaego = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->