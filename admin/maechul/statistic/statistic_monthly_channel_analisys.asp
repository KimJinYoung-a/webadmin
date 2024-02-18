<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 채널통계-월별
' History : 2016.07.25 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_analisys.asp" -->
<%
dim i, yyyy, tmpyyyymm, mm_MaechulProfit, tot_MaechulProfit, mm_beforeitemcostsum
dim mm_itemcostsum, mm_buycashsum, mm_ordercnt, mm_itemnosum, tot_itemcnt
	yyyy = requestcheckvar(request("yyyy"),4)

if yyyy="" then yyyy = year(date)

dim cStatistic
Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectyyyy = yyyy
	cStatistic.fStatistic_monthly_channel()
%>

<script type='text/javascript'>

function searchSubmit(){
    frm.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">검색</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="30">
				* 날짜 : <% DrawyearBoxdynamic "yyyy", yyyy, " onchange='searchSubmit();'" %>
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="검색" onClick="searchSubmit();"></td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<p>

* 주문건수는 월별로 출고된 상품이 있는 주문의 건수입니다.(2개월에 걸쳐 상품이 출고되면 각각 1건씩.)

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>판매월</td>
	<td>구매채널</td>
    <td>구매총액[상품]<br>(상품쿠폰적용)</td>
    <td>주문건수</td>
    <td>객단가<br>(주문)</td>
    <td>상품수량</td>
    <td>객단가<br>(상품)</td>
    <td>매출수익</td>
	<td>매출수익률</td>
    <td>채널비중</td>
    <td>전월대비<br>성장률</td>
</tr>
<%
if cStatistic.FTotalCount > 0 then

For i = 0 To cStatistic.FTotalCount -1

if tmpyyyymm <> cStatistic.flist(i).Fyyyymm and i <> 0 then
%>
	<tr align="center" bgcolor="#f1f1f1">
		<td colspan=2><%= tmpyyyymm %> 합계</td>
		<td align="right">
			<% '/매출액 %>
			<%= FormatNumber(mm_itemcostsum,0) %>
		</td>
		<td align="right">
			<% '/주문건수 %>
			<%= FormatNumber(mm_ordercnt,0) %>
		</td>
		<td align="right">
			<%
			'/객단가(주문)
			if mm_itemcostsum<>0 and mm_ordercnt<>0 then
			%>
				<%= FormatNumber(mm_itemcostsum/mm_ordercnt,0) %>
			<% else %>
				0
			<% end if %>
		</td>
		<td align="right">
			<% '/상품수량 %>
			<%= FormatNumber(mm_itemnosum,0) %>
		</td>
		<td align="right">
			<%
			'/객단가(상품)
			if mm_itemcostsum<>0 and mm_itemnosum<>0 then
			%>
				<%= FormatNumber(mm_itemcostsum/mm_itemnosum,0) %>
			<% else %>
				0
			<% end if %>
		</td>
		<td align="right">
			<% '/매출수익 %>
			<%= FormatNumber(mm_MaechulProfit,0) %>
		</td>
		<td>
			<%
			'/매출수익률
			if mm_itemcostsum<>0 then
			%>
				<%= round( ((( mm_itemcostsum-mm_buycashsum ) / mm_itemcostsum )*100) ,2) %>%
			<% else %>
				<%= round( ((( mm_itemcostsum-mm_buycashsum ) / 1 )*100) ,2) %>%
			<% end if %>
		</td>
		<td></td>
		<td>
			<%
			'/전월대비 성장률
			if mm_itemcostsum<>0 and mm_beforeitemcostsum<>0 then
			%>
				<%= round( (( mm_itemcostsum/mm_beforeitemcostsum )*100) -100 ,2) %>%
			<% else %>
				0%
			<% end if %>
		</td>
	</tr>
<%
	mm_itemcostsum = 0
	mm_beforeitemcostsum = 0
	mm_buycashsum = 0
	mm_ordercnt = 0
	mm_itemnosum = 0
	mm_MaechulProfit = 0
end if

tmpyyyymm = cStatistic.flist(i).Fyyyymm
mm_itemcostsum = mm_itemcostsum + cStatistic.flist(i).Fitemcostsum
mm_beforeitemcostsum = mm_beforeitemcostsum + cStatistic.flist(i).Fbeforeitemcostsum
mm_buycashsum = mm_buycashsum + cStatistic.flist(i).fbuycashsum
mm_ordercnt = mm_ordercnt + cStatistic.flist(i).fordercnt
mm_itemnosum = mm_itemnosum + cStatistic.flist(i).fitemnosum
mm_MaechulProfit = mm_MaechulProfit + cStatistic.flist(i).fMaechulProfit
tot_MaechulProfit = tot_MaechulProfit + mm_MaechulProfit
%>

<tr bgcolor="#FFFFFF" align="center">
	<td>
		<%= cStatistic.flist(i).Fyyyymm %>
	</td>
	<td>
		<%= getchannelname(cStatistic.flist(i).Fchannel) %>
	</td>
	<td align="right">
		<% '/매출액 %>
		<%= FormatNumber(cStatistic.flist(i).Fitemcostsum,0) %>
	</td>
	<td align="right">
		<% '/주문건수 %>
		<%= FormatNumber(cStatistic.flist(i).Fordercnt,0) %>
	</td>
	<td align="right">
		<% '/객단가(주문) %>
		<%= FormatNumber(cStatistic.flist(i).forderunit,0) %>
	</td>
	<td align="right">
		<% '/상품수량 %>
		<%= FormatNumber(cStatistic.flist(i).Fitemnosum,0) %>
	</td>
	<td align="right">
		<% '/객단가(상품) %>
		<%= FormatNumber(cStatistic.flist(i).fitemunit,0) %>
	</td>
	<td align="right">
		<% '매출수익 %>
		<%= FormatNumber(cStatistic.flist(i).fMaechulProfit,0) %>
	</td>
	<td>
		<% '매출수익률 %>
		<%= cStatistic.flist(i).FMaechulProfitPer %>%
	</td>
	<td>
		<%
		'채널비중
		if cStatistic.flist(i).fchannelitemcostsum<>0 and cStatistic.flist(i).Fitemcostsum<>0 then
		%>
			<%= round((cStatistic.flist(i).Fitemcostsum/cStatistic.flist(i).fchannelitemcostsum)*100,2) %>%
		<% else %>
			0%
		<% end if %>
	</td>
	<td>
		<% '/전월대비<br>성장률 %>
		<%= cStatistic.flist(i).fbeforemmper %>%
	</td>
</tr>
<%
Next
%>

<tr align="center" bgcolor="#f1f1f1">
	<td colspan=2><%= tmpyyyymm %> 합계</td>
	<td align="right">
		<% '/매출액 %>
		<%= FormatNumber(mm_itemcostsum,0) %>
	</td>
	<td align="right">
		<% '/주문건수 %>
		<%= FormatNumber(mm_ordercnt,0) %>
	</td>
	<td align="right">
		<%
		'/객단가(주문)
		if mm_itemcostsum<>0 and mm_ordercnt<>0 then
		%>
			<%= FormatNumber(mm_itemcostsum/mm_ordercnt,0) %>
		<% else %>
			0
		<% end if %>
	</td>
	<td align="right">
		<% '/상품수량 %>
		<%= FormatNumber(mm_itemnosum,0) %>
	</td>
	<td align="right">
		<%
		'/객단가(상품)
		if mm_itemcostsum<>0 and mm_itemnosum<>0 then
		%>
			<%= FormatNumber(mm_itemcostsum/mm_itemnosum,0) %>
		<% else %>
			0
		<% end if %>
	</td>
	<td align="right">
		<% '/매출수익 %>
		<%= FormatNumber(mm_MaechulProfit,0) %>
	</td>
	<td>
		<%
		'/매출수익률
		if mm_itemcostsum<>0 then
		%>
			<%= round( ((( mm_itemcostsum-mm_buycashsum ) / mm_itemcostsum )*100) ,2) %>%
		<% else %>
			<%= round( ((( mm_itemcostsum-mm_buycashsum ) / 1 )*100) ,2) %>%
		<% end if %>
	</td>
	<td></td>
	<td>
		<%
		'/전월대비 성장률
		if mm_itemcostsum<>0 and mm_beforeitemcostsum<>0 then
		%>
			<%= round( (( mm_itemcostsum/mm_beforeitemcostsum )*100) -100 ,2) %>%
		<% else %>
			0%
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="#f1f1f1">
	<td colspan=2>총합계</td>
	<td align="right">
		<% '/매출액 %>
		<%= FormatNumber(cStatistic.totitemcostsum,0) %>
	</td>
	<td align="right">
		<% '/주문건수 %>
		<%= FormatNumber(cStatistic.totordercnt,0) %>
	</td>
	<td align="right">
		<%
		'/객단가(주문)
		if cStatistic.totitemcostsum<>0 and cStatistic.totordercnt<>0 then
		%>
			<%= FormatNumber(cStatistic.totitemcostsum/cStatistic.totordercnt,0) %>
		<% else %>
			0
		<% end if %>
	</td>
	<td align="right">
		<% '/상품수량 %>
		<%= FormatNumber(cStatistic.totitemnosum,0) %>
	</td>
	<td align="right">
		<%
		'/객단가(상품)
		if cStatistic.totitemcostsum<>0 and cStatistic.totitemnosum<>0 then
		%>
			<%= FormatNumber(cStatistic.totitemcostsum/cStatistic.totitemnosum,0) %>
		<% else %>
			0
		<% end if %>
	</td>
	<td align="right">
		<% '/매출수익 %>
		<%= FormatNumber(tot_MaechulProfit,0) %>
	</td>
	<td>
		<%
		'/매출수익률
		if mm_itemcostsum<>0 then
		%>
			<%= round( ((( cStatistic.totitemcostsum-cStatistic.totbuycashsum ) / cStatistic.totitemcostsum )*100) ,2) %>%
		<% else %>
			<%= round( ((( cStatistic.totitemcostsum-cStatistic.totbuycashsum ) / 1 )*100) ,2) %>%
		<% end if %>
	</td>
	<td></td>
	<td>
		<%
		'/전월대비 성장률
		if cStatistic.totitemcostsum<>0 and cStatistic.totbeforeitemcostsum<>0 then
		%>
			<%= round( (( cStatistic.totitemcostsum/cStatistic.totbeforeitemcostsum )*100) -100 ,2) %>%
		<% else %>
			0%
		<% end if %>
	</td>
</tr>

<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="20">매출이 없습니다.</td>
	</tr>
<% end if %>

</table>

<%
set cStatistic = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
