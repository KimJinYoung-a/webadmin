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
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/maechul/statistic/statisticCls_dw.asp" -->
<%
dim i, yyyy, tmpyyyymm, mm_MaechulProfit, tot_MaechulProfit, mm_beforeitemcostsum
dim mm_itemTotalSum, mm_itemOrdercnt, mm_itemavrPrice
	yyyy = requestcheckvar(request("yyyy"),4)

if yyyy="" then yyyy = year(date)

dim cStatistic
Set cStatistic = New cStaticTotalClass_list
	cStatistic.FRectyyyy = yyyy
	cStatistic.fStatistic_monthly_userlevel()
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

<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>판매월</td>
    <td>구매총액</td>
    <td>주문건수</td>
    <td>객단가</td>
    <td>회원등급</td>
</tr>
<%
if cStatistic.FTotalCount > 0 then
For i = 0 To cStatistic.FTotalCount -1
if ((tmpyyyymm <> cStatistic.flist(i).Fyyyymm) and (i <> 0)) then
%>
	<tr align="center" bgcolor="#f1f1f1">
		<td><%= tmpyyyymm %> 합계</td>
		<td><%= FormatNumber(mm_itemTotalSum,0) %></td>
		<td><%= FormatNumber(mm_itemOrdercnt,0) %></td>
		<td><%= FormatNumber(mm_itemavrPrice,0) %></td>
		<td></td>
	<tr>
<%
	mm_itemTotalSum = 0
	mm_itemOrdercnt = 0
	mm_itemavrPrice = 0
end if


tmpyyyymm = cStatistic.flist(i).Fyyyymm
mm_itemTotalSum = mm_itemTotalSum + cStatistic.flist(i).FTotalSum
mm_itemOrdercnt = mm_itemOrdercnt + cStatistic.flist(i).FOrdercnt
mm_itemavrPrice = mm_itemavrPrice + cStatistic.flist(i).FavrPrice
%>
<tr bgcolor="#FFFFFF" align="center">
	<td><%= cStatistic.flist(i).Fyyyymm %></td>
	<td><%= FormatNumber(cStatistic.flist(i).FTotalSum,0) %></td>
	<td><%= FormatNumber(cStatistic.flist(i).FOrdercnt,0) %></td>
	<td><%= FormatNumber(cStatistic.flist(i).FavrPrice,0) %></td>
	<td>
		<font color="<%= getUserLevelColor(cStatistic.flist(i).Fuserlevel) %>">
		<%= getUserLevelStr(cStatistic.flist(i).Fuserlevel ) %>
		</font>
	</td>
</tr>
<%
Next
%>
<tr align="center" bgcolor="#f1f1f1">
	<td><%= tmpyyyymm %> 합계</td>
	<td><%= FormatNumber(mm_itemTotalSum,0) %></td>
	<td><%= FormatNumber(mm_itemOrdercnt,0) %></td>
	<td><%= FormatNumber(mm_itemavrPrice,0) %></td>
	<td></td>
<tr>
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
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
