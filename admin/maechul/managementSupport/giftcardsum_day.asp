<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기프트 카드 매출 통계
' History : 2012.11.08 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/giftcard/giftcardsum_cls.asp" -->

<%
Dim i, yyyy1,mm1,yyyy2,mm2, fromDate ,toDate ,csell, onoffgubun
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	onoffgubun     = request("onoffgubun")
	
if (yyyy1="") then yyyy1 = Cstr(Year( dateadd("m",-3,date()) ))
if (mm1="") then mm1 = Cstr(Month( dateadd("m",-3,date()) ))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))

fromDate = DateSerial(yyyy1, mm1, "01")
toDate = DateSerial(yyyy2, mm2, LastDayOfThisMonth(yyyy2,mm2) +1)

Set csell = New cgiftcardsum_list
	csell.FRectStartdate = fromDate
	csell.FRectEndDate = toDate
	csell.FRectonoffgubun = onoffgubun
	csell.FPageSize = 100
	csell.FCurrPage	= 1
	csell.fgiftcardsum_sell_day()

%>

<script language="javascript">

function searchSubmit()
{

	frm.submit();
}

function pop_sell_list(yyyy1, mm1, dd1, yyyy2, mm2, dd2, accountdiv){
	var pop_sell_list = window.open('/admin/maechul/managementsupport/giftcardsum_sell_list.asp?yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&accountdiv='+accountdiv+'&onoffgubun=<%=onoffgubun%>&menupos=<%=menupos%>','pop_sell_list','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_sell_list.focus();
}

function pop_use_list(yyyy1, mm1, dd1, yyyy2, mm2, dd2, jukyocd){
	var pop_use_list = window.open('/admin/maechul/managementsupport/giftcardsum_use_list.asp?yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&jukyocd='+jukyocd+'&onoffgubun=<%=onoffgubun%>&menupos=<%=menupos%>','pop_use_list','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_use_list.focus();
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
			<td height="25">
				기간 : <% DrawYMBoxdynamic "yyyy1",yyyy1,"mm1",mm1,"" %> ~ <% DrawYMBoxdynamic "yyyy2",yyyy2,"mm2",mm2,"" %>
				사용구분 : <% drawonoffgubun "onoffgubun",onoffgubun," onchange='javascript:searchSubmit();'" %>
			</td>
		</tr>
	    </table>
	</td>	
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="검색" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<Br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
	</td>
	<td align="right">	
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		검색결과 : <b><%= csell.FresultCount %></b> ※ 총 100건까지 검색 됩니다.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>날짜</td>
    <td>적립액(판매내역)</td>
    <td>고객사용액</td>
    <td>환불</td>
    <td>회원탈퇴</td>
    <td>소멸</td>
</tr>
<%
dim totsellCash, totuseCash, totrefundCash, totuseroutCash, totdelcash
	totsellCash = 0
	totuseCash = 0
	totrefundCash = 0
	totuseroutCash = 0
	totdelcash = 0
	
if csell.FresultCount > 0 then
	
For i = 0 To csell.FresultCount -1

totsellCash = totsellCash + csell.fitemlist(i).fsellCash
totuseCash = totuseCash + csell.fitemlist(i).fuseCash
totrefundCash = totrefundCash + csell.fitemlist(i).frefundCash
totuseroutCash = totuseroutCash + csell.fitemlist(i).fuseroutCash
totdelcash = totdelcash + csell.fitemlist(i).fdelcash
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<%= csell.fitemlist(i).fYYYYMMdd %>
	</td>
	<td>
		<a href="javascript:pop_sell_list('<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fsellCash,0) %></a>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','useCash');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fuseCash,0) %></a>
	</td>
		
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','refundCash');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).frefundCash,0) %></a>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','useroutCash');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fuseroutCash,0) %></a>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','delcash');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fdelcash,0) %></a>
	</td>
</tr>	
<% next %>

<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>합계</td>
    <td><%= FormatNumber(totsellCash,0) %></td> 
    <td><%= FormatNumber(totuseCash,0) %></td>
    <td><%= FormatNumber(totrefundCash,0) %></td> 
    <td><%= FormatNumber(totuseroutCash,0) %></td>
    <td><%= FormatNumber(totdelcash,0) %></td>         
</tr>

<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="25">등록된 내용이 없습니다.</td>
</tr>
<% end if %>
</table>

<% 
Set csell = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->