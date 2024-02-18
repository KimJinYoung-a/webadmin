<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  온라인 포인트 통계
' History : 2013.01.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/mileage/pointsum_on_cls.asp" -->

<%
Dim i, yyyy1,mm1,yyyy2,mm2, fromDate ,toDate ,csell
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")

if (yyyy1="") then yyyy1 = Cstr(Year( dateadd("m",-3,date()) ))
if (mm1="") then mm1 = Cstr(Month( dateadd("m",-3,date()) ))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))

fromDate = left(DateSerial(yyyy1, mm1,"01"),7)
toDate = left(DateSerial(yyyy2, mm2+1,"01"),7)

Set csell = New cpointsum_on_list
	csell.FRectStartdate = fromDate
	csell.FRectEndDate = toDate
	csell.FPageSize = 100
	csell.FCurrPage	= 1
	csell.FRectonoffgubun = "ON"
	csell.fpointsum_sell_month_on()


dim item_M, item_N, item_S, item_T, item_U
dim tot_M, tot_N, tot_S, tot_T, tot_U

dim item_S_60, tot_S_60

dim item_X, item_Y, item_Z
dim item_XN, item_YN, item_ZN
dim item_ZZ

%>

<script language="javascript">

function searchSubmit()
{

	frm.submit();
}

function pop_use_list(yyyy1, mm1, dd1, yyyy2, mm2, dd2, jukyocd){
	var pop_use_list = window.open('/admin/maechul/managementsupport/pointsum_day_on.asp?yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&jukyocd='+jukyocd+'&menupos=<%=menupos%>','pop_use_list','width=1024,height=768,scrollbars=yes,resizable=yes');
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
    <td height="25" width="60">날짜</td>
    <td>이월잔액<br>(A)</td>
    <td>ON<br>적립액(B)</td>
	<td>OFF-&gt;ON<br>전환적립(C)</td>
    <td>고객사용액<br>(D)</td>
    <td>회원탈퇴<br>(E)</td>
    <td>소멸<br>(F)</td>
    <td>월잔액(R)</td>
	<td>사용률<br>(S=D/(B+C))</td>
	<td>최근60개월<br>사용률<br>(V)</td>

	<td>충당액<br>(X=R*V)</td>
	<td>원가율<br>(Y)</td>
	<td>최종충당액<br>(Z=X*Y)</td>
	<td>월별충당액</td>
</tr>
<%
dim totbeforeremainpoint, totgainpoint, totspendpoint, totofflineshiftpoint, totuseroutpoint, totdelpoint, totremaincash
	totbeforeremainpoint = 0
	totgainpoint = 0
	totspendpoint = 0
	totofflineshiftpoint = 0
	totuseroutpoint = 0
	totdelpoint = 0
	totremaincash = 0

dim totgainpoint60, totspendpoint60, totofflineshiftpoint60

totgainpoint60 = 0
totspendpoint60 = 0
totofflineshiftpoint60 = 0

if csell.FresultCount > 0 then

For i = 0 To csell.FresultCount -1

totbeforeremainpoint = totbeforeremainpoint + csell.fitemlist(i).fbeforeremainpoint
totgainpoint = totgainpoint + csell.fitemlist(i).fgainpoint
totspendpoint = totspendpoint + csell.fitemlist(i).fspendpoint
totofflineshiftpoint = totofflineshiftpoint + csell.fitemlist(i).fofflineshiftpoint
totuseroutpoint = totuseroutpoint + csell.fitemlist(i).fuseroutpoint
totdelpoint = totdelpoint + csell.fitemlist(i).fdelpoint
totremaincash = totremaincash + csell.fitemlist(i).fremaincash

totgainpoint60 = totgainpoint60 + csell.fitemlist(i).Fgainpoint60mon
totspendpoint60 = totspendpoint60 + csell.fitemlist(i).Fspendpoint60mon
totofflineshiftpoint60 = totofflineshiftpoint60 + csell.fitemlist(i).Fofflineshiftpoint60mon

item_M = csell.fitemlist(i).fgainpoint + csell.fitemlist(i).fofflineshiftpoint
item_N = csell.fitemlist(i).fspendpoint + csell.fitemlist(i).fuseroutpoint + csell.fitemlist(i).fdelpoint

item_S = csell.fitemlist(i).fspendpoint * -1 * 100.0 / (csell.fitemlist(i).fgainpoint + csell.fitemlist(i).fofflineshiftpoint)
item_S_60 = csell.fitemlist(i).Fspendpoint60mon * -1 * 100.0 / (csell.fitemlist(i).Fgainpoint60mon + csell.fitemlist(i).Fofflineshiftpoint60mon)

item_T = 100.0 - item_S
item_U = (csell.fitemlist(i).fremaincash - csell.fitemlist(i).fbeforeremainpoint) * item_S / 100.0

item_X = csell.fitemlist(i).fremaincash * (csell.fitemlist(i).Fspendpoint60mon * -1 / (csell.fitemlist(i).Fgainpoint60mon + csell.fitemlist(i).Fofflineshiftpoint60mon))
item_Y = csell.fitemlist(i).Fcostpricepercent
item_Z = item_X * item_Y / 100

if (i = (csell.FresultCount - 1)) then
	item_ZZ = NULL
else
	item_XN = csell.fitemlist(i + 1).fremaincash * (csell.fitemlist(i + 1).Fspendpoint60mon * -1 / (csell.fitemlist(i + 1).Fgainpoint60mon + csell.fitemlist(i + 1).Fofflineshiftpoint60mon))
	item_YN = csell.fitemlist(i + 1).Fcostpricepercent
	item_ZN = item_XN * item_YN / 100

	item_ZZ = item_Z - item_ZN
end if

%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td height="25">
		<%= csell.fitemlist(i).fYYYYMM %>
	</td>
	<td>
		<%= FormatNumber(csell.fitemlist(i).fbeforeremainpoint,0) %>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth(left(csell.fitemlist(i).fYYYYMM,4),mid(csell.fitemlist(i).fYYYYMM,6,2)) %>','gainpoint');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fgainpoint,0) %></a></a>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth(left(csell.fitemlist(i).fYYYYMM,4),mid(csell.fitemlist(i).fYYYYMM,6,2)) %>','offlineshiftpoint');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fofflineshiftpoint,0) %></a>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth(left(csell.fitemlist(i).fYYYYMM,4),mid(csell.fitemlist(i).fYYYYMM,6,2)) %>','spendpoint');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fspendpoint,0) %></a>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth(left(csell.fitemlist(i).fYYYYMM,4),mid(csell.fitemlist(i).fYYYYMM,6,2)) %>','useroutpoint');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fuseroutpoint,0) %></a>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth(left(csell.fitemlist(i).fYYYYMM,4),mid(csell.fitemlist(i).fYYYYMM,6,2)) %>','delpoint');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fdelpoint,0) %></a>
	</td>
	<td>
		<%= FormatNumber(csell.fitemlist(i).fremaincash,0) %>
	</td>
	<td>
		<%= FormatNumber(item_S,2) %> %
	</td>
	<td>
		<%= FormatNumber(item_S_60,2) %> %
	</td>

	<td>
		<%= FormatNumber(item_X,0) %>
	</td>
	<td>
		<%= FormatNumber(item_Y,2) %> %
	</td>
	<td>
		<%= FormatNumber(item_Z,0) %>
	</td>
	<td>
		<% if Not IsNull(item_ZZ) then %>
			<%= FormatNumber(item_ZZ,0) %>
		<% end if %>
	</td>
</tr>
<% next %>

<%
tot_M = totgainpoint + totofflineshiftpoint
tot_N = totspendpoint + totuseroutpoint + totdelpoint
tot_S = totspendpoint * -1 * 100.0 / (totgainpoint + totofflineshiftpoint)
tot_S_60 = totspendpoint60 * -1 * 100.0 / (totgainpoint60 + totofflineshiftpoint60)
tot_T = 100.0 - tot_S
tot_U = (totremaincash - totbeforeremainpoint) * tot_S / 100.0
%>

<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td height="25">합계</td>
    <td><%= FormatNumber(totbeforeremainpoint,0) %></td>
    <td><%= FormatNumber(totgainpoint,0) %></td>
	<td><%= FormatNumber(totofflineshiftpoint,0) %></td>
    <td><%= FormatNumber(totspendpoint,0) %></td>
    <td><%= FormatNumber(totuseroutpoint,0) %></td>
    <td><%= FormatNumber(totdelpoint,0) %></td>
    <td><%= FormatNumber(totremaincash,0) %></td>
	<td><%= FormatNumber(tot_S,2) %></td>
	<td><%= FormatNumber(tot_S_60,2) %></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
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
