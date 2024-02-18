<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인 포인트 통계
' History : 2012.12.21 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/point/pointsum_off_cls.asp" -->

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

fromDate = DateSerial(yyyy1, mm1, "01")
toDate = DateSerial(yyyy2, mm2, LastDayOfThisMonth(yyyy2,mm2) +1)
	
Set csell = New cpointsum_off_list
	csell.FRectStartdate = fromDate
	csell.FRectEndDate = toDate
	csell.FRectonoffgubun = "OFF"
	csell.FPageSize = 100
	csell.FCurrPage	= 1
	csell.fpointsum_sell_day_off()

%>

<script language="javascript">

function searchSubmit()
{

	frm.submit();
}

function pop_use_list(yyyy1, mm1, dd1, yyyy2, mm2, dd2, pointcode){
	var pop_use_list = window.open('/admin/maechul/managementsupport/depositsum_use_list.asp?yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&pointcode='+pointcode+'&menupos=<%=menupos%>','pop_use_list','width=1024,height=768,scrollbars=yes,resizable=yes');
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
    <td>날짜</td>
    <td>적립액</td>
    <td>고객사용액</td>
    <td>온라인전환</td>
    <td>회원탈퇴</td>
    <td>소멸</td>
</tr>
<%
dim totgainpoint, totspendpoint, totonlineshiftpoint, totuseroutpoint, totdelpoint
	totgainpoint = 0
	totspendpoint = 0
	totonlineshiftpoint = 0
	totuseroutpoint = 0
	totdelpoint = 0
	
if csell.FresultCount > 0 then
	
For i = 0 To csell.FresultCount -1

totgainpoint = totgainpoint + csell.fitemlist(i).fgainpoint
totspendpoint = totspendpoint + csell.fitemlist(i).fspendpoint
totonlineshiftpoint = totonlineshiftpoint + csell.fitemlist(i).fonlineshiftpoint
totuseroutpoint = totuseroutpoint + csell.fitemlist(i).fuseroutpoint
totdelpoint = totdelpoint + csell.fitemlist(i).fdelpoint
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<%= csell.fitemlist(i).fYYYYMMdd %>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','gainpoint');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fgainpoint,0) %></a>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','spendpoint');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fspendpoint,0) %></a>
	</td>
		
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','onlineshiftpoint');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fonlineshiftpoint,0) %></a>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','useroutpoint');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fuseroutpoint,0) %></a>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','<%= left(csell.fitemlist(i).fYYYYMMdd,4) %>','<%= mid(csell.fitemlist(i).fYYYYMMdd,6,2) %>','<%= right(csell.fitemlist(i).fYYYYMMdd,2) %>','delpoint');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fdelpoint,0) %></a>
	</td>
</tr>	
<% next %>

<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>합계</td>
    <td><%= FormatNumber(totgainpoint,0) %></td>
    <td><%= FormatNumber(totspendpoint,0) %></td>
    <td><%= FormatNumber(totonlineshiftpoint,0) %></td>
    <td><%= FormatNumber(totuseroutpoint,0) %></td>
    <td><%= FormatNumber(totdelpoint,0) %></td>        
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