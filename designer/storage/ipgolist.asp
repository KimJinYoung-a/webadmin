<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->

<%
dim code
dim page,designer, statecd, onoffgubun, divcode, rackipgoyn
dim yyyy1,mm1
dim fromDate, toDate

page = requestCheckVar(request("page"),20)
designer = session("ssBctID")
statecd = requestCheckVar(request("statecd"),30)
code = requestCheckVar(request("code"),50)				' 입고 코드
onoffgubun = requestCheckVar(request("onoffgubun"),20)	' 온/오프 구분
divcode = requestCheckVar(request("divcode"),20)		' 매입 구분
rackipgoyn = requestCheckVar(request("rackipgoyn"),10)	'


'// 입고일 검색에 필요한 변수 대입

yyyy1 = requestCheckVar(request("yyyy1"),10)
mm1	  = requestCheckVar(request("mm1"),10)

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))

fromDate = CStr(DateSerial(yyyy1, mm1, 01))
toDate = CStr(DateSerial(yyyy1, mm1 + 1, 01))

if onoffgubun="" then onoffgubun="all"

if page="" then page=1

dim oipchul
set oipchul = new CIpChulStorage
oipchul.FPageSize=300
oipchul.FRectCode = code
oipchul.FRectDivcode = divcode
oipchul.FRectExecuteDtStart = fromDate
oipchul.FRectExecuteDtEnd   = toDate


if code="" then
oipchul.FRectCodeGubun = "ST"  ''입고
oipchul.FRectSocID = designer
oipchul.FRectOnOffGubun = onoffgubun
end if
oipchul.GetIpChulgoList

dim i
dim totalsellcash,totalsuply
totalsellcash = 0
totalsuply	  = 0
%>
<script language='javascript'>
function PopIpgoSheet(v,itype){
	var popwin;
	popwin = window.open('popipgosheet.asp?idx=' + v + '&itype=' + itype,'popipgosheet','width=760,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function ExcelSheet(v,itype){
	window.open('popipgosheet_excel.asp?idx=' + v + '&itype=' + itype + '&xl=on');
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td>
	        	대상년월 :&nbsp;<% DrawYMBox yyyy1,mm1 %>
	        </td>
	        <td align="right">
	        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">입고코드</td>
		<td width="100">브랜드ID</td>
		<td width=80>입고처리자</td>
		<td width=80>예정일</td>
		<td width=80>입고일</td>
		<td width="100">소비자가총액</td>
		<td width="100">공급가총액</td>
		<td width=70>계약구분</td>
		<td width=80>내역서출력</td>
		<td>비고</td>
	</tr>
	<% if oipchul.FResultCount >0 then %>
	<% for i=0 to oipchul.FResultcount-1 %>
	<%
	totalsellcash = totalsellcash + oipchul.FItemList(i).Ftotalsellcash
	totalsuply	  = totalsuply + oipchul.FItemList(i).Ftotalsuplycash
	%>
	<tr align=center bgcolor="#FFFFFF">
		<td><a href="javascript:PopIpgoSheet('<%= oipchul.FItemList(i).Fid %>','');"><%= oipchul.FItemList(i).Fcode %></a></td>
		<td><%= oipchul.FItemList(i).Fsocid %></td>
		<td><%= oipchul.FItemList(i).Fchargename %></td>
		<td><font color="#777777"><%= Left(oipchul.FItemList(i).Fscheduledt,10) %></font></td>
		<td><%= Left(oipchul.FItemList(i).Fexecutedt,10) %></td>
		<td align=right><font color="<%= oipchul.FItemList(i).GetMinusColor(oipchul.FItemList(i).Ftotalsellcash) %>"><%= FormatNumber(oipchul.FItemList(i).Ftotalsellcash,0) %></font></td>
		<td align=right><b><font color="<%= oipchul.FItemList(i).GetMinusColor(oipchul.FItemList(i).Ftotalsuplycash) %>"><%= FormatNumber(oipchul.FItemList(i).Ftotalsuplycash,0) %></font></b></td>
		<td><font color="<%= oipchul.FItemList(i).GetDivCodeColor %>"><%= oipchul.FItemList(i).GetDivCodeName %></font></td>
		<td>
	    	<a href="javascript:PopIpgoSheet('<%= oipchul.FItemList(i).Fid %>','');"><img src="/images/iexplorer.gif" border="0"></a>
	        <a href="javascript:ExcelSheet('<%= oipchul.FItemList(i).Fid %>','');"><img src="/images/iexcel.gif" border="0"></a>
	    </td>
	    <td></td>
	</tr>
	<% next %>
	<tr align=center bgcolor="#FFFFFF">
		<td>총계</td>
		<td colspan=4></td>
		<td align=right><%= formatNumber(totalsellcash,0) %></td>
		<td align=right><b><%= formatNumber(totalsuply,0) %></b></td>
		<td></td>
		<td></td>
		<td></td>
	</tr>

	<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan=11 align="center">[ 검색결과가 없습니다. ]</td>
	</tr>
	<% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->



<%
set oipchul = Nothing
%>


<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->