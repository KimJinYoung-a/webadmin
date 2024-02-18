<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/userjoincls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, startdateStr, nextdateStr, page
dim eventinclude, research
dim rpttype

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
rpttype = request("rpttype")
page = request("page")
eventinclude = request("eventinclude")
research = request("research")

if (research="") then eventinclude="on"
if page="" then page=1

nowdateStr = CStr(now())


if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

if (rpttype="") then rpttype="day"

startdateStr = yyyy1 + "-" + mm1 + "-" + dd1
nextdateStr = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

dim oneuserjoin
set oneuserjoin = new UserJoinClass
oneuserjoin.FRectStart = startdateStr
oneuserjoin.FRectEnd =  nextdateStr
oneuserjoin.FRectGroup = rpttype
oneuserjoin.FRectEventInclude = eventinclude
oneuserjoin.GetUserJoinBySex

dim i
%>

<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<input type="checkbox" name="eventinclude" <% if eventinclude="on" then response.write "checked" %> >결로 유입 가입 포함
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>
<%
const MAXBARSIZE = 500
dim totno, MsexPercent,WsexPercent

totno = oneuserjoin.FManNo + oneuserjoin.FWoManNo

if totno<>0 then
	MsexPercent = CInt(oneuserjoin.FManNo/totno*100)
	WsexPercent = CInt(oneuserjoin.FWoManNo/totno*100)
else
	MsexPercent = 0
	WsexPercent = 0
end if
%>
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
    	<td colspan="4">1. 성별 가입비율</td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td width="100">전체</td>
    	<td width="100" align="right"><%= FormatNumber(totno,0) %></td>
    	<td width="100" align="right">100 (%)</td>
    	<td></td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td>남성</td>
    	<td align="right"><%= FormatNumber(oneuserjoin.FManNo,0) %></td>
    	<td align="right"><%= MsexPercent %> (%)</td>
    	<td><img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%= CInt(MAXBARSIZE * MsexPercent / 100) %>"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td>여성</td>
    	<td align="right"><%= FormatNumber(oneuserjoin.FWoManNo,0) %></td>
    	<td align="right"><%= WsexPercent %> (%)</td>
    	<td><img src="http://partner.10x10.co.kr/images/dot2.gif" height="4" width="<%= CInt(MAXBARSIZE * WsexPercent / 100) %>"></td>
    </tr>
</table>
<br>
<%
oneuserjoin.GetUserJoinByNai
%>

<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
    	<td colspan="5">2. 연령별 가입비율</td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td width="100">전체</td>
    	<td width="100" align="right">
    		<%= FormatNumber(oneuserjoin.FNaiMaster.FManTotal,0) %><br>
    		<%= FormatNumber(oneuserjoin.FNaiMaster.FWoManTotal,0) %>
    	</td>
    	<td width="50" align="right">
    		<%= oneuserjoin.FNaiMaster.GetManTotalPercent %> (%)<br>
    		<%= oneuserjoin.FNaiMaster.GetWoManTotalPercent %> (%)
    	</td>
    	<td width="50" align="right">100 (%)</td>
    	<td></td>
    </tr>
    <% for i=0 to 5  %>
    <tr bgcolor="#FFFFFF">
    	<td width="100"><%= oneuserjoin.FNaiMaster.FItemList(i).FNaiStr %></td>
    	<td width="100" align="right">
    		<%'= FormatNumber(oneuserjoin.FNaiMaster.FItemList(i).FManCount,0) %>
    		<%'= FormatNumber(oneuserjoin.FNaiMaster.FItemList(i).FWoManCount,0) %>
			<% = oneuserjoin.FNaiMaster.FItemList(i).FManCount + oneuserjoin.FNaiMaster.FItemList(i).FWoManCount  %>

    	</td>
    	<td width="50" align="right">
    		<%= oneuserjoin.FNaiMaster.GetManPercent(i) %> (%)<br>
    		<%= oneuserjoin.FNaiMaster.GetWoManPercent(i) %> (%)
    	</td>
    	<td width="50" align="right"><%= oneuserjoin.FNaiMaster.GetTotPercent(i) %> (%)</td>
    	<td>
    		<img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%= CInt(MAXBARSIZE * oneuserjoin.FNaiMaster.GetManPercent(i) / 100) %>"><br>
    		<img src="http://partner.10x10.co.kr/images/dot2.gif" height="4" width="<%= CInt(MAXBARSIZE * oneuserjoin.FNaiMaster.GetWoManPercent(i) / 100) %>">
    	</td>
    </tr>
    <% next %>
</table>
<br>
<%
dim tmppercent
oneuserjoin.GetUserJoinByArea
%>

<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
    	<td colspan="5">3. 지역별 가입비율</td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td width="120">전체</td>
    	<td width="100" align="right"><%= FormatNumber(oneuserjoin.FTotalUsercount,0) %></td>
    	<td width="100" align="right">100 (%)</td>
    	<td></td>
    </tr>
    <% for i=0 to oneuserjoin.FResultCount -1 %>
    <%
    if oneuserjoin.FTotalUsercount=0 then
    	tmppercent = 0
    else
    	tmppercent = CInt(oneuserjoin.FItemList(i).FCount/oneuserjoin.FTotalUsercount*100)
    end if
    %>
    <tr bgcolor="#FFFFFF">
    	<td width="120"><%= oneuserjoin.FItemList(i).GetArea %> </td>
    	<td width="100" align="right"><%= FormatNumber(oneuserjoin.FItemList(i).FCount,0) %></td>
    	<td width="100" align="right"><%= tmppercent %> (%)</td>
    	<td><img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%= CInt(MAXBARSIZE * tmppercent / 100) %>"></td>
    </tr>
    <% next %>
</table>
<%
set oneuserjoin = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->