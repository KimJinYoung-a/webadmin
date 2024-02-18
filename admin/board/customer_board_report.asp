<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/board/lib/classes/customer_board_reportcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, startdateStr, nextdateStr
dim i

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

nowdateStr = CStr(now())


if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

startdateStr = yyyy1 + "-" + mm1 + "-" + dd1
nextdateStr = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)


dim oreport

set oreport = new CReportMaster
oreport.FRectStart = startdateStr
oreport.FRectEnd =  nextdateStr
oreport.SearchReport

dim flashvar
flashvar = "startdate=" + startdateStr + "&enddate=" + nextdateStr
%>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >±â°£ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %></td>
		<td class="a" align="right"><a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a></td>
	</tr>
	</form>
</table>
<br>
<!--
<table width="800" cellspacing="0" cellpadding="0" border="0" class="a">
<tr>
	<td align="center">
		<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="550" height="290">
		  <param name="movie" value="customer_chart.swf?link=<% = flashvar %>">
		  <param name="quality" value="high">
		  <embed src="customer_chart.swf?link=<% = flashvar %>" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="550" height="290"></embed>
		</object>
	</td>
</tr>
</table>
<br>
-->
<% if oreport.FResultCount >0 then %>
<table width="100%" cellspacing="1" class="a" bgcolor="#3d3d3d">

    <tr bgcolor="#DDDDFF">
<% for i=0 to oreport.FResultCount-1 %>
    	<td width="90" align="center"><%= oreport.FMasterItemList(i).GetQadivName %></td>
<% next %>
    </tr>

    <tr bgcolor="#FFFFFF">
<% for i=0 to oreport.FResultCount-1 %>
    	<td width="90" align="center"><%= oreport.FMasterItemList(i).Fcount %></td>
<% next %>
    </tr>
</table>
<% end if %>
<%
set oreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->