<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/resending_reportcls.asp"-->
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
oreport.SearchObaesongReport


dim totalcount
totalcount =0
%>
<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		</td>
		<td class="a" align="right"><a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a></td>
	</tr>
	</form>
</table>
<br>
<table cellspacing="1" cellpadding="3" class="a" bgcolor="#3d3d3d">
    <tr bgcolor="#DDDDFF">
        <td width="100">사유1</td>
        <td width="100">사유2</td>
        <td >건수</td>
    </tr>
    <% if oreport.FResultCount < 0 then %>
    <tr bgcolor="#FFFFFF">
        <td colspan="3" align="center">[검색결과가 없습니다.]</td>
    </tr>
    <% else %>
    <% for i=0 to oreport.FResultCount-1 %>
    <% totalcount = totalcount + oreport.FMasterItemList(i).Fcount %>
    <tr bgcolor="#FFFFFF">
        <td ><%= oreport.FMasterItemList(i).Fgubun01Name %></td>
        <td ><%= oreport.FMasterItemList(i).Fgubun02Name %></td>
        <td align="right"><%= oreport.FMasterItemList(i).Fcount %></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
        <td >Total</td>
        <td ></td>
        <td align="right"><%= totalcount %></td>
    </tr>
    <% end if %>
</table>
<%
set oreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->