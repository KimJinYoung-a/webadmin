<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/event_reportcls.asp"-->
<%
response.write "사용중지"
dbget.close()	:	response.End
dim eventid,i

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, startdateStr, nextdateStr, page
dim rpttype,addstand
Dim egubun, oldlist

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
egubun = request("egubun")
oldlist = request("oldlist")

nowdateStr = CStr(now())


if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

startdateStr = yyyy1 + "-" + mm1 + "-" + dd1
nextdateStr = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

eventid = request("eventid")


dim oreport
set oreport = new CReportMaster
oreport.FRectEventid = eventid
oreport.FRectStart = startdateStr
oreport.FRectEnd =  nextdateStr
oreport.FRectGubun = egubun
oreport.FRectOldJumun = oldlist
oreport.SearchJointEventReport

dim totalprice,totalea
totalprice = 0
totalea = 0
%>

<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >6개월이전내역
		기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %><br>
		이벤트 번호 : <input type="text" name="eventid" size="30" value="<% = eventid %>">콤마(,)마로 구분해 주세요 마지막콤마 필요없음 ex)1234,5678,9101</td>
		<td class="a" align="right"><a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a></td>
	</tr>
	</form>
</table>
<br>
<table width="800" cellspacing="1" class="a" bgcolor="#3d3d3d">
    <tr bgcolor="#DDDDFF">
    	<td align="center">구매일</td>
    	<td align="center">아이템번호</td>
    	<td align="center">총구매가</td>
    	<td align="center">총구매갯수</td>
    	<td align="center">그래프</td>
    </tr>
<% if oreport.FResultCount < 0 then %>
<% else %>
<% for i=0 to oreport.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
    	<td width="90" align="center"><%= oreport.FMasterItemList(i).Fselldate %></td>
    	<td width="70" align="center"><%= oreport.FMasterItemList(i).Fitemid %></td>
    	<td width="70" align="right"><%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %></td>
    	<td width="70" align="right"><%= oreport.FMasterItemList(i).Fsellcnt %>개</td>
    	<td>
    		<img src="http://partner.10x10.co.kr/images/dot1.gif" height="4" width="<%= Clng(oreport.FMasterItemList(i).Fsellcnt/oreport.maxc*100) %>">
    	</td>
    </tr>
<%
totalprice = totalprice + oreport.FMasterItemList(i).Fselltotal
totalea = totalea + oreport.FMasterItemList(i).Fsellcnt
%>
<% next %>
<% end if %>
	<tr bgcolor="#FFFFFF">
		<td align="center">총계</td>
		<td></td>
		<td  align="right"><% = Formatnumber(totalprice,0) %></td>
		<td  align="right"><% = totalea %>개</td>
		<td></td>
	</tr>
</table>

<%
dim ototal
set ototal = new CReportTotal
ototal.FRectEventid = eventid
ototal.Fstartday = startdateStr
ototal.Fendday = nextdateStr
ototal.FRectOldJumun = oldlist
ototal.SearchJointEventReportTotal
%>
<table width="300" cellspacing="1" class="a" bgcolor="#3d3d3d">
    <tr bgcolor="#DDDDFF">
    	<td align="center">아이템번호</td>
    	<td align="center">총구매가</td>
    	<td align="center">총구매갯수</td>
    </tr>
<% if ototal.FResultCount < 0 then %>
<% else %>
<% for i=0 to ototal.FResultCount-1 %>
    <tr bgcolor="#FFFFFF">
    	<td width="100" align="center"><%= ototal.FMasterItemList(i).Fitem %></td>
    	<td width="100" align="right"><%= FormatNumber(ototal.FMasterItemList(i).FTotalPrice,0) %></td>
    	<td width="100" align="right"><%= ototal.FMasterItemList(i).FTotalEa %>개</td>
    </tr>
<% next %>
<% end if %>
</table>
<%
set oreport = Nothing
set ototal = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->