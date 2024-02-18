<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 회원로그인 현황
' History : 2008.02.05 한용민 생성
'			2018.07.25 정태훈 수정 (회원등급 개편 적용)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/userjoin/userlogincls.asp"-->
<%
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2,defaultdate1 ,i ,ouserloginlist,ouserloginlist_date, datetime,loginSex ,loginlevel
dim datetimecount ,datetimeMalecount,datetimefemalecount, yellowcount,greencount,bluecount,silvercount,goldcount,orangecount,staffcount, vvipcount
dim FAMILYcount, BIZcount
	defaultdate1 = dateadd("d",-7,year(now) & "-" &month(now) & "-" & day(now))		'날짜값이 없을때 기본값으로 7이전까지 검색	
	yyyy1 = requestcheckvar(getNumeric(request("yyyy1")),4)
	mm1 = requestcheckvar(getNumeric(request("mm1")),2)
	dd1 = requestcheckvar(getNumeric(request("dd1")),2)
	yyyy2 = requestcheckvar(getNumeric(request("yyyy2")),4)
	mm2 = requestcheckvar(getNumeric(request("mm2")),2)
	dd2 = requestcheckvar(getNumeric(request("dd2")),2)
	datetime = requestcheckvar(request("datetime"),16)
	loginSex = requestcheckvar(request("loginSex"),2)
	menupos = requestcheckvar(getNumeric(request("menupos")),10)

yellowcount=0
greencount=0
bluecount=0
silvercount=0
goldcount=0
orangecount=0
staffcount=0
vvipcount=0
datetimecount = 0
datetimeMalecount = 0
datetimefemalecount = 0
FAMILYcount = 0
BIZcount = 0

if yyyy1 = "" then yyyy1 = left(defaultdate1,4)
if mm1 = "" then mm1 = mid(defaultdate1,6,2)
if dd1 = "" then dd1 = right(defaultdate1,2)
if yyyy2 = "" then yyyy2 = year(now)
if mm2 = "" then mm2 = month(now)
if dd2 = "" then dd2 = day(now)
if datetime = "" then datetime = "date"

' 일별
if datetime = "date" then
	set ouserloginlist_date = new cuserloginlist
		ouserloginlist_date.frectdatetime = datetime
		ouserloginlist_date.frectloginSex = loginSex
		ouserloginlist_date.FRectStartdate = dateserial(yyyy1,mm1,dd1)
		ouserloginlist_date.FRectEndDate = dateserial(yyyy2,mm2,dd2)
		ouserloginlist_date.fuserloginlist_date()

' 월별
elseif datetime="month" then
	set ouserloginlist_date = new cuserloginlist
		ouserloginlist_date.frectdatetime = datetime
		ouserloginlist_date.frectloginSex = loginSex
		ouserloginlist_date.FRectStartdate = yyyy1&"-"&mm1
		ouserloginlist_date.FRectEndDate = year(dateadd("m",+1,yyyy2&"-"&mm2)) & "-" & Format00(2,month(dateadd("m",+1,yyyy2&"-"&mm2)))
		ouserloginlist_date.fuserloginlist_monthly()

' 시간별
else
	set ouserloginlist = new cuserloginlist
		ouserloginlist.frectdatetime = datetime
		ouserloginlist.frectloginSex = loginSex
		ouserloginlist.FRectStartdate = dateserial(yyyy1,mm1,dd1)
		ouserloginlist.FRectEndDate = dateserial(yyyy2,mm2,dd2)
		ouserloginlist.fuserloginlist()
end if

dim ouserloginlist_graph
set ouserloginlist_graph = new cuserloginlist
	ouserloginlist_graph.FRectStartdate = dateserial(yyyy1,mm1,dd1)
	ouserloginlist_graph.FRectEndDate = dateserial(yyyy2,mm2,dd2)
	'ouserloginlist_graph.fuserloginlist_graph()

dim ouserloginlist_graph2
set ouserloginlist_graph2 = new cuserloginlist
	ouserloginlist_graph2.FRectStartdate = dateserial(yyyy1,mm1,dd1)
	ouserloginlist_graph2.FRectEndDate = dateserial(yyyy2,mm2,dd2)
	'ouserloginlist_graph2.fuserloginlist_graph2()
	
'그래프
dim sTotal1,sTotal2, strXML1, strXML2, strTemp1,strTemp2
' strTemp1 =	"<?xml version='1.0' encoding='EUC-KR' ?>" &_
' 			"<chart chartBottomMargin='2' formatNumberScale='0' showLimits='0' divLineThickness='1' decimalPrecision='1' chartTopMargin='2' showShadow='1' canvasBorderColor='CBCBCB' animation='1' baseFontColor='666666' bgColor='FCFCFC' formatNumber='1' nameTBDistance='25' legendBorderColor='FFFFFF' canvasPadding='3' legendBgColor='FFFFFF' chartRightMargin='2' legendPadding='2' legendShadow='0' pieYScale='70' divLineIsDashed='1' showPercentValues='1' showBorder='0' pieSliceDepth='10' legendBorderThickness='0' placeValuesInside='1' chartLeftMargin='0' canvasBorderThickness='1' baseFontSize='11' divLineDashGap='3' setAdaptiveYMin='1' plotBorderAlpha='20' >"
' strXML1 = strTemp1

' for i=0 to ouserloginlist_graph.ftotalcount -1
' 	sTotal1 = sTotal1 + clng(ouserloginlist_graph.FItemList(i).floginDate_count)
' 	strXML1 = strXML1 & "<set value='" & ouserloginlist_graph.FItemList(i).floginDate_count & "' name='" &ouserloginlist_graph.FItemList(i).floginage & "' />"
' next
' strTemp1 = "<styles>" &_
' 		"<definition>" &_
' 		"<style name='shadow215' type='shadow' angle='215' distance='3'/>" &_
' 		"<style name='shadow45' type='shadow' angle='45' distance='3'/>" &_
' 		"</definition>" &_
' 		"<application>" &_
' 		"<apply toObject='DATAPLOTCOLUMN' styles='shadow215' />" &_
' 		"<apply toObject='DATAPLOTLINE' styles='shadow215' />" &_
' 		"<apply toObject='DATAPLOT' styles='shadow215' />" &_
' 		"</application>" &_
' 		"</styles>" &_
' 		"</chart>"
' strXML1 = strXML1 & strTemp1

' strTemp2 =	"<?xml version='1.0' encoding='EUC-KR' ?>" &_
' 			"<chart chartBottomMargin='2' formatNumberScale='0' showLimits='0' divLineThickness='1' decimalPrecision='1' chartTopMargin='2' showShadow='1' canvasBorderColor='CBCBCB' animation='1' baseFontColor='666666' bgColor='FCFCFC' formatNumber='1' nameTBDistance='25' legendBorderColor='FFFFFF' canvasPadding='3' legendBgColor='FFFFFF' chartRightMargin='2' legendPadding='2' legendShadow='0' pieYScale='70' divLineIsDashed='1' showPercentValues='1' showBorder='0' pieSliceDepth='10' legendBorderThickness='0' placeValuesInside='1' chartLeftMargin='0' canvasBorderThickness='1' baseFontSize='11' divLineDashGap='3' setAdaptiveYMin='1' plotBorderAlpha='20' >"
' strXML2 = strTemp2

' for i=0 to ouserloginlist_graph2.ftotalcount -1
' 	sTotal2 = sTotal2 + clng(ouserloginlist_graph2.FItemList(i).floginDate_count)
' 	strXML2 = strXML2 & "<set value='" & ouserloginlist_graph2.FItemList(i).floginDate_count & "' name='" &ouserloginlist_graph2.FItemList(i).floginarea & "' />"
' next
' strTemp2 = "<styles>" &_
' 		"<definition>" &_
' 		"<style name='shadow215' type='shadow' angle='215' distance='3'/>" &_
' 		"<style name='shadow45' type='shadow' angle='45' distance='3'/>" &_
' 		"</definition>" &_
' 		"<application>" &_
' 		"<apply toObject='DATAPLOTCOLUMN' styles='shadow215' />" &_
' 		"<apply toObject='DATAPLOTLINE' styles='shadow215' />" &_
' 		"<apply toObject='DATAPLOT' styles='shadow215' />" &_
' 		"</application>" &_
' 		"</styles>" &_
' 		"</chart>"
' strXML2 = strXML2 & strTemp2

%>

<script language="javascript" src="/lib/util/chart/FusionCharts.js"></script>
<script type="text/javascript">

function loginsex_submit(){
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<form action="" name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 날짜 :
		<% if datetime="month" then %>
			<% DrawYMBoxdynamic "yyyy1", yyyy1, "mm1", mm1, "" %> - <% DrawYMBoxdynamic "yyyy2", yyyy2, "mm2", mm2, "" %>
		<% else %>
			<% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		<% end if %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 그룹 :
		<input type="radio" name="datetime" value="month" <% if datetime="month" then response.write "checked" %>>월별
		<input type="radio" name="datetime" value="date" <% if datetime="date" then response.write "checked" %>>일별
		<input type="radio" name="datetime" value="time" <% if datetime="time" then response.write "checked" %>>시간별
		
	</td>
</tr>
</table>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<font color="red">- 해당 조회 기간중 처음 방문한 최초 1회만 카운트되며 이후의 방문은 인식하지 않습니다.<br> 
		&nbsp;&nbsp;예를 들어 오전에 1번 방문하고 오후에 1번 방문하여도 중복으로 제거되므로 일 실방문자 수는 1이 됩니다.</font>	
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<% if false and ouserloginlist_graph.ftotalcount > 0 then %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td >나이 비율(%)</td>			
		<td >지역 비율(%)</td>
	</tr>
	<tr bgcolor="#FFFFFF" align="center">
		<td >
			<div id="chartdiv1" align="center"></div>
			<script type="text/javascript">	
				var chart = new FusionCharts("/lib/util/chart/Pie3D.swf", "chartdiv1", "320", "200", "0", "0");
				chart.setDataXML("<%=strXML1%>");
				chart.render("chartdiv1");
			</script>
		</td>	
		<td >
			<div id="chartdiv2" align="center"></div>
			<script type="text/javascript">	
				var chart = new FusionCharts("/lib/util/chart/Pie3D.swf", "chartdiv2", "320", "200", "0", "0");
				chart.setDataXML("<%=strXML2%>");
				chart.render("chartdiv2");
			</script>
		</td>			
	</tr>			
	</table>
	<br>
<% end if %>

<% if datetime="date" or datetime="month" then %>
	<% if ouserloginlist_date.ftotalcount > 0 then %>			
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">		
			<td align="left" colspan="15">
				<input type="checkbox" name="loginSex" value="on" onclick="loginsex_submit();" <% if loginSex="on" then response.write "checked" %>>성별표시 
			</td>
		</tr>	
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td >시간</td>
			<td >접속자</td>
			<td >WHITE<br />(Yellow)</td>
			<td >RED<br />(Green)</td>
			
			<td ><br />(Blue)</td>
			<td >VIP<br />(VIP Silver)</td>
			<td >VIP gold</td>
			<td >VVIP</td>
			<td ><br />(Orange)</td>
			<td >STAFF</td>
			<td >FAMILY</td>
			<td >BIZ</td>
			<% if loginSex = "on" then %>
				<td >남성</td>
				<td >여성</td>				
			<% end if %>			   
		</tr>
		<% for i = 0 to ouserloginlist_date.ftotalcount -1 %>
		<tr bgcolor="#FFFFFF" align="center">
			<td ><%= ouserloginlist_date.FItemList(i).floginDate %></td>
			<td ><%= formatNumber(ouserloginlist_date.FItemList(i).floginDate_count,0) %></td>
			<td ><%= formatNumber(chkIIF(cLng(ouserloginlist_date.FItemList(i).fWhite)>0,ouserloginlist_date.FItemList(i).fWhite,ouserloginlist_date.FItemList(i).fyellow),0) %></td>
			<td ><%= formatNumber(chkIIF(cLng(ouserloginlist_date.FItemList(i).fRed)>0,ouserloginlist_date.FItemList(i).fRed,ouserloginlist_date.FItemList(i).fgreen),0) %></td>
			
			<td ><%= formatNumber(ouserloginlist_date.FItemList(i).fblue,0) %></td>
			<td ><%= formatNumber(chkIIF(cLng(ouserloginlist_date.FItemList(i).fVIP)>0,ouserloginlist_date.FItemList(i).fVIP,ouserloginlist_date.FItemList(i).fsilver),0) %></td>
			<td ><%= formatNumber(ouserloginlist_date.FItemList(i).fgold,0) %></td>
			<td ><%= formatNumber(ouserloginlist_date.FItemList(i).fvvip,0) %></td>
			<td ><%= formatNumber(ouserloginlist_date.FItemList(i).forange,0) %></td>
			<td ><%= formatNumber(ouserloginlist_date.FItemList(i).fstaff,0) %></td>			
			<td ><%= formatNumber(ouserloginlist_date.FItemList(i).fFAMILY,0) %></td>
			<td ><%= formatNumber(ouserloginlist_date.FItemList(i).fBIZ,0) %></td>
			<% if loginSex = "on" then %>
				<td ><%= formatNumber(ouserloginlist_date.FItemList(i).fMaleCnt,0) %></td>
				<td ><%= formatNumber(ouserloginlist_date.FItemList(i).fFemaleCnt,0) %></td>
			<% end if %>			
			<% 
			datetimecount = datetimecount + clng(ouserloginlist_date.FItemList(i).floginDate_count)
			datetimeMalecount = datetimeMalecount + clng(ouserloginlist_date.FItemList(i).fMaleCnt)
			datetimefemalecount = datetimefemalecount + clng(ouserloginlist_date.FItemList(i).fFemaleCnt)
			yellowcount = yellowcount + clng(ouserloginlist_date.FItemList(i).fyellow) + clng(ouserloginlist_date.FItemList(i).fWhite)
			greencount = greencount + clng(ouserloginlist_date.FItemList(i).fgreen) + clng(ouserloginlist_date.FItemList(i).fRed)
			bluecount = bluecount + clng(ouserloginlist_date.FItemList(i).fblue)
			silvercount = silvercount + clng(ouserloginlist_date.FItemList(i).fsilver) + clng(ouserloginlist_date.FItemList(i).fVIP)
			goldcount = goldcount + clng(ouserloginlist_date.FItemList(i).fgold)
			vvipcount = vvipcount + clng(ouserloginlist_date.FItemList(i).fvvip)
			orangecount = orangecount + clng(ouserloginlist_date.FItemList(i).forange)
			staffcount = staffcount + clng(ouserloginlist_date.FItemList(i).fstaff)	
			FAMILYcount = FAMILYcount + clng(ouserloginlist_date.FItemList(i).fFAMILY)
			BIZcount = BIZcount + clng(ouserloginlist_date.FItemList(i).fBIZ)
			%>   
		</tr>	
		<% next %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td >접속자 합계</td>
			<td ><%= formatNumber(datetimecount,0) %></td>
			<td ><%= formatNumber(yellowcount,0) %></td>
			<td ><%= formatNumber(greencount,0) %></td>
			<td ><%= formatNumber(bluecount,0) %></td>
			<td ><%= formatNumber(silvercount,0) %></td>
			<td ><%= formatNumber(goldcount,0) %></td>
			<td ><%= formatNumber(vvipcount,0) %></td>
			<td ><%= formatNumber(orangecount,0) %></td> 
			<td ><%= formatNumber(staffcount,0) %></td>			
			<td ><%= formatNumber(FAMILYcount,0) %></td>
			<td ><%= formatNumber(BIZcount,0) %></td>
			<% if loginSex = "on" then %>
				<td ><%= formatNumber(datetimeMalecount,0) %></td>
				<td ><%= formatNumber(datetimefemalecount,0) %></td>		
			<% end if %>	
		</tr>			
		</table>
	<% else %>
		<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
		<tr align="center" bgcolor="#FFFFFF">
			<td >검색 결과가 없습니다.</td>
		</tr>
		</table>
	<% end if %>
	
<% else %>

	<% if ouserloginlist.ftotalcount > 0 then %>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">		
			<td align="left" colspan="4">&nbsp;	
				<input type="checkbox" name="loginSex" value="on" onclick="loginsex_submit();" <% if loginSex="on" then response.write "checked" %>>성별표시 
			</td>
		</tr>	
		<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
			<td>시간</td>
		   <td>접속자</td>
			<% if loginSex = "on" then %>
				<td>남성</td>
				<td>여성</td>				
			<% end if %>		   
		</tr>
		<% for i = 0 to ouserloginlist.ftotalcount -1 %>
		<tr bgcolor="#FFFFFF" align="center">
			<td ><%= ouserloginlist.FItemList(i).floginDate&"시" %></td>
			<td ><%= ouserloginlist.FItemList(i).floginDate_count %></td>
			<% if loginSex = "on" then %>
				<td ><%= ouserloginlist.FItemList(i).fMaleCnt %></td>
				<td ><%= ouserloginlist.FItemList(i).fFemaleCnt %></td>
			<% end if %>				   
		</tr>
		<%
		if ouserloginlist.FItemList(i).floginDate_count <> "" then 
			datetimecount = datetimecount + cint(ouserloginlist.FItemList(i).floginDate_count)
		end if
		if ouserloginlist.FItemList(i).fMaleCnt <> "" then
			datetimeMalecount = datetimeMalecount + cint(ouserloginlist.FItemList(i).fMaleCnt)
		end if
		if ouserloginlist.FItemList(i).fFemaleCnt <> "" then
			datetimefemalecount = datetimefemalecount + cint(ouserloginlist.FItemList(i).fFemaleCnt)			
		end if
		
		next 
		%>
		<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
			<td >접속자 합계</td>
			<td  colspan="<% if loginsex="" then response.write "2" %>"><%= datetimecount %></td> 
			<% if loginSex = "on" then %>
				<td ><%= datetimeMalecount %></td>
				<td ><%= datetimefemalecount %></td>		
			<% end if %>			
		</tr>			
		</table>
	<% else %>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	    <tr bgcolor="#FFFFFF" align="center">
	    	<td>검색 결과가 없습니다.</td>
	    </tr>
		</table>
	<% end if %>
<% end if %>

</form>

<%
if datetime="date" or datetime="month" then
	set ouserloginlist_date = nothing
else	
	set ouserloginlist = nothing
end if

set ouserloginlist_graph = nothing
set ouserloginlist_graph2 = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->