<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 회원탈퇴 현황
' History : 2008.02.15 한용민 개발
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/userjoin/userwithdrawcls.asp"-->
 
<%
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2,defaultdate1 ,i
dim fwdrawReason_total,fwdrawReason_01_total,fwdrawReason_02_total,fwdrawReason_03_total,fwdrawReason_04_total,fwdrawReason_05_total,fwdrawReason_06_total
dim fwdrawReason_07_total,fwithdrowtotalcount_total,fmancount_total,fgirlcount_total
	defaultdate1 = dateadd("d",-30,year(now) & "-" &month(now) & "-" & day(now))		'날짜값이 없을때 기본값으로 30이전까지 검색	
	yyyy1 = request("yyyy1")
	if yyyy1 = "" then yyyy1 = left(defaultdate1,4)
	mm1 = request("mm1")
	if mm1 = "" then mm1 = mid(defaultdate1,6,2)
	dd1 = request("dd1")
	if dd1 = "" then dd1 = right(defaultdate1,2)	
	yyyy2 = request("yyyy2")
	if yyyy2 = "" then yyyy2 = year(now)
	mm2 = request("mm2")
	if mm2 = "" then 
		mm2 = month(now)
	end if
	dd2 = request("dd2")
	if dd2 = "" then dd2 = day(now)
	menupos =  request("menupos")

dim ouserwithdrawlist		
	set ouserwithdrawlist = new cuserwithdrawlist
	ouserwithdrawlist.FRectStartdate = dateserial(yyyy1,mm1,dd1)
	ouserwithdrawlist.FRectEndDate = dateserial(yyyy2,mm2,dd2)
	ouserwithdrawlist.fuserwithdrawlist()

dim ouserwithdraw_sexgraph		'성별(그래프)용
	set ouserwithdraw_sexgraph = new cuserwithdrawlist
	ouserwithdraw_sexgraph.FRectStartdate = dateserial(yyyy1,mm1,dd1)
	ouserwithdraw_sexgraph.FRectEndDate = dateserial(yyyy2,mm2,dd2)
	ouserwithdraw_sexgraph.fuserwithdraw_sexgraph()

dim ouserwithdraw_areagraph		'탈퇴사유(그래프)용
	set ouserwithdraw_areagraph = new cuserwithdrawlist
	ouserwithdraw_areagraph.FRectStartdate = dateserial(yyyy1,mm1,dd1)
	ouserwithdraw_areagraph.FRectEndDate = dateserial(yyyy2,mm2,dd2)
	ouserwithdraw_areagraph.fuserwithdraw_areagraph()

'################################################################################################################
'그래프
dim sTotal1,sTotal2, strXML1, strXML2, strTemp1,strTemp2,frectwdrawReason

	strTemp1 =	"<?xml version='1.0' encoding='EUC-KR' ?>" &_
				"<chart chartBottomMargin='2' formatNumberScale='0' showLimits='0' divLineThickness='1' decimalPrecision='1' chartTopMargin='2' showShadow='1' canvasBorderColor='CBCBCB' animation='1' baseFontColor='666666' bgColor='FCFCFC' formatNumber='1' nameTBDistance='25' legendBorderColor='FFFFFF' canvasPadding='3' legendBgColor='FFFFFF' chartRightMargin='2' legendPadding='2' legendShadow='0' pieYScale='70' divLineIsDashed='1' showPercentValues='1' showBorder='0' pieSliceDepth='10' legendBorderThickness='0' placeValuesInside='1' chartLeftMargin='0' canvasBorderThickness='1' baseFontSize='11' divLineDashGap='3' setAdaptiveYMin='1' plotBorderAlpha='20' >"
	strXML1 = strTemp1
	
	for i=0 to ouserwithdraw_sexgraph.ftotalcount -1
		sTotal1 = sTotal1 + clng(ouserwithdraw_sexgraph.FItemList(i).fwdrawCount)
		strXML1 = strXML1 & "<set value='" & ouserwithdraw_sexgraph.FItemList(i).fwdrawCount & "' name='" &ouserwithdraw_sexgraph.FItemList(i).fwdrawSex & "' />"
	next
	strTemp1 = "<styles>" &_
			"<definition>" &_
			"<style name='shadow215' type='shadow' angle='215' distance='3'/>" &_
			"<style name='shadow45' type='shadow' angle='45' distance='3'/>" &_
			"</definition>" &_
			"<application>" &_
			"<apply toObject='DATAPLOTCOLUMN' styles='shadow215' />" &_
			"<apply toObject='DATAPLOTLINE' styles='shadow215' />" &_
			"<apply toObject='DATAPLOT' styles='shadow215' />" &_
			"</application>" &_
			"</styles>" &_
			"</chart>"
	strXML1 = strXML1 & strTemp1

	strTemp2 =	"<?xml version='1.0' encoding='EUC-KR' ?>" &_
				"<chart chartBottomMargin='2' formatNumberScale='0' showLimits='0' divLineThickness='1' decimalPrecision='1' chartTopMargin='2' showShadow='1' canvasBorderColor='CBCBCB' animation='1' baseFontColor='666666' bgColor='FCFCFC' formatNumber='1' nameTBDistance='25' legendBorderColor='FFFFFF' canvasPadding='3' legendBgColor='FFFFFF' chartRightMargin='2' legendPadding='2' legendShadow='0' pieYScale='70' divLineIsDashed='1' showPercentValues='1' showBorder='0' pieSliceDepth='10' legendBorderThickness='0' placeValuesInside='1' chartLeftMargin='0' canvasBorderThickness='1' baseFontSize='11' divLineDashGap='3' setAdaptiveYMin='1' plotBorderAlpha='20' >"
	strXML2 = strTemp2
	
	for i=0 to ouserwithdraw_areagraph.ftotalcount -1
		if ouserwithdraw_areagraph.FItemList(i).fwdrawReason = "01" then 
			frectwdrawReason = "상품품질불만"
		elseif ouserwithdraw_areagraph.FItemList(i).fwdrawReason = "02" then 
			frectwdrawReason = "이용빈도낮음"				
		elseif ouserwithdraw_areagraph.FItemList(i).fwdrawReason = "03" then 
			frectwdrawReason = "배송지연"	
		elseif ouserwithdraw_areagraph.FItemList(i).fwdrawReason = "04" then 
			frectwdrawReason = "개인정보유출"	
		elseif ouserwithdraw_areagraph.FItemList(i).fwdrawReason = "05" then 
			frectwdrawReason = "교환/환불/품질불만"	
		elseif ouserwithdraw_areagraph.FItemList(i).fwdrawReason = "06" then 
			frectwdrawReason = "기타"	
		elseif ouserwithdraw_areagraph.FItemList(i).fwdrawReason = "07" then 
			frectwdrawReason = "a/s불만"															
		else
			frectwdrawReason = "사유없음"		
		end if
		sTotal2 = sTotal2 + clng(ouserwithdraw_areagraph.FItemList(i).fwdrawCount)
		strXML2 = strXML2 & "<set value='" & ouserwithdraw_areagraph.FItemList(i).fwdrawCount & "' name='" & frectwdrawReason & "' />"
	next
	frectwdrawReason = ""
	strTemp2 = "<styles>" &_
			"<definition>" &_
			"<style name='shadow215' type='shadow' angle='215' distance='3'/>" &_
			"<style name='shadow45' type='shadow' angle='45' distance='3'/>" &_
			"</definition>" &_
			"<application>" &_
			"<apply toObject='DATAPLOTCOLUMN' styles='shadow215' />" &_
			"<apply toObject='DATAPLOTLINE' styles='shadow215' />" &_
			"<apply toObject='DATAPLOT' styles='shadow215' />" &_
			"</application>" &_
			"</styles>" &_
			"</chart>"
	strXML2 = strXML2 & strTemp2
'################################################################################################################	
%>

<script language="javascript" src="/lib/util/chart/FusionCharts.js"></script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form action="" name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		날짜 <% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>		
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">			
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>	
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">			
		</td>
		<td align="right">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<% if ouserwithdrawlist.ftotalcount > 0 then %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF" align="center">
		<td >성별 비율(%)</td>			
		<td >탈퇴사유 비율(%)</td>
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
	<table width="100%" border="0" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA" align="center">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td ></td>			
		<td colspan="3">총인원</td>
		<td colspan="8">탈퇴사유</td>		
	</tr>		
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>날짜</td>			
		<td >총</td>
		<td >남</td>
		<td >여</td>
		<td >상품품질불만</td>
		<td >이용빈도낮음</td>
		<td >배송지연</td>
		<td >개인정보유출</td>
		<td >교환/환불/품질불만</td>			      
		<td >기타</td>
		<td >A/S불만</td>
		<td >사유없음</td>
	</tr>
	<% for i = 0 to ouserwithdrawlist.ftotalcount -1 %>
	<tr bgcolor="#FFFFFF" align="center">
		<td ><%= ouserwithdrawlist.FItemList(i).fwdrawDate %></td>
		<td ><%= ouserwithdrawlist.FItemList(i).fwithdrowtotalcount %></td>
		<td ><%= ouserwithdrawlist.FItemList(i).fmancount %></td>
		<td ><%= ouserwithdrawlist.FItemList(i).fgirlcount %></td>
		<td ><%= ouserwithdrawlist.FItemList(i).fwdrawReason_01 %></td>
		<td ><%= ouserwithdrawlist.FItemList(i).fwdrawReason_02 %></td>			
		<td ><%= ouserwithdrawlist.FItemList(i).fwdrawReason_03 %></td>			
		<td ><%= ouserwithdrawlist.FItemList(i).fwdrawReason_04 %></td> 
		<td ><%= ouserwithdrawlist.FItemList(i).fwdrawReason_05 %></td>
		<td ><%= ouserwithdrawlist.FItemList(i).fwdrawReason_06 %></td>
		<td ><%= ouserwithdrawlist.FItemList(i).fwdrawReason_07 %></td>						
		<td ><%= ouserwithdrawlist.FItemList(i).fwdrawReason %></td>
	</tr>
	<%
	if ouserwithdrawlist.FItemList(i).fwdrawReason <>"" then fwdrawReason_total = fwdrawReason_total + clng(ouserwithdrawlist.FItemList(i).fwdrawReason)		
	if ouserwithdrawlist.FItemList(i).fwdrawReason_01 <>"" then fwdrawReason_01_total = fwdrawReason_01_total + clng(ouserwithdrawlist.FItemList(i).fwdrawReason_01)
	if ouserwithdrawlist.FItemList(i).fwdrawReason_02 <>"" then fwdrawReason_02_total = fwdrawReason_02_total + clng(ouserwithdrawlist.FItemList(i).fwdrawReason_02)
	if ouserwithdrawlist.FItemList(i).fwdrawReason_03 <>"" then fwdrawReason_03_total = fwdrawReason_03_total + clng(ouserwithdrawlist.FItemList(i).fwdrawReason_03)
	if ouserwithdrawlist.FItemList(i).fwdrawReason_04 <>"" then fwdrawReason_04_total = fwdrawReason_04_total + clng(ouserwithdrawlist.FItemList(i).fwdrawReason_04)
	if ouserwithdrawlist.FItemList(i).fwdrawReason_05 <>"" then fwdrawReason_05_total = fwdrawReason_05_total + clng(ouserwithdrawlist.FItemList(i).fwdrawReason_05)
	if ouserwithdrawlist.FItemList(i).fwdrawReason_06 <>"" then fwdrawReason_06_total = fwdrawReason_06_total + clng(ouserwithdrawlist.FItemList(i).fwdrawReason_06)
	if ouserwithdrawlist.FItemList(i).fwdrawReason_07 <>"" then fwdrawReason_07_total = fwdrawReason_07_total + clng(ouserwithdrawlist.FItemList(i).fwdrawReason_07)
	if ouserwithdrawlist.FItemList(i).fwithdrowtotalcount <>"" then fwithdrowtotalcount_total = fwithdrowtotalcount_total + clng(ouserwithdrawlist.FItemList(i).fwithdrowtotalcount)
	if ouserwithdrawlist.FItemList(i).fmancount <>"" then fmancount_total = fmancount_total + clng(ouserwithdrawlist.FItemList(i).fmancount)
	if ouserwithdrawlist.FItemList(i).fgirlcount <>"" then fgirlcount_total = fgirlcount_total + clng(ouserwithdrawlist.FItemList(i).fgirlcount)				
	next
	%>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td >합계</td>			
		<td ><%= fwithdrowtotalcount_total %></td>
		<td ><%= fmancount_total %></td>
		<td ><%= fgirlcount_total %></td>
		<td ><%= fwdrawReason_01_total %></td>
		<td ><%= fwdrawReason_02_total %></td>
		<td ><%= fwdrawReason_03_total %></td>
		<td ><%= fwdrawReason_04_total %></td>
		<td ><%= fwdrawReason_05_total %></td>			      
		<td ><%= fwdrawReason_06_total %></td>
		<td ><%= fwdrawReason_07_total %></td>
		<td ><%= fwdrawReason_total %></td>
	</tr>	
	</table>
<% else %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF">
		<td >검색 결과가 없습니다.</td>
	</tr>
	</table>
<% end if %>		

<%
set ouserwithdrawlist = nothing
set ouserwithdraw_sexgraph = nothing
set ouserwithdraw_areagraph = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->