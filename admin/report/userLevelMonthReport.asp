<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 회원등급별통계
' Hieditor : 2009.04.17 허진원 생성
'			 2016.07.20 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
dim yyyy, mm, i
	yyyy = requestcheckvar(Request("yyyy"),4)
	mm = requestcheckvar(Request("mm"),2)
dim tot_userlevelcount, tot_iOSexistscount, tot_ANDPushexistscount, tot_ANDALLY, tot_ANDALLN, tot_iOSALLY, tot_iOSALLN, tot_ANDPushY, tot_ANDPushN
dim tot_iOSPushY, tot_iOSPushN, tot_emailokY, tot_emailokN, tot_smsokY, tot_smsokN
dim tot_ANDPushYSmsY, tot_iOSPushYSmsY
	tot_userlevelcount=0
	tot_ANDPushexistscount=0
	tot_iOSexistscount=0
	tot_ANDALLY=0
	tot_ANDALLN=0
	tot_iOSALLY=0
	tot_iOSALLN=0
	tot_ANDPushY=0
	tot_ANDPushN=0
	tot_iOSPushY=0
	tot_iOSPushN=0
	tot_emailokY=0
	tot_emailokN=0
	tot_smsokY=0
	tot_smsokN=0
	tot_ANDPushYSmsY=0
	tot_iOSPushYSmsY=0

if yyyy="" then yyyy=year(date)
if mm="" then mm=Format00(2,month(date))

dim oreport
set oreport = new CUserLevelMonth
	oreport.FRectyyyymm = yyyy & "-" & mm
	oreport.GetLevelList

dim oagreeY
set oagreeY = new CUserLevelMonth
	oagreeY.FRectyyyymm = yyyy & "-" & mm
	oagreeY.GetLevelagreeList

dim oagreeNoMem
set oagreeNoMem = new CUserLevelMonth
	oagreeNoMem.FRectyyyymm = yyyy & "-" & mm
	oagreeNoMem.GetNonMemeberPushAgreeList

dim oHOLD
set oHOLD = new CUserLevelMonth
	oHOLD.GetUserHOLD_count

'각 비율 및 그래프 산출
dim sTotal, nTotal, strXML, strTemp, sAxisDate
sAxisDate = left(Date,7)

if oreport.FResultCount>0 then
	strTemp =	"<?xml version='1.0' encoding='EUC-KR' ?>" &_
				"<chart chartBottomMargin='2' formatNumberScale='0' showLimits='0' divLineThickness='1' decimalPrecision='1' chartTopMargin='2' showShadow='1' canvasBorderColor='CBCBCB' animation='1' lineThickness='3' baseFontColor='666666' bgColor='FFFFFF' formatNumber='1' legendBorderColor='FFFFFF' canvasPadding='3' legendBgColor='FFFFFF' chartRightMargin='2' legendPadding='2' legendShadow='0' divLineIsDashed='1' showBorder='0' legendBorderThickness='0' placeValuesInside='1' chartLeftMargin='0' canvasBorderThickness='1' baseFontSize='11' divLineDashGap='3' setAdaptiveYMin='1' anchorRadius='4' plotBorderAlpha='20' >"
	strXML = strTemp

	'날짜 카테고리
	strXML = strXML & "<categories>"
	for i=0 to oreport.FResultCount -1
		strXML = strXML & "<category name='" & oreport.FItemList(i).FAxisDate & "' showName='1' showLine='1' />"
		sAxisDate = oreport.FItemList(i).FAxisDate
	next
	strXML = strXML & "</categories>"

	'오렌지등급
	strXML = strXML & "<dataset seriesName='Orange' color='F8941D' showValues='0' >"
	for i=0 to oreport.FResultCount -1
		strXML = strXML & "<set value='" & oreport.FItemList(i).FOrange & "' />"
	next
	strXML = strXML & "</dataset>"
	'옐로우등급
	strXML = strXML & "<dataset seriesName='" & chkIIF(sAxisDate<"2018-08","Yellow","White") & "' color='FFCE00' showValues='0' >"
	for i=0 to oreport.FResultCount -1
		strXML = strXML & "<set value='" & oreport.FItemList(i).FYellow & "' />"
	next
	strXML = strXML & "</dataset>"
	'그린등급
	strXML = strXML & "<dataset seriesName='" & chkIIF(sAxisDate<"2018-08","Green","Red") & "' color='6EE111' showValues='0' >"
	for i=0 to oreport.FResultCount -1
		strXML = strXML & "<set value='" & oreport.FItemList(i).FGreen & "' />"
	next
	strXML = strXML & "</dataset>"
	'블루등급
	strXML = strXML & "<dataset seriesName='" & chkIIF(sAxisDate<"2018-08","Blue","VIP") & "' color='0093FF' showValues='0' >"
	for i=0 to oreport.FResultCount -1
		strXML = strXML & "<set value='" & oreport.FItemList(i).FBlue & "' />"
	next
	strXML = strXML & "</dataset>"
	'VIP Silver등급
	strXML = strXML & "<dataset seriesName='" & chkIIF(sAxisDate<"2018-08","Silver","Gold") & "' color='FF0175' showValues='0' >"
	for i=0 to oreport.FResultCount -1
		strXML = strXML & "<set value='" & oreport.FItemList(i).FSilver & "' />"
	next
	strXML = strXML & "</dataset>"
	'VIP Gold등급
	strXML = strXML & "<dataset seriesName='" & chkIIF(sAxisDate<"2018-08","Gold","VVIP") & "' color='E35D86' showValues='0' >"
	for i=0 to oreport.FResultCount -1
		strXML = strXML & "<set value='" & oreport.FItemList(i).FGold & "' />"
	next
	strXML = strXML & "</dataset>"
	'VVIP 등급
	strXML = strXML & "<dataset seriesName='VVIP' color='red' showValues='0' >"
	for i=0 to oreport.FResultCount -1
		strXML = strXML & "<set value='" & oreport.FItemList(i).fVVIP & "' />"
	next
	strXML = strXML & "</dataset>"
	'스탭등급
	strXML = strXML & "<dataset seriesName='Staff' color='A80000' showValues='0' >"
	for i=0 to oreport.FResultCount -1
		strXML = strXML & "<set value='" & oreport.FItemList(i).FStaff & "' />"
	next
	'FAMILY등급
	strXML = strXML & "<dataset seriesName='FAMILY' color='A60000' showValues='0' >"
	for i=0 to oreport.FResultCount -1
		strXML = strXML & "<set value='" & oreport.FItemList(i).fFAMILY & "' />"
	next
	'BIZ등급
	strXML = strXML & "<dataset seriesName='BIZ' color='A70000' showValues='0' >"
	for i=0 to oreport.FResultCount -1
		strXML = strXML & "<set value='" & oreport.FItemList(i).fBIZ & "' />"
	next
	strXML = strXML & "</dataset>"

	strTemp = "<trendLines></trendLines>" &_
				"<styles>" &_
				"	<definition>" &_
				"		<style name='shadow215' type='shadow' angle='215' distance='3'/>" &_
				"		<style name='shadow45' type='shadow' angle='45' distance='3'/>" &_
				"	</definition>" &_
				"	<application>" &_
				"		<apply toObject='DATAPLOTCOLUMN' styles='shadow215' />" &_
				"		<apply toObject='DATAPLOTLINE' styles='shadow215' />" &_
				"		<apply toObject='DATAPLOT' styles='shadow215' />" &_
				"	</application>" &_
				"</styles>" &_
				"</chart>"
	strXML = strXML & strTemp
end if
%>

<script type="text/javascript">

function exceldown(){
	frm.target="view";
	frm.action="/admin/report/userLevelMonthReport_excel.asp";
	frm.submit();
	frm.target="";
	frm.action="";
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 기간 : <% Call DrawYMBoxdynamic("yyyy", yyyy, "mm", mm, "") %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<iframe id="view" name="view" src="" width="0" height="0" frameborder="0" scrolling="no"></iframe>
<Br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<font color="red">부하가 있는 매뉴 입니다. 여러번 누르지 마시고 기다려주세요.</font>
	</td>
	<td align="right">	
		<input type="button" value="엑셀다운" onClick="exceldown();" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="18">
		검색결과 : <b><%= oagreeY.FresultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td rowspan="3">회원등급</td>
	<td rowspan="3">고객수</td>
	<td colspan="5">iOS</td>
	<td colspan="5">Android</td>
	<td colspan="2">SMS</td>
	<td colspan="2">E-mail</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td rowspan="2">보유수</td>
	<td colspan="2">PUSH(Y)</td>
	<td colspan="2">PUSH(N)</td>
	<td rowspan="2">보유수</td>
	<td colspan="2">PUSH(Y)</td>
	<td colspan="2">PUSH(N)</td>
	<td rowspan="2">SMS(Y)</td>
	<td rowspan="2">SMS(N)</td>
	<td rowspan="2">E-mail(Y)</td>
	<td rowspan="2">E-mail(N)</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>SMS(Y)</td>
	<td>SMS(N)</td>
	<td>SMS(Y)</td>
	<td>SMS(N)</td>
	<td>SMS(Y)</td>
	<td>SMS(N)</td>
	<td>SMS(Y)</td>
	<td>SMS(N)</td>
</tr>
<% if oagreeY.FresultCount>0 then %>
	<% for i=0 to oagreeY.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oagreeY.FItemList(i).fuserlevelname %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fuserlevelcount,0) %></td>

		<td><%= FormatNumber(oagreeY.FItemList(i).fiOSexistscount,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fiOSPushYSmsY,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fiOSPushY-oagreeY.FItemList(i).fiOSPushYSmsY,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fiOSPushN-oagreeY.FItemList(i).fiOSALLN,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fiOSALLN,0) %></td>

		<td><%= FormatNumber(oagreeY.FItemList(i).fANDPushexistscount,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fANDPushYSmsY,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fANDPushY-oagreeY.FItemList(i).fANDPushYSmsY,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fANDPushN-oagreeY.FItemList(i).fANDALLN,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fANDALLN,0) %></td>

		<td><%= FormatNumber(oagreeY.FItemList(i).fsmsokY,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).fsmsokN,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).femailokY,0) %></td>
		<td><%= FormatNumber(oagreeY.FItemList(i).femailokN,0) %></td>
	</tr>
	<%
	tot_userlevelcount = tot_userlevelcount + oagreeY.FItemList(i).fuserlevelcount
	tot_ANDPushexistscount = tot_ANDPushexistscount + oagreeY.FItemList(i).fANDPushexistscount
	tot_iOSexistscount = tot_iOSexistscount + oagreeY.FItemList(i).fiOSexistscount
	tot_ANDALLY = tot_ANDALLY + oagreeY.FItemList(i).fANDALLY
	tot_ANDALLN = tot_ANDALLN + oagreeY.FItemList(i).fANDALLN
	tot_iOSALLY = tot_iOSALLY + oagreeY.FItemList(i).fiOSALLY
	tot_iOSALLN = tot_iOSALLN + oagreeY.FItemList(i).fiOSALLN
	tot_ANDPushY = tot_ANDPushY + oagreeY.FItemList(i).fANDPushY
	tot_ANDPushN = tot_ANDPushN + oagreeY.FItemList(i).fANDPushN
	tot_iOSPushY = tot_iOSPushY + oagreeY.FItemList(i).fiOSPushY
	tot_iOSPushN = tot_iOSPushN + oagreeY.FItemList(i).fiOSPushN
	tot_emailokY = tot_emailokY + oagreeY.FItemList(i).femailokY
	tot_emailokN = tot_emailokN + oagreeY.FItemList(i).femailokN
	tot_smsokY = tot_smsokY + oagreeY.FItemList(i).fsmsokY
	tot_smsokN = tot_smsokN + oagreeY.FItemList(i).fsmsokN
	tot_ANDPushYSmsY = tot_ANDPushYSmsY + oagreeY.FItemList(i).fANDPushYSmsY
	tot_iOSPushYSmsY = tot_iOSPushYSmsY + oagreeY.FItemList(i).fiOSPushYSmsY
	next
	%>

	<tr align="center" bgcolor="#FFFFFF">
		<td>합계</td>
		<td><%= FormatNumber(tot_userlevelcount,0) %></td>

		<td><%= FormatNumber(tot_iOSexistscount,0) %></td>
		<td><%= FormatNumber(tot_iOSPushYSmsY,0) %></td>
		<td><%= FormatNumber(tot_iOSPushY-tot_iOSPushYSmsY,0) %></td>
		<td><%= FormatNumber(tot_iOSPushN-tot_iOSALLN,0) %></td>
		<td><%= FormatNumber(tot_iOSALLN,0) %></td>

		<td><%= FormatNumber(tot_ANDPushexistscount,0) %></td>
		<td><%= FormatNumber(tot_ANDPushYSmsY,0) %></td>
		<td><%= FormatNumber(tot_ANDPushY-tot_ANDPushYSmsY,0) %></td>
		<td><%= FormatNumber(tot_ANDPushN-tot_ANDALLN,0) %></td>
		<td><%= FormatNumber(tot_ANDALLN,0) %></td>

		<td><%= FormatNumber(tot_smsokY,0) %></td>
		<td><%= FormatNumber(tot_smsokN,0) %></td>
		<td><%= FormatNumber(tot_emailokY,0) %></td>
		<td><%= FormatNumber(tot_emailokN,0) %></td>
	</tr>

<% if oagreeNoMem.FresultCount>0 then %>
	<% for i=0 to oagreeNoMem.FresultCount-1 %>
	<tr align="center" bgcolor="#F0F0F0">
		<td>비회원 푸시</td>
		<td>-</td>
		<td><%= FormatNumber(oagreeNoMem.FItemList(i).fiOSexistscount,0) %></td>
		<td>-</td>
		<td><%= FormatNumber(oagreeNoMem.FItemList(i).fiOSPushY,0) %></td>
		<td>-</td>
		<td><%= FormatNumber(oagreeNoMem.FItemList(i).fiOSPushN,0) %></td>
		<td><%= FormatNumber(oagreeNoMem.FItemList(i).fANDPushexistscount,0) %></td>
		<td>-</td>
		<td><%= FormatNumber(oagreeNoMem.FItemList(i).fANDPushY,0) %></td>
		<td>-</td>
		<td><%= FormatNumber(oagreeNoMem.FItemList(i).fANDPushN,0) %></td>
		<td>-</td>
		<td>-</td>
		<td>-</td>
		<td>-</td>
	</tr>
<%
		next
	end if
%>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>휴면계정 합계</td>
		<td>
			<% if oHOLD.ftotalcount>0 then %>
				<%= FormatNumber(oHOLD.FOneItem.fUserHOLD_count,0) %>
			<% else %>
				0
			<% end if %>
		</td>
		<td colspan=14></td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="14" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<% 'if false then %>
<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="12">
		검색결과 : <b><%= oreport.FresultCount %></b>
		&nbsp;&nbsp;&nbsp;&nbsp; ※ <%= year(dateadd("m",-1,dateserial(yyyy,mm,"01"))) %>년 <%= month(dateadd("m",-1,dateserial(yyyy,mm,"01"))) %>월말 기준
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">년/월<br />(개편이전)</td>
	<td width="90"><br />(Orange)</td>
	<td width="90">White<br />(Yellow)</td>
	<td width="90">Red<br />(Green)</td>
	<td width="90">Blue<br />(VIP)</td>
	<td width="90">VIP Gold<br />(VIP Silver)</td>
	<td width="90">VVIP<br />(VIP Gold)</td>
	<td width="90"><br />(VVIP)</td>
	<td width="90">Staff</td>
	<td width="90">FAMILY</td>
	<td width="90">BIZ</td>
	<td width="93" bgcolor="#E0E0E0">소계</td>
</tr>

<% if oreport.FresultCount>0 then %>
	<% for i=0 to oreport.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%=oreport.FItemList(i).FAxisDate%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FOrange,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FYellow,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FGreen,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FBlue,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FSilver,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FGold,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FVVIP,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FStaff,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).fFAMILY,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).fBIZ,0)%></td>
		<td bgcolor="#FAFAFA"><%=FormatNumber(oreport.FItemList(i).FTotal,0)%></td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="12" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% 'end if %>
</table>

<% if false then %>
<script language="javascript" src="/lib/util/chart/FusionCharts.js"></script>
<table width="823" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr>
	<td align="center" style="padding-top:10px;">
		<div id="chartdiv" align="center"></div>
		<script type="text/javascript">	
			var chart = new FusionCharts("/lib/util/chart/StackedColumn2D.swf", "chartdiv", "700", "400", "0", "0");
			chart.setDataXML("<%=strXML%>");
			chart.render("chartdiv");
		</script>
	</td>
</tr>
</table>
<% end if %>
<% end if %>

<%
set oreport = nothing
set oagreeY = nothing
set oagreeNoMem = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->