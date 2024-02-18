<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 회원등급별 매출
' Hieditor : 2009.04.17 허진원 생성
'			 2016.07.20 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
dim sDt, eDt, tNo, tDiv, chkOld, isBanpum, makerid
	sDt = Request("startDate")
	eDt = Request("endDate")
	tNo = Request("tNo")
	tDiv = Request("tDiv")
	chkOld = Request("chkOld")
	isBanpum = Request("isBanpum")
	makerid = Request("makerid")

'기본 1주일 세팅
if sDt="" then sDt=dateAdd("ww",-1,date())
if eDt="" then eDt=date()
if tNo="" then tNo="1"
if tDiv="" then tDiv="week"
if chkOld="" then chkOld="N"
if isBanpum="" then isBanpum="all"

'// 내용 접수
dim oreport
set oreport = new CUserLevelSell
	oreport.FRectSdate = sDt
	oreport.FRectEdate = eDt
	oreport.FRectOld = chkOld
	oreport.FRectMinusInc = isBanpum
	oreport.FRectMakerid = makerid
	if makerid<>"" then
		oreport.GetLevelListWithDetail
	else
		oreport.GetLevelList
	end if
	

'각 비율 및 그래프 산출
dim sTotal, nTotal, strXML1, strXML2, strTemp, i, uTotal

if oreport.FResultCount>0 then
	strTemp =	"<?xml version='1.0' encoding='EUC-KR' ?>" &_
				"<chart chartBottomMargin='2' formatNumberScale='0' showLimits='0' divLineThickness='1' decimalPrecision='1' chartTopMargin='2' showShadow='1' canvasBorderColor='CBCBCB' animation='1' baseFontColor='666666' bgColor='FCFCFC' formatNumber='1' nameTBDistance='25' legendBorderColor='FFFFFF' canvasPadding='3' legendBgColor='FFFFFF' chartRightMargin='2' legendPadding='2' legendShadow='0' pieYScale='70' divLineIsDashed='1' showPercentValues='1' showBorder='0' pieSliceDepth='10' legendBorderThickness='0' placeValuesInside='1' chartLeftMargin='0' canvasBorderThickness='1' baseFontSize='11' divLineDashGap='3' setAdaptiveYMin='1' plotBorderAlpha='20' >"
	strXML1 = strTemp
	strXML2 = strTemp
	for i=0 to oreport.FResultCount -1
		sTotal = sTotal + oreport.FItemList(i).FSellTotal
		nTotal = nTotal + oreport.FItemList(i).FSellCount
		uTotal = uTotal + oreport.FItemList(i).Funiqcnt
		strXML1 = strXML1 & "<set value='" & oreport.FItemList(i).FSellTotal & "' name='" & getUserLevelStr(oreport.FItemList(i).FUserLevel) & "' />"
		strXML2 = strXML2 & "<set value='" & oreport.FItemList(i).FSellCount & "' name='" & getUserLevelStr(oreport.FItemList(i).FUserLevel) & "' />"
	next
	strTemp = "<styles>" &_
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
	strXML1 = strXML1 & strTemp
	strXML2 = strXML2 & strTemp
end if
%>
<script type="text/javascript">

// 선택된 기간을 최근기준으로 날짜 입력
function presetTerm()
	select case document.frm.tDiv.value
		case "day"
			document.frm.startDate.value = dateAdd("d",((document.frm.tNo.value - 1) * -1),document.frm.endDate.value)
			document.all.strSDt.innerText = document.frm.startDate.value
		case "week"
			document.frm.startDate.value = dateAdd("ww",(document.frm.tNo.value * -1),document.frm.endDate.value)
			document.all.strSDt.innerText = document.frm.startDate.value
		case "month"
			document.frm.startDate.value = dateAdd("m",(document.frm.tNo.value * -1),document.frm.endDate.value)
			document.all.strSDt.innerText = document.frm.startDate.value
	end select
end function

</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript" src="/lib/util/chart/FusionCharts.js"></script>

<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<input id="startDate" name="startDate" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startDate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="endDate" name="endDate" value="<%=eDt%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="endDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "startDate", trigger    : "startDate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "endDate", trigger    : "endDate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		(<input type="checkbox" name="chkOld" value="Y" <% if chkOld="Y" then Response.Write "checked"%>> 6개월 이전 자료)
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	    <select name="isBanpum" class="select">
			<option value="all" <%=CHKIIF(isBanpum="all","selected","")%>>반품포함</option>
			<option value="plus" <%=CHKIIF(isBanpum="plus","selected","")%>>반품제외</option>
			<option value="minus" <%=CHKIIF(isBanpum="minus","selected","")%>>반품건만</option>
		</select>
				
		<input type="text" name="tNo" size="2" value="<%=tNo%>" style="text-align:right;">
		<select name="tDiv">
			<option value="day">일</option>
			<option value="week">주일</option>
			<option value="month">개월</option>
		</select>
		<script language=javascript>document.frm.tDiv.value="<%=tDiv%>";</script>
		<input type="button" value="적용" onclick="vbscript:presetTerm()" class="button">
		&nbsp;/&nbsp;
		브랜드 :
		<% drawSelectBoxDesigner "makerid",makerid %>
	</td>
</tr>
</table>
<!-- 검색 끝 -->

</form>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		반품,교환주문 제외건(정상주문건 만), 마일리지 포함
	</td>
	<td align="right">	

	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oreport.FResultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="188" rowspan="2">회원등급</td>
	<td width="228" colspan="2">매출</td>
	<td width="228" colspan="2">건수</td>
	<td width="50" rowspan="2">Uniq고객건수</td>
	<td width="106" rowspan="2">객단가(원)</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="139">매출액(원)</td>
	<td width="89">비율(%)</td>
	<td width="139">건수</td>
	<td width="89">비율(%)</td>
</tr>
<% if oreport.FResultCount>0 then %>
	<% for i=0 to oreport.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= getUserLevelStr(oreport.FItemList(i).FUserLevel) %></td>
		<td><%=FormatNumber(oreport.FItemList(i).FSellTotal,0)%></td>
		<td><%=FormatNumber((oreport.FItemList(i).FSellTotal/sTotal)*100,2)%>%</td>
		<td><%=FormatNumber(oreport.FItemList(i).FSellCount,0)%></td>
		<td><%=FormatNumber((oreport.FItemList(i).FSellCount/nTotal)*100,2)%>%</td>
		<td><%=FormatNumber(oreport.FItemList(i).Funiqcnt,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FSellAvr,0)%></td>
	</tr>
	<% next %>
	
	<tr align="center" bgcolor="#FAFAFA">
		<td>계</td>
		<td><%=FormatNumber(sTotal,0)%></td>
		<td>100%</td>
		<td><%=FormatNumber(nTotal,0)%></td>
		<td>100%</td>
		<td><%=FormatNumber(uTotal,0)%></td>
		<td><%=FormatNumber((sTotal/nTotal),0)%></td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>

</table>

<% if oreport.FResultCount>0 then %>
	<table width="800" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
	<tr>
		<td align="center" style="padding-top:10px;">
			<table width="640" border="0" cellpadding="3" cellspacing="2" class="a">
			<tr align="center">
				<td width="320" bgcolor="#E0E0E0">매출</td>
				<td width="320" bgcolor="#E0E0E0">건수</td>
			</tr>
			<tr>
				<td>
					<div id="chartdiv1" align="center"></div>
					<script type="text/javascript">	
						var chart = new FusionCharts("/lib/util/chart/Pie3D.swf", "chartdiv1", "320", "200", "0", "0");
						chart.setDataXML("<%=strXML1%>");
						chart.render("chartdiv1");
					</script>
				</td>
				<td>
					<div id="chartdiv2" align="center"></div>
					<script type="text/javascript">	
						var chart = new FusionCharts("/lib/util/chart/Pie3D.swf", "chartdiv2", "320", "200", "0", "0");
						chart.setDataXML("<%=strXML2%>");
						chart.render("chartdiv2");
					</script>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	</table>
<% end if %>

<%
set oreport = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
