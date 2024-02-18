<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  wishApp 일일데이터
' History : 2014.07.03 원승현 개발
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/report/wishAppDailycls.asp"-->

<%
	Dim defaultdate1, yyyy1, mm1, dd1, yyyy2, mm2, dd2, wishAppDailylist, i, strTemp, strXML, ChartViDi, strDay
	Dim strWeb, strMobile, strApp, strWebLen, strMobileLen, strAppLen, strDate, strDateLen, striOs, striOsLen, strAnd, strAndLen
    Dim strNidiOs, strNidiOsLen, strNidAnd, strNidAndLen

	defaultdate1 = dateadd("d",-10,year(now) & "-" &month(now) & "-" & day(now))		'날짜값이 없을때 기본값으로 10이전까지 검색	
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


	set wishAppDailylist = new CwishAppDaily
	wishAppDailylist.FRectFromDate = dateserial(yyyy1,mm1,dd1)
	wishAppDailylist.FRectToDate = dateserial(yyyy2,mm2,dd2)
	wishAppDailylist.GetwishAppDailyReport()


	If yyyy1 <> "" And yyyy2 <> "" Then
		ChartViDi = DateDiff("d", yyyy1&"-"&mm1&"-"&dd1, yyyy2&"-"&mm2&"-"&dd2)
	End If

	'// 일자
	if wishAppDailylist.ftotalcount > 0 Then
		strDate = ""
		for i = 0 to wishAppDailylist.ftotalcount -1 
			strDate = strDate & "{'label': '"&wishAppDailylist.FItemList(i).Fregdate&"'},"
		Next
			strDateLen = Len(strDate)
			strDate = Left(strDate, strDateLen - 1)
	End If

	'// PCWEB로그인
	if wishAppDailylist.ftotalcount > 0 Then
		strWeb = ""
		for i = 0 to wishAppDailylist.ftotalcount -1 
			strWeb = strWeb & "{'value': '"&wishAppDailylist.FItemList(i).FlogWeb&"'},"
		Next
			strWebLen = Len(strWeb)
			strWeb = Left(strWeb, strWebLen - 1)
	End If

	'// 모바일 로그인
	if wishAppDailylist.ftotalcount > 0 Then
		strMobile = ""
		for i = 0 to wishAppDailylist.ftotalcount -1 
			strMobile = strMobile & "{'value': '"&wishAppDailylist.FItemList(i).FlogMobile&"'},"
		Next
			strMobileLen = Len(strMobile)
			strMobile = Left(strMobile, strMobileLen - 1)
	End If

	'// 앱 로그인
	if wishAppDailylist.ftotalcount > 0 Then
		strApp = ""
		for i = 0 to wishAppDailylist.ftotalcount -1 
			strApp = strApp & "{'value': '"&wishAppDailylist.FItemList(i).FlogApp&"'},"
		Next
			strAppLen = Len(strApp)
			strApp = Left(strApp, strAppLen - 1)
	End If

	'// ios 등록 pushid
	if wishAppDailylist.ftotalcount > 0 Then
		striOs = ""
		for i = 0 to wishAppDailylist.ftotalcount -1 
			striOs = striOs & "{'value': '"&wishAppDailylist.FItemList(i).Fappios&"'},"
			strNidiOs = strNidiOs & "{'value': '"&wishAppDailylist.FItemList(i).FAppIosNid&"'},"
		Next
			striOsLen = Len(striOs)
			striOs = Left(striOs, striOsLen - 1)
			
			strNidiOsLen = Len(strNidiOs)
			strNidiOs = Left(strNidiOs, strNidiOsLen - 1)
	End If

	'// Android 등록 pushid
	if wishAppDailylist.ftotalcount > 0 Then
		strAnd = ""
		for i = 0 to wishAppDailylist.ftotalcount -1 
			strAnd = strAnd & "{'value': '"&wishAppDailylist.FItemList(i).Fappand&"'},"
			strNidAnd = strNidAnd & "{'value': '"&wishAppDailylist.FItemList(i).FAppAndNid&"'},"
		Next
			strAndLen = Len(strAnd)
			strAnd = Left(strAnd, strAndLen - 1)
			
			strNidAndLen = Len(strNidAnd)
			strNidAnd = Left(strNidAnd, strNidAndLen - 1)
	End If
	
	
	

%>

<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>

<script type="text/javascript">
  FusionCharts.ready(function(){
    var myChart = new FusionCharts({
        "type": "msline",
        "renderAt": "chartContainer",
        "width": "100%",
        "height": "300",
        "dataFormat": "json",
        "dataSource":  {
   "chart": {
      "caption": "로그인",
      "subCaption": "",
      "xAxisName": "일자",
      "showborder": "0",
      "yAxisName": "명",
      "paletteColors": "#0000cd,#dc143c,#008000",
      "bgAlpha": "0",
      "borderAlpha": "20",
      "canvasBorderAlpha": "0",
      "usePlotGradientColor": "0",
      "plotBorderAlpha": "10",
      "legendBorderAlpha": "0",
      "legendShadow": "0",
      "captionpadding": "20",
      "showXAxisLines": "1",
      "axisLineAlpha": "25",
      "divLineAlpha": "10",
      "showValues": "0",
      "showAlternateHGridColor": "0",
      "animation": "1",
      "showYAxisValues": "1",
      "yAxisNamePadding": "0",
      "showtooltip": "1",
	  "formatNumberScale":"0",
	  "rotateYAxisName":"0"


   },
   "categories": [
      {
         "category": [
			<%=strDate%>
         ]
      }
   ],
   "dataset": [
      {
         "seriesname": "로그인 웹",
         "data": [
            <%=strWeb%>
         ]
      },
      {
         "seriesname": "로그인 모바일",
         "data": [
            <%=strMobile%>
         ]
      },
      {
         "seriesname": "로그인 위시앱",
         "data": [
            <%=strApp%>
         ]
      }
   ]
	}
  });
myChart.render();
})
</script>



<script type="text/javascript">
  FusionCharts.ready(function(){
    var myChart2 = new FusionCharts({
        "type": "msline",
        "renderAt": "chartContainer2",
        "width": "100%",
        "height": "300",
        "dataFormat": "json",
        "dataSource":  {
   "chart": {
      "caption": "App등록(pushid)수",
      "subCaption": "",
      "xAxisName": "일자",
      "showborder": "0",
      "yAxisName": "명",
      "paletteColors": "#800080,#008080",
      "bgAlpha": "0",
      "borderAlpha": "20",
      "canvasBorderAlpha": "0",
      "usePlotGradientColor": "0",
      "plotBorderAlpha": "10",
      "legendBorderAlpha": "0",
      "legendShadow": "0",
      "captionpadding": "20",
      "showXAxisLines": "1",
      "axisLineAlpha": "25",
      "divLineAlpha": "10",
      "showValues": "0",
      "showAlternateHGridColor": "0",
      "animation": "1",
      "showYAxisValues": "1",
      "yAxisNamePadding": "0",
      "showtooltip": "1",
	  "formatNumberScale":"0",
	  "rotateYAxisName":"0"


   },
   "categories": [
      {
         "category": [
			<%=strDate%>
         ]
      }
   ],
   "dataset": [
      {
         "seriesname": "iOs",
         "data": [
            <%=striOs%>
         ]
      },
      {
         "seriesname": "Android",
         "data": [
            <%=strAnd%>
         ]
      }
   ]
	}
  });
myChart2.render();
})
</script>

<script type="text/javascript">
  FusionCharts.ready(function(){
    var myChart3 = new FusionCharts({
        "type": "msline",
        "renderAt": "chartContainer3",
        "width": "100%",
        "height": "300",
        "dataFormat": "json",
        "dataSource":  {
   "chart": {
      "caption": "App등록(NID)수",
      "subCaption": "",
      "xAxisName": "일자",
      "showborder": "0",
      "yAxisName": "명",
      "paletteColors": "#800080,#008080",
      "bgAlpha": "0",
      "borderAlpha": "20",
      "canvasBorderAlpha": "0",
      "usePlotGradientColor": "0",
      "plotBorderAlpha": "10",
      "legendBorderAlpha": "0",
      "legendShadow": "0",
      "captionpadding": "20",
      "showXAxisLines": "1",
      "axisLineAlpha": "25",
      "divLineAlpha": "10",
      "showValues": "0",
      "showAlternateHGridColor": "0",
      "animation": "1",
      "showYAxisValues": "1",
      "yAxisNamePadding": "0",
      "showtooltip": "1",
	  "formatNumberScale":"0",
	  "rotateYAxisName":"0"


   },
   "categories": [
      {
         "category": [
			<%=strDate%>
         ]
      }
   ],
   "dataset": [
      {
         "seriesname": "iOs",
         "data": [
            <%=strNidiOs%>
         ]
      },
      {
         "seriesname": "Android",
         "data": [
            <%=strNidAnd%>
         ]
      }
   ]
	}
  });
myChart3.render();
})
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form action="" name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
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

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<font color="red">- 하루전 데이터까지만 검색가능합니다.<br>- 데이터는 2014년 4월1일부터 검색 가능합니다.</font>	
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<% if wishAppDailylist.ftotalcount > 0 then %>			
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan="2">날짜</td>
		<td colspan="4">로그인</td>
		<td colspan="2">팔로잉</td>
		<td colspan="4">위시상품</td>
		<td colspan="3">위시폴더</td>
		<td colspan="3">App 등록(pushid)수</td>
		<td colspan="3">App 등록(NID)수</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>웹</td>
		<td>모바일</td>
		<td colspan="2">위시앱</td>
		<td>회원수</td>
		<td>대상수</td>
		<td>전일대증분</td>
		<td>전체</td>
		<td>공개</td>
		<td>공개율</td>
		<td>전체</td>
		<td>공개</td>
		<td>공개율</td>
		<td>IOS</td>
		<td>Android</td>
		<td>계</td>
		<td>IOS</td>
		<td>Android</td>
		<td>계</td>
	</tr>
	<% for i = 0 to wishAppDailylist.ftotalcount -1 %>
		<tr bgcolor="#FFFFFF" align="center">
			<td><%=wishAppDailylist.FItemList(i).Fregdate%></td>
			<td><%=FormatNumber(wishAppDailylist.FItemList(i).FlogWeb,0)%></td>
			<td><%=FormatNumber(wishAppDailylist.FItemList(i).FlogMobile,0)%></td>
			<td><%=FormatNumber(wishAppDailylist.FItemList(i).FlogApp,0)%></td>
			<td><%=FormatNumber((CDbl(wishAppDailylist.FItemList(i).FlogApp)/(CDbl(wishAppDailylist.FItemList(i).FlogWeb)+CDbl(wishAppDailylist.FItemList(i).FlogMobile)+CDbl(wishAppDailylist.FItemList(i).FlogApp)))*100, 2)%>%</td>
			<td>
				<%
					If len(wishAppDailylist.FItemList(i).FfollowuCnt) = 0 Or IsNull(wishAppDailylist.FItemList(i).FfollowuCnt) Then
						response.write "0"
					Else
						response.write FormatNumber(wishAppDailylist.FItemList(i).FfollowuCnt, 0)
					End If
				%>
			</td>
			<td>
				<%
					If Len(wishAppDailylist.FItemList(i).FfollowpCnt)=0 Or IsNull(wishAppDailylist.FItemList(i).FfollowpCnt) Then
						response.write "0"
					Else
						response.write FormatNumber(wishAppDailylist.FItemList(i).FfollowpCnt,0)
					End If
				%>
			</td>
			<td><%=FormatNumber(wishAppDailylist.FItemList(i).FprevDayPM,0)%></td>
			<td><%=FormatNumber(wishAppDailylist.FItemList(i).FwishpdAll,0)%></td>
			<td><%=FormatNumber(wishAppDailylist.FItemList(i).FwishpdView,0)%></td>
			<td><%=FormatNumber((CDbl(wishAppDailylist.FItemList(i).FwishpdView)/Cdbl(wishAppDailylist.FItemList(i).FwishpdAll))*100, 2)%>%</td>
			<td><%=FormatNumber(wishAppDailylist.FItemList(i).FwishfdAll,0)%></td>
			<td><%=FormatNumber(wishAppDailylist.FItemList(i).FwishfdView,0)%></td>
			<td><%=FormatNumber((CDbl(wishAppDailylist.FItemList(i).FwishfdView)/Cdbl(wishAppDailylist.FItemList(i).FwishfdAll))*100, 2)%>%</td>
			<td><%=FormatNumber(wishAppDailylist.FItemList(i).Fappios,0)%></td>
			<td><%=FormatNumber(wishAppDailylist.FItemList(i).Fappand,0)%></td>
			<td><%=FormatNumber(CDbl(wishAppDailylist.FItemList(i).Fappios)+CDbl(wishAppDailylist.FItemList(i).Fappand),0)%></td>
			
			<td><%=FormatNumber(wishAppDailylist.FItemList(i).FAppIosNid,0)%></td>
			<td><%=FormatNumber(wishAppDailylist.FItemList(i).FAppAndNid,0)%></td>
			<td><%=FormatNumber(CDbl(wishAppDailylist.FItemList(i).FAppIosNid)+CDbl(wishAppDailylist.FItemList(i).FAppAndNid),0)%></td>
			
		</tr>
	<% next %>
	</table>
<% else %>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#FFFFFF">
		<td >검색 결과가 없습니다.</td>
	</tr>
	</table>
<% end if %>

<% If ChartViDi < 300 Then %> <!-- 기존 31 로 되어 있었음?? -->
	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
	<tr>
		<td align="center" style="padding-top:10px;width:50%">
			<div id="chartContainer"></div>
		</td>
	</tr>
	<tr>
	    <td align="center" style="padding-top:10px;width:50%">
			<div id="chartContainer2"></div>
		</td>
	</tr>
	<tr>
	    <td align="center" style="padding-top:10px;width:50%">
			<div id="chartContainer3"></div>
		</td>
	</tr>
	</table>

<% End If %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->