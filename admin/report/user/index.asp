<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 기존, 신규회원 첫구매 현황
' History : 2014.06.24 원승현 개발
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/report/firstordercls.asp"-->

<%
	Dim defaultdate1, yyyy1, mm1, dd1, yyyy2, mm2, dd2, orderfirstlist, i, oldOrdFstTotalCnt, newOrdFstTotalCnt, strTemp, strXML, ChartViDi
	Dim strDate, strDateLen, strolder, strolderLen, strnew, strnewLen

	oldOrdFstTotalCnt = 0
	newOrdFstTotalCnt = 0

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


	set orderfirstlist = new CFirstOrder
	orderfirstlist.FRectFromDate = dateserial(yyyy1,mm1,dd1)
	orderfirstlist.FRectToDate = dateserial(yyyy2,mm2,dd2)
	orderfirstlist.GetFirstOrderReport()


	If yyyy1 <> "" And yyyy2 <> "" Then

		ChartViDi = DateDiff("d", yyyy1&"-"&mm1&"-"&dd1, yyyy2&"-"&mm2&"-"&dd2)

	End If


	'// 일자
	if orderfirstlist.ftotalcount > 0 Then
		strDate = ""
		for i = 0 to orderfirstlist.ftotalcount -1
			strDate = strDate & "{'label': '"&orderfirstlist.FItemList(i).FdataDate&"'},"
		Next
			strDateLen = Len(strDate)
			strDate = Left(strDate, strDateLen - 1)
	End If


	'// 기존회원 첫구매자
	if orderfirstlist.ftotalcount > 0 Then
		strolder = ""
		for i = 0 to orderfirstlist.ftotalcount -1
			strolder = strolder & "{'value': '"&CDbl(orderfirstlist.FItemList(i).FoldOrdFst)-CDbl(orderfirstlist.FItemList(i).FnewOrdFst)&"'},"
		Next
			strolderLen = Len(strolder)
			strolder = Left(strolder, strolderLen - 1)
	End If


	'// 신규회원 첫구매자
	if orderfirstlist.ftotalcount > 0 Then
		strnew = ""
		for i = 0 to orderfirstlist.ftotalcount -1
			strnew = strnew & "{'value': '"&orderfirstlist.FItemList(i).FnewOrdFst&"'},"
		Next
			strnewLen = Len(strnew)
			strnew = Left(strnew, strnewLen - 1)
	End If
%>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>

<script type="text/javascript">
  FusionCharts.ready(function(){
    var myChart = new FusionCharts({
        "type": "mscolumn3d",
        "renderAt": "chartContainer",
        "width": "100%",
        "height": "300",
        "dataFormat": "json",
        "dataSource":  {
   "chart": {
      "caption": "첫구매통계",
      "subCaption": "",
      "xAxisName": "일자",
      "showborder": "0",
      "yAxisName": "명",
      "paletteColors": "#6baa01,#008ee4",
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
         "seriesname": "기존회원 첫구매자",
         "data": [
            <%=strolder%>
         ]
      },
      {
         "seriesname": "신규회원 첫구매자",
         "data": [
            <%=strnew%>
         ]
      }
   ]
	}
  });
myChart.render();
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
		<font color="red">- 하루전 데이터까지만 검색가능합니다.<br>- 데이터는 2014년 1월1일부터 검색 가능합니다.<br>- 기존회원 첫구매자는 기존에 가입한 회원들중 해당일자에 처음 구매한 회원수 입니다.<br>- 신규회원 첫구매자는 해당일자에 가입하고 해당일자에 바로 구매한 회원수 입니다.</font>
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<% if orderfirstlist.ftotalcount > 0 then %>
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td >일자</td>
		<td >기존회원 첫구매자</td>
		<td >신규회원 첫구매자</td>
		<td >첫구매자 합계</td>
	</tr>
	<% for i = 0 to orderfirstlist.ftotalcount -1 %>
	<tr bgcolor="#FFFFFF" align="center">
		<td><%=orderfirstlist.FItemList(i).FdataDate%></td>
		<td><%=FormatNumber((CLng(orderfirstlist.FItemList(i).FoldOrdFst)-CLng(orderfirstlist.FItemList(i).FnewOrdFst)), 0)%>명</td>
		<td><%=FormatNumber(orderfirstlist.FItemList(i).FnewOrdFst, 0)%>명</td>
		<td bgcolor="<%= adminColor("tabletop") %>"><%=FormatNumber(orderfirstlist.FItemList(i).FoldOrdFst, 0)%>명</td>
	</tr>
	<%
		oldOrdFstTotalCnt = (CLng(orderfirstlist.FItemList(i).FoldOrdFst)-CLng(orderfirstlist.FItemList(i).FnewOrdFst)) + oldOrdFstTotalCnt
		newOrdFstTotalCnt = CLng(orderfirstlist.FItemList(i).FnewOrdFst) + newOrdFstTotalCnt
	%>
	<% next %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td >총합</td>
		<td ><%=FormatNumber(oldOrdFstTotalCnt,0)%>명</td>
		<td ><%=FormatNumber(newOrdFstTotalCnt,0)%>명</td>
		<td><%=FormatNumber(CLng(oldOrdFstTotalCnt)+CLng(newOrdFstTotalCnt),0)%>명</td>
	</tr>
	</table>
<% else %>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#FFFFFF">
		<td >검색 결과가 없습니다.</td>
	</tr>
	</table>
<% end if %>

<% If ChartViDi < 31 Then %>
	<% if orderfirstlist.ftotalcount>0 then %>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
	<tr>
		<td align="center" style="padding-top:10px;">
			<div id="chartContainer" align="center" ></div>

		</td>
	</tr>
	</table>
	<% end if %>
<% End If %>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
