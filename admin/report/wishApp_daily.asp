<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  wishApp ���ϵ�����
' History : 2014.07.03 ������ ����
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

	defaultdate1 = dateadd("d",-10,year(now) & "-" &month(now) & "-" & day(now))		'��¥���� ������ �⺻������ 10�������� �˻�	
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

	'// ����
	if wishAppDailylist.ftotalcount > 0 Then
		strDate = ""
		for i = 0 to wishAppDailylist.ftotalcount -1 
			strDate = strDate & "{'label': '"&wishAppDailylist.FItemList(i).Fregdate&"'},"
		Next
			strDateLen = Len(strDate)
			strDate = Left(strDate, strDateLen - 1)
	End If

	'// PCWEB�α���
	if wishAppDailylist.ftotalcount > 0 Then
		strWeb = ""
		for i = 0 to wishAppDailylist.ftotalcount -1 
			strWeb = strWeb & "{'value': '"&wishAppDailylist.FItemList(i).FlogWeb&"'},"
		Next
			strWebLen = Len(strWeb)
			strWeb = Left(strWeb, strWebLen - 1)
	End If

	'// ����� �α���
	if wishAppDailylist.ftotalcount > 0 Then
		strMobile = ""
		for i = 0 to wishAppDailylist.ftotalcount -1 
			strMobile = strMobile & "{'value': '"&wishAppDailylist.FItemList(i).FlogMobile&"'},"
		Next
			strMobileLen = Len(strMobile)
			strMobile = Left(strMobile, strMobileLen - 1)
	End If

	'// �� �α���
	if wishAppDailylist.ftotalcount > 0 Then
		strApp = ""
		for i = 0 to wishAppDailylist.ftotalcount -1 
			strApp = strApp & "{'value': '"&wishAppDailylist.FItemList(i).FlogApp&"'},"
		Next
			strAppLen = Len(strApp)
			strApp = Left(strApp, strAppLen - 1)
	End If

	'// ios ��� pushid
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

	'// Android ��� pushid
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
      "caption": "�α���",
      "subCaption": "",
      "xAxisName": "����",
      "showborder": "0",
      "yAxisName": "��",
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
         "seriesname": "�α��� ��",
         "data": [
            <%=strWeb%>
         ]
      },
      {
         "seriesname": "�α��� �����",
         "data": [
            <%=strMobile%>
         ]
      },
      {
         "seriesname": "�α��� ���þ�",
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
      "caption": "App���(pushid)��",
      "subCaption": "",
      "xAxisName": "����",
      "showborder": "0",
      "yAxisName": "��",
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
      "caption": "App���(NID)��",
      "subCaption": "",
      "xAxisName": "����",
      "showborder": "0",
      "yAxisName": "��",
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

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form action="" name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">			
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<font color="red">- �Ϸ��� �����ͱ����� �˻������մϴ�.<br>- �����ʹ� 2014�� 4��1�Ϻ��� �˻� �����մϴ�.</font>	
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- �׼� �� -->

<% if wishAppDailylist.ftotalcount > 0 then %>			
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan="2">��¥</td>
		<td colspan="4">�α���</td>
		<td colspan="2">�ȷ���</td>
		<td colspan="4">���û�ǰ</td>
		<td colspan="3">��������</td>
		<td colspan="3">App ���(pushid)��</td>
		<td colspan="3">App ���(NID)��</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>��</td>
		<td>�����</td>
		<td colspan="2">���þ�</td>
		<td>ȸ����</td>
		<td>����</td>
		<td>���ϴ�����</td>
		<td>��ü</td>
		<td>����</td>
		<td>������</td>
		<td>��ü</td>
		<td>����</td>
		<td>������</td>
		<td>IOS</td>
		<td>Android</td>
		<td>��</td>
		<td>IOS</td>
		<td>Android</td>
		<td>��</td>
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
		<td >�˻� ����� �����ϴ�.</td>
	</tr>
	</table>
<% end if %>

<% If ChartViDi < 300 Then %> <!-- ���� 31 �� �Ǿ� �־���?? -->
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