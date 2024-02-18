<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/admin/dataanalysis/chart/chartCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim mxChartSeries : mxChartSeries = 5
Dim oChart, vArr1, vArr2, i, j, k
Dim vSDate, vEDate, vChannel, grptype, datebase

vSDate = requestCheckvar(request("startdate"),10)
vEDate = requestCheckvar(request("enddate"),10)
vChannel = requestCheckvar(request("channel"),10)
grptype = requestCheckvar(request("grptype"),32)
datebase = requestCheckvar(request("datebase"),10)

if (grptype="") then grptype="d" ''d / m

    
If vSDate = "" Then
	vSDate = dateadd("d",-31,Date())
End If

If vEDate = "" Then
	vEDate = dateadd("d",-1,Date())
End If

if (datebase="") then datebase="ipkumdt"

SET oChart = new CChart
	oChart.FRectSDate = vSDate
	oChart.FRectEDate = vEDate
	oChart.FRectChannel = vChannel
	oChart.FRectGroupType = grptype
    oChart.FRectDateBase = datebase

    vArr1 = oChart.fnNvSellp_Trend()
	


Dim precate,posN
dim vArrTitle
dim vArrPos 
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>

<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>

<script>
$(function() {
	var CAL_Start = new Calendar({
		inputField : "startdate", trigger    : "startdate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_End.args.min = date;
			CAL_End.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
	var CAL_End = new Calendar({
		inputField : "enddate", trigger    : "enddate_trigger",
		onSelect: function() {
			var date = Calendar.intToDate(this.selection.get());
			CAL_Start.args.max = date;
			CAL_Start.redraw();
			this.hide();
		}, bottomBar: true, dateFormat: "%Y-%m-%d"
	});
});

function goSearch(){
	if($("#sdate").val() == ""){
		alert("시작일을 입력하세요");	
		return false;
	}
	if($("#edate").val()== ""){
		alert("종료일을 입력하세요");	
		return false;
	}
	document.frm1.submit();
}
</script>

<script type='text/javascript'>//<![CDATA[
window.onload=function(){
    
<% if isArray(vArr1) then %>
<%
    vArrTitle = Array("NV 매출비중/자사","자사몰수익율","NV 수익율")
    vArrPos = Array(13,14,15)
%>
// "VN 매출비중 및 수익율",
FusionCharts.ready(function () {
    var vstrChart1 = new FusionCharts({
        type: 'msline', //'', 
        renderAt: 'chart-container0',
        width: '800',
        height: '400',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "VN 매출비중 및 수익율",
                "subCaption": "",
                "xAxisName": "날짜",
                "yAxisName": "%",
                "theme": "fint",
                "showSum": "1",
                "showValues": "<%=CHKIIF(UBound(vArr1,2)>14,0,1)%>",
                //Setting automatic calculation of div lines to off
  //              "adjustDiv": "0",
                //Manually defining y-axis lower and upper limit
                //"yAxisMaxvalue": "35000",	//y축 맥스값
                //"yAxisMinValue": "5000",		//y축 민값
                //Setting number of divisional lines to 9
                //"numDivLines": "9"				//0~맥스 사이 표시되어지는 수치갯수
  //              "anchorBgHoverColor": "#96d7fa",
  //              "anchorBorderHoverThickness" : "4",
  //              "anchorHoverRadius":"7"
            },
            // X축 
            "categories": [
                {
                    "category": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
							    if (precate<>vArr1(0,i)) then
    								Response.Write "{" & vbCrLf
    								Response.Write """label"": """&vArr1(0,i)&"""" & vbCrLf
    								Response.Write "}"
    								If i <> UBound(vArr1,2) Then
    									Response.Write ","
    								End If
    								Response.Write vbCrLf
    								precate=vArr1(0,i)
							    end if
							Next
						End If
						%>
                    ]
                }
            ],            
            "dataset": [
                <% for k=LBound(vArrTitle) to UBound(vArrTitle) %>
                {
                    "seriesname": "<%=vArrTitle(k)%>",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
							    posN = vArrPos(k)
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(posN,i)&"""" & vbCrLf
								Response.Write "}"
								If i <> UBound(vArr1,2) Then
									Response.Write ","
								End If
								Response.Write vbCrLf
							Next
						End If
						%>
                    ]
                }
                <% if (k<UBound(vArrTitle)) then response.write "," %>
                <% next %>
            ]
        }
    }).render();
});


<%
    vArrTitle = Array("자사몰매출","NV 매출","제휴몰 매출")
    vArrPos = Array(7,11,18)
%>

FusionCharts.ready(function () {
    var vstrChart1 = new FusionCharts({
        type: 'stackedcolumn2d', //'', 
        renderAt: 'chart-container1',
        width: '800',
        height: '400',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "자사몰 매출액 / NV 매출액 / 제휴몰 매출",
                "subCaption": "",
                "xAxisName": "날짜",
                "yAxisName": "매출액",
                "theme": "fint",
                "showSum": "<%=CHKIIF(UBound(vArr1,2)>14,0,1)%>",
                "showValues": "<%=CHKIIF(UBound(vArr1,2)>14,0,1)%>",
                //Setting automatic calculation of div lines to off
  //              "adjustDiv": "0",
                //Manually defining y-axis lower and upper limit
                //"yAxisMaxvalue": "35000",	//y축 맥스값
                //"yAxisMinValue": "5000",		//y축 민값
                //Setting number of divisional lines to 9
                //"numDivLines": "9"				//0~맥스 사이 표시되어지는 수치갯수
  //              "anchorBgHoverColor": "#96d7fa",
  //              "anchorBorderHoverThickness" : "4",
  //              "anchorHoverRadius":"7"
            },
            // X축 
            "categories": [
                {
                    "category": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
							    'if (precate<>vArr1(0,i)) then
    								Response.Write "{" & vbCrLf
    								Response.Write """label"": """&vArr1(0,i)&"""" & vbCrLf
    								Response.Write "}"
    								If i <> UBound(vArr1,2) Then
    									Response.Write ","
    								End If
    								Response.Write vbCrLf
    							'	precate=vArr1(0,i)
							    'end if
							Next
						End If
						%>
                    ]
                }
            ],            
            "dataset": [
                <% for k=LBound(vArrTitle) to UBound(vArrTitle) %>
                {
                    "seriesname": "<%=vArrTitle(k)%>",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
							    posN = vArrPos(k)
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(posN,i)&"""" & vbCrLf
								Response.Write "}"
								If i <> UBound(vArr1,2) Then
									Response.Write ","
								End If
								Response.Write vbCrLf
							Next
						End If
						%>
                    ]
                }
                <% if (k<UBound(vArrTitle)) then response.write "," %>
                <% next %>
            ]
        }
    }).render();
});



<%
    vArrTitle = Array("자사몰매출","NV 매출","제휴몰 매출")
    vArrPos = Array(3,11,18)
%>
// "VN 매출액",
FusionCharts.ready(function () {
    var vstrChart2 = new FusionCharts({
        type: 'msline', //'', 
        renderAt: 'chart-container2',
        width: '800',
        height: '400',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "자사몰 매출액 / NV 매출액 / 제휴몰 매출",
                "subCaption": "",
                "xAxisName": "날짜",
                "yAxisName": "매출액",
                "theme": "fint",
                "showSum": "1",
                "showValues": "<%=CHKIIF(UBound(vArr1,2)>14,0,1)%>",
                //Setting automatic calculation of div lines to off
  //              "adjustDiv": "0",
                //Manually defining y-axis lower and upper limit
                //"yAxisMaxvalue": "35000",	//y축 맥스값
                //"yAxisMinValue": "5000",		//y축 민값
                //Setting number of divisional lines to 9
                //"numDivLines": "9"				//0~맥스 사이 표시되어지는 수치갯수
  //              "anchorBgHoverColor": "#96d7fa",
  //              "anchorBorderHoverThickness" : "4",
  //              "anchorHoverRadius":"7"
            },
            // X축 
            "categories": [
                {
                    "category": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
							    if (precate<>vArr1(0,i)) then
    								Response.Write "{" & vbCrLf
    								Response.Write """label"": """&vArr1(0,i)&"""" & vbCrLf
    								Response.Write "}"
    								If i <> UBound(vArr1,2) Then
    									Response.Write ","
    								End If
    								Response.Write vbCrLf
    								precate=vArr1(0,i)
							    end if
							Next
						End If
						%>
                    ]
                }
            ],            
            "dataset": [
                <% for k=LBound(vArrTitle) to UBound(vArrTitle) %>
                {
                    "seriesname": "<%=vArrTitle(k)%>",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
							    posN = vArrPos(k)
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(posN,i)&"""" & vbCrLf
								Response.Write "}"
								If i <> UBound(vArr1,2) Then
									Response.Write ","
								End If
								Response.Write vbCrLf
							Next
						End If
						%>
                    ]
                }
                <% if (k<UBound(vArrTitle)) then response.write "," %>
                <% next %>
            ]
        }
    }).render();
});
<% end if %>




}//]]>
</script>
<%
SET oChart = nothing
%>
<body>
<form name="frm1" method="post" action="/admin/dataanalysis/chart/nv_trend.asp">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#F4F4F4">
    <td width="50" rowspan="2" bgcolor="#EEEEEE">검색<br>조건</td>
	<td align="left">
	<select name="datebase">
        <option value="ipkumdt" <%=CHKIIF(datebase="ipkumdt","selected","")%> >결제일
        <option value="oipkumdt" <%=CHKIIF(datebase="oipkumdt","selected","")%> >원결제일
        <option value="beasongdt" <%=CHKIIF(datebase="beasongdt","selected","")%> >출고일
    </select>
     : 
	    <input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
		~<input id='enddate' name='enddate' value='<%= vEDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
    &nbsp;&nbsp;
    
    채널 : <% call drawConversionChannelSelectBoxII("channel",vChannel) %>
    &nbsp;&nbsp;
    날짜구분 : 
    <input type="radio" name="grptype" value="d" <%=CHKIIF(grptype="d","checked","") %> >일
    <input type="radio" name="grptype" value="m" <%=CHKIIF(grptype="m","checked","") %> >월
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="goSearch(document.frm1);">
	</td>
</tr>

</table>
</form>

* 매출로그기준, 1일 지연데이터<br />
* 출고일기준의 경우 배송비는 정산작업 후 반영됩니다.<br />
*EP쿠폰사용액(rdsite 관계없이 실제 Naver 쿠폰 사용액)<br />
* 매출로그결제일(취소일자(==결제일)반영됨), 매출로그원결제일(취소일자무관 원결제일)
<% dim sum1,sum2,sum3,sum4,sum5,sum6,sum7 %>
<table cellpadding="0" cellspacing="0" border="0" class="a" align="center">
<tr bgcolor="#FFFFFF">
    <% if isArray(vArr1) then %>
    
    <td width="900" valign="top">
        <!-- yyyymm	판매액	구매총액	매출총액	매입총액	NV제외_판매액	NV제외_구매총액	NV제외_매출총액	NV제외_매입총액	NV_판매액	NV_구매총액	NV_매출총액	NV_매입총액-->
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4">
            <td>날짜</td>
            <td>자사판매액<br>(Nv포함)</td>
            <td>자사주문수<br>(Nv포함)</td>
            <td>자사구매총액<br>(Nv포함)</td>
            <td>자사매출총액<br>(Nv포함)</td>
            <td>자사매입총액<br>(Nv포함)</td>
            
            <td>NV 판매액</td>
            <td>NV 주문수</td>
            <td>NV 구매총액</td>
            <td>NV 매출총액</td>
            <td>NV 매입총액</td>
            
            <td>NV<br>매출비중</td>
            <td>자사몰<br>수익율2</td>
            <td>NV<br>수익율2</td>
            <td>EP쿠폰<br>사용액</td>
        </tr>
        <% For i = 0 To UBound(vArr1,2) %>
        <tr bgcolor="#FFFFFF" align="right">
            <td align="left"><%=vArr1(0,i)%></td>
            <td><%=FormatNumber(vArr1(1,i),0)%></td>
            <td><%=FormatNumber(vArr1(21,i),0)%></td>
            <td><%=FormatNumber(vArr1(2,i),0)%></td>
            <td><%=FormatNumber(vArr1(3,i),0)%></td>
            <td><%=FormatNumber(vArr1(4,i),0)%></td>
            
            <td><%=FormatNumber(vArr1(9,i),0)%></td>
            <td><%=FormatNumber(vArr1(23,i),0)%></td>
            <td><%=FormatNumber(vArr1(10,i),0)%></td>
            <td><%=FormatNumber(vArr1(11,i),0)%></td>
            <td><%=FormatNumber(vArr1(12,i),0)%></td>
            
            <td><%=FormatNumber(vArr1(13,i),2)%></td>
            <td><%=FormatNumber(vArr1(14,i),2)%></td>
            <td><%=FormatNumber(vArr1(15,i),2)%></td>
            <td><%=FormatNumber(vArr1(20,i),0)%></td>
        </tr>
        <% next %>
        </table>
    </td>
    <% end if %>
    
	<td valign="top">
	    <div id="chart-container0">FusionCharts will render here</div>
	    <div id="chart-container1">FusionCharts will render here</div>
	    <div id="chart-container2">FusionCharts will render here</div>
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->