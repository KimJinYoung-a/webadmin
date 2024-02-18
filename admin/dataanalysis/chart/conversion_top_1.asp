<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/dataanalysis/chart/chartCls.asp" -->
<%
Dim oChart, vArr1, vArr2, i, j
Dim vSDate, vEDate, vChannel, vpType, vpValue, vOrdType, sTp, vpUpType

vSDate = requestCheckvar(request("sdate"),10)
vEDate = requestCheckvar(request("edate"),10)
vChannel = requestCheckvar(request("channel"),10)
vpType = requestCheckvar(request("ptype"),32)
vpValue = requestCheckvar(request("pvalue"),64)
vOrdType = requestCheckvar(request("ordtype"),32)
vpUpType = requestCheckvar(request("puptype"),32)

if (vpType="") then vpType="pRtr"
if (vOrdType="") then vOrdType="C" ''건수(C) , 금액(S), 수익(G)

if (sTp="") then sTp="1" 
    
If vSDate = "" Then
	vSDate = dateadd("d",-7,Date())
End If

If vEDate = "" Then
	vEDate = date()
End If

SET oChart = new CChart
	oChart.FRectSDate = vSDate
	oChart.FRectEDate = vEDate
	oChart.FRectChannel = vChannel
	oChart.FRectPType = vpType
	oChart.FPageSize = CHKIIF(vpType="gaparam",50,50)
	oChart.FRectOrderType = vOrdType
	oChart.FRectPValue = vpValue
	oChart.FRectUPTypeValue = vpUpType
	vArr2 = oChart.fnConversionTopByType
	
	oChart.FPageSize = 5
	vArr1 = oChart.fnConversionTopByType_Trend
SET oChart = nothing

Dim iChartCaption : iChartCaption = "전환 타입별 주문건수"
Dim iChartSubCaption : iChartSubCaption = ""
Dim ixAxisName : ixAxisName = "" ''날짜
Dim yAxisName : yAxisName = "주문건수"
Dim iDataSetPosArr : iDataSetPosArr = Array(2)
Dim iDataSetHeadArr : iDataSetHeadArr = Array("주문건수")

if (vOrdType="2") then
    iDataSetPosArr = Array(6)
    iDataSetHeadArr = Array("구매총액")
    
    iChartCaption = "전환 타입별 구매총액"
    yAxisName = "구매총액"
end if

Dim pTypeName : pTypeName = getpTypeName(vpType)
Dim mxChartSeries : mxChartSeries = 5

%>
<html>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>
<script>
function jsPopCal(sName){
	var winCal;

	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}
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
FusionCharts.ready(function () {
    var vstrChart1 = new FusionCharts({
        type: 'msline',
        renderAt: 'chart-container1',
        width: '1100',
        height: '500',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "<%=iChartCaption%>",
                "subCaption": "<%=iChartSubCaption%>",
                "xAxisName": "<%=ixAxisName%>",
                "yAxisName": "<%=yAxisName%>",
                "theme": "fint",
                "showValues": "1",
                //Setting automatic calculation of div lines to off
                "adjustDiv": "0",
                //Manually defining y-axis lower and upper limit
                //"yAxisMaxvalue": "35000",	//y축 맥스값
                //"yAxisMinValue": "5000",		//y축 민값
                //Setting number of divisional lines to 9
                //"numDivLines": "9"				//0~맥스 사이 표시되어지는 수치갯수
                "anchorBgHoverColor": "#96d7fa",
                "anchorBorderHoverThickness" : "4",
                "anchorHoverRadius":"7"
            },
            "categories": [
                {
                    "category": [
						<%
						dim precate
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
            <% if isArray(vArr2) then %>
            <% for j=0 To UBound(vArr2,2) %>
                <% if j<mxChartSeries then %>
                {
                    "seriesname": "<%=vArr2(0,j)%>",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
							    if (vArr1(1,i)=vArr2(0,j)) then
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(iDataSetPosArr(0),i)&"""" & vbCrLf
								Response.Write "}"
								If i <> UBound(vArr1,2) Then
									Response.Write ","
								End If
								Response.Write vbCrLf
							    end if
							Next
						End If
						%>
                    ]
                }
                <% if j<UBound(vArr2,2) then %>,<% end if %>
                <% end if %>
            <% next %> 
            <% end if %>
            ]
        }
    }).render();


});
}//]]>
</script>
<body>
<form name="frm1" method="post" >
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#F4F4F4">
    <td width="50" rowspan="2" bgcolor="#EEEEEE">검색<br>조건</td>
	<td align="left">
	조회날짜 : 
		<input type="text" name="sdate" id="sdate" value="<%=vSDate%>" onClick="jsPopCal('sdate');" style="cursor:pointer;" size="10" maxlength="10" readonly>&nbsp;~&nbsp;
		<input type="text" name="edate" id="edate" value="<%=vEDate%>" onClick="jsPopCal('edate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
    &nbsp;&nbsp;
    
    채널 : <% call drawConversionChannelSelectBox("channel",vChannel) %>
    &nbsp;&nbsp;
    
    전환타입 : <% call drawConversionTypeSelectBox("ptype",vpType) %>
    &nbsp;&nbsp;
    
    <% if (vpType="gaparam") then %>
        <% call drawConversionTypeGroupSelectBox("puptype",vpUpType, vpType) %>
        &nbsp;&nbsp;
    <% end if %> 
    <input type="radio" name="ordtype" value="C" <%=CHKIIF(vOrdType="C","checked","") %> >주문건수순
    <input type="radio" name="ordtype" value="S" <%=CHKIIF(vOrdType="S","checked","") %> >구매총액순
    <input type="radio" name="ordtype" value="G" <%=CHKIIF(vOrdType="G","checked","") %> >매출수익순
    &nbsp;&nbsp;
    <%= pTypeName %> : <input type="text" name="pvalue" size="20" value="<%=vpValue%>">
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="goSearch(document.frm1);">
	</td>
</tr>

</table>
</form>
<br />
<table cellpadding="0" cellspacing="0" border="0" class="a" align="center">
<tr bgcolor="#FFFFFF">
    <% if isArray(vArr2) then %>
    <td width="500">
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4">
            <td><%=pTypeName%></td>
            <td></td>
            <td>주문건수</td>
            <td>구매총액</td>
            <td>매출수익</td>
        </tr>
        <% For i = 0 To UBound(vArr2,2) %>
        <tr bgcolor="#FFFFFF" align="right">
            <td align="left"><%=vArr2(0,i)%></td>
            <td align="left"><%=vArr2(11,i)%></td>
            <td><%=FormatNumber(vArr2(1,i),0)%></td>
            <td><%=FormatNumber(vArr2(5,i),0)%></td>
            <td><%=FormatNumber(vArr2(9,i),0)%></td>
        </tr>
        <% next %>
        </table>
    </td>
    <% end if %>
	<td>
		<div id="chart-container1">FusionCharts will render here</div>
	</td>
</tr>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->