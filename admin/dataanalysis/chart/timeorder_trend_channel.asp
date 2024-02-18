<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/dataanalysis/chart/chartCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim oChart, vArr1, i, j, k
Dim vSDate, vEDate, vChannel, sTp, dTp

vSDate = requestCheckvar(request("startdate"),10)
vEDate = requestCheckvar(request("enddate"),10)
vChannel = requestCheckvar(request("channel"),10)
sTp  = requestCheckvar(request("sTp"),10)
dTp  = requestCheckvar(request("dTp"),10)

if (sTp="") then sTp="2" ''건수(1) , 금액(2)
if (dTp="") then dTp="d" ''시간(h) , 일간(d) , 주간 (w), 월간 (m)
     
If vSDate = "" Then
	vSDate = LEFT(dateadd("d",-14,Date()),10)
End If

If vEDate = "" Then
	vEDate = LEFT(date(),10)
End If

SET oChart = new CChart
	oChart.FRectSDate = vSDate
	oChart.FRectEDate = vEDate
	oChart.FRectChannel = vChannel
	oChart.FRectGroupType = dTp
	vArr1 = oChart.fnTimeMeachul_trend_channel
SET oChart = nothing

Dim iChartCaption : iChartCaption = "채널별 주문건수"
Dim iChartSubCaption : iChartSubCaption = ""
Dim ixAxisName : ixAxisName = "" ''날짜
Dim yAxisName : yAxisName = "주문건수"
Dim iDataSetPosArr : iDataSetPosArr = Array(4,6,8,10) 
Dim iDataSetHeadArr : iDataSetHeadArr = Array("주문건수-Pc","주문건수-Mob","주문건수-App","주문건수-Out") 

''Dim iDataSetPosArr2 : iDataSetPosArr2 = Array(2,6,4,10) 
''Dim iDataSetHeadArr2 : iDataSetHeadArr2 = Array("주문건수-금년","누적예상","주문건수-전년","전년비<br>(%)") 

if (sTp="2") then
    iDataSetPosArr = Array(5,7,9,11) '',8
    iDataSetHeadArr = Array("구매총액-Pc","구매총액-Mob","구매총액-App","구매총액-Out")
    
    ''iDataSetPosArr2 = Array(3,7,5,11) '',8
    ''iDataSetHeadArr2 = Array("구매총액-금년","누적예상","구매총액-전년","전년비<br>(%)")
    
    iChartCaption = "채널별 구매총액"
    yAxisName = "구매총액"
end if

'dim SumArr()
'redim SumArr(UBound(iDataSetPosArr2))
'dim SumArrType : SumArrType = Array(0,9,0,9)

%>

<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>


<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>

<script type="text/javascript">
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
FusionCharts.ready(function () {
    var vstrChart1 = new FusionCharts({
        type: 'msline',
        renderAt: 'chart-container1',
        width: '1200',
        height: '500',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "<%=iChartCaption%>",
                "subCaption": "<%=iChartSubCaption%>",
                "xAxisName": "<%=ixAxisName%>",
                "yAxisName": "<%=yAxisName%>",
                "theme": "fint",
                "showValues": "0",
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
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
								Response.Write "{" & vbCrLf
								Response.Write """label"": """&vArr1(0,i)&"""" & vbCrLf
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
            ],            
            "dataset": [
            <% for j=LBound(iDataSetPosArr) to Ubound(iDataSetPosArr) %>
                {
                    "seriesname": "<%=iDataSetHeadArr(j)%>",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(iDataSetPosArr(j),i)&"""" & vbCrLf
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
                <% if j<Ubound(iDataSetPosArr) then %>,<% end if %>
            <% next %> 
            
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
	조회날짜(주문일) : 
	    
	    <input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
		~<input id='enddate' name='enddate' value='<%= vEDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
			
	    
    &nbsp;&nbsp;
    
    채널 :
    <select name="channel" >
        <option value="" <%=CHKIIF(vChannel="","selected","")%>>ALL</option>
        <option value="pc" <%=CHKIIF(vChannel="pc","selected","")%>>WEB</option>
        <option value="mw" <%=CHKIIF(vChannel="mw","selected","")%>>MOB</option>
        <option value="app" <%=CHKIIF(vChannel="app","selected","")%>>APP</option>
    </select>
    
    &nbsp;&nbsp;
    <input type="radio" name="dTp" value="h" <%=CHKIIF(dTp="h","checked","") %> >시간별
    <input type="radio" name="dTp" value="d" <%=CHKIIF(dTp="d","checked","") %> >일별
    <input type="radio" name="dTp" value="m" <%=CHKIIF(dTp="m","checked","") %> >월별
    
    &nbsp;&nbsp;
    <input type="radio" name="sTp" value="1" <%=CHKIIF(sTp="1","checked","") %> >주문건수
    <input type="radio" name="sTp" value="2" <%=CHKIIF(sTp="2","checked","") %> >구매총액
    
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="goSearch(document.frm1);">
	</td>
</tr>

</table>
</form>
<br />
* 약 1시간 지연데이터
* 반품 교환건은 포함되지 않음
* 제휴,해외,3pl은 포함되지 않음
* 무통장 결제 이전 주문 포함됨(차후 취소될 수 있음)
<p>
<table cellpadding="0" cellspacing="0" border="0" class="a" align="center" >
<tr bgcolor="#FFFFFF">
	<td>
		<div id="chart-container1">FusionCharts will render here</div>
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->