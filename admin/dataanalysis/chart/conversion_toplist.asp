<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/dataanalysis/chart/chartCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim oChart, vArr1, vArr2, i, j
Dim vArrR, vArrB, vArrE, vArrC
Dim vArrW, vArrM, vArrA
Dim vSDate, vEDate, vChannel, vpType, vpValue, vOrdType, sTp, vpUpType
Dim dPageSize : dPageSize = 100
Dim crtPageSize :  crtPageSize = 5

vSDate = requestCheckvar(request("startdate"),10)
vEDate = requestCheckvar(request("enddate"),10)
vChannel = requestCheckvar(request("channel"),10)
vpType = requestCheckvar(request("ptype"),32)
vpValue = requestCheckvar(request("pvalue"),64)
vOrdType = requestCheckvar(request("ordtype"),32)
vpUpType = requestCheckvar(request("puptype"),32)

''if (vpType="") then vpType="pRtr"
if (vpType<>"") and (vChannel<>"") then dPageSize=20
if (vOrdType="") then vOrdType="C" ''건수(C) , 금액(S), 수익(G)

if (sTp="") then sTp="1" 
    
If vSDate = "" Then
	vSDate = dateadd("d",-1,Date())
End If

If vEDate = "" Then
	vEDate = date()
End If

SET oChart = new CChart
	oChart.FRectSDate = vSDate
	oChart.FRectEDate = vEDate
	oChart.FRectChannel = vChannel
	oChart.FRectPType = vpType
	oChart.FPageSize = dPageSize
	oChart.FRectOrderType = vOrdType
	oChart.FRectPValue = vpValue
	oChart.FRectUPTypeValue = vpUpType
	
	if (vpType="") then
	    oChart.FRectPType =  "pRtr"
	    vArrR = oChart.fnConversionTopByType
	    
	    oChart.FRectPType =  "pBtr"
	    vArrB = oChart.fnConversionTopByType
	    
	    oChart.FRectPType =  "pCtr"
	    vArrC = oChart.fnConversionTopByType
	    
	    oChart.FRectPType =  "pEtr"
	    vArrE = oChart.fnConversionTopByType
	else
	    vArr2 = oChart.fnConversionTopByType
	    
	    
	    if (vChannel="") then
	        oChart.FRectPType =  vpType
	        oChart.FPageSize = dPageSize
	        oChart.FRectChannel = "pc"
	        vArrW = oChart.fnConversionTopByType
	        
	        oChart.FRectPType =  vpType
	        oChart.FPageSize = dPageSize
	        oChart.FRectChannel = "mw"
	        vArrM = oChart.fnConversionTopByType
	        
	        oChart.FRectPType =  vpType
	        oChart.FPageSize = dPageSize
	        oChart.FRectChannel = "app"
	        vArrA = oChart.fnConversionTopByType
	    else
	        oChart.FPageSize = 5
	        vArr1 = oChart.fnConversionTopByType_Trend    
	    end if
	end if
	
	
	
	
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

Dim sumOrdCnt,sumOrdSum,SumGainSum
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

function popConvDetail(iptype,ipvalue,ichannel){
    var frm = document.frm1;
    
    
    var iURL = "/admin/dataanalysis/chart/conversion_type_detail.asp?startdate="+frm.startdate.value+"&enddate="+frm.enddate.value
    iURL=iURL+"&channel="+ichannel+"&ptype="+iptype+"&pvalue="+ipvalue+"&ordtype="+frm.ordtype.value;
    
    var popwin = window.open(iURL,"popConvDetail","width=1600,height=600,scrollbars=yes,resizable=yes")
    popwin.focus();
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
								if (j=0) then 
								    Response.Write ",""color"": ""#FF0000""" & vbCrLf
								end if
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
	    <input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
		~<input id='enddate' name='enddate' value='<%= vEDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
			
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
    <% if (vpType<>"") then %>
    <%= pTypeName %> : <input type="text" name="pvalue" size="20" value="<%=vpValue%>">
    <% end if %>
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="검색" onClick="goSearch(document.frm1);">
	</td>
</tr>

</table>
</form>
<br />
<% if (vpType="") then %>
<table cellpadding="0" cellspacing="0" border="0" class="a" align="left">
<tr bgcolor="#FFFFFF">
    <td width="18%" valign="top">
    <% if isArray(vArrR) then %>
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <% sumOrdCnt =0:sumOrdSum =0:SumGainSum =0 %>
        <tr bgcolor="#F4F4F4">
            <td width="100"><%= getpTypeName("pRtr")%></td>
            <td></td>
            <td width="40">주문<br>건수</td>
            <td width="60">구매총액</td>
            <td width="60">매출수익</td>
        </tr>
        <% For i = 0 To UBound(vArrR,2) %>
        <tr bgcolor="#FFFFFF" align="right">
            <td align="left" onClick="popConvDetail('pRtr','<%=vArrR(0,i)%>','<%=vChannel%>')" style="cursor:pointer"><%=vArrR(0,i)%></td>
            <td align="left"><%=vArrR(11,i)%></td>
            <td><%=FormatNumber(vArrR(1,i),0)%></td>
            <td><%=FormatNumber(vArrR(5,i),0)%></td>
            <td><%=FormatNumber(vArrR(9,i),0)%></td>
        </tr>
        <%
        sumOrdCnt=sumOrdCnt+vArrR(1,i)
        sumOrdSum=sumOrdSum+vArrR(5,i)
        SumGainSum=SumGainSum+vArrR(9,i)
        %>
        <% next %>
        <thead>
        <tr bgcolor="#FFFFFF" align="right">
            <th></th>
            <th></th>
            <th><%=FormatNumber(sumOrdCnt,0)%></th>
            <th><%=FormatNumber(sumOrdSum,0)%></th>
            <th><%=FormatNumber(SumGainSum,0)%></th>
        </tr>
        </thead>
        </table>
    <% end if %>
    </td>
	<td width="23%" valign="top">
	<% if isArray(vArrB) then %>
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <% sumOrdCnt =0:sumOrdSum =0:SumGainSum =0 %>
        <tr bgcolor="#F4F4F4">
            <td width="100"><%= getpTypeName("pBtr")%></td>
            <td></td>
            <td width="40">주문<br>건수</td>
            <td width="60">구매총액</td>
            <td width="60">매출수익</td>
        </tr>
        <% For i = 0 To UBound(vArrB,2) %>
        <tr bgcolor="#FFFFFF" align="right">
            <td align="left" onClick="popConvDetail('pBtr','<%=vArrB(0,i)%>','<%=vChannel%>')" style="cursor:pointer"><%=vArrB(0,i)%></td>
            <td align="left"><%=vArrB(11,i)%></td>
            <td><%=FormatNumber(vArrB(1,i),0)%></td>
            <td><%=FormatNumber(vArrB(5,i),0)%></td>
            <td><%=FormatNumber(vArrB(9,i),0)%></td>
        </tr>
        <%
        sumOrdCnt=sumOrdCnt+vArrB(1,i)
        sumOrdSum=sumOrdSum+vArrB(5,i)
        SumGainSum=SumGainSum+vArrB(9,i)
        %>
        <% next %>
        <thead>
        <tr bgcolor="#FFFFFF" align="right">
            <th></th>
            <th></th>
            <th><%=FormatNumber(sumOrdCnt,0)%></th>
            <th><%=FormatNumber(sumOrdSum,0)%></th>
            <th><%=FormatNumber(SumGainSum,0)%></th>
        </tr>
        </thead>
        </table>
    <% end if %>    
	</td>
	<td width="27%" valign="top">
	<% if isArray(vArrE) then %>
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <% sumOrdCnt =0:sumOrdSum =0:SumGainSum =0 %>
        <tr bgcolor="#F4F4F4">
            <td width="40"><%= getpTypeName("pEtr")%></td>
            <td></td>
            <td width="40">주문<br>건수</td>
            <td width="60">구매총액</td>
            <td width="60">매출수익</td>
        </tr>
        <% For i = 0 To UBound(vArrE,2) %>
        <tr bgcolor="#FFFFFF" align="right">
            <td align="left" onClick="popConvDetail('pEtr','<%=vArrE(0,i)%>','<%=vChannel%>')" style="cursor:pointer"><%=vArrE(0,i)%></td>
            <td align="left"><%=vArrE(11,i)%></td>
            <td><%=FormatNumber(vArrE(1,i),0)%></td>
            <td><%=FormatNumber(vArrE(5,i),0)%></td>
            <td><%=FormatNumber(vArrE(9,i),0)%></td>
        </tr>
        <%
        sumOrdCnt=sumOrdCnt+vArrE(1,i)
        sumOrdSum=sumOrdSum+vArrE(5,i)
        SumGainSum=SumGainSum+vArrE(9,i)
        %>
        <% next %>
        <thead>
        <tr bgcolor="#FFFFFF" align="right">
            <th></th>
            <th></th>
            <th><%=FormatNumber(sumOrdCnt,0)%></th>
            <th><%=FormatNumber(sumOrdSum,0)%></th>
            <th><%=FormatNumber(SumGainSum,0)%></th>
        </tr>
        </thead>
        </table>
    <% end if %>    
	</td>
	<td width="26%" valign="top">
	<% if isArray(vArrC) then %>
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <% sumOrdCnt =0:sumOrdSum =0:SumGainSum =0 %>
        <tr bgcolor="#F4F4F4">
            <td width="60"><%= getpTypeName("pCtr")%></td>
            <td></td>
            <td width="40">주문<br>건수</td>
            <td width="60">구매총액</td>
            <td width="60">매출수익</td>
        </tr>
        <% For i = 0 To UBound(vArrC,2) %>
        <tr bgcolor="#FFFFFF" align="right">
            <td align="left" onClick="popConvDetail('pCtr','<%=vArrC(0,i)%>','<%=vChannel%>')" style="cursor:pointer"><%=vArrC(0,i)%></td>
            <td align="left"><%=vArrC(11,i)%></td>
            <td><%=FormatNumber(vArrC(1,i),0)%></td>
            <td><%=FormatNumber(vArrC(5,i),0)%></td>
            <td><%=FormatNumber(vArrC(9,i),0)%></td>
        </tr>
        <%
        sumOrdCnt=sumOrdCnt+vArrC(1,i)
        sumOrdSum=sumOrdSum+vArrC(5,i)
        SumGainSum=SumGainSum+vArrC(9,i)
        %>
        <% next %>
        <thead>
        <tr bgcolor="#FFFFFF" align="right">
            <th></th>
            <th></th>
            <th><%=FormatNumber(sumOrdCnt,0)%></th>
            <th><%=FormatNumber(sumOrdSum,0)%></th>
            <th><%=FormatNumber(SumGainSum,0)%></th>
        </tr>
        </thead>
        </table>
    <% end if %>    
	</td>
</tr>
</table>
<% else %>
<table cellpadding="0" cellspacing="0" border="0" class="a" align="center">
<tr bgcolor="#FFFFFF">
    <% if isArray(vArr2) then %>
    <td valign="top">
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <% sumOrdCnt =0:sumOrdSum =0:SumGainSum =0 %>
        <tr bgcolor="#F4F4F4">
            <td><%=pTypeName%>(ALL)</td>
            <td></td>
            <td>주문<br>건수</td>
            <td>구매총액</td>
            <td>매출수익</td>
        </tr>
        <% For i = 0 To UBound(vArr2,2) %>
        <tr bgcolor="#FFFFFF" align="right">
            <td align="left" onClick="popConvDetail('<%=vpType%>','<%=vArr2(0,i)%>','<%=vChannel%>')" style="cursor:pointer"><%=vArr2(0,i)%></td>
            <td align="left"><%=vArr2(11,i)%></td>
            <td><%=FormatNumber(vArr2(1,i),0)%></td>
            <td><%=FormatNumber(vArr2(5,i),0)%></td>
            <td><%=FormatNumber(vArr2(9,i),0)%></td>
        </tr>
        <%
        sumOrdCnt=sumOrdCnt+vArr2(1,i)
        sumOrdSum=sumOrdSum+vArr2(5,i)
        SumGainSum=SumGainSum+vArr2(9,i)
        %>
        <% next %>
        <thead>
        <tr bgcolor="#FFFFFF" align="right">
            <th></th>
            <th></th>
            <th><%=FormatNumber(sumOrdCnt,0)%></th>
            <th><%=FormatNumber(sumOrdSum,0)%></th>
            <th><%=FormatNumber(SumGainSum,0)%></th>
        </tr>
        </thead>
        </table>
    </td>
    <% end if %>
    <% if (vpType<>"") and (vChannel="") then %>
        <% if isArray(vArrW) then %>
        <td width="24%" valign="top">
            <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
            <% sumOrdCnt =0:sumOrdSum =0:SumGainSum =0 %>
            <tr bgcolor="#F4F4F4">
                <td><%=pTypeName%>(WEB)</td>
                <td></td>
                <td>주문<br>건수</td>
                <td>구매총액</td>
                <td>매출수익</td>
            </tr>
            <% For i = 0 To UBound(vArrW,2) %>
            <tr bgcolor="#FFFFFF" align="right">
                <td align="left" onClick="popConvDetail('<%=vpType%>','<%=vArrW(0,i)%>','pc')" style="cursor:pointer"><%=vArrW(0,i)%></td>
                <td align="left"><%=vArrW(11,i)%></td>
                <td><%=FormatNumber(vArrW(1,i),0)%></td>
                <td><%=FormatNumber(vArrW(5,i),0)%></td>
                <td><%=FormatNumber(vArrW(9,i),0)%></td>
            </tr>
            <%
            sumOrdCnt=sumOrdCnt+vArrW(1,i)
            sumOrdSum=sumOrdSum+vArrW(5,i)
            SumGainSum=SumGainSum+vArrW(9,i)
            %>
            <% next %>
            <thead>
            <tr bgcolor="#FFFFFF" align="right">
                <th></th>
                <th></th>
                <th><%=FormatNumber(sumOrdCnt,0)%></th>
                <th><%=FormatNumber(sumOrdSum,0)%></th>
                <th><%=FormatNumber(SumGainSum,0)%></th>
            </tr>
            </thead>
            </table>
        </td>
        <% end if %>
        <% if isArray(vArrM) then %>
        <td width="24%" valign="top">
            <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
            <% sumOrdCnt =0:sumOrdSum =0:SumGainSum =0 %>
            <tr bgcolor="#F4F4F4">
                <td><%=pTypeName%>(MOB)</td>
                <td></td>
                <td>주문<br>건수</td>
                <td>구매총액</td>
                <td>매출수익</td>
            </tr>
            <% For i = 0 To UBound(vArrM,2) %>
            <tr bgcolor="#FFFFFF" align="right">
                <td align="left" onClick="popConvDetail('<%=vpType%>','<%=vArrM(0,i)%>','mw')" style="cursor:pointer"><%=vArrM(0,i)%></td>
                <td align="left"><%=vArrM(11,i)%></td>
                <td><%=FormatNumber(vArrM(1,i),0)%></td>
                <td><%=FormatNumber(vArrM(5,i),0)%></td>
                <td><%=FormatNumber(vArrM(9,i),0)%></td>
            </tr>
            <%
            sumOrdCnt=sumOrdCnt+vArrM(1,i)
            sumOrdSum=sumOrdSum+vArrM(5,i)
            SumGainSum=SumGainSum+vArrM(9,i)
            %>
            <% next %>
            <thead>
            <tr bgcolor="#FFFFFF" align="right">
                <th></th>
                <th></th>
                <th><%=FormatNumber(sumOrdCnt,0)%></th>
                <th><%=FormatNumber(sumOrdSum,0)%></th>
                <th><%=FormatNumber(SumGainSum,0)%></th>
            </tr>
            </thead>
            </table>
        </td>
        <% end if %>
        <% if isArray(vArrA) then %>
        <td width="24%" valign="top">
            <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
            <% sumOrdCnt =0:sumOrdSum =0:SumGainSum =0 %>
            <tr bgcolor="#F4F4F4">
                <td><%=pTypeName%>(APP)</td>
                <td></td>
                <td>주문<br>건수</td>
                <td>구매총액</td>
                <td>매출수익</td>
            </tr>
            <% For i = 0 To UBound(vArrA,2) %>
            <tr bgcolor="#FFFFFF" align="right">
                <td align="left" onClick="popConvDetail('<%=vpType%>','<%=vArrA(0,i)%>','app')" style="cursor:pointer"><%=vArrA(0,i)%></td>
                <td align="left"><%=vArrA(11,i)%></td>
                <td><%=FormatNumber(vArrA(1,i),0)%></td>
                <td><%=FormatNumber(vArrA(5,i),0)%></td>
                <td><%=FormatNumber(vArrA(9,i),0)%></td>
            </tr>
            <%
            sumOrdCnt=sumOrdCnt+vArrA(1,i)
            sumOrdSum=sumOrdSum+vArrA(5,i)
            SumGainSum=SumGainSum+vArrA(9,i)
            %>
            <% next %>
            <thead>
            <tr bgcolor="#FFFFFF" align="right">
                <th></th>
                <th></th>
                <th><%=FormatNumber(sumOrdCnt,0)%></th>
                <th><%=FormatNumber(sumOrdSum,0)%></th>
                <th><%=FormatNumber(SumGainSum,0)%></th>
            </tr>
            </thead>
            </table>
        </td>
        <% end if %>
    <% else %>
	<td>
		<div id="chart-container1">FusionCharts will render here</div>
	</td>
	<% end if %>
</tr>
</table>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->