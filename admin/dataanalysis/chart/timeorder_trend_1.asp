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
Dim oChart, vArr1, i, j, k, ii
Dim iDataSetPosArrCC(), iDataSetHeadArrCC()
Dim iDataSetPosArrCh(), iDataSetHeadArrCh(), titleArrCh()
Dim vSDate, vEDate, vChannel, sTp, pTp

vSDate = requestCheckvar(request("startdate"),10)
vEDate = requestCheckvar(request("enddate"),10)
vChannel = requestCheckvar(request("channel"),10)
sTp  = requestCheckvar(request("sTp"),10)
pTp  = requestCheckvar(request("pTp"),10)

if (sTp="") then sTp="2" ''건수(1) , 금액(2)
if (pTp="") then pTp="py" ''전년(py) , 전월(pm), 전주(pw), 전일(pd)
    
If vSDate = "" Then
	vSDate = LEFT(dateadd("d",-1,Date()),10)
End If

If vEDate = "" Then
	vEDate = LEFT(date(),10)
End If

SET oChart = new CChart
	oChart.FRectSDate = vSDate
	oChart.FRectEDate = vEDate
	oChart.FRectChannel = vChannel
	oChart.FRectCompTerms = pTp
	vArr1 = oChart.fnTimeMeachul_trend
	
SET oChart = nothing

Dim iChartCaption : iChartCaption = "시간별 주문건수"
Dim iChartSubCaption : iChartSubCaption = ""
Dim ixAxisName : ixAxisName = "" ''날짜
Dim yAxisName : yAxisName = "주문건수"
Dim iDataSetPosArr : iDataSetPosArr = Array(2,4) 
Dim iDataSetHeadArr : iDataSetHeadArr = Array("주문건수-금년","주문건수-이전") 

Dim iDataSetPosArr2 : iDataSetPosArr2 = Array(0,2,6,4,8,10) 
Dim iDataSetHeadArr2 : iDataSetHeadArr2 = Array("날짜","주문건수-금년","누적예상","주문건수-이전","누적-이전","이전비<br>(%)")   '',"날짜-이전"

redim titleArrCh(2): titleArrCh(0)="WEB":titleArrCh(1)="MOB":titleArrCh(2)="APP"
redim iDataSetPosArrCC(2), iDataSetHeadArrCC(2)
redim iDataSetPosArrCh(2), iDataSetHeadArrCh(2)

iDataSetPosArrCC(0) = Array(12,14) 
iDataSetHeadArrCC(0) = Array("주문건수-금년","주문건수-이전") 
iDataSetPosArrCC(1) = Array(22,24) 
iDataSetHeadArrCC(1) = Array("주문건수-금년","주문건수-이전") 
iDataSetPosArrCC(2) = Array(32,34) 
iDataSetHeadArrCC(2) = Array("주문건수-금년","주문건수-이전") 

iDataSetPosArrCh(0) = Array(0,12,16,14,18,20) 
iDataSetHeadArrCh(0) = Array("날짜","주문건수-금년","누적예상","주문건수-이전","누적-이전","이전비<br>(%)")
iDataSetPosArrCh(1) = Array(0,22,26,24,28,30) 
iDataSetHeadArrCh(1) = Array("날짜","주문건수-금년","누적예상","주문건수-이전","누적-이전","이전비<br>(%)")
iDataSetPosArrCh(2) = Array(0,32,36,34,38,40) 
iDataSetHeadArrCh(2) = Array("날짜","주문건수-금년","누적예상","주문건수-이전","누적-이전","이전비<br>(%)")

if (sTp="2") then
    iDataSetPosArr = Array(3,5) '',8
    iDataSetHeadArr = Array("구매총액-금년","구매총액-이전")
    
    iDataSetPosArr2 = Array(0,3,7,5,9,11) '',8
    iDataSetHeadArr2 = Array("날짜","구매총액-금년","누적예상","구매총액-이전","누적-이전","이전비<br>(%)") '',"날짜-이전"
    
    iDataSetPosArrCC(0) = Array(13,15) 
    iDataSetHeadArrCC(0) = Array("구매총액-금년","구매총액-이전")
    iDataSetPosArrCC(1) = Array(23,25) 
    iDataSetHeadArrCC(1) = Array("주문건수-금년","주문건수-이전") 
    iDataSetPosArrCC(2) = Array(33,35) 
    iDataSetHeadArrCC(2) = Array("주문건수-금년","주문건수-이전") 

    iDataSetPosArrCh(0) = Array(0,13,17,15,19,21) '',8
    iDataSetHeadArrCh(0) = Array("날짜","구매총액-금년","누적예상","구매총액-이전","누적-이전","이전비<br>(%)") 
    iDataSetPosArrCh(1) = Array(0,23,27,25,29,31) '',8
    iDataSetHeadArrCh(1) = Array("날짜","구매총액-금년","누적예상","구매총액-이전","누적-이전","이전비<br>(%)") 
    iDataSetPosArrCh(2) = Array(0,33,37,35,39,41) '',8
    iDataSetHeadArrCh(2) = Array("날짜","구매총액-금년","누적예상","구매총액-이전","누적-이전","이전비<br>(%)") 
    
    iChartCaption = "시간별 구매총액"
    yAxisName = "구매총액"
end if

dim SumArr()
redim SumArr(UBound(iDataSetPosArr2))
dim SumArrType : SumArrType = Array(9,0,1,0,1,9) 

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
                //"anchorHoverRadius":"7",
                "anchorRadius":"1"
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
								if (j=0) then 
								    Response.Write ",""color"": ""#FF0000""" & vbCrLf
								end if
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

    <% if (vChannel="")  then %>
    <% for ii=LBound(iDataSetPosArrCh) to UBound(iDataSetPosArrCh) %>
    var vstrChartA<%=ii%> = new FusionCharts({
        type: 'msline',
        renderAt: 'chart-containerA<%=ii%>',
        width: '1200',
        height: '500',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "<%=iChartCaption%> -<%=titleArrCh(ii)%>",
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
                //"anchorHoverRadius":"7",
                "anchorRadius":"1"
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
            <% for j=LBound(iDataSetPosArrCC(ii)) to Ubound(iDataSetPosArrCC(ii)) %>
                {
                    "seriesname": "<%=iDataSetPosArrCC(ii)(j)%>",
                    "data": [
						<%
						If isArray(vArr1) Then
						   
							For i = 0 To UBound(vArr1,2)
						
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(iDataSetPosArrCC(ii)(j),i)&"""" & vbCrLf
								if (j=0) then 
								    Response.Write ",""color"": ""#FF0000""" & vbCrLf
								end if
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
                <% if j<Ubound(iDataSetPosArrCC(ii)) then %>,<% end if %>
            <% next %> 
             
            ]
        }
    }).render();
    <% next %>
    <% end if %>
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
    <input type="radio" name="sTp" value="1" <%=CHKIIF(sTp="1","checked","") %> >주문건수
    <input type="radio" name="sTp" value="2" <%=CHKIIF(sTp="2","checked","") %> >구매총액
    &nbsp;&nbsp;
    |
    &nbsp;&nbsp;
    <input type="radio" name="pTp" value="py" <%=CHKIIF(pTp="py","checked","") %> >전년동요일(-52 weeks)
    <input type="radio" name="pTp" value="pm" <%=CHKIIF(pTp="pm","checked","") %> >전월(-4 weeks)
    <input type="radio" name="pTp" value="pw" <%=CHKIIF(pTp="pw","checked","") %> >전주(-7 days)
    <input type="radio" name="pTp" value="pd" <%=CHKIIF(pTp="pd","checked","") %> >전일(-1 day)
    
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
* 전년 동요일 기준 (d-364일)
<p>
<table cellpadding="0" cellspacing="0" border="0" class="a" align="center" width="100%">
<tr bgcolor="#FFFFFF">
    <% if isArray(vArr1) then %>
    <td width="560">
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4">
            <% for k=Lbound(iDataSetHeadArr2) to Ubound(iDataSetHeadArr2) %>
            <td><%=iDataSetHeadArr2(k)%></td>
            <% next %>
        </tr> 
        <% For i = 0 To UBound(vArr1,2) %>
        <% if LEFT(vArr1(0,i),10)=vEDate then %>
        <tr bgcolor="#FFFFFF" align="right">
           
            <% for k=Lbound(iDataSetPosArr2) to Ubound(iDataSetPosArr2) %>
            <% 
                if SumArrType(k)=0 then
                    SumArr(k)=SumArr(k)+CDBL(vArr1(iDataSetPosArr2(k),i)) 
                end if
            %>
            <td>
                <% 
                if SumArrType(k)<=1 then
                    response.write FormatNumber(vArr1(iDataSetPosArr2(k),i),0)
                else
                    response.write vArr1(iDataSetPosArr2(k),i)
                end if
                %>
            </td>
            <% next %>
        </tr>
        <% end if %>
        <% next %>
        <tr bgcolor="#FFFFFF" align="right">
            <% for k=Lbound(SumArr) to Ubound(SumArr) %>
            <td>
                <% if SumArrType(k)=0 then %>
                <%=FormatNumber(SumArr(k),0)%>
                <% else %>
            
                <% end if %>
            </td>
            <% SumArr(k) =0 %>
            <% next %>
        </tr>
        </table>
    </td>
    <% end if %>
	<td>
		<div id="chart-container1">FusionCharts will render here</div>
	</td>
</tr>

<% if (vChannel="") then %>
<% for ii=LBound(titleArrCh) to UBound(titleArrCh) %>
<tr bgcolor="#FFFFFF" >
    <% if isArray(vArr1) then %>
    <td width="560" style="padding-top:20px">
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4">
            <% for k=Lbound(iDataSetHeadArrCh(ii)) to Ubound(iDataSetHeadArrCh(ii)) %>
            <td><%=iDataSetHeadArrCh(ii)(k)%></td>
            <% next %>
        </tr> 
        <% For i = 0 To UBound(vArr1,2) %>
        <% if LEFT(vArr1(0,i),10)=vEDate then %>
        <tr bgcolor="#FFFFFF" align="right">
           
            <% for k=Lbound(iDataSetPosArrCh(ii)) to Ubound(iDataSetPosArrCh(ii)) %>
            <% 
                if SumArrType(k)=0 then
                    SumArr(k)=SumArr(k)+CDBL(vArr1(iDataSetPosArrCh(ii)(k),i)) 
                end if
            %>
            <td>
                <% 
                if SumArrType(k)<=1 then
                    response.write FormatNumber(vArr1(iDataSetPosArrCh(ii)(k),i),0)
                else
                    response.write vArr1(iDataSetPosArrCh(ii)(k),i)
                end if
                %>
            </td>
            <% next %>
        </tr>
        <% end if %>
        <% next %>
        <tr bgcolor="#FFFFFF" align="right">
            <% for k=Lbound(SumArr) to Ubound(SumArr) %>
            <td>
                <% if SumArrType(k)=0 then %>
                <%=FormatNumber(SumArr(k),0)%>
                <% else %>
            
                <% end if %>
            </td>
            <% SumArr(k) =0 %>
            <% next %>
        </tr>
        </table>
    </td>
    <% end if %>
	<td>
		<div id="chart-containerA<%=ii%>">FusionCharts will render here</div>
	</td>
</tr>
<% next %>
<% end if %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->