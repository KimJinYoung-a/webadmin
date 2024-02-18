<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/maechul/fusionchart/maechul_class.asp" -->
<%
	Dim cNaIt, vArr1, vArr2, i, j, vItemID, vSDate, vEDate, vItemName, vTotalCount1, vTotalCount2, vPrePrice2, vPrice2, vMaxPrice2
	Dim vArrJ1D, vJust1Day, k, vArr3, m, vTotalCount3
	vItemID = requestCheckvar(request("itemid"),10)
	vSDate = requestCheckvar(request("sdate"),10)
	vEDate = requestCheckvar(request("edate"),10)
	
	If vItemID <> "" Then
		If Not isNumeric(vItemID) Then
			Response.Write "<script>alert('잘못된 상품코드입니다.');location.href='/admin/maechul/item_graph.asp';</script>"
			Response.End
		End If
	End If
	
	If vSDate = "" Then
		vSDate = FormatDate(DateAdd("d",-14,now()),"0000-00-00")
	End If
	
	If vEDate = "" Then
		vEDate = FormatDate(now(),"0000-00-00")
	End If
	'response.write vItemID & "<br>"
	'response.write vSDate & "<br>"
	'response.write vEDate & "<br>"
	'response.end
	
	Set cNaIt = new Cmaechul_list
	If vItemID <> "" Then
		cNaIt.FRectItemID = vItemID
		cNaIt.FRectSDate = vSDate
		cNaIt.FRectEDate = vEDate
		vArr1 = cNaIt.fnNaverMaechulByItem		'### 네이버랭킹
		vItemName = cNaIt.FNaItemName
		vTotalCount1 = cNaIt.FTotalCount
		
		vArr2 = cNaIt.fnItemSellcashHistory	'### 일별 가격 변동
		vTotalCount2 = cNaIt.FTotalCount
		vArrJ1D = cNaIt.FArrJust1Day			'### 저스트원데이
		
		vArr3 = cNaIt.fnCouponMasterList		'### 쿠폰마스터 리스트
		vTotalCount3 = cNaIt.FTotalCount
		
		
		vPrePrice2 = 0
		vPrice2 = 0
		vMaxPrice2 = 0
		
		
		If isArray(vArr2) Then
			For j = 0 To UBound(vArr2,2)
				If vMaxPrice2 < vArr2(2,j) Then
					vMaxPrice2 = vArr2(2,j)
				End If
			Next
		End If
	End If
	Set cNaIt = Nothing
	
	If isArray(vArrJ1D) Then
		For k = 0 To UBound(vArrJ1D,2)
			vJust1Day = vJust1Day & vArrJ1D(0,k)
			If k <> UBound(vArrJ1D,2) Then
				vJust1Day = vJust1Day & ","
			End If
		Next
	End If

	'response.write vJust1Day
%>
<html>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>
<script>
function jsPopCal(sName){
	var winCal;

	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function goSearch(){
	if(frm1.itemid.value == ""){
		alert("상품코드를 입력하세요.");
		frm1.itemid.focus();
		return;
	}
	if(isNaN(frm1.itemid.value)){
		alert("상품코드를 숫자로만 입력하세요.");
		frm1.itemid.value = "";
		frm1.itemid.focus();
		return;
	}
	
	frm1.submit();
}
</script>
<% If vItemID <> "" Then %>
<script type='text/javascript'>//<![CDATA[
window.onload=function(){
FusionCharts.ready(function () {
    var vstrChart1 = new FusionCharts({
        type: 'msline',
        renderAt: 'chart-container1',
        width: '1200',
        height: '400',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "<%=vItemName%>",
                "subCaption": "네이버 & 다음 랭킹에 따른 판매",
                //"xAxisName": "Day",
                //"yAxisName": "No. of Visitors",
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
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
								Response.Write "{" & vbCrLf
								Response.Write """label"": """&vArr1(1,i)&"("&vArr1(2,i)&")""" & vbCrLf
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
                {
                    "seriesname": "My Rank",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(3,i)&"""" & vbCrLf
								Response.Write "}"
								If i <> UBound(vArr1,2) Then
									Response.Write ","
								End If
								Response.Write vbCrLf
							Next
						End If
						%>
                    ]
                }, 
                {
                    "seriesname": "주문횟수",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(4,i)&"""" & vbCrLf
								Response.Write "}"
								If i <> UBound(vArr1,2) Then
									Response.Write ","
								End If
								Response.Write vbCrLf
							Next
						End If
						%>
                    ]
                },
                {
                    "seriesname": "NP_DAUM_sellCNT",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(5,i)&"""" & vbCrLf
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
            ]
        }
    }).render();

    var vstrChart2 = new FusionCharts({
        type: 'msline',
        renderAt: 'chart-container2',
        width: '1200',
        height: '250',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                //"caption": "<%=vItemName%>",
                "subCaption": "일별 가격 변동 Log(동일 날짜 여러번 변경일 경우 최종 가격 기준)",
                //"xAxisName": "Day",
                //"yAxisName": "No. of Visitors",
                "theme": "fint",
                "showValues": "1",
                //Setting automatic calculation of div lines to off
                "adjustDiv": "0",
                //Manually defining y-axis lower and upper limit
                "yAxisMaxvalue": "<%=vMaxPrice2*2%>",
                "yAxisMinValue": "0",
                //Setting number of divisional lines to 9
                "numDivLines": "0",
                "anchorBgHoverColor": "#96d7fa",
                "anchorBorderHoverThickness" : "4",
                "anchorHoverRadius":"7",
                "formatNumberScale":"0",         // 천단위자동 변환 여부; 0:안함, 1:자동변환
                "formatNumber":"1"               // 천단위 쉼표 표시여부
            },            
            "categories": [
                {
                    "category": [
						<%
						If isArray(vArr2) Then
							For j = 0 To UBound(vArr2,2)
								Response.Write "{" & vbCrLf
								Response.Write """label"": """&vArr2(0,j)&"("&vArr2(1,j)&")""" & vbCrLf
								Response.Write "}"
								If j <> UBound(vArr2,2) Then
									Response.Write ","
								End If
								
								If InStr(vJust1Day,vArr2(0,j)) > 0 Then
		                        Response.Write "{" & vbCrLf
		                        Response.Write "	""vline"": ""true""," & vbCrLf
		                        Response.Write "	""lineposition"": ""0""," & vbCrLf
		                        Response.Write "	""color"": ""#6baa01""," & vbCrLf
		                        Response.Write "	""labelHAlign"": ""center""," & vbCrLf
		                        Response.Write "	""labelPosition"": ""0""," & vbCrLf
		                        Response.Write "	""label"": ""Just 1 Day""," & vbCrLf
		                        Response.Write "	""dashed"": ""1""" & vbCrLf
		                        Response.Write "}"
									If j <> UBound(vArr2,2) Then
										Response.Write ","
									End If
								End If
								Response.Write vbCrLf
							Next
						End If
						%>
                    ]
                }
            ],            
            "dataset": [
                {
                    "seriesname": "판매가",
                    "data": [
						<%
						If isArray(vArr2) Then
							For j = 0 To UBound(vArr2,2)
								If vArr2(2,j) = -1 Then
									vPrice2 = vPrePrice2
								Else
									vPrice2 = vArr2(2,j)
								End IF
								
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vPrice2&"""" & vbCrLf
								Response.Write "}"
								If j <> UBound(vArr2,2) Then
									Response.Write ","
								End If
								Response.Write vbCrLf
								
								If vArr2(2,j) <> -1 Then
									vPrePrice2 = vArr2(2,j)
								End IF
							Next
						End If
						%>
                    ]
                }
            ]
        }
    }).render();
    
    var topStores = new FusionCharts({
        type: 'msline',
        renderAt: 'chart-container3',
        width: '1200',
        height: '300',
        dataFormat: 'json',
        dataSource: {
		    "chart": {
		        "caption": "날짜별 보너스 쿠폰 시작일",
		        "showvalues": "0",
		        "anchorradius": "7",
		        "slantlabels": "1",
		        "linethickness": "5",
		        "connectnulldata": "0",
		        "xtlabelmanagement": "0",
		        "showborder": "0",
                "formatNumberScale":"0",
                "formatNumber":"1"
		    },
		    "categories": [
		        {
		            "category": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
								Response.Write "{" & vbCrLf
								Response.Write """label"": """&vArr1(1,i)&"("&vArr1(2,i)&")""" & vbCrLf
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
					<%
					If isArray(vArr3) Then
						For m = 0 To UBound(vArr3,2)
							Response.Write "					{" & vbCrLf
							Response.Write "					""seriesname"": """&vArr3(1,m)&"""," & vbCrLf
							Response.Write "					""data"": [" & vbCrLf
							
							For i = 0 To UBound(vArr1,2)
								'If vArr1(1,i) >= vArr3(2,m) and vArr1(1,i) <= vArr3(3,m) Then
								If vArr1(1,i) = vArr3(2,m) Then
									Response.Write "					{" & vbCrLf
									Response.Write "					""value"": """&vArr3(0,m)&"""" & vbCrLf
									Response.Write "					}"
								Else
									Response.Write "					{" & vbCrLf
									Response.Write "					""value"": """"" & vbCrLf
									Response.Write "					}"
								End If
								
								If i <> UBound(vArr1,2) Then
									Response.Write ","
								End If
								Response.Write vbCrLf
							Next
							Response.Write "					]" & vbCrLf
							If m <> UBound(vArr3,2) Then
								Response.Write "					}," & vbCrLf
							End If
						Next
					End If
					%>
					}
		    ]
        }
    })
    .render();

});
}//]]>
</script>
<% End If %>
<body>
<form name="frm1" method="post" action="item_graph1.asp">
<table width="100%" class="a">
<tr>
	<td height="30" align="center">
		상품코드 : 
		<input type="text" name="itemid" value="<%=vItemID%>" size="10" maxlength="10">&nbsp;&nbsp;
		조회날짜 : 
		<input type="text" name="sdate" value="<%=vSDate%>" onClick="jsPopCal('sdate');" style="cursor:pointer;" size="10" maxlength="10" readonly>&nbsp;~&nbsp;
		<input type="text" name="edate" value="<%=vEDate%>" onClick="jsPopCal('edate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
		&nbsp;&nbsp;&nbsp;
		<input type="button" value="조  회" onClick="goSearch();">
		<% If vItemID <> "" Then %>
		[<a href='http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=vItemID%>' target='_blank'>상품상세보기</a>]
		<% End If %>
	</td>
</tr>
</table>
</form>
<br />
<% If vItemID <> "" Then %>
<table cellpadding="0" cellspacing="0" border="0" class="a" align="center">
<tr bgcolor="#FFFFFF">
	<td>
		<div id="chart-container1">FusionCharts will render here</div>
		<br />
		<div id="chart-container2">FusionCharts will render here</div>
		<br />
		<div id="chart-container3" style="text-align:center;">FusionCharts will render here</div>
	</td>
</tr>
</table>
<% End If %>
</body>
</html>
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->