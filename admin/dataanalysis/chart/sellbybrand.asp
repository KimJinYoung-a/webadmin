<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/admin/dataanalysis/chart/chartCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim mxChartSeries : mxChartSeries = 5
Dim oChart, vArr1, vArr2, vArr3, i, j, k, dispCate
Dim vSDate, vEDate, vChannel, vpValue, vOrdType, sTp', vpUpType, rdsitegrp

vSDate = requestCheckvar(request("startdate"),10)
vEDate = requestCheckvar(request("enddate"),10)
vChannel = requestCheckvar(request("channel"),10)
'rdsitegrp = requestCheckvar(request("rdsitegrp"),32)
vpValue = requestCheckvar(request("pvalue"),64)
vOrdType = requestCheckvar(request("ordtype"),32)
dispCate = requestCheckvar(request("disp"),16)
'vpUpType = requestCheckvar(request("puptype"),32)

if (vOrdType="") then vOrdType="S" ''�Ǽ�(C) , �ݾ�(S), ����(G)

if (sTp="") then sTp="1"

If vSDate = "" Then
    if (vpValue="") then
	    vSDate = dateadd("d",-7,Date())
    else
        vSDate = dateadd("d",-31,Date())
    end if
End If

If vEDate = "" Then
	vEDate = date()
End If

SET oChart = new CChart
	oChart.FRectSDate = vSDate
	oChart.FRectEDate = vEDate
	oChart.FRectChannel = vChannel
'	oChart.FRectRdsiteGrp = rdsitegrp
	oChart.FPageSize = CHKIIF(vpValue<>"",100,100)
	oChart.FRectOrderType = vOrdType
	oChart.FRectPValue = vpValue
	oChart.FRectSubChartTopN = mxChartSeries
    oChart.FRectDispCate = dispCate
'	oChart.FRectUPTypeValue = vpUpType

    if (vpValue<>"") then
        mxChartSeries = 1
        vArr2 = oChart.fnBrandBestSell_DW
        vArr1 = oChart.fnBrandSellTop_Trend_DW(vpValue)
        vArr3 = oChart.fnBrandSellTop_Trend_Monthly_DW(vpValue)
    else
	    vArr2 = oChart.fnBrandSellTop_DW(vArr1)
	end if


Dim iChartCaption : iChartCaption = "�귣�庰 �ֹ��Ǽ�"
Dim iChartSubCaption : iChartSubCaption = ""
Dim ixAxisName : ixAxisName = "" ''��¥
Dim yAxisName : yAxisName = "�ֹ��Ǽ�"
Dim iDataSetPosArr
Dim iDataSetHeadArr
Dim iDataSeriseArr
Dim epPOSN : epPOSN= -1

if (UCASE(vChannel)="TEN") then
    iDataSeriseArr = Array("WEB","MOB","APP")
elseif (UCASE(vChannel)="TEN_LK") then
    iDataSeriseArr = Array("WEB","W_LK","MOB","M_LK","APP","A_LK")
elseif (vChannel="") then
    iDataSeriseArr = Array("WEB","W_LK","MOB","M_LK","APP","A_LK","OUT","FRN")
else
    iDataSeriseArr = Array("WEB","MOB","APP","OUT","FRN")
end if
if (vOrdType="C") then
    if (UCASE(vChannel)="TEN") then
        iDataSetPosArr = Array(5,8,11)
        epPOSN = 20
    elseif (UCASE(vChannel)="TEN_LK") then
        iDataSetPosArr = Array(5,8,11,14,17,20)
        epPOSN = 29
    elseif (vChannel="") then
        iDataSetPosArr = Array(5,8,11,14,17,20,23,26)
        epPOSN = 29
    else
        iDataSetPosArr = Array(5,8,11,14,17)
        epPOSN = 20
    end if
    iDataSetHeadArr = Array("�ֹ��Ǽ�")


    iChartCaption = "�귣�庰 �ֹ��Ǽ�"
    yAxisName = "�ֹ��Ǽ�"
elseif (vOrdType="S") then
    if (UCASE(vChannel)="TEN") then
        iDataSetPosArr = Array(6,9,12)
        epPOSN = 21
    elseif (UCASE(vChannel)="TEN_LK") then
        iDataSetPosArr = Array(6,9,12,15,18,21)
        epPOSN = 30
    elseif (vChannel="") then
        iDataSetPosArr = Array(6,9,12,15,18,21,24,27)
        epPOSN = 30
    else
        iDataSetPosArr = Array(6,9,12,15,18)
        epPOSN = 21
    end if
    iDataSetHeadArr = Array("�����Ѿ�")

    iChartCaption = "�귣�庰 �����Ѿ�"
    yAxisName = "�����Ѿ�"
elseif (vOrdType="G") then
    if (UCASE(vChannel)="TEN") then
        iDataSetPosArr = Array(7,10,13)
        epPOSN = 22
    elseif (UCASE(vChannel)="TEN_LK") then
        iDataSetPosArr = Array(7,10,13,16,19,22)
        epPOSN = 31
    elseif (vChannel="") then
        iDataSetPosArr = Array(7,10,13,16,19,22,25,28)
        epPOSN = 31
    else
        iDataSetPosArr = Array(7,10,13,16,19)
        epPOSN = 22
    end if
    iDataSetHeadArr = Array("�������")

    iChartCaption = "�귣�庰 �������"
    yAxisName = "�������"
end if



Dim pTypeName : pTypeName = "�귣��ID"
Dim chrtN

dim precate, imakerid
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.ocean.js"></script>

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
		alert("�������� �Է��ϼ���");
		return false;
	}
	if($("#edate").val()== ""){
		alert("�������� �Է��ϼ���");
		return false;
	}
	document.frm1.submit();
}
</script>

<script type='text/javascript'>//<![CDATA[
window.onload=function(){

<% if isArray(vArr2) then %>
<%
dim posN : posN = 2
redim pielabelValArr(UBound(iDataSeriseArr))

if vOrdType="S" then posN=4
if vOrdType="G" then posN=7
%>

<% if (vpValue="") then %>
// �߼���Ʈ
FusionCharts.ready(function () {
    var vstrChart<%=chrtN%> = new FusionCharts({
        type: 'msline', //'',
        renderAt: 'chart-container0',
        width: '1100',
        height: '400',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "<%=imakerid%> �Ϻ� �߼� (�ִ� 180��)",
                "subCaption": "<%=iChartSubCaption%>",
                "xAxisName": "<%=ixAxisName%>",
                "yAxisName": "<%=yAxisName%>",
                "theme": "fint",
                "showSum": "1",
                "showValues": "1",
                //Setting automatic calculation of div lines to off
  //              "adjustDiv": "0",
                //Manually defining y-axis lower and upper limit
                //"yAxisMaxvalue": "35000",	//y�� �ƽ���
                //"yAxisMinValue": "5000",		//y�� �ΰ�
                //Setting number of divisional lines to 9
                //"numDivLines": "9"				//0~�ƽ� ���� ǥ�õǾ����� ��ġ����
  //              "anchorBgHoverColor": "#96d7fa",
  //              "anchorBorderHoverThickness" : "4",
  //              "anchorHoverRadius":"7"
            },
            // X��
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
                <% for k=0 to mxChartSeries-1 %>
                <% if UBound(vArr2,2)>=k then %>
                {
                    "seriesname": "<%=vArr2(0,k)%>",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
							    if (vArr1(1,i)=vArr2(0,k)) then  ''�귣�尡 ������
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(posN,i)&"""" & vbCrLf
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
                <% if (k<UBound(vArr2,2)) then response.write "," %>
                <% end if %>
                <% next %>
            ]
        }
    }).render();
});
<% end if %>

<% if (vpValue<>"") then %>
<%
posN=2
if vOrdType="S" then posN=3
if vOrdType="G" then posN=4
%>

// �߼���Ʈ(����)
FusionCharts.ready(function () {
    var vstrChartMonthlyBrand = new FusionCharts({
        type: 'msline', //'',
        renderAt: 'chart-container3',
        width: '1100',
        height: '400',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "<%=vpValue%> ���� �߼� (�ִ� 18����)",
                "subCaption": "<%=iChartSubCaption%>",
                "xAxisName": "<%=ixAxisName%>",
                "yAxisName": "<%=yAxisName%>",
                "theme": "ocean",
                "showSum": "1",
                "showValues": "1"
            },
            // X��
            "categories": [
                {
                    "category": [
						<%
						If isArray(vArr3) Then
							For i = 0 To UBound(vArr3,2)
                                Response.Write "{" & vbCrLf
                                Response.Write """label"": """&vArr3(0,i)&"""" & vbCrLf
                                Response.Write "}"
                                If i <> UBound(vArr3,2) Then
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
                    "seriesname": "<%=vpValue%>",
                    "data": [
						<%
						If isArray(vArr3) Then
							For i = 0 To UBound(vArr3,2)
							    for k=0 to Ubound(iDataSeriseArr)
							        pielabelValArr(k) = pielabelValArr(k) + vArr3(iDataSetPosArr(k),i)
							    next
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr3(posN,i)&"""" & vbCrLf
								Response.Write "}"
								If i <> UBound(vArr3,2) Then
									Response.Write ","
								End If
								Response.Write vbCrLf
							Next
						End If
						%>
                    ]
                }
                <% if epPOSN>=0 then %>
                ,{
                    "seriesname": "Naver EP",
                    "data": [
						<%
						If isArray(vArr3) Then
							For i = 0 To UBound(vArr3,2)
							    ' for k=0 to Ubound(iDataSeriseArr)
							    '     pielabelValArr(k) = pielabelValArr(k) + vArr3(iDataSetPosArr(k),i)
							    ' next
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr3(epPOSN,i)&"""" & vbCrLf
								Response.Write "}"
								If i <> UBound(vArr3,2) Then
									Response.Write ","
								End If
								Response.Write vbCrLf
							Next
						End If
						%>
                    ]
                }
                <% end if %>

            ]
        }
    }).render();
});

// �߼���Ʈ(�Ϻ�)
FusionCharts.ready(function () {
    var vstrChart<%=chrtN%> = new FusionCharts({
        type: 'msline', //'',
        renderAt: 'chart-container0',
        width: '1100',
        height: '400',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "<%=vpValue%> �Ϻ� �߼� (�ִ� 180��)",
                "subCaption": "<%=iChartSubCaption%>",
                "xAxisName": "<%=ixAxisName%>",
                "yAxisName": "<%=yAxisName%>",
                "theme": "fint",
                "showSum": "1",
                "showValues": "1",
                //Setting automatic calculation of div lines to off
  //              "adjustDiv": "0",
                //Manually defining y-axis lower and upper limit
                //"yAxisMaxvalue": "35000",	//y�� �ƽ���
                //"yAxisMinValue": "5000",		//y�� �ΰ�
                //Setting number of divisional lines to 9
                //"numDivLines": "9"				//0~�ƽ� ���� ǥ�õǾ����� ��ġ����
  //              "anchorBgHoverColor": "#96d7fa",
  //              "anchorBorderHoverThickness" : "4",
  //              "anchorHoverRadius":"7"
            },
            // X��
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
                {
                    "seriesname": "<%=vpValue%>",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
							    for k=0 to Ubound(iDataSeriseArr)
							        pielabelValArr(k) = pielabelValArr(k) + vArr1(iDataSetPosArr(k),i)
							    next
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
                <% if epPOSN>=0 then %>
                ,{
                    "seriesname": "Naver EP",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
							    ' for k=0 to Ubound(iDataSeriseArr)
							    '     pielabelValArr(k) = pielabelValArr(k) + vArr1(iDataSetPosArr(k),i)
							    ' next
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(epPOSN,i)&"""" & vbCrLf
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
                <% end if %>

            ]
        }
    }).render();
});

FusionCharts.ready(function () {
    var vstrChart0 = new FusionCharts({
        type: 'stackedcolumn2d', //'',
        renderAt: 'chart-container1',
        width: '1100',
        height: '400',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "<%=vpValue%> �Ϻ� ä�� ���� �߼�",
                "subCaption": "<%=iChartSubCaption%>",
                "xAxisName": "<%=ixAxisName%>",
                "yAxisName": "<%=yAxisName%>",
                "theme": "fint",
                "showSum": "1",
                "showValues": "1",
                //Setting automatic calculation of div lines to off
  //              "adjustDiv": "0",
                //Manually defining y-axis lower and upper limit
                //"yAxisMaxvalue": "35000",	//y�� �ƽ���
                //"yAxisMinValue": "5000",		//y�� �ΰ�
                //Setting number of divisional lines to 9
                //"numDivLines": "9"				//0~�ƽ� ���� ǥ�õǾ����� ��ġ����
  //              "anchorBgHoverColor": "#96d7fa",
  //              "anchorBorderHoverThickness" : "4",
  //              "anchorHoverRadius":"7"
            },
            // X��
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
                <% for k=0 to Ubound(iDataSeriseArr) %>
                {
                    "seriesname": "<%=iDataSeriseArr(k)%>",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
							    'if (vArr1(1,i)=vArr2(0,chrtN-1)) then  ''�귣�尡 ������
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(iDataSetPosArr(k),i)&"""" & vbCrLf
								Response.Write "}"
								If i <> UBound(vArr1,2) Then
									Response.Write ","
								End If
								Response.Write vbCrLf
							    'end if
							Next
						End If
						%>
                    ]
                }
                <% if (k<Ubound(iDataSeriseArr)) then response.write "," %>
                <% next %>

            ]
        }
    }).render();
});

// pie ��Ʈ(ä��)
FusionCharts.ready(function () {
    var vstrChart<%=chrtN%> = new FusionCharts({
        type: 'pie2d', //'',
        renderAt: 'chart-container2',
        width: '1100',
        height: '400',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "<%=imakerid%> ä�κ���",
                "subCaption": "<%=iChartSubCaption%>",
                "numberPrefix": "",
                "showPercentInTooltip": "0",
                "decimals": "1",
                "useDataPlotColorForLabels": "1",
                "theme": "fint"
            },

            "data": [
				<%
				If isArray(vArr1) Then
					For i = Lbound(iDataSetPosArr) To UBound(iDataSetPosArr)
						Response.Write "{" & vbCrLf
						Response.Write """label"": """&iDataSeriseArr(i)&"""," & vbCrLf
						Response.Write """value"": """&pielabelValArr(i)&"""" & vbCrLf
						Response.Write "}"
						If i <> UBound(iDataSetPosArr) Then
							Response.Write ","
						End If
						Response.Write vbCrLf
					Next
				End If
				%>
            ]
        }
    }).render();
});
<% end if %>

<% end if %>



}//]]>
</script>
<%
SET oChart = nothing
%>
<body>
<form name="frm1" method="post" action="/admin/dataanalysis/chart/sellbybrand.asp">
<input type="hidden" name="menupos" value="<%=menupos%>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#F4F4F4">
    <td width="50" rowspan="2" bgcolor="#EEEEEE">�˻�<br>����</td>
	<td align="left">
	��ȸ��¥ :
	    <input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
		~<input id='enddate' name='enddate' value='<%= vEDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
    &nbsp;&nbsp;

    ä�� : <% call drawConversionChannelSelectBoxII("channel",vChannel) %>
    &nbsp;&nbsp;

    <% if (FALSE) then %>
    rdsiteŸ�� : <% call drawConversionTypeGroupSelectBox2("rdsitegrp",rdsitegrp,"rdsite",2,"") %>
    &nbsp;&nbsp;
    <% end if %>
    <input type="radio" name="ordtype" value="C" <%=CHKIIF(vOrdType="C","checked","") %> >�ֹ��Ǽ���
    <input type="radio" name="ordtype" value="S" <%=CHKIIF(vOrdType="S","checked","") %> >�����Ѿ׼�
    <input type="radio" name="ordtype" value="G" <%=CHKIIF(vOrdType="G","checked","") %> >������ͼ�
    &nbsp;&nbsp;
    <%= pTypeName %> : <input type="text" name="pvalue" size="20" value="<%=vpValue%>">

    <% if (vpValue<>"") then %>
    &nbsp;&nbsp;
    <input type="button" value="���޼���LOG" onClick="window.open('/admin/etc/outmall/index.asp?research=on&menupos=1742&makerid=<%=vpValue%>','_outmallsellyn','');">
    <% end if %>
    </br> ����ī�װ�: <!-- #include virtual="/common/module/dispCateSelectBox.asp"-->
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="�˻�" onClick="goSearch(document.frm1);">
	</td>
</tr>

</table>
</form>

* �ֹ��� ����, �ֹ����� ����, 1�ð� ����������<br /><br />
<% dim sum1,sum2,sum3,sum4,sum5,sum6,sum7,sum10,sum11 %>
<table cellpadding="0" cellspacing="0" border="0" class="a" align="center">
<tr bgcolor="#FFFFFF">
    <% if isArray(vArr2) then %>
    <% if (vpValue<>"") then %>
    <td width="700" valign="top">
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4">
            <td>Rnk</td>
            <td>��ǰ�ڵ�</td>
            <td>�̹���</td>
            <td>�ֹ�<br>�Ǽ�</td>
            <td>��ǰ<br>����</td>
            <td>�����Ѿ�</td>
            <td>��޾�</td>
            <td>���Ծ�</td>
            <td>�������1</td>
            <td>�������2</td>
            <td>���԰�</td>
            <td>�ǻ����</td>
        </tr>
        <% For i = 0 To UBound(vArr2,2) %>
        <%
        sum1 = sum1 + vArr2(1,i)
        sum2 = sum2 + vArr2(2,i)
        sum3 = sum3 + vArr2(3,i)
        sum4 = sum4 + vArr2(4,i)
        sum5 = sum5 + vArr2(5,i)
        sum6 = sum6 + vArr2(6,i)
        sum7 = sum7 + vArr2(7,i)
        sum10 = sum10 + vArr2(10,i)
        sum11 = sum11 + vArr2(11,i)
        %>
        <tr bgcolor="#FFFFFF" align="right">
            <td align="center"><%=i+1%></td>
            <td ><%=vArr2(0,i)%></td>
            <td align="center"><%=vArr2(9,i)%></td>
            <td><%=FormatNumber(vArr2(1,i),0)%></td>
            <td><%=FormatNumber(vArr2(2,i),0)%></td>
            <td><%=FormatNumber(vArr2(3,i),0)%></td>
            <td><%=FormatNumber(vArr2(4,i),0)%></td>
            <td><%=FormatNumber(vArr2(5,i),0)%></td>
            <td><%=FormatNumber(vArr2(6,i),0)%></td>
            <td><%=FormatNumber(vArr2(7,i),0)%></td>
            <td><%=FormatNumber(vArr2(10,i),0)%></td>
            <td><%=FormatNumber(vArr2(11,i),0)%></td>
        </tr>
        <% next %>
        <tr bgcolor="#F4F4F4" align="right">
            <td align="center">�հ�</td>
            <td ></td>
            <td align="center"></td>
            <td><%=FormatNumber(sum1,0)%></td>
            <td><%=FormatNumber(sum2,0)%></td>
            <td><%=FormatNumber(sum3,0)%></td>
            <td><%=FormatNumber(sum4,0)%></td>
            <td><%=FormatNumber(sum5,0)%></td>
            <td><%=FormatNumber(sum6,0)%></td>
            <td><%=FormatNumber(sum7,0)%></td>
            <td><%=FormatNumber(sum10,0)%></td>
            <td><%=FormatNumber(sum11,0)%></td>
        </tr>
        </table>
    </td>
    <% else %>
    <td width="700" valign="top">
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4">
            <td><%=pTypeName%></td>
            <td>�ֹ�<br>�Ǽ�</td>
            <td>��ǰ<br>����</td>
            <td>�����Ѿ�</td>
            <td>��޾�</td>
            <td>���Ծ�</td>
            <td>�������1</td>
            <td>�������2</td>
            <td>��</td>
        </tr>
        <% For i = 0 To UBound(vArr2,2) %>
        <tr bgcolor="#FFFFFF" align="right">
            <td align="left"><%=vArr2(0,i)%></td>
            <td><%=FormatNumber(vArr2(1,i),0)%></td>
            <td><%=FormatNumber(vArr2(2,i),0)%></td>
            <td><%=FormatNumber(vArr2(3,i),0)%></td>
            <td><%=FormatNumber(vArr2(4,i),0)%></td>
            <td><%=FormatNumber(vArr2(5,i),0)%></td>
            <td><%=FormatNumber(vArr2(6,i),0)%></td>
            <td><%=FormatNumber(vArr2(7,i),0)%></td>
            <td><a target="_branddtl" href="/admin/dataanalysis/chart/sellbybrand.asp?menupos=<%=menupos%>&startdate=<%=vSDate%>&enddate=<%=vEDate%>&channel=<%=vChannel%>&pvalue=<%=vArr2(0,i)%>&ordtype=<%=vOrdType%>">����</a></td>
        </tr>
        <% next %>
        </table>
    </td>
    <% end if %>
    <% end if %>
	<td valign="top">
	<% if (vpValue<>"") then %>
        <div id="chart-container0">FusionCharts will render here</div>
	    <div id="chart-container1">FusionCharts will render here</div>
	    <div id="chart-container2">FusionCharts will render here</div>
        <div id="chart-container3">FusionCharts will render here</div>
    <% else %>
        <% if isArray(vArr2) then %>
            <div id="chart-container0">FusionCharts will render here</div>
        <% end if %>
	<% end if %>
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->