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
    
<% if isArray(vArr1) then %>
<%
    vArrTitle = Array("NV �������/�ڻ�","�ڻ��������","NV ������")
    vArrPos = Array(13,14,15)
%>
// "VN ������� �� ������",
FusionCharts.ready(function () {
    var vstrChart1 = new FusionCharts({
        type: 'msline', //'', 
        renderAt: 'chart-container0',
        width: '800',
        height: '400',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "VN ������� �� ������",
                "subCaption": "",
                "xAxisName": "��¥",
                "yAxisName": "%",
                "theme": "fint",
                "showSum": "1",
                "showValues": "<%=CHKIIF(UBound(vArr1,2)>14,0,1)%>",
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
    vArrTitle = Array("�ڻ������","NV ����","���޸� ����")
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
                "caption": "�ڻ�� ����� / NV ����� / ���޸� ����",
                "subCaption": "",
                "xAxisName": "��¥",
                "yAxisName": "�����",
                "theme": "fint",
                "showSum": "<%=CHKIIF(UBound(vArr1,2)>14,0,1)%>",
                "showValues": "<%=CHKIIF(UBound(vArr1,2)>14,0,1)%>",
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
    vArrTitle = Array("�ڻ������","NV ����","���޸� ����")
    vArrPos = Array(3,11,18)
%>
// "VN �����",
FusionCharts.ready(function () {
    var vstrChart2 = new FusionCharts({
        type: 'msline', //'', 
        renderAt: 'chart-container2',
        width: '800',
        height: '400',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "�ڻ�� ����� / NV ����� / ���޸� ����",
                "subCaption": "",
                "xAxisName": "��¥",
                "yAxisName": "�����",
                "theme": "fint",
                "showSum": "1",
                "showValues": "<%=CHKIIF(UBound(vArr1,2)>14,0,1)%>",
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
    <td width="50" rowspan="2" bgcolor="#EEEEEE">�˻�<br>����</td>
	<td align="left">
	<select name="datebase">
        <option value="ipkumdt" <%=CHKIIF(datebase="ipkumdt","selected","")%> >������
        <option value="oipkumdt" <%=CHKIIF(datebase="oipkumdt","selected","")%> >��������
        <option value="beasongdt" <%=CHKIIF(datebase="beasongdt","selected","")%> >�����
    </select>
     : 
	    <input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
		~<input id='enddate' name='enddate' value='<%= vEDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
    &nbsp;&nbsp;
    
    ä�� : <% call drawConversionChannelSelectBoxII("channel",vChannel) %>
    &nbsp;&nbsp;
    ��¥���� : 
    <input type="radio" name="grptype" value="d" <%=CHKIIF(grptype="d","checked","") %> >��
    <input type="radio" name="grptype" value="m" <%=CHKIIF(grptype="m","checked","") %> >��
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="�˻�" onClick="goSearch(document.frm1);">
	</td>
</tr>

</table>
</form>

* ����αױ���, 1�� ����������<br />
* ����ϱ����� ��� ��ۺ�� �����۾� �� �ݿ��˴ϴ�.<br />
*EP��������(rdsite ������� ���� Naver ���� ����)<br />
* ����αװ�����(�������(==������)�ݿ���), ����α׿�������(������ڹ��� ��������)
<% dim sum1,sum2,sum3,sum4,sum5,sum6,sum7 %>
<table cellpadding="0" cellspacing="0" border="0" class="a" align="center">
<tr bgcolor="#FFFFFF">
    <% if isArray(vArr1) then %>
    
    <td width="900" valign="top">
        <!-- yyyymm	�Ǹž�	�����Ѿ�	�����Ѿ�	�����Ѿ�	NV����_�Ǹž�	NV����_�����Ѿ�	NV����_�����Ѿ�	NV����_�����Ѿ�	NV_�Ǹž�	NV_�����Ѿ�	NV_�����Ѿ�	NV_�����Ѿ�-->
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4">
            <td>��¥</td>
            <td>�ڻ��Ǹž�<br>(Nv����)</td>
            <td>�ڻ��ֹ���<br>(Nv����)</td>
            <td>�ڻ籸���Ѿ�<br>(Nv����)</td>
            <td>�ڻ�����Ѿ�<br>(Nv����)</td>
            <td>�ڻ�����Ѿ�<br>(Nv����)</td>
            
            <td>NV �Ǹž�</td>
            <td>NV �ֹ���</td>
            <td>NV �����Ѿ�</td>
            <td>NV �����Ѿ�</td>
            <td>NV �����Ѿ�</td>
            
            <td>NV<br>�������</td>
            <td>�ڻ��<br>������2</td>
            <td>NV<br>������2</td>
            <td>EP����<br>����</td>
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