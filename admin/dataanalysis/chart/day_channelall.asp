<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/dataanalysis/chart/chartCls.asp" -->
<%
Dim oChart, vArr1, i, vArr2
Dim vSDate, vEDate

vSDate = requestCheckvar(request("sdate"),10)
vEDate = requestCheckvar(request("edate"),10)

If vSDate = "" Then
	vSDate = LEFT(Date(), 7) & "-01"
End If

If vEDate = "" Then
	vEDate = date()
End If

SET oChart = new CChart
	oChart.FRectSDate = vSDate
	oChart.FRectEDate = vEDate
	vArr1 = oChart.fnDayChannelAll
SET oChart = nothing
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
		alert("�������� �Է��ϼ���");	
		return false;
	}
	if($("#edate").val()== ""){
		alert("�������� �Է��ϼ���");	
		return false;
	}
	frm1.submit();
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
                "caption": "�Ϻ� ä�� ��ü",
                "subCaption": "ä�κ� �ֹ��Ǽ� ����",
                "xAxisName": "��¥",
                "yAxisName": "�ֹ��Ǽ�",
                "theme": "fint",
                "showValues": "1",
                //Setting automatic calculation of div lines to off
                "adjustDiv": "0",
                //Manually defining y-axis lower and upper limit
                //"yAxisMaxvalue": "35000",	//y�� �ƽ���
                //"yAxisMinValue": "5000",		//y�� �ΰ�
                //Setting number of divisional lines to 9
                //"numDivLines": "9"				//0~�ƽ� ���� ǥ�õǾ����� ��ġ����
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
                {
                    "seriesname": "�ֹ��Ǽ�-mw",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(1,i)&"""" & vbCrLf
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
                    "seriesname": "�ֹ��Ǽ�-pc",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(2,i)&"""" & vbCrLf
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
                    "seriesname": "�ֹ��Ǽ�-app",
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
                }
            ]
        }
    }).render();


    var vstrChart2 = new FusionCharts({
        type: 'msline',
        renderAt: 'chart-container2',
        width: '1200',
        height: '500',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "�Ϻ� ä�� ��ü",
                "subCaption": "Ÿ�Ժ� �����Ѿ� ����",
                "xAxisName": "��¥",
                "yAxisName": "�����Ѿ�",
                "theme": "fint",
                "showValues": "1",
                //Setting automatic calculation of div lines to off
                "adjustDiv": "0",
                //Manually defining y-axis lower and upper limit
                //"yAxisMaxvalue": "35000",	//y�� �ƽ���
                //"yAxisMinValue": "5000",		//y�� �ΰ�
                //Setting number of divisional lines to 9
                //"numDivLines": "9"				//0~�ƽ� ���� ǥ�õǾ����� ��ġ����
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
                {
                    "seriesname": "�����Ѿ�-mw",
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
                    "seriesname": "�����Ѿ�-pc",
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
                },
                {
                    "seriesname": "�����Ѿ�-app",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(6,i)&"""" & vbCrLf
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

});
}//]]>
</script>
<body>
<form name="frm1" method="post" action="day_channelall.asp">
<table width="100%" class="a">
<tr>
	<td height="30" align="center">
		��ȸ��¥ : 
		<input type="text" name="sdate" id="sdate" value="<%=vSDate%>" onClick="jsPopCal('sdate');" style="cursor:pointer;" size="10" maxlength="10" readonly>&nbsp;~&nbsp;
		<input type="text" name="edate" id="edate" value="<%=vEDate%>" onClick="jsPopCal('edate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
		&nbsp;&nbsp;&nbsp;
		<input type="button" value="��  ȸ" onClick="goSearch();">
	</td>
</tr>
</table>
</form>
<br />
<table cellpadding="0" cellspacing="0" border="0" class="a" align="center">
<tr bgcolor="#FFFFFF">
    <td>
    * ���� ä�κ� ����� ���̰� �ֽ��ϴ�. (��ȯ Ÿ�Կ� ���� ���� �߼���  ǥ�õ˴ϴ�)
    <br>* ���� �߼��� <a href="/admin/dataanalysis/chart/timeorder_trend_channel.asp"><strong>ä�κ� ��������</strong></a> �̰����� ���� �� �ֽ��ϴ�.
    </td>
</tr>  
<tr bgcolor="#FFFFFF">
	<td>
		<div id="chart-container1">FusionCharts will render here</div>
		<br />
		<div id="chart-container2">FusionCharts will render here</div>
	</td>
</tr>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->