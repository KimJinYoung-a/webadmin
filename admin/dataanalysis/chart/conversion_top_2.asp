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
Dim oChart, vArr1, vArr2, vArrR, vArrB, vArrE, vArrC, i, j
Dim vSDate, vEDate, vChannel, vpType, vpValue, vOrdType, sTp, vpUpType
Dim dPageSize : dPageSize = 50

vSDate = requestCheckvar(request("startdate"),10)
vEDate = requestCheckvar(request("enddate"),10)
vChannel = requestCheckvar(request("channel"),10)
vpType = requestCheckvar(request("ptype"),32)
vpValue = requestCheckvar(request("pvalue"),64)
vOrdType = requestCheckvar(request("ordtype"),32)
vpUpType = requestCheckvar(request("puptype"),32)

''if (vpType="") then vpType="pRtr"
if (vOrdType="") then vOrdType="C" ''�Ǽ�(C) , �ݾ�(S), ����(G)

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
	    
	    oChart.FPageSize = 5
	    vArr1 = oChart.fnConversionTopByType_Trend
	end if
	
	
	
	
SET oChart = nothing

Dim iChartCaption : iChartCaption = "��ȯ Ÿ�Ժ� �ֹ��Ǽ�"
Dim iChartSubCaption : iChartSubCaption = ""
Dim ixAxisName : ixAxisName = "" ''��¥
Dim yAxisName : yAxisName = "�ֹ��Ǽ�"
Dim iDataSetPosArr : iDataSetPosArr = Array(2)
Dim iDataSetHeadArr : iDataSetHeadArr = Array("�ֹ��Ǽ�")

if (vOrdType="2") then
    iDataSetPosArr = Array(6)
    iDataSetHeadArr = Array("�����Ѿ�")
    
    iChartCaption = "��ȯ Ÿ�Ժ� �����Ѿ�"
    yAxisName = "�����Ѿ�"
end if

Dim pTypeName : pTypeName = getpTypeName(vpType)
Dim mxChartSeries : mxChartSeries = 5

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
    <td width="50" rowspan="2" bgcolor="#EEEEEE">�˻�<br>����</td>
	<td align="left">
	��ȸ��¥ : 
	    <input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
		~<input id='enddate' name='enddate' value='<%= vEDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
			
    &nbsp;&nbsp;
    
    ä�� : <% call drawConversionChannelSelectBox("channel",vChannel) %>
    &nbsp;&nbsp;
    
    ��ȯŸ�� : <% call drawConversionTypeSelectBox("ptype",vpType) %>
    &nbsp;&nbsp;
    
    <% if (vpType="gaparam") then %>
        <% call drawConversionTypeGroupSelectBox("puptype",vpUpType, vpType) %>
        &nbsp;&nbsp;
    <% end if %> 
    <input type="radio" name="ordtype" value="C" <%=CHKIIF(vOrdType="C","checked","") %> >�ֹ��Ǽ���
    <input type="radio" name="ordtype" value="S" <%=CHKIIF(vOrdType="S","checked","") %> >�����Ѿ׼�
    <input type="radio" name="ordtype" value="G" <%=CHKIIF(vOrdType="G","checked","") %> >������ͼ�
    &nbsp;&nbsp;
    <% if (vpType<>"") then %>
    <%= pTypeName %> : <input type="text" name="pvalue" size="20" value="<%=vpValue%>">
    <% end if %>
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="�˻�" onClick="goSearch(document.frm1);">
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
        <tr bgcolor="#F4F4F4">
            <td width="100"><%= getpTypeName("pRtr")%></td>
            <td></td>
            <td width="40">�ֹ�<br>�Ǽ�</td>
            <td width="60">�����Ѿ�</td>
            <td width="60">�������</td>
        </tr>
        <% For i = 0 To UBound(vArrR,2) %>
        <tr bgcolor="#FFFFFF" align="right">
            <td align="left"><%=vArrR(0,i)%></td>
            <td align="left"><%=vArrR(11,i)%></td>
            <td><%=FormatNumber(vArrR(1,i),0)%></td>
            <td><%=FormatNumber(vArrR(5,i),0)%></td>
            <td><%=FormatNumber(vArrR(9,i),0)%></td>
        </tr>
        <% next %>
        </table>
    <% end if %>
    </td>
	<td width="23%" valign="top">
	<% if isArray(vArrB) then %>
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4">
            <td width="100"><%= getpTypeName("pBtr")%></td>
            <td></td>
            <td width="40">�ֹ�<br>�Ǽ�</td>
            <td width="60">�����Ѿ�</td>
            <td width="60">�������</td>
        </tr>
        <% For i = 0 To UBound(vArrB,2) %>
        <tr bgcolor="#FFFFFF" align="right">
            <td align="left"><%=vArrB(0,i)%></td>
            <td align="left"><%=vArrB(11,i)%></td>
            <td><%=FormatNumber(vArrB(1,i),0)%></td>
            <td><%=FormatNumber(vArrB(5,i),0)%></td>
            <td><%=FormatNumber(vArrB(9,i),0)%></td>
        </tr>
        <% next %>
        </table>
    <% end if %>    
	</td>
	<td width="27%" valign="top">
	<% if isArray(vArrE) then %>
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4">
            <td width="40"><%= getpTypeName("pEtr")%></td>
            <td></td>
            <td width="40">�ֹ�<br>�Ǽ�</td>
            <td width="60">�����Ѿ�</td>
            <td width="60">�������</td>
        </tr>
        <% For i = 0 To UBound(vArrE,2) %>
        <tr bgcolor="#FFFFFF" align="right">
            <td align="left"><%=vArrE(0,i)%></td>
            <td align="left"><%=vArrE(11,i)%></td>
            <td><%=FormatNumber(vArrE(1,i),0)%></td>
            <td><%=FormatNumber(vArrE(5,i),0)%></td>
            <td><%=FormatNumber(vArrE(9,i),0)%></td>
        </tr>
        <% next %>
        </table>
    <% end if %>    
	</td>
	<td width="26%" valign="top">
	<% if isArray(vArrC) then %>
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4">
            <td width="60"><%= getpTypeName("pCtr")%></td>
            <td></td>
            <td width="40">�ֹ�<br>�Ǽ�</td>
            <td width="60">�����Ѿ�</td>
            <td width="60">�������</td>
        </tr>
        <% For i = 0 To UBound(vArrC,2) %>
        <tr bgcolor="#FFFFFF" align="right">
            <td align="left"><%=vArrC(0,i)%></td>
            <td align="left"><%=vArrC(11,i)%></td>
            <td><%=FormatNumber(vArrC(1,i),0)%></td>
            <td><%=FormatNumber(vArrC(5,i),0)%></td>
            <td><%=FormatNumber(vArrC(9,i),0)%></td>
        </tr>
        <% next %>
        </table>
    <% end if %>    
	</td>
</tr>
</table>
<% else %>
<table cellpadding="0" cellspacing="0" border="0" class="a" align="center">
<tr bgcolor="#FFFFFF">
    <% if isArray(vArr2) then %>
    <td width="500">
        <table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4">
            <td><%=pTypeName%></td>
            <td></td>
            <td>�ֹ�<br>�Ǽ�</td>
            <td>�����Ѿ�</td>
            <td>�������</td>
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
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->