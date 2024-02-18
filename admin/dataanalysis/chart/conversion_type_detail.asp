<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/dataanalysis/chart/chartCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<h1>������..</h1>
<%
Dim oChart, vArr1, vArr2, i, j
Dim vSDate, vEDate, vChannel, vpType, vpValue, vOrdType, sTp, vpUpType

vSDate = requestCheckvar(request("startdate"),10)
vEDate = requestCheckvar(request("enddate"),10)
vChannel = requestCheckvar(request("channel"),10)
vpType = requestCheckvar(request("ptype"),32)
vpValue = requestCheckvar(request("pvalue"),64)
vOrdType = requestCheckvar(request("ordtype"),32)
vpUpType = requestCheckvar(request("puptype"),32)

if (vpType="") then vpType="pEtr"
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
	oChart.FPageSize = CHKIIF(vpType="gaparam",50,50)
	oChart.FRectOrderType = vOrdType
	oChart.FRectPValue = vpValue
	oChart.FRectUPTypeValue = vpUpType
	oChart.FPageSize = 100
	if (vpValue<>"") then
	    vArr1 = oChart.fnConversionTopByType_Trend
    end if
	
	vArr2 = oChart.fnConversionTopByType_Item
SET oChart = nothing

Dim iChartCaption : iChartCaption = "��ȯ Ÿ�Ժ� �ֹ��Ǽ�"
Dim iChartSubCaption : iChartSubCaption = ""
Dim ixAxisName : ixAxisName = "" ''��¥
Dim yAxisName : yAxisName = "�ֹ��Ǽ�"
Dim iDataSetPosArr : iDataSetPosArr = Array(2)
Dim iDataSetHeadArr : iDataSetHeadArr = Array("�ֹ��Ǽ�")

if (vOrdType="S") then
    iDataSetPosArr = Array(6)
    iDataSetHeadArr = Array("�����Ѿ�")
    
    iChartCaption = "��ȯ Ÿ�Ժ� �����Ѿ�"
    yAxisName = "�����Ѿ�"
end if

Dim pTypeName : pTypeName = getpTypeName(vpType)
Dim mxChartSeries : mxChartSeries = 5
Dim vSum1,vSum2,vSum3
Dim isellStr, iLimitStr, priceStr, iSellyn, iLimityn, iLimitNo
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
        width: '800',
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
                <% if isArray(vArr1) then %>
                {
                    "seriesname": "<%=vArr1(0,j)%>",
                    "data": [
						<%
						If isArray(vArr1) Then
							For i = 0 To UBound(vArr1,2)
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr1(iDataSetPosArr(0),i)&"""" & vbCrLf
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
                <% if j<UBound(vArr1,2) then %>,<% end if %>
                <% end if %>
            ]
        }
    }).render();


});
}//]]>
</script>
<body>
    * ��ǰ�� �ֹ��Ǽ� �հ�� ��ȯ Ÿ�Ժ� �հ�� �ٸ� �� ����. (ex ���� �˻���� ���� ��ǰ�� ���ֹ����� �ֹ��� ���)
<p>
<form name="frm1" method="get" >
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
    <%= pTypeName %> : <input type="text" name="pvalue" size="20" value="<%=vpValue%>">
     
    &nbsp;&nbsp;
    
    <% if (vpType="gaparam") then %>
        <% call drawConversionTypeGroupSelectBox("puptype",vpUpType, vpType) %>
        &nbsp;&nbsp;
    <% end if %> 
    <input type="radio" name="ordtype" value="C" <%=CHKIIF(vOrdType="C","checked","") %> >�ֹ��Ǽ���
    <input type="radio" name="ordtype" value="S" <%=CHKIIF(vOrdType="S","checked","") %> >�����Ѿ׼�
    <input type="radio" name="ordtype" value="G" <%=CHKIIF(vOrdType="G","checked","") %> >������ͼ�
    &nbsp;&nbsp;
   
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="�˻�" onClick="goSearch(document.frm1);">
	</td>
</tr>

</table>
</form>
<br />
<table width="100%" cellpadding="2" cellspacing="5" border="0" class="a" align="center">
<tr bgcolor="#FFFFFF">
    <% if isArray(vArr2) then %>
    <td valign="top">
        <table width="100%" cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4">
            
            <td>��ǰ�ڵ�</td>
            <td>��ǰ��</td>
            <td>�ֹ��Ǽ�</td>
            <td>�����Ѿ�</td>
            <td>�������</td>
            <td>�ǸŻ���</td>
        </tr>
        <% For i = 0 To UBound(vArr2,2) %>
        <%
        vSum1=vSum1+vArr2(1,i)
        vSum2=vSum2+vArr2(5,i)
        vSum3=vSum3+vArr2(9,i)
        
        isellStr    =""
        iLimitStr   =""
        priceStr    = ""
        
        iSellyn = vArr2(13,i)
        iLimityn = vArr2(14,i)
        iLimitNo = vArr2(15,i)-vArr2(16,i)
        if (iLimitNo<1) then iLimitNo=0
            
        
        if (iSellyn<>"Y") then isellStr="<strong><font color='#FF0000'>ǰ��</font></strong>"
        if (iSellyn="S") then isellStr="<strong><font color='#CC3333'>�Ͻ�ǰ��</font></strong>"
        if (iLimityn="Y") then iLimitStr="<font color='#3333CC'>����("&iLimitNo&")</font>"
            
        %>
        <tr bgcolor="#FFFFFF" align="right">
            <td align="left"><%=vArr2(11,i)%></td>
            <td align="left"><%=vArr2(12,i)%></td>
            <td><%=FormatNumber(vArr2(1,i),0)%></td>
            <td><%=FormatNumber(vArr2(5,i),0)%></td>
            <td><%=FormatNumber(vArr2(9,i),0)%></td>
            <td align="left"><%=isellStr%><%if(iLimitStr<>"")then response.write "<br>"&iLimitStr%></td>
        </tr>
        <% next %>
        <tr bgcolor="#F4F4F4" align="right">
            <td align="left">�հ�</td>
            <td align="left"></td>
            <td><%=FormatNumber(vSum1,0)%></td>
            <td><%=FormatNumber(vSum2,0)%></td>
            <td><%=FormatNumber(vSum3,0)%></td>
            <td></td>
        </tr>
        </table>
    </td>
    <% end if %>
	<td valign="top">
		<div id="chart-container1">FusionCharts will render here</div>
	</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->