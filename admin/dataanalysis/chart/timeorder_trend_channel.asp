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

if (sTp="") then sTp="2" ''�Ǽ�(1) , �ݾ�(2)
if (dTp="") then dTp="d" ''�ð�(h) , �ϰ�(d) , �ְ� (w), ���� (m)
     
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

Dim iChartCaption : iChartCaption = "ä�κ� �ֹ��Ǽ�"
Dim iChartSubCaption : iChartSubCaption = ""
Dim ixAxisName : ixAxisName = "" ''��¥
Dim yAxisName : yAxisName = "�ֹ��Ǽ�"
Dim iDataSetPosArr : iDataSetPosArr = Array(4,6,8,10) 
Dim iDataSetHeadArr : iDataSetHeadArr = Array("�ֹ��Ǽ�-Pc","�ֹ��Ǽ�-Mob","�ֹ��Ǽ�-App","�ֹ��Ǽ�-Out") 

''Dim iDataSetPosArr2 : iDataSetPosArr2 = Array(2,6,4,10) 
''Dim iDataSetHeadArr2 : iDataSetHeadArr2 = Array("�ֹ��Ǽ�-�ݳ�","��������","�ֹ��Ǽ�-����","�����<br>(%)") 

if (sTp="2") then
    iDataSetPosArr = Array(5,7,9,11) '',8
    iDataSetHeadArr = Array("�����Ѿ�-Pc","�����Ѿ�-Mob","�����Ѿ�-App","�����Ѿ�-Out")
    
    ''iDataSetPosArr2 = Array(3,7,5,11) '',8
    ''iDataSetHeadArr2 = Array("�����Ѿ�-�ݳ�","��������","�����Ѿ�-����","�����<br>(%)")
    
    iChartCaption = "ä�κ� �����Ѿ�"
    yAxisName = "�����Ѿ�"
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
    <td width="50" rowspan="2" bgcolor="#EEEEEE">�˻�<br>����</td>
	<td align="left">
	��ȸ��¥(�ֹ���) : 
	    
	    <input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
		~<input id='enddate' name='enddate' value='<%= vEDate %>' class='text' size='10' maxlength='10' />
		<img src='http://webadmin.10x10.co.kr/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
			
	    
    &nbsp;&nbsp;
    
    ä�� :
    <select name="channel" >
        <option value="" <%=CHKIIF(vChannel="","selected","")%>>ALL</option>
        <option value="pc" <%=CHKIIF(vChannel="pc","selected","")%>>WEB</option>
        <option value="mw" <%=CHKIIF(vChannel="mw","selected","")%>>MOB</option>
        <option value="app" <%=CHKIIF(vChannel="app","selected","")%>>APP</option>
    </select>
    
    &nbsp;&nbsp;
    <input type="radio" name="dTp" value="h" <%=CHKIIF(dTp="h","checked","") %> >�ð���
    <input type="radio" name="dTp" value="d" <%=CHKIIF(dTp="d","checked","") %> >�Ϻ�
    <input type="radio" name="dTp" value="m" <%=CHKIIF(dTp="m","checked","") %> >����
    
    &nbsp;&nbsp;
    <input type="radio" name="sTp" value="1" <%=CHKIIF(sTp="1","checked","") %> >�ֹ��Ǽ�
    <input type="radio" name="sTp" value="2" <%=CHKIIF(sTp="2","checked","") %> >�����Ѿ�
    
    </td>
    <td width="50" rowspan="2" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="�˻�" onClick="goSearch(document.frm1);">
	</td>
</tr>

</table>
</form>
<br />
* �� 1�ð� ����������
* ��ǰ ��ȯ���� ���Ե��� ����
* ����,�ؿ�,3pl�� ���Ե��� ����
* ������ ���� ���� �ֹ� ���Ե�(���� ��ҵ� �� ����)
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