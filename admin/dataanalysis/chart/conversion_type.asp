<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/admin/dataanalysis/chart/chartCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim oChart, oChart2, vArr1, vArr2, i, j
Dim vSDate, vEDate, vChannel, sTp

vSDate = requestCheckvar(request("sdate"),10)
vEDate = requestCheckvar(request("edate"),10)
vChannel = requestCheckvar(request("channel"),10)
sTp  = requestCheckvar(request("sTp"),10)

if (sTp="") then sTp="1" ''�Ǽ�(1) , �ݾ�(2)
    
If vSDate = "" Then
	vSDate = dateadd("m",-1,Date())
End If

If vEDate = "" Then
	vEDate = date()
End If

SET oChart = new CChart
	oChart.FRectSDate = vSDate
	oChart.FRectEDate = vEDate
	oChart.FRectChannel = vChannel
	vArr1 = oChart.fnDayChannelByType
SET oChart = nothing

''2019/08/08 �Ϻ����� �߰�.
SET oChart2 = new CChart
	oChart2.FRectSDate = vSDate
	oChart2.FRectEDate = vEDate
	oChart2.FRectChannel = vChannel

	vArr2 = oChart2.fnDailyMeachul_vs_Conversion_DW
SET oChart2 = nothing

Dim iChartCaption : iChartCaption = "��ȯ Ÿ�Ժ� �ֹ��Ǽ�"
Dim iChartSubCaption : iChartSubCaption = ""
Dim ixAxisName : ixAxisName = "" ''��¥
Dim yAxisName : yAxisName = "�ֹ��Ǽ�"
Dim iDataSetPosArr : iDataSetPosArr = Array(1,2,3,4,5,6,7)
Dim iDataSetHeadArr : iDataSetHeadArr = Array("�ֹ��Ǽ�-�˻�","�ֹ��Ǽ�-ī�װ�","�ֹ��Ǽ�-�귣��","�ֹ��Ǽ�-�̺�Ʈ","�ֹ��Ǽ�-rc","�ֹ��Ǽ�-gaparam","�ֹ��Ǽ�-rdsitedirect")

if (sTp="2") then
    iDataSetPosArr = Array(8,9,10,11,12,13,14)
    iDataSetHeadArr = Array("�����Ѿ�-�˻�","�����Ѿ�-ī�װ�","�����Ѿ�-�귣��","�����Ѿ�-�̺�Ʈ","�����Ѿ�-rc","�����Ѿ�-gaparam","�����Ѿ�-rdsitedirect")
    
    iChartCaption = "��ȯ Ÿ�Ժ� �����Ѿ�"
    yAxisName = "�����Ѿ�"
end if

Dim iDataSetPosArr2 : iDataSetPosArr2 = Array(0,1) 
Dim iDataSetHeadArr2 : iDataSetHeadArr2 = Array("��¥","�ֹ��Ǽ�-��ü")  

if (sTp="2") then
	iDataSetPosArr2 = Array(0,2) 
	iDataSetHeadArr2 = Array("��¥","�����Ѿ�-��ü")  
end if
dim SumArr()
redim SumArr(UBound(iDataSetPosArr2))
dim SumArrType : SumArrType = Array(9,0) 
Dim font_html1, font_html2, k
%>
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
				<% if isArray(vArr2) then %>
					{
						"seriesname": "<%=iDataSetHeadArr2(1)%>",
						"data": [
							<%
							For i = 0 To UBound(vArr2,2)
								Response.Write "{" & vbCrLf
								Response.Write """value"": """&vArr2(iDataSetPosArr2(1),i)&"""" & vbCrLf
								Response.Write "}"
								If i <> UBound(vArr2,2) Then
									Response.Write ","
								End If
								Response.Write vbCrLf
							Next
							%>
						]
					},
				<% end if %>
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
	��ȸ��¥ : 
		<input type="text" name="sdate" id="sdate" value="<%=vSDate%>" onClick="jsPopCal('sdate');" style="cursor:pointer;" size="10" maxlength="10" readonly>&nbsp;~&nbsp;
		<input type="text" name="edate" id="edate" value="<%=vEDate%>" onClick="jsPopCal('edate');" style="cursor:pointer;" size="10" maxlength="10" readonly>
    &nbsp;&nbsp;
    
    ä�� :
    <select name="channel" >
        <option value="" <%=CHKIIF(vChannel="","selected","")%>>ALL</option>
        <option value="pc" <%=CHKIIF(vChannel="pc","selected","")%>>WEB</option>
        <option value="mw" <%=CHKIIF(vChannel="mw","selected","")%>>MOB</option>
        <option value="app" <%=CHKIIF(vChannel="app","selected","")%>>APP</option>
    </select>
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
* �ڻ�� ����, �ֹ����� ����, ��ǰ ��ȯ ���� (��ȯŸ�Ժ� ���� ��ü����� ��ġ������ �ʽ��ϴ�.)
<br />
<p>
<table cellpadding="0" cellspacing="0" border="0" class="a" align="center">
<tr bgcolor="#FFFFFF">
	<td style="vertical-align:top;">
		<table cellpadding="2" cellspacing="1" border="0" class="a" align="center" style="width:300px;" bgcolor="#999999">
<%
	Dim ArrLength : ArrLength = Ubound(iDataSetHeadArr)
	Dim SumDataArr
	ReDim SumDataArr(ArrLength)
	Dim sumData : sumData = 0
	Dim totalRevenue : totalRevenue = 0

	For i = 0 To ArrLength

		If isArray(vArr1) Then
			For j = 0 To UBound(vArr1,2)
				SumDataArr(i) = SumDataArr(i) + CDbl(vArr1(iDataSetPosArr(i), j))
			Next
		End If

		totalRevenue = totalRevenue + SumDataArr(i)

	Next
%>
<%
	For i = Lbound(iDataSetHeadArr) To Ubound(iDataSetHeadArr)
%>
			<tr bgcolor="#FFFFFF">
				<td><%=iDataSetHeadArr(i)%></td>
				<td style="text-align:right;"><%=CurrFormat(SumDataArr(i))%></td>
				<td style="text-align:right;">
					<%
						If totalRevenue > 0 Then
							Response.Write FormatPercent(SumDataArr(i) / totalRevenue, 2)
						End If
					%>
				</td>
			</tr>
<%
	Next
%>
			<tr bgcolor="#FFFFFF">
				<td style="text-align:right;">total : </td>
				<td style="text-align:right;"><%=CurrFormat(totalRevenue)%></td>
				<td style="text-align:right;"></td>
			</tr>
		</table>

		<p>

<% IF(FALSE) then %>
		<% If isArray(vArr2) Then %>
		<table cellpadding="3" cellspacing="1" class="a" align="center" bgcolor="#999999">
        <tr bgcolor="#F4F4F4">
            <% for k=Lbound(iDataSetHeadArr2) to Ubound(iDataSetHeadArr2) %>
            <td><%=iDataSetHeadArr2(k)%></td>
            <% next %>
        </tr> 
        <% For i = 0 To UBound(vArr2,2) %>
        <tr bgcolor="#FFFFFF" align="right">
           
            <% for k=Lbound(iDataSetPosArr2) to Ubound(iDataSetPosArr2) %>
            <% 
                if SumArrType(k)=0 then
                 '   SumArr(k)=SumArr(k)+CDBL(vArr1(iDataSetPosArr2(k),i)) 
                end if
                
                if (k=0) then
                    ' if datepart("w",vArr1(iDataSetPosArr2(k),i))=1 then
                    '     font_html1 = "<font color='red'>"
                    '     font_html2 = "</font>"
                    ' elseif datepart("w",vArr1(iDataSetPosArr2(k),i))=7 then
                    '     font_html1 = "<font color='blue'>"
                    '     font_html2 = "</font>"
                    ' end if
                end if
            %>
            <td >
                <%=font_html1%>
                <% 
                if SumArrType(k)<=1 then
                    response.write FormatNumber(vArr2(iDataSetPosArr2(k),i),0)
                else
                    response.write vArr2(iDataSetPosArr2(k),i)
                end if
                %>
                <%=font_html2%>
            </td>
            <% next %>
        </tr>
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
		<% end if %>
<% end if %>
	</td>
	<td>
		<div id="chart-container1">FusionCharts will render here</div>
	</td>
</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->