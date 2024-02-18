<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ�� �����߼�
' History : 2019.04.15 ������ ����
'			2022.10.07 �ѿ�� ����(��������, ǥ���ڵ�� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<!-- #include virtual="/admin/dataanalysis/chart/chartCls.asp" -->
<!-- #include virtual="/admin/maechul/fusionchart/maechul_class.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim itemid : itemid = requestCheckvar(getNumeric(trim(request("itemid"))),10)
Dim itemoption : itemoption = requestCheckvar(trim(request("itemoption")),4)
Dim vSDate : vSDate = requestCheckvar(trim(request("startdate")),10)
Dim vEDate : vEDate = requestCheckvar(trim(request("enddate")),10)
Dim vOrdType : vOrdType = requestCheckvar(trim(request("ordtype")),32)
Dim i, k, vArr1, vArr2, vArr3, vMakerid, vArrEpNotMakerid, vArrEpNotItemid

'��ǰ�ڵ� ��ȿ�� �˻�
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		itemid = ""
	end if
end if

if (vOrdType="") then vOrdType="S" ''�Ǽ�(C) , �ݾ�(S), ����(G)

If vSDate = "" Then
    vSDate = FormatDate(DateAdd("d",-14,now()),"0000-00-00")
End If
	
If vEDate = "" Then
    vEDate = FormatDate(now(),"0000-00-00")
End If

dim oitem
set oitem = new CItemInfo
oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItemInfo

	if (oitem.FResultCount>0) then
		vMakerid = oitem.FOneItem.FMakerid
	end if
end if


dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if

Dim isOptionExists : isOptionExists = (oitemoption.FResultCount>0)

Dim oChart
SET oChart = new CChart
oChart.FRectSDate = vSDate
oChart.FRectEDate = vEDate
'oChart.FRectChannel = vChannel
'oChart.FRectRdsiteGrp = rdsitegrp
'oChart.FPageSize = CHKIIF(vpValue<>"",100,30)
oChart.FRectOrderType = vOrdType

vArr1 = oChart.fnItemSellTrend_DW(itemid)
vArr2 = oChart.fnGetItemInfoHistory(itemid) 
vArr3 = oChart.fnItemUserAcqTrend_DW(itemid)

Dim iChartSubCaption
Dim ixAxisName : ixAxisName = "��¥"
Dim yAxisName, yAxisName2, yAxisName3

Dim iDataSeriseArr, iDataSetPosArr,  iDataSeriseArr2, iDataSetPosArr2,  iDataSeriseArr3, iDataSetPosArr3
iDataSeriseArr = Array("�ڻ��-Nv����","NV-ep","���޸�")

if (vOrdType="C") then
	iDataSetPosArr = Array(5,8,11)
	yAxisName = "�ֹ��Ǽ�"
elseif (vOrdType="S") then	
	iDataSetPosArr = Array(6,9,12)
	yAxisName = "�����Ѿ�"
elseif (vOrdType="G") then
	iDataSetPosArr = Array(7,10,13)
	yAxisName = "�������"
end if

iDataSeriseArr2 = Array("���ݷα�","���ǸŰ�","���̹�Rank")
iDataSetPosArr2 = Array(2,5,6)
yAxisName2 = "�ǸŰ�"

iDataSeriseArr3 = Array("view","wish","cart")
iDataSetPosArr3 = Array(1,2,3)
yAxisName3 = "�Ǽ�"

Dim iSellStDate 
if (oitem.FResultCount>0) then
	iSellStDate=oitem.FOneItem.FSellStdate
	if isNULL(iSellStDate) then 
		iSellStDate=""
	else
		iSellStDate=LEFT(iSellStDate,10)
	end if
end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>

<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript' src="/js/jsCal/js/jscal2.js"></script>
<script type='text/javascript' src="/js/jsCal/js/lang/ko.js"></script>

<script type='text/javascript'>

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

function showoption(comp){
	var ioptdiv = document.getElementById("idoptlist");
	if (comp.value=="�ɼ�ǥ��"){
		ioptdiv.style.display = "block";
		comp.value="�ɼǼ���";
	}else{
		ioptdiv.style.display = "none";
		comp.value="�ɼ�ǥ��";
	}
}

function jsPopCal(sName){
	var winCal;

	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function goSearch(){
	if(frm1.itemid.value == ""){
		alert("��ǰ�ڵ带 �Է��ϼ���.");
		frm1.itemid.focus();
		return;
	}
	if(isNaN(frm1.itemid.value)){
		alert("��ǰ�ڵ带 ���ڷθ� �Է��ϼ���.");
		frm1.itemid.value = "";
		frm1.itemid.focus();
		return;
	}
	
	frm1.submit();
}
</script>

<% If isArray(vArr1) Then %>
<script type='text/javascript'>//<![CDATA[
window.onload=function(){

	FusionCharts.ready(function () {
		var vstrChart1 = new FusionCharts({
			type: 'msline', //'', 
			renderAt: 'chart-container1',
			width: '1100',
			height: '400',
			dataFormat: 'json',
			dataSource: {
				"chart": {
					"caption": "<%=itemid%> �Ϻ� ä�� ���� �߼�",
					"subCaption": "<%=iChartSubCaption%>",
					"xAxisName": "<%=ixAxisName%>",
					"yAxisName": "<%=yAxisName%>",
					"theme": "fint",
					"showSum": "1",
					"showValues": "<%=CHKIIF(UBound(vArr1,2)>90,"0","1")%>",
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
					"numberScaleValue": "10000",
    				"numberScaleUnit": "��"
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

// using MultiAxisLine 
		var vstrChart2 = new FusionCharts({
			//type: 'msline', //'', 
			type: 'mscombidy2d', //'mscombi2d',
			renderAt: 'chart-container2',
			width: '1100',
			height: '400',
			dataFormat: 'json',
			dataSource: {
				"chart": {
					"caption": "<%=itemid%> �Ϻ� ���� ����α�",
					"subCaption": "<%=iChartSubCaption%>",
					"xAxisName": "<%=ixAxisName%>",
					"pYAxisName": "<%=yAxisName2%>",
					"sYAxisName": "���̹� rank",
					"sYAxisMaxValue": "30",
					"theme": "fint", //fusion
					"showSum": "1",
					"showValues": "0",
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
					"numberScaleValue": "10000",
					"numberScaleUnit": "��"
				},
				// X�� 
				"categories": [
					{
						"category": [
							<%
							If isArray(vArr2) Then
								For i = 0 To UBound(vArr2,2)
									'if (precate<>vArr1(0,i)) then
										Response.Write "{" & vbCrLf
										Response.Write """label"": """&vArr2(0,i)&"""" & vbCrLf
										Response.Write "}"

										''ǰ���Ȱ�� ǥ������.
										if vArr2(3,i)<>"Y" then
											Response.Write ","
											Response.Write "{" & vbCrLf
											Response.Write "	""vline"": ""true""," & vbCrLf
											Response.Write "	""lineposition"": ""0""," & vbCrLf
											''Response.Write "	""color"": ""#6baa01""," & vbCrLf
											Response.Write "	""color"": ""#BBBBBB""," & vbCrLf
											Response.Write "	""labelHAlign"": ""center""," & vbCrLf
											Response.Write "	""labelPosition"": ""0.9""," & vbCrLf
											Response.Write "	""label"": """&vArr2(3,i)&"""," & vbCrLf
											Response.Write "	""dashed"": ""1""" & vbCrLf
											Response.Write "}"
											
										end if

										if (CStr(vArr2(0,i))=CStr(iSellStDate)) then
											Response.Write ","
											Response.Write "{" & vbCrLf
											Response.Write "	""vline"": ""true""," & vbCrLf
											Response.Write "	""lineposition"": ""0""," & vbCrLf
											Response.Write "	""color"": ""#6baa01""," & vbCrLf
											Response.Write "	""labelHAlign"": ""center""," & vbCrLf
											Response.Write "	""labelPosition"": ""0.5""," & vbCrLf
											Response.Write "	""label"": ""�ǸŽ���""," & vbCrLf
											Response.Write "	""dashed"": ""1""" & vbCrLf
											Response.Write "}"
										end if

										' if (vArr2(6,i)>0) then
										' 	Response.Write ","
										' 	Response.Write "{" & vbCrLf
										' 	Response.Write "	""vline"": ""true""," & vbCrLf
										' 	Response.Write "	""lineposition"": ""0""," & vbCrLf
										' 	Response.Write "	""color"": ""#FFaa01""," & vbCrLf
										' 	Response.Write "	""labelHAlign"": ""center""," & vbCrLf
										' 	Response.Write "	""labelPosition"": """&(1-vArr2(6,i)*1.0/100)-0.2&"""," & vbCrLf
										' 	Response.Write "	""label"": """&vArr2(6,i)&"""," & vbCrLf
										' 	Response.Write "	""dashed"": ""1""" & vbCrLf
										' 	Response.Write "}"
										' end if

										If i <> UBound(vArr2,2) Then
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
					<% for k=0 to Ubound(iDataSeriseArr2) %>
					{
						"seriesname": "<%=iDataSeriseArr2(k)%>",
						<% if k=2 then %>
						"parentYAxis": "S",
						"renderAs": "line",
						"lineThickness": "1",
						<% else %>
						"renderAs": "line",
						<% end if %>
						"data": [
							<%
							If isArray(vArr2) Then
								For i = 0 To UBound(vArr2,2)
									'if (vArr1(1,i)=vArr2(0,chrtN-1)) then  ''�귣�尡 ������
									Response.Write "{" & vbCrLf
									Response.Write """value"": """&CHKIIF(vArr2(iDataSetPosArr2(k),i)<0,"",vArr2(iDataSetPosArr2(k),i))&"""" & vbCrLf
									Response.Write "}"
									If i <> UBound(vArr2,2) Then
										Response.Write ","
									End If
									Response.Write vbCrLf
									'end if
								Next
							End If
							%>
						]
					}
					<% if (k<Ubound(iDataSeriseArr2)) then response.write "," %>
					<% next %>
					
				]
			}
		}).render();


		var vstrChart3 = new FusionCharts({
			type: 'msline', //'', 
			renderAt: 'chart-container3',
			width: '1100',
			height: '400',
			dataFormat: 'json',
			dataSource: {
				"chart": {
					"caption": "<%=itemid%> ��ȸ,����,��ٱ��ϰǼ�",
					"subCaption": "<%=iChartSubCaption%>",
					"xAxisName": "<%=ixAxisName%>",
					"yAxisName": "<%=yAxisName3%>",
					"theme": "fint",
					"showSum": "1",
					"showValues": "0",
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
							If isArray(vArr3) Then
								For i = 0 To UBound(vArr3,2)
									'if (precate<>vArr1(0,i)) then
										Response.Write "{" & vbCrLf
										Response.Write """label"": """&vArr3(0,i)&"""" & vbCrLf
										Response.Write "}"
										If i <> UBound(vArr3,2) Then
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
					<% for k=0 to Ubound(iDataSeriseArr3) %>
					{
						"seriesname": "<%=iDataSeriseArr3(k)%>",
						"data": [
							<%
							If isArray(vArr3) Then
								For i = 0 To UBound(vArr3,2)
									'if (vArr1(1,i)=vArr3(0,chrtN-1)) then  ''�귣�尡 ������
									Response.Write "{" & vbCrLf
									Response.Write """value"": """&vArr3(iDataSetPosArr3(k),i)&"""" & vbCrLf
									Response.Write "}"
									If i <> UBound(vArr3,2) Then
										Response.Write ","
									End If
									Response.Write vbCrLf
									'end if
								Next
							End If
							%>
						]
					}
					<% if (k<Ubound(iDataSeriseArr3)) then response.write "," %>
					<% next %>
					
				]
			}
		}).render();
	});

}//]]>
</script>
<% End If %>


<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="#FFFFFF">
	    <td align="center" width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
        <td>
        	��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="9" maxlength="16" onKeyPress="if (event.keyCode == 13) document.frm.submit();">&nbsp;&nbsp;
<% if (FALSE) then %>
        	<% if oitemoption.FResultCount>0 then %>
			&nbsp;
			�ɼǼ��� :
			<select class="select" name="itemoption">
				<option  value="">----
				<% for i=0 to oitemoption.FResultCount-1 %>
				<option value="<%= oitemoption.FITemList(i).FItemOption %>" <% if itemoption=oitemoption.FITemList(i).FItemOption then response.write "selected" %> ><%= oitemoption.FITemList(i).FOptionName %>
				<% next %>
				</select>
			<% end if %>
<% end if %>
			&nbsp;
			�˻��Ⱓ(�ֹ���) :

			<input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
			<img src='/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
			~<input id='enddate' name='enddate' value='<%= vEDate %>' class='text' size='10' maxlength='10' />
			<img src='/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
			
			
			&nbsp;

            <input type="radio" name="ordtype" value="C" <%=CHKIIF(vOrdType="C","checked","") %> >�ֹ��Ǽ�
            <input type="radio" name="ordtype" value="S" <%=CHKIIF(vOrdType="S","checked","") %> >�����Ѿ�
            <input type="radio" name="ordtype" value="G" <%=CHKIIF(vOrdType="G","checked","") %> >�������
			
			
            
        </td>
        <td align="center" width="50" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
</table>
</form>
<!-- �˻� �� -->
<br>
<% if (oitem.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=5 width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
      	<td width="60" bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
      	<td width="35%">
      		10 <b><%= CHKIIF(oitem.FOneItem.FItemID>=1000000,Format00(8,oitem.FOneItem.FItemID),Format00(6,oitem.FOneItem.FItemID)) %></b> <%= itemoption %>
      	</td>
      	<td width="60" bgcolor="<%= adminColor("tabletop") %>">�ǸŽ�����</td>
      	<td><%=iSellStDate%></td>
		<td align="right">
		<input type="button" value="��ǰ���� web"  onClick="window.open('http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=ItemID%>','_viewitem','');">

		<input type="button" value="���ݺ���LOG"  onClick="window.open('/admin/etc/extsitejungsan_check.asp?itemid=<%=ItemID%>','_itemlog','');">

		<input type="button" value="���޼���LOG" onClick="window.open('/admin/etc/outmall/index.asp?research=on&menupos=1742&makerid=<%=vMakerid%>','_outmallsellyn','');">
		</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">�귣��ID</td>
      	<td><%= oitem.FOneItem.FMakerid %></td>
      	<td bgcolor="<%= adminColor("tabletop") %>">�Ǹſ���</td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FSellyn) %>"><%= oitem.FOneItem.FSellyn %></font></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
      	<td><%= oitem.FOneItem.FItemName %></td>
      	<td bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FIsUsing) %>"><%= oitem.FOneItem.FIsUsing %></font></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">�ǸŰ�</td>
      	<td>
      		<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
      		&nbsp;&nbsp;
      		<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
      	    <% if oitem.FOneItem.FSellcash<>0 then %>
			<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
			<% end if %>
			&nbsp;&nbsp;
			<!-- ���ο���/�������뿩�� -->
			<% if (oitem.FOneItem.FSailYn="Y") then %>
			    <font color=red>
			    <% if (oitem.FOneItem.Forgprice<>0) then %>
			        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %> %
			    <% end if %>
			     ����
			    </font>
			<% end if %>

			<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>

			    <font color=green><%= oitem.FOneItem.GetCouponDiscountStr %> ����
			    (<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>)</font>
			<% end if %>

      	</td>
      	<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
      	<td colspan=2>
      		<% if oitem.FOneItem.Fdanjongyn="Y" then %>
			<font color="#33CC33">����</font>
			<% elseif oitem.FOneItem.Fdanjongyn="S" then %>
			<font color="#33CC33">�Ͻ�ǰ��</font>
			<% else %>
			������
			<% end if %>
		</td>
    </tr>
    
    <tr bgcolor="#FFFFFF">
        <td bgcolor="<%= adminColor("tabletop") %>">�ɼ�</td>
        <td><%=CHKIIF(isOptionExists,"��"&oitemoption.FResultCount,"-")%>
		<% if (isOptionExists) then %>
		&nbsp;&nbsp;<input type="button" value="�ɼ�ǥ��" onClick="showoption(this);">
		<% end if %>
		</td>
        <td bgcolor="<%= adminColor("tabletop") %>">��������</td>
        <td><font color="<%= ynColor(oitem.FOneItem.Flimityn) %>"><%= oitem.FOneItem.Flimityn %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
        <td>���� ����� (<b><%= oitem.FOneItem.GetLimitStockNo %></b>)</td>
    </tr>
</table>
<div id="idoptlist" name="idoptlist" style="display:none">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" >
    <% if oitemoption.FResultCount>1 then %>
	    <% for i=0 to oitemoption.FResultCount -1 %>
            <% if oitemoption.FITemList(i).FOptIsUsing<>"Y" then %>
		    <tr bgcolor="#FFFFFF">
                <% if (i=0) then %><td rowspan="<%=oitemoption.FResultCount%>" width="110" ></td><% end if %>
		      	<td bgcolor="<%= adminColor("tabletop") %>" width="60"><font color="#AAAAAA">�ɼǸ� :</font></td>
		      	<td width="35%"><font color="#AAAAAA"><%= oitemoption.FITemList(i).FOptionName %></font></td>
		      	<td bgcolor="<%= adminColor("tabletop") %>" width="60"><font color="#AAAAAA">�������� : </font></td>
		      	<td ><font color="#AAAAAA"><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</font></td>
		      	<td width="120">���� ����� (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
		    </tr>
		    <% else %>
				<tr bgcolor="<%=CHKIIF(oitemoption.FITemList(i).Fitemoption=itemoption,"#EEEEEE","#FFFFFF")%>">
					<% if (i=0) then %><td rowspan="<%=oitemoption.FResultCount%>" width="110" ></td><% end if %>
					<td width="60">�ɼǸ�</td>
					<td width="35%"><%= oitemoption.FITemList(i).FOptionName %></td>
					<td width="60">��������</td>
					<td ><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
					<td width="120">���� ����� (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
				</tr>
		    <% end if %>
	    <% next %>
    <% end if %>
</table>
</div>
<% end if %>
<%
SET oitem = Nothing
SET oitemoption = Nothing
%>
<br />

<% If ItemID <> "" Then %>
<table cellpadding="0" cellspacing="0" border="0" class="a" align="center">
<tr bgcolor="#FFFFFF">
	<td>
		<div id="chart-container1" style="text-align:center;">FusionCharts will render here</div>
		<br />
		<div id="chart-container2" style="text-align:center;">FusionCharts will render here</div>
		<br />
		<div id="chart-container3" style="text-align:center;">FusionCharts will render here</div>
	</td>
</tr>
</table>
<% End If %>

<% If ItemID <> "" Then %>
	<table cellpadding="3" cellspacing="1" border="0" class="a" align="center" width="1200" bgcolor="#444444">
	<% if IsArray(vArrEpNotMakerid) then %>
	<tr bgcolor="#FFFFFF">
		<td width="100">EP���� ���� ó��<br>by �귣��</td>
		<td>
				<table cellpadding="2" cellspacing="1" border="0" class="a" align="center" width="95%" bgcolor="#444444">
					<tr bgcolor="#EEEEEE">
					<td width="120">�귣��ID</td><td width="100">mall</td><td width="100">����</td><td width="120">�����</td><td  width="120">����������</td><td  width="120">�����</td><td>���������</td>
					</tr>
				<% for i=0 to uBound(vArrEpNotMakerid,2) %>
					<tr bgcolor="#FFFFFF">
					<td><%= vArrEpNotMakerid(0,i) %></td><td><%= vArrEpNotMakerid(1,i) %></td>
					<td><%=CHKIIF(vArrEpNotMakerid(2,i)="N","<font color=red>����</font>","OK")%></td>
					<td><%= vArrEpNotMakerid(3,i) %></td><td><%= vArrEpNotMakerid(4,i) %></td><td><%= vArrEpNotMakerid(5,i) %></td><td><%= vArrEpNotMakerid(6,i) %></td>
					</tr>
				<% next %>
				</table>
		</td>
		<td></td>
	</tr> 
	<% end if %>

	<% if IsArray(vArrEpNotItemid) then %>
	<tr bgcolor="#FFFFFF">   
		<td width="100">EP���� ���� ó��<br>by ��ǰ</td>
		<td>    
				<table cellpadding="2" cellspacing="1" border="0" class="a" align="center" width="95%" bgcolor="#444444">
					<tr bgcolor="#EEEEEE">
					<td width="120">��ǰ�ڵ�</td><td width="100">mall</td><td width="100">����</td><td  width="120">�����</td><td  width="120">����������</td><td  width="120">�����</td><td>���������</td>
					</tr>
				<% for i=0 to uBound(vArrEpNotItemid,2) %>
					<tr bgcolor="#FFFFFF">
					<td><%= vArrEpNotItemid(0,i) %></td><td><%= vArrEpNotItemid(1,i) %></td>
					<td><%=CHKIIF(vArrEpNotItemid(2,i)="Y","<font color=red>����</font>","OK")%></td>
					<td><%= vArrEpNotItemid(3,i) %></td><td><%= vArrEpNotItemid(4,i) %></td><td><%= vArrEpNotItemid(5,i) %></td><td><%= vArrEpNotItemid(6,i) %></td>
					</tr>
				<% next %>
				</table>
		</td>
		<td>
		<input type="button" value="EP���ܼ���" onClick="window.open('/admin/etc/naverEp/notinitemid.asp?page=1&research=on&menupos=1614&itemid=<%=vItemID%>','_epitemsellyn','');"> 
		</td>
	</tr>
	</table>
	<% end if %>
<% End If %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->