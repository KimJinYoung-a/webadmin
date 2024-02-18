<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품별 매출추세
' History : 2019.04.15 서동석 생성
'			2022.10.07 한용민 수정(오류수정, 표준코드로 변경)
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

'상품코드 유효성 검사
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		itemid = ""
	end if
end if

if (vOrdType="") then vOrdType="S" ''건수(C) , 금액(S), 수익(G)

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
Dim ixAxisName : ixAxisName = "날짜"
Dim yAxisName, yAxisName2, yAxisName3

Dim iDataSeriseArr, iDataSetPosArr,  iDataSeriseArr2, iDataSetPosArr2,  iDataSeriseArr3, iDataSetPosArr3
iDataSeriseArr = Array("자사몰-Nv제외","NV-ep","제휴몰")

if (vOrdType="C") then
	iDataSetPosArr = Array(5,8,11)
	yAxisName = "주문건수"
elseif (vOrdType="S") then	
	iDataSetPosArr = Array(6,9,12)
	yAxisName = "구매총액"
elseif (vOrdType="G") then
	iDataSetPosArr = Array(7,10,13)
	yAxisName = "매출수익"
end if

iDataSeriseArr2 = Array("가격로그","실판매가","네이버Rank")
iDataSetPosArr2 = Array(2,5,6)
yAxisName2 = "판매가"

iDataSeriseArr3 = Array("view","wish","cart")
iDataSetPosArr3 = Array(1,2,3)
yAxisName3 = "건수"

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
	if (comp.value=="옵션표시"){
		ioptdiv.style.display = "block";
		comp.value="옵션숨김";
	}else{
		ioptdiv.style.display = "none";
		comp.value="옵션표시";
	}
}

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
					"caption": "<%=itemid%> 일별 채널 매출 추세",
					"subCaption": "<%=iChartSubCaption%>",
					"xAxisName": "<%=ixAxisName%>",
					"yAxisName": "<%=yAxisName%>",
					"theme": "fint",
					"showSum": "1",
					"showValues": "<%=CHKIIF(UBound(vArr1,2)>90,"0","1")%>",
					//Setting automatic calculation of div lines to off
	//              "adjustDiv": "0",
					//Manually defining y-axis lower and upper limit
					//"yAxisMaxvalue": "35000",	//y축 맥스값
					//"yAxisMinValue": "5000",		//y축 민값
					//Setting number of divisional lines to 9
					//"numDivLines": "9"				//0~맥스 사이 표시되어지는 수치갯수
	//              "anchorBgHoverColor": "#96d7fa",
	//              "anchorBorderHoverThickness" : "4",
	//              "anchorHoverRadius":"7"
					"numberScaleValue": "10000",
    				"numberScaleUnit": "만"
				},
				// X축 
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
									'if (vArr1(1,i)=vArr2(0,chrtN-1)) then  ''브랜드가 같으면
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
					"caption": "<%=itemid%> 일별 가격 변경로그",
					"subCaption": "<%=iChartSubCaption%>",
					"xAxisName": "<%=ixAxisName%>",
					"pYAxisName": "<%=yAxisName2%>",
					"sYAxisName": "네이버 rank",
					"sYAxisMaxValue": "30",
					"theme": "fint", //fusion
					"showSum": "1",
					"showValues": "0",
					//Setting automatic calculation of div lines to off
	//              "adjustDiv": "0",
					//Manually defining y-axis lower and upper limit
					//"yAxisMaxvalue": "35000",	//y축 맥스값
					//"yAxisMinValue": "5000",		//y축 민값
					//Setting number of divisional lines to 9
					//"numDivLines": "9"				//0~맥스 사이 표시되어지는 수치갯수
	//              "anchorBgHoverColor": "#96d7fa",
	//              "anchorBorderHoverThickness" : "4",
	//              "anchorHoverRadius":"7"
					"numberScaleValue": "10000",
					"numberScaleUnit": "만"
				},
				// X축 
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

										''품절된경우 표시하자.
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
											Response.Write "	""label"": ""판매시작""," & vbCrLf
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
									'if (vArr1(1,i)=vArr2(0,chrtN-1)) then  ''브랜드가 같으면
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
					"caption": "<%=itemid%> 조회,위시,장바구니건수",
					"subCaption": "<%=iChartSubCaption%>",
					"xAxisName": "<%=ixAxisName%>",
					"yAxisName": "<%=yAxisName3%>",
					"theme": "fint",
					"showSum": "1",
					"showValues": "0",
					//Setting automatic calculation of div lines to off
	//              "adjustDiv": "0",
					//Manually defining y-axis lower and upper limit
					//"yAxisMaxvalue": "35000",	//y축 맥스값
					//"yAxisMinValue": "5000",		//y축 민값
					//Setting number of divisional lines to 9
					//"numDivLines": "9"				//0~맥스 사이 표시되어지는 수치갯수
	//              "anchorBgHoverColor": "#96d7fa",
	//              "anchorBorderHoverThickness" : "4",
	//              "anchorHoverRadius":"7"
				},
				// X축 
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
									'if (vArr1(1,i)=vArr3(0,chrtN-1)) then  ''브랜드가 같으면
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


<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="#FFFFFF">
	    <td align="center" width="50" bgcolor="#EEEEEE">검색<br>조건</td>
        <td>
        	상품코드 : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="9" maxlength="16" onKeyPress="if (event.keyCode == 13) document.frm.submit();">&nbsp;&nbsp;
<% if (FALSE) then %>
        	<% if oitemoption.FResultCount>0 then %>
			&nbsp;
			옵션선택 :
			<select class="select" name="itemoption">
				<option  value="">----
				<% for i=0 to oitemoption.FResultCount-1 %>
				<option value="<%= oitemoption.FITemList(i).FItemOption %>" <% if itemoption=oitemoption.FITemList(i).FItemOption then response.write "selected" %> ><%= oitemoption.FITemList(i).FOptionName %>
				<% next %>
				</select>
			<% end if %>
<% end if %>
			&nbsp;
			검색기간(주문일) :

			<input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
			<img src='/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
			~<input id='enddate' name='enddate' value='<%= vEDate %>' class='text' size='10' maxlength='10' />
			<img src='/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
			
			
			&nbsp;

            <input type="radio" name="ordtype" value="C" <%=CHKIIF(vOrdType="C","checked","") %> >주문건수
            <input type="radio" name="ordtype" value="S" <%=CHKIIF(vOrdType="S","checked","") %> >구매총액
            <input type="radio" name="ordtype" value="G" <%=CHKIIF(vOrdType="G","checked","") %> >매출수익
			
			
            
        </td>
        <td align="center" width="50" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
<% if (oitem.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
    	<td rowspan=5 width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
      	<td width="60" bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
      	<td width="35%">
      		10 <b><%= CHKIIF(oitem.FOneItem.FItemID>=1000000,Format00(8,oitem.FOneItem.FItemID),Format00(6,oitem.FOneItem.FItemID)) %></b> <%= itemoption %>
      	</td>
      	<td width="60" bgcolor="<%= adminColor("tabletop") %>">판매시작일</td>
      	<td><%=iSellStDate%></td>
		<td align="right">
		<input type="button" value="상품보기 web"  onClick="window.open('http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%=ItemID%>','_viewitem','');">

		<input type="button" value="가격변동LOG"  onClick="window.open('/admin/etc/extsitejungsan_check.asp?itemid=<%=ItemID%>','_itemlog','');">

		<input type="button" value="제휴설정LOG" onClick="window.open('/admin/etc/outmall/index.asp?research=on&menupos=1742&makerid=<%=vMakerid%>','_outmallsellyn','');">
		</td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">브랜드ID</td>
      	<td><%= oitem.FOneItem.FMakerid %></td>
      	<td bgcolor="<%= adminColor("tabletop") %>">판매여부</td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FSellyn) %>"><%= oitem.FOneItem.FSellyn %></font></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">상품명</td>
      	<td><%= oitem.FOneItem.FItemName %></td>
      	<td bgcolor="<%= adminColor("tabletop") %>">사용여부</td>
      	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FIsUsing) %>"><%= oitem.FOneItem.FIsUsing %></font></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">판매가</td>
      	<td>
      		<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
      		&nbsp;&nbsp;
      		<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
      	    <% if oitem.FOneItem.FSellcash<>0 then %>
			<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
			<% end if %>
			&nbsp;&nbsp;
			<!-- 할인여부/쿠폰적용여부 -->
			<% if (oitem.FOneItem.FSailYn="Y") then %>
			    <font color=red>
			    <% if (oitem.FOneItem.Forgprice<>0) then %>
			        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %> %
			    <% end if %>
			     할인
			    </font>
			<% end if %>

			<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>

			    <font color=green><%= oitem.FOneItem.GetCouponDiscountStr %> 쿠폰
			    (<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>)</font>
			<% end if %>

      	</td>
      	<td bgcolor="<%= adminColor("tabletop") %>">단종여부</td>
      	<td colspan=2>
      		<% if oitem.FOneItem.Fdanjongyn="Y" then %>
			<font color="#33CC33">단종</font>
			<% elseif oitem.FOneItem.Fdanjongyn="S" then %>
			<font color="#33CC33">일시품절</font>
			<% else %>
			생산중
			<% end if %>
		</td>
    </tr>
    
    <tr bgcolor="#FFFFFF">
        <td bgcolor="<%= adminColor("tabletop") %>">옵션</td>
        <td><%=CHKIIF(isOptionExists,"총"&oitemoption.FResultCount,"-")%>
		<% if (isOptionExists) then %>
		&nbsp;&nbsp;<input type="button" value="옵션표시" onClick="showoption(this);">
		<% end if %>
		</td>
        <td bgcolor="<%= adminColor("tabletop") %>">한정여부</td>
        <td><font color="<%= ynColor(oitem.FOneItem.Flimityn) %>"><%= oitem.FOneItem.Flimityn %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
        <td>한정 비교재고 (<b><%= oitem.FOneItem.GetLimitStockNo %></b>)</td>
    </tr>
</table>
<div id="idoptlist" name="idoptlist" style="display:none">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" >
    <% if oitemoption.FResultCount>1 then %>
	    <% for i=0 to oitemoption.FResultCount -1 %>
            <% if oitemoption.FITemList(i).FOptIsUsing<>"Y" then %>
		    <tr bgcolor="#FFFFFF">
                <% if (i=0) then %><td rowspan="<%=oitemoption.FResultCount%>" width="110" ></td><% end if %>
		      	<td bgcolor="<%= adminColor("tabletop") %>" width="60"><font color="#AAAAAA">옵션명 :</font></td>
		      	<td width="35%"><font color="#AAAAAA"><%= oitemoption.FITemList(i).FOptionName %></font></td>
		      	<td bgcolor="<%= adminColor("tabletop") %>" width="60"><font color="#AAAAAA">한정여부 : </font></td>
		      	<td ><font color="#AAAAAA"><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</font></td>
		      	<td width="120">한정 비교재고 (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
		    </tr>
		    <% else %>
				<tr bgcolor="<%=CHKIIF(oitemoption.FITemList(i).Fitemoption=itemoption,"#EEEEEE","#FFFFFF")%>">
					<% if (i=0) then %><td rowspan="<%=oitemoption.FResultCount%>" width="110" ></td><% end if %>
					<td width="60">옵션명</td>
					<td width="35%"><%= oitemoption.FITemList(i).FOptionName %></td>
					<td width="60">한정여부</td>
					<td ><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
					<td width="120">한정 비교재고 (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
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
		<td width="100">EP연동 예외 처리<br>by 브랜드</td>
		<td>
				<table cellpadding="2" cellspacing="1" border="0" class="a" align="center" width="95%" bgcolor="#444444">
					<tr bgcolor="#EEEEEE">
					<td width="120">브랜드ID</td><td width="100">mall</td><td width="100">연동</td><td width="120">등록일</td><td  width="120">최종수정일</td><td  width="120">등록자</td><td>최종등록자</td>
					</tr>
				<% for i=0 to uBound(vArrEpNotMakerid,2) %>
					<tr bgcolor="#FFFFFF">
					<td><%= vArrEpNotMakerid(0,i) %></td><td><%= vArrEpNotMakerid(1,i) %></td>
					<td><%=CHKIIF(vArrEpNotMakerid(2,i)="N","<font color=red>안함</font>","OK")%></td>
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
		<td width="100">EP연동 예외 처리<br>by 상품</td>
		<td>    
				<table cellpadding="2" cellspacing="1" border="0" class="a" align="center" width="95%" bgcolor="#444444">
					<tr bgcolor="#EEEEEE">
					<td width="120">상품코드</td><td width="100">mall</td><td width="100">연동</td><td  width="120">등록일</td><td  width="120">최종수정일</td><td  width="120">등록자</td><td>최종등록자</td>
					</tr>
				<% for i=0 to uBound(vArrEpNotItemid,2) %>
					<tr bgcolor="#FFFFFF">
					<td><%= vArrEpNotItemid(0,i) %></td><td><%= vArrEpNotItemid(1,i) %></td>
					<td><%=CHKIIF(vArrEpNotItemid(2,i)="Y","<font color=red>안함</font>","OK")%></td>
					<td><%= vArrEpNotItemid(3,i) %></td><td><%= vArrEpNotItemid(4,i) %></td><td><%= vArrEpNotItemid(5,i) %></td><td><%= vArrEpNotItemid(6,i) %></td>
					</tr>
				<% next %>
				</table>
		</td>
		<td>
		<input type="button" value="EP예외설정" onClick="window.open('/admin/etc/naverEp/notinitemid.asp?page=1&research=on&menupos=1614&itemid=<%=vItemID%>','_epitemsellyn','');"> 
		</td>
	</tr>
	</table>
	<% end if %>
<% End If %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->