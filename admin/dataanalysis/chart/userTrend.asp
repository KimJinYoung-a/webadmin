<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<!-- #include virtual="/admin/dataanalysis/chart/chartCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
''<!-- #include virtual="/admin/maechul/fusionchart/maechul_class.asp" -->
Dim ulevel : ulevel = requestCheckvar(request("ulevel"),1)
Dim vSDate : vSDate = requestCheckvar(request("startdate"),10)
Dim vEDate : vEDate = requestCheckvar(request("enddate"),10)
Dim grptp : grptp = requestCheckvar(request("grptp"),10)

Dim vChannel : vChannel = requestCheckvar(request("channel"),10)
Dim vOrdType : vOrdType = requestCheckvar(request("ordtype"),32)
Dim i, k, vList1, vArr1, vArr2, vArr3, vMakerid, vArrEpNotMakerid, vArrEpNotItemid

if (vOrdType="") then vOrdType="S" ''건수(C) , 금액(S), 수익(G)

If vSDate = "" Then
    vSDate = FormatDate(DateAdd("d",-14,now()),"0000-00-00")
End If
	
If vEDate = "" Then
    vEDate = FormatDate(now(),"0000-00-00")
End If

if (grptp="") then
    grptp="L"
end if

Dim oChart
SET oChart = new CChart
oChart.FRectSDate = vSDate
oChart.FRectEDate = vEDate
oChart.FRectChannel = vChannel
oChart.FRectUserLevel = ulevel
'oChart.FPageSize = CHKIIF(vpValue<>"",100,30)
'oChart.FRectOrderType = vOrdType



if (grptp="L") then
    vList1 = oChart.fnUserActiveTrendSumUserLevel_DW()
    vArr1 = oChart.fnUserActiveTrendByUserLevel_DW()
else
    vList1 = oChart.fnUserActiveTrendSumChannel_DW()
    vArr1 = oChart.fnUserActiveTrendChannel_DW()
end if

Dim ixAxisName : ixAxisName = "날짜"
Dim kk,yAxisName(4), iDataSeriseArr(4), iDataSetPosArr(4), iChartSubCaption(4)

if (grptp="L") then
    yAxisName(0) = "회원등급별 로그인수"
    iChartSubCaption(0) = "(유저당 채널별 1일 1건)"
    iDataSeriseArr(0) = Array("합계","White","Red","VIP","VIPGold","VVIP","STAFF")
    iDataSetPosArr(0) = Array(1,5,9,13,17,21,25)

    yAxisName(1) = "회원등급별 위시상품수"
    iChartSubCaption(1) = "(유저당 채널별 상품별 1일 1건)"
    iDataSeriseArr(1) = Array("합계","White","Red","VIP","VIPGold","VVIP","STAFF")
    iDataSetPosArr(1) = Array(2,6,10,14,18,22,26)

    yAxisName(2) = "회원등급별 장바구니상품수"
    iChartSubCaption(2) = "(유저당 채널별 상품별 1일 1건)"
    iDataSeriseArr(2) = Array("합계","White","Red","VIP","VIPGold","VVIP","STAFF")
    iDataSetPosArr(2) = Array(3,7,11,15,19,23,27)

    yAxisName(3) = "회원등급별 조회상품수"
    iChartSubCaption(3) = "(유저당 채널별 상품별 1일 1건)"
    iDataSeriseArr(3) = Array("합계","White","Red","VIP","VIPGold","VVIP","STAFF")
    iDataSetPosArr(3) = Array(4,8,12,16,20,24,28)

    yAxisName(4) = "회원등급별 매출추이"
    iChartSubCaption(4) = ""
    iDataSeriseArr(4) = Array("합계","White","Red","VIP","VIPGold","VVIP","STAFF","비회원")
    if (vOrdType="C") then
        iDataSetPosArr(4) = Array(29,34,39,44,49,54,59,64)
    elseif (vOrdType="V") then ''매출총액
        iDataSetPosArr(4) = Array(31,36,41,46,51,56,61,66)
    elseif (vOrdType="G") then
        iDataSetPosArr(4) = Array(33,38,43,48,53,58,63,68)
    else '' S 구매총액
        iDataSetPosArr(4) = Array(30,35,40,45,50,55,60,65)
    end if
else
    yAxisName(0) = "채널별 로그인수"
    iChartSubCaption(0) = "(유저당 채널별 1일 1건)"
    iDataSeriseArr(0) = Array("합계","로그인-App","로그인-Mob","로그인-Pc")
    iDataSetPosArr(0) = Array(1,2,3,4)

    yAxisName(1) = "채널별 위시상품수"
    iChartSubCaption(1) = "(유저당 채널별 상품별 1일 1건)"
    iDataSeriseArr(1) = Array("합계","위시상품수-App","위시상품수-Mob","위시상품수-Pc")
    iDataSetPosArr(1) = Array(5,6,7,8)

    yAxisName(2) = "채널별 장바구니상품수"
    iChartSubCaption(2) = "(유저당 채널별 상품별 1일 1건)"
    iDataSeriseArr(2) = Array("합계","장바구니상품수-App","장바구니상품수-Mob","장바구니상품수-Pc")
    iDataSetPosArr(2) = Array(9,10,11,12)

    yAxisName(3) = "채널별 조회상품수"
    iChartSubCaption(3) = "(유저당 채널별 상품별 1일 1건)"
    iDataSeriseArr(3) = Array("합계","조회상품수-App","조회상품수-Mob","조회상품수-Pc")
    iDataSetPosArr(3) = Array(13,14,15,16)

    yAxisName(4) = "채널별 매출추이"
    iChartSubCaption(4) = ""
    iDataSeriseArr(4) = Array("합계","조회상품수-App","조회상품수-Mob","조회상품수-Pc")

    if (vOrdType="C") then
        iDataSetPosArr(4) = Array(17,18,19,20)
    elseif (vOrdType="V") then ''매출총액
        iDataSetPosArr(4) = Array(25,26,27,28)
    elseif (vOrdType="G") then
        iDataSetPosArr(4) = Array(33,34,35,36)
    else '' S 구매총액
        iDataSetPosArr(4) = Array(21,22,23,24)
    end if
end if

Dim sum1,sum2,sum3,sum4,sum5,sum6,sum7,sum8,sum9
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

function enDisableComp(comp){
    var pfrm = comp.form;
    if (comp.value=="L"){
        pfrm.ulevel.value="";
        pfrm.ulevel.disabled = true;
        pfrm.channel.disabled = false;
    }else{
        pfrm.channel.value="";
        pfrm.channel.disabled = true;
        pfrm.ulevel.disabled = false;
    }
}
</script>

<% If isArray(vArr1) Then %>
<script type='text/javascript'>//<![CDATA[
window.onload=function(){
    <% if (grptp="L") then%>
        enDisableComp(document.frm.grptp[0]);
    <% else %>
        enDisableComp(document.frm.grptp[1]);
    <% end if %>

	FusionCharts.ready(function () {
        <% for kk=Lbound(yAxisName) to Ubound(yAxisName) %>
		var vstrChart<%=kk%> = new FusionCharts({
			type: 'msline', //'', 
			renderAt: 'chart-container<%=kk%>',
			width: '1100',
			height: '400',
			dataFormat: 'json',
			dataSource: {
				"chart": {
					"caption": " <%=yAxisName(kk)%> 추세",
					"subCaption": "<%=iChartSubCaption(kk)%>",
					"xAxisName": "<%=ixAxisName%>",
					"yAxisName": "<%=yAxisName(kk)%>",
					"theme": "fint",
					"showSum": "1",
					"showValues": "<%=CHKIIF(UBound(vArr1,2)>=60,0,1)%>",
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
					<% for k=0 to Ubound(iDataSeriseArr(kk)) %>
					{
						"seriesname": "<%=iDataSeriseArr(kk)(k)%>",
						"data": [
							<%
							If isArray(vArr1) Then
								For i = 0 To UBound(vArr1,2)
									'if (vArr1(1,i)=vArr2(0,chrtN-1)) then  ''브랜드가 같으면
									Response.Write "{" & vbCrLf
									Response.Write """value"": """&vArr1(iDataSetPosArr(kk)(k),i)&"""" & vbCrLf
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
					<% if (k<Ubound(iDataSeriseArr(kk))) then response.write "," %>
					<% next %>
					
				]
			}
		}).render();
        <% next %>
	});

}//]]>
</script>
<% End If %>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">

	<tr height="25" bgcolor="#FFFFFF">
	    <td align="center" width="50" bgcolor="#EEEEEE">검색<br>조건</td>
        <td>
            
			&nbsp;
			검색기간(주문일) :

			<input id='startdate' name='startdate' value='<%= vSDate %>' class='text' size='10' maxlength='10' />
			<img src='/images/calicon.gif' id='startdate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
			~<input id='enddate' name='enddate' value='<%= vEDate %>' class='text' size='10' maxlength='10' />
			<img src='/images/calicon.gif' id='enddate_trigger' border='0' style='cursor:pointer' align='absmiddle' />
			
            
            &nbsp;
            그루핑 : 
            <input type="radio" name="grptp" value="L" <%=CHKIIF(grptp="L","checked","") %> onClick="enDisableComp(this)">회원등급
            <input type="radio" name="grptp" value="C" <%=CHKIIF(grptp="C","checked","") %> onClick="enDisableComp(this)">채널
            
            &nbsp;
            회원등급 : 
            <% Call DrawselectboxUserLevel("ulevel", ulevel, "") %>
            &nbsp;
			채널 :
            <select name="channel" >
                <option value="" <%=CHKIIF(vChannel="","selected","")%>>ALL(채널별 합계)</option>
                <option value="pc" <%=CHKIIF(vChannel="pc","selected","")%>>WEB</option>
                <option value="mw" <%=CHKIIF(vChannel="mw","selected","")%>>MOB</option>
                <option value="app" <%=CHKIIF(vChannel="app","selected","")%>>APP</option>
            </select>
			&nbsp;
            |
            &nbsp;
            <input type="radio" name="ordtype" value="C" <%=CHKIIF(vOrdType="C","checked","") %> >주문건수
            <input type="radio" name="ordtype" value="S" <%=CHKIIF(vOrdType="S","checked","") %> >구매총액
            <input type="radio" name="ordtype" value="V" <%=CHKIIF(vOrdType="V","checked","") %> >매출총액
            <input type="radio" name="ordtype" value="G" <%=CHKIIF(vOrdType="G","checked","") %> >매출수익
			
			
            
        </td>
        <td align="center" width="50" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>
<table cellpadding="2" cellspacing="1" border="0" class="a" align="center" width="1100" bgcolor="#444444">
<% if IsArray(vList1) then %>
<% for i=0 to uBound(vList1,2) %>
    <%
    sum1 = sum1 + vList1(1,i)
    sum2 = sum2 + vList1(2,i)
    sum3 = sum3 + vList1(3,i)
    sum4 = sum4 + vList1(4,i)
    sum5 = sum5 + vList1(5,i)
    sum6 = sum6 + vList1(6,i)
    sum7 = sum7 + vList1(7,i)
    sum8 = sum8 + vList1(8,i)
    sum9 = sum9 + vList1(9,i)
    %>
    <tr bgcolor="#FFFFFF" align="right">
    <td align="center"><%= vList1(0,i) %></td>
    <td><%= FormatNumber(vList1(1,i),0) %></td>
    <td><%= FormatNumber(vList1(2,i),0) %></td>
    <td><%= FormatNumber(vList1(3,i),0) %></td>
    <td><%= FormatNumber(vList1(4,i),0) %></td>
    <td><%= FormatNumber(vList1(5,i),0) %></td>
    <td><%= FormatNumber(vList1(6,i),0) %></td>
    <td><%= FormatNumber(vList1(7,i),0) %></td>
    <td><%= FormatNumber(vList1(8,i),0) %></td>
    <td><%= FormatNumber(vList1(9,i),0) %></td>
    </tr>
<% next %>
    <thead>
    <tr bgcolor="#EEEEEE" align="center">
        <td width="120"><%=CHKIIF(grptp="L","회원등급","채널")%></td>
        <td width="110">로그인수(인)</td>
        <td width="110">위시수(건)</td>
        <td width="110">장바구니수(건)</td>
        <td width="110">상품조회(건)</td>
        <td  width="80">주문건수</td>
        <td  width="120">구매총액</td>
        <td  width="120">매출액</td>
        <td  width="120">매입액</td>
        <td  width="120">매출수익</td>
    </tr>
    <tr bgcolor="#FFFFFF" align="right">
		<th align="center">합계</th>
        <th><%= FormatNumber(sum1,0) %></th>
        <th><%= FormatNumber(sum2,0) %></th>
        <th><%= FormatNumber(sum3,0) %></th>
        <th><%= FormatNumber(sum4,0) %></th>
        <th><%= FormatNumber(sum5,0) %></th>
        <th><%= FormatNumber(sum6,0) %></th>
        <th><%= FormatNumber(sum7,0) %></th>
        <th><%= FormatNumber(sum8,0) %></th>
        <th><%= FormatNumber(sum9,0) %></th>
    </tr>
    </thead>
<% end if %>
</table>


<br />
<p>
<% for kk=Lbound(yAxisName) to Ubound(yAxisName) %>
<div id="chart-container<%=kk%>" style="text-align:center;">FusionCharts will render here</div>
<br />
<% next %>



<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbSTSclose.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->