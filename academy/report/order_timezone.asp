<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스 매출집계-일별
' History : 2016.09.20 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/report/maechul/statisticCls.asp" -->
<%
Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay
Dim vTot_CountPlus, vTot_CountMinus, vTot_MaechulPlus, vTot_MaechulMinus, vTot_Subtotalprice, vTot_Miletotalprice, vTot_subtotalprice_notexists_sumPaymentEtc
dim vTot_MaechulCountSum, vTot_MaechulPriceSum, vTot_sumPaymentEtc, page, pagesize, vSorting
dim sellchnl
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= RequestCheckvar(request("sitename"),16)
	vDateGijun	= NullFillWith(RequestCheckvar(request("date_gijun"),16),"regdate")
	'vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",-13,now())))
	'vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",-13,now())))
	'vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",-13,now())))
	vSYear		= NullFillWith(RequestCheckvar(request("syear"),4),Year(now()))
	vSMonth		= NullFillWith(RequestCheckvar(request("smonth"),2),Month(now()))
	vSDay		= NullFillWith(RequestCheckvar(request("sday"),2),"01")
	vEYear		= NullFillWith(RequestCheckvar(request("eyear"),4),Year(now))
	vEMonth		= NullFillWith(RequestCheckvar(request("emonth"),2),Month(now))
	vEDay		= NullFillWith(RequestCheckvar(request("eday"),2),Day(now))
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	vSorting	= NullFillWith(RequestCheckvar(request("sorting"),32),"ddateD")

if (page = "") then
	page = 1
end if

if (pagesize = "") then
	pagesize = 3000
end if

Set cStatistic = New cacademyStatic_list
	cStatistic.FCurrPage = page
	cStatistic.FPageSize = pagesize
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	cStatistic.facademyStatistic_TimeZonelist()
%>

<script type='text/javascript'>

function downloadexcel(){
    document.frm.target = "view"; 
    document.frm.action = "/academy/report/maechul/statistic_timezone_excel.asp";  
	document.frm.submit();
    document.frm.target = "";
    document.frm.action = "";
}

function searchSubmit(){
    frm.submit();
}

</script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>
<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sorting" value="<%= vsorting %>">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">검색 조건</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="30">
				* 기간 :
				<select name="date_gijun" class="select">
					<option value="regdate" <%=CHKIIF(vDateGijun="regdate","selected","")%>>주문일</option>
					<option value="ipkumdate" <%=CHKIIF(vDateGijun="ipkumdate","selected","")%>>결제일</option>
				</select>
				&nbsp;
				<% DrawDateBoxdynamic vSYear,"syear",vEYear,"eyear",vSMonth,"smonth",vEMonth,"emonth",vSDay,"sday",vEDay,"eday" %>
			</td>
		</tr>
		<tr>
		    <td>
		    	* 사이트구분 : <% drawradio_academy_sitename "sitename", vSiteName, "", "Y" %>
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="검색" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		* 검색 기간이 길어지면 상당히 느려집니다. 그러니 검색 버튼을 클릭한 뒤 아무 반응이 없어보인다고 재차 검색버튼을 클릭하지 마세요.
	</td>
	<td align="right">	
		<input type="button" onclick="downloadexcel();" value="엑셀다운로드" class="button">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>시간대</td>
    <td>주문건수</td>
    <td>금액</td>
</tr>
<% if cStatistic.FTotalCount > 0 then %>
	<% For i = 0 To cStatistic.FTotalCount -1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%=cStatistic.FItemList(i).FTimeZone %>시</td>
		<td align="center"><%=cStatistic.FItemList(i).FCount %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FMaeChul,0) %></td>
		<%
		vTot_MaechulCountSum	= vTot_MaechulCountSum + CLng(FormatNumber(cStatistic.FItemList(i).FCount,0))
		vTot_MaechulPriceSum	= vTot_MaechulPriceSum + CLng(FormatNumber(cStatistic.FItemList(i).FMaeChul,0))
		%>
	</tr>
	<% Next %>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<td bgcolor="#E6B9B8"></td>
		<td bgcolor="#E6B9B8"><%= FormatNumber(vTot_MaechulCountSum,0) %></td>
		<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><%= FormatNumber(vTot_MaechulPriceSum,0) %></td>
	</tr>
<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="25">등록된 내용이 없습니다.</td>
	</tr>
<% end if %>

</table>

<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<script>
<!--
FusionCharts.ready(function () {
    var revenueChart = new FusionCharts({
        type: 'msline',
        renderAt: 'chart-container',
        width: '<%=35*cStatistic.FTotalCount%>',
        height: '500',
        dataFormat: 'json',
        dataSource: {
            "chart": {
				"caption": "시간대별 매출집계",
                "subCaption": "<%=vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay) & " ~ " & vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay) %>",
				"formatnumberscale": "0",
				"paletteColors": "#0075c2,#1aaf5d",
				"labelDisplay": "auto",
                "bgcolor": "#ffffff",
                "showBorder": "0",
                "showShadow": "0",
                "showCanvasBorder": "0",
                "usePlotGradientColor": "0",
                "legendBorderAlpha": "0",
                "legendShadow": "0",
                "showAxisLines": "0",
                "showAlternateHGridColor": "0",
                "divlineThickness": "1",
                "divLineIsDashed": "1",
                "divLineDashLen": "1",
                "divLineGapLen": "1",
                "xAxisName": "Day",
                "showValues": "0"
            },            
            "categories": [
                {
                    "category": [
						<%
						if cStatistic.FTotalCount > 0 then
							For i = 0 To cStatistic.FTotalCount -1
								Response.Write "{" & vbCrLf
								Response.Write """label"": """ & cStatistic.FItemList(i).FTimeZone & """," & vbCrLf
								Response.Write "}"
								If i <> cStatistic.FTotalCount-1 Then
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
                    "seriesname": "매출",
                    "data": [
						<%
						if cStatistic.FTotalCount > 0 then
							For i = 0 To cStatistic.FTotalCount -1
								Response.Write "{" & vbCrLf
								Response.Write """value"": """ & cStatistic.FItemList(i).FMaeChul & """" & vbCrLf
								Response.Write "}"
								If i <> cStatistic.FTotalCount-1 Then
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
//-->
</script>
<div id="chart-container">FusionCharts will render here</div>
<%
Set cStatistic = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->