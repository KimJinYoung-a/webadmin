<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/report/maechul/statisticCls.asp" -->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2, vIsOldOrder
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate
dim vSiteName, vSorting

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)
yyyy2 = RequestCheckvar(request("yyyy2"),4)
mm2 = RequestCheckvar(request("mm2"),2)
dd2 = RequestCheckvar(request("dd2"),2)
vSiteName 	= RequestCheckvar(request("sitename"),16)
vSorting	= NullFillWith(RequestCheckvar(request("sorting"),32),"ddateD")

dim tNo, tDiv, chkOld, isBanpum
tNo = Request("tNo")
tDiv = Request("tDiv")
chkOld = Request("chkOld")
isBanpum = RequestCheckvar(Request("isBanpum"),10)

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = "1"

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

'// 내용 접수
dim oreport
set oreport = new CUserLevelSell
	oreport.FRectSdate = fromDate
	oreport.FRectEdate = toDate
	oreport.FRectMinusInc = isBanpum
	oreport.FRectSiteName = vSiteName
	oreport.FRectSorting = vSorting
	oreport.GetLevelList

'각 비율 및 그래프 산출
dim sTotal, nTotal,  i, uTotal

if oreport.FResultCount>0 then
	for i=0 to oreport.FResultCount -1
		sTotal = sTotal + oreport.FItemList(i).FSellTotal
		nTotal = nTotal + oreport.FItemList(i).FSellCount
		uTotal = uTotal + oreport.FItemList(i).Funiqcnt
	next
end if
%>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>
<script type="text/javascript">
<!--
function downloadexcel(){
    document.frm.target = "view"; 
    document.frm.action = "/academy/report/maechul/userLevelSellReport_excel.asp";  
	document.frm.submit();
    document.frm.target = ""; 
    document.frm.action = "";  
}

function jstrSort(vsorting){
	var tmpSorting = document.getElementById("img"+vsorting)

	if (-1 < tmpSorting.src.indexOf("_alpha")){
		frm.sorting.value= vsorting+"D";
	}else if (-1 < tmpSorting.src.indexOf("_bot")){
		frm.sorting.value= vsorting+"A";
	}else{
		frm.sorting.value= vsorting+"D";
	}
	document.frm.submit();
}
//-->
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="sorting" value="<%= vsorting %>">
	<tr>
		<td class="a" >
		검색기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;&nbsp;* 사이트구분 : <% drawradio_academy_sitename "sitename", vSiteName, "", "Y" %>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
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
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oreport.FResultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="188" rowspan="2" onClick="jstrSort('userlevel'); return false;" style="cursor:hand;">회원등급 <img src="/images/list_lineup<%=CHKIIF(vSorting="userlevelD","_bot","_top")%><%=CHKIIF(instr(vSorting,"userlevel")>0,"_on","")%>.png" id="imguserlevel"></td>
	<td width="228" colspan="2" onClick="jstrSort('maechul'); return false;" style="cursor:hand;">매출 <img src="/images/list_lineup<%=CHKIIF(vSorting="maechulD","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechul")>0,"_on","")%>.png" id="imgmaechul"></td>
	<td width="228" colspan="2" onClick="jstrSort('sellcnt'); return false;" style="cursor:hand;">건수 <img src="/images/list_lineup<%=CHKIIF(vSorting="sellcntD","_bot","_top")%><%=CHKIIF(instr(vSorting,"sellcnt")>0,"_on","")%>.png" id="imgsellcnt"></td>
	<td width="50" rowspan="2" onClick="jstrSort('uniqcnt'); return false;" style="cursor:hand;">Uniq고객건수 <img src="/images/list_lineup<%=CHKIIF(vSorting="uniqcntD","_bot","_top")%><%=CHKIIF(instr(vSorting,"uniqcnt")>0,"_on","")%>.png" id="imguniqcnt"></td>
	<td width="106" rowspan="2" onClick="jstrSort('customerprice'); return false;" style="cursor:hand;">객단가(원) <img src="/images/list_lineup<%=CHKIIF(vSorting="customerpriceD","_bot","_top")%><%=CHKIIF(instr(vSorting,"customerprice")>0,"_on","")%>.png" id="imgcustomerprice"></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="139">매출액(원)</td>
	<td width="89">비율(%)</td>
	<td width="139">건수</td>
	<td width="89">비율(%)</td>
</tr>
<% if oreport.FResultCount>0 then %>
	<% for i=0 to oreport.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%=oreport.FItemList(i).GetUserLevelStr %></td>
		<td><%=FormatNumber(oreport.FItemList(i).FSellTotal,0)%></td>
		<td><%=FormatNumber((oreport.FItemList(i).FSellTotal/sTotal)*100,2)%>%</td>
		<td><%=FormatNumber(oreport.FItemList(i).FSellCount,0)%></td>
		<td><%=FormatNumber((oreport.FItemList(i).FSellCount/nTotal)*100,2)%>%</td>
		<td><%=FormatNumber(oreport.FItemList(i).Funiqcnt,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FSellAvr,0)%></td>
	</tr>
	<% next %>
	
	<tr align="center" bgcolor="#FAFAFA">
		<td>계</td>
		<td><%=FormatNumber(sTotal,0)%></td>
		<td>100%</td>
		<td><%=FormatNumber(nTotal,0)%></td>
		<td>100%</td>
		<td><%=FormatNumber(uTotal,0)%></td>
		<td><%=FormatNumber((sTotal/nTotal),0)%></td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
<script>
<!--
FusionCharts.ready(function () {
    var revenueChart = new FusionCharts({
        type: 'doughnut2d',
        renderAt: 'chart-container',
        width: '450',
        height: '450',
        dataFormat: 'json',
        dataSource: {
            "chart": {
                "caption": "회원등급별 매출현황",
                "subCaption": "<%=fromDate & " ~ " & toDate %>",
                "numberPrefix": "w ",
                "paletteColors": "#0075c2,#1aaf5d,#f2c500,#f45b00,#8e0000",
                "bgColor": "#ffffff",
                "showBorder": "0",
                "use3DLighting": "0",
				"formatNumberScale": "0",
                "showShadow": "0",
                "enableSmartLabels": "0",
                "startingAngle": "310",
                "showLabels": "0",
                "showPercentValues": "1",
                "showLegend": "1",
                "legendShadow": "0",
                "legendBorderAlpha": "0",
                "defaultCenterLabel": "Total revenue: <%=FormatNumber(sTotal,0)%>",
                "centerLabel": "Revenue from $label: $value",
                "centerLabelBold": "1",
                "showTooltip": "1",
                "decimals": "0",
                "captionFontSize": "14",
                "subcaptionFontSize": "10",
                "subcaptionFontBold": "0"
            },
            "data": [
				<%
				if oreport.FresultCount > 0 then
					For i = 0 To oreport.FresultCount -1
						Response.Write "{" & vbCrLf
						Response.Write """label"": """ & oreport.FItemList(i).GetUserLevelStr & """," & vbCrLf
						Response.Write """value"": """ & oreport.FItemList(i).FSellTotal & """" & vbCrLf
						Response.Write "}"
						If i <> oreport.FresultCount-1 Then
							Response.Write ","
						End If
						Response.Write vbCrLf
					Next
				End If
				%>
            ]
        }
    }).render();
});
//-->
</script>
<div id="chart-container">FusionCharts will render here</div>
<iframe id="view" name="view" src="" width=1000 height=500 frameborder="0" scrolling="no"></iframe>
<%
set oreport = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbACADEMYclose.asp" -->