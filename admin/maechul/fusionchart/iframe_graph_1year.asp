<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim vYear, vGubun, vParam
	dim menupos, yyyy1, yyyy2,dateview1 , datecancle,bancancle,accountdiv,sitename,i, mm1, mm2, defaultdate1, monthday
	dim ipkumdatesucc
	menupos 		= request("menupos")
	yyyy1 			= request("yyyy1")
	yyyy2 			= request("yyyy2")
	dateview1 		= request("dateview1")
	datecancle 		= request("datecancle")
	bancancle 		= request("bancancle") 
	accountdiv 		= request("accountdiv")			
	sitename 		= request("sitename") 
	ipkumdatesucc 	= request("ipkumdatesucc")
	mm1 			= request("mm1")
	mm2 			= request("mm2")
	monthday		= request("monthday")
	
	defaultdate1 = dateadd("d",-60,year(now) & "-" &TwoNumber(month(now)) & "-" & day(now))		'날짜값이 없을때 기본값으로 60이전까지 검색
	if yyyy2 = "" then yyyy2 = year(now)
	if yyyy1 = "" then yyyy1 = CInt(yyyy2)-2
	if mm1 = "" then mm1 = "01"
	if mm2 = "" then mm2 = month(now)
	mm2 = TwoNumber(mm2)
	if bancancle = "" then bancancle = "1"
	if dateview1 = "" then dateview1 = "yes"
	
	vParam = "yyyy1="&yyyy1&"&yyyy2="&yyyy2&"&datecancle="&datecancle&"&bancancle="&bancancle&"&accountdiv="&accountdiv&"&sitename="&sitename&"&dateview1="&dateview1&"&ipkumdatesucc="&ipkumdatesucc&"&mm1="&mm1&"&mm2="&mm2&"&monthday="&monthday&""
	vParam = Replace(vParam, "&", "^^")
%>
<html>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>
<body>
<!-- 그래프 시작-->	

<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
<tr bgcolor="#FFFFFF">
	<td style="padding:5 0 5 0"><center><font size="3">[<b><%=yyyy2%>년 월 매출</b>]</font></center></td>
</tr>
<tr>
	<td style="padding:5 0 5 0">
		<div id="chartdiv0" align="center"></div>
		<script type="text/javascript">	
			FusionCharts.ready(function(){
				var myChart = new FusionCharts({
					"type": "mscolumn3dlinedy",
					"width":"750",
					"height":"350",
					"dataFormat": "xml"
				});
				myChart.setXMLUrl("/admin/maechul/fusionchart/graph_xmllist_big_imsi.asp?param=^^<%=vParam%>");
				myChart.render("chartdiv0");
			});
		</script>
	</td>
</tr>
</table>
<!-- 그래프 끝-->
</body>
</html>
