<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/maechul/fusionchart/maechul_class.asp" -->
<%
	Dim vYear, vGubun, vGraph, vParam, vParamOrig
	vGubun = Request("gubun")
	vGraph = Request("graph")
	
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
	
	If vGubun = "" Then
		vGubun = "1"
	End If
	If vGraph = "" Then
		vGraph = "1"
	End If

	vYear = yyyy1 & " ~ " & yyyy2
	
	vParam = "yyyy1="&yyyy1&"&yyyy2="&yyyy2&"&datecancle="&datecancle&"&bancancle="&bancancle&"&accountdiv="&accountdiv&"&sitename="&sitename&"&dateview1="&dateview1&"&ipkumdatesucc="&ipkumdatesucc&"&mm1="&mm1&"&mm2="&mm2&"&monthday="&monthday&""
	vParamOrig = vParam
	vParam = Replace(vParam, "&", "^^")
%>
<html>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>
<body LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0>
<!-- 그래프 시작-->	
<table cellpadding="0" cellspacing="0" border="0" class="a">
<tr bgcolor="#FFFFFF">
<%
	If vGubun = "1" Then
%>
		<td>
			<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td width="25%">&nbsp;<!--<input type="button" value="Print" onClick="print()" class="button">//--></td>
				<td width="50%" align="center" style="padding:7 0 7 0"><font size="3">[<b><%= vYear %> <% If monthday = "d" Then Response.Write "&nbsp;&nbsp;" & mm1 & "월" End If %> 실금액 통계</b>]</font></td>
				<td width="25%">
					<input type="button" value="선형" onClick="location.href='?gubun=1&graph=1&<%=vParamOrig%>'" class="button">&nbsp;<input type="button" value="막대형" onClick="location.href='?gubun=1&graph=2&<%=vParamOrig%>'" class="button">
				</td>
			</tr>
			</table>
			<div id="chartdiv1" align="center"></div>
			<script type="text/javascript">
			
			FusionCharts.ready(function(){
				var myChart = new FusionCharts({
					"type": "<%=GraphFile2(vGraph)%>",
					"width":"750",
					"height":"350",
					"dataFormat": "xml"
				});
				myChart.setXMLUrl("/admin/maechul/fusionchart/graph_xmllist_small_imsi.asp?param=^^<%=vParam%>^^gubun=<%=vGubun%>");
				myChart.render("chartdiv1");
			});


//			var chart = new FusionCharts("/admin/maechul/fusionchart/<%=GraphFile(vGraph)%>", "chartdiv1", "750", "350", "0", "0");
//			chart.setDataURL("/admin/maechul/fusionchart/graph_xmllist_small.asp?param=^^<%=vParam%>^^gubun=<%=vGubun%>");
//			chart.render("chartdiv1");
			</script>
		</td>
<%
	ElseIf vGubun = "2" Then
%>
		<td>
			<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td width="25%">&nbsp;<!--<input type="button" value="Print" onClick="print()" class="button">//--></td>
				<td width="50%" align="center" style="padding:7 0 7 0"><font size="3">[<b><%= vYear %> <% If monthday = "d" Then Response.Write "&nbsp;&nbsp;" & mm1 & "월" End If %> 순수익 통계</b>]</font></td>
				<td width="25%">
					<input type="button" value="선형" onClick="location.href='?gubun=2&graph=1&<%=vParamOrig%>'" class="button">&nbsp;<input type="button" value="막대형" onClick="location.href='?gubun=2&graph=2&<%=vParamOrig%>'" class="button">
				</td>
			</tr>
			</table>
			<div id="chartdiv2" align="center"></div>
			<script type="text/javascript">	

			FusionCharts.ready(function(){
				var myChart2 = new FusionCharts({
					"type": "<%=GraphFile2(vGraph)%>",
					"width":"750",
					"height":"350",
					"dataFormat": "xml"
				});
				myChart2.setXMLUrl("/admin/maechul/fusionchart/graph_xmllist_small_imsi.asp?param=^^<%=vParam%>^^gubun=<%=vGubun%>");
				myChart2.render("chartdiv2");
			});

//			var chart2 = new FusionCharts("/admin/maechul/fusionchart/<%=GraphFile(vGraph)%>", "chartdiv2", "750", "350", "0", "0");
//			chart2.setDataURL("/admin/maechul/fusionchart/graph_xmllist_small.asp?param=^^<%=vParam%>^^gubun=<%=vGubun%>");
//			chart2.render("chartdiv2");
			</script>
		</td>
<%
	ElseIf vGubun = "3" Then
%>
		<td>
			<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td width="25%">&nbsp;<!--<input type="button" value="Print" onClick="print()" class="button">//--></td>
				<td width="50%" align="center" style="padding:7 0 7 0"><font size="3">[<b><%= vYear %> <% If monthday = "d" Then Response.Write "&nbsp;&nbsp;" & mm1 & "월" End If %> 총건수 통계</b>]</font></td>
				<td width="25%">
					<input type="button" value="선형" onClick="location.href='?gubun=3&graph=1&<%=vParamOrig%>'" class="button">&nbsp;<input type="button" value="막대형" onClick="location.href='?gubun=3&graph=2&<%=vParamOrig%>'" class="button">
				</td>
			</tr>
			</table>
			<div id="chartdiv3" align="center"></div>
			<script type="text/javascript">

			
			FusionCharts.ready(function(){
				var myChart3 = new FusionCharts({
					"type": "<%=GraphFile2(vGraph)%>",
					"width":"750",
					"height":"350",
					"dataFormat": "xml"
				});
				myChart3.setXMLUrl("/admin/maechul/fusionchart/graph_xmllist_small_imsi.asp?param=^^<%=vParam%>^^gubun=<%=vGubun%>");
				myChart3.render("chartdiv3");
			});

//			var chart3 = new FusionCharts("/admin/maechul/fusionchart/<%=GraphFile(vGraph)%>", "chartdiv3", "750", "350", "0", "0");
//			chart3.setDataURL("/admin/maechul/fusionchart/graph_xmllist_small.asp?param=^^<%=vParam%>^^gubun=<%=vGubun%>");
//			chart3.render("chartdiv3");
			</script>
		</td>
<%
	End If
%>
</tr>
</table>
<!-- 그래프 끝-->
</body>
</html>
