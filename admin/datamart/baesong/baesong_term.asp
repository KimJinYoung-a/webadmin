<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
	Dim vParameter, vLink
	dim yyyy1, yyyy2,dateview1 , datecancle,bancancle,accountdiv,sitename,i, mm1, mm2, defaultdate1, monthday
	dim ipkumdatesucc, vParam
	
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
	if mm1 = "" then mm1 = month(now)
	if mm2 = "" then mm2 = month(now)
	mm2 = TwoNumber(mm2)
	if bancancle = "" then bancancle = "1"
	if dateview1 = "" then dateview1 = "yes"
	if monthday = "" then monthday = "m"


	vParameter = "yyyy1="&yyyy1&"&yyyy2="&yyyy2&"&datecancle="&datecancle&"&bancancle="&bancancle&"&accountdiv="&accountdiv&"&sitename="&sitename&"&dateview1="&dateview1&"&ipkumdatesucc="&ipkumdatesucc&"&mm1="&mm1&"&mm2="&mm2&"&monthday="&monthday&""
	vLink = "/admin/datamart/baesong"
%>
<link rel="stylesheet" href="/css/scm.css" type="text/css">

<!-- 그래프 시작-->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td>
			<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="#999999">
				<tr>
					<td width="400" style="padding:5; border-top:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999"  background="/images/menubar_1px.gif">
						<font color="#333333"><b>DATAMART&gt;&gt;배송 소요일 분석</b></font>
					</td>
					<td align="right" style="border-bottom:1px solid #999999" bgcolor="#F4F4F4"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td style="padding:5; border-bottom:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999" bgcolor="#FFFFFF">
			<table cellpadding="0" cellspacing="0" border="0" class="a">
			<tr bgcolor="#FFFFFF">
				<td width="50%" valign="top">
					<b>배송소요일분석</b>
					<br>상품별 : 한 상품의 모든 옵션 배송일 평균값.
					<br>상품+옵션별 : 한 상품의 각 옵션별 배송일 평균값.
					<br>상품별 : 한 브랜드별 배송일 평균값.
				</td>
				<td width="50%" valign="top">
					<b>상품및브랜드일별세부분석</b>
					<br>각 날짜별로 브랜드, 상품, 옵션으로 Grouping 된 배송소요일 리스트.
				</td>
			</tr>
			</table>
		</td>
	</tr>
</table>
<Br>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td style="padding:0 0 30 0"><iframe id="graph0" name="graph0" src="<%=vLink%>/iframe_baesong_term_graph.asp" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td style="padding:0 0 20 0"><iframe id="graph1" name="graph1" src="<%=vLink%>/iframe_baesong_term_list.asp" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe></td>
</tr>
</table>
<!-- 그래프 끝-->
</body>
</html>
