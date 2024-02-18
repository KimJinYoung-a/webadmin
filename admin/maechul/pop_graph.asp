<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
	Dim vParameter, vLink
	dim yyyy1, yyyy2,dateview1 , datecancle,bancancle,accountdiv,sitename,i, mm1, mm2, defaultdate1, monthday
	dim ipkumdatesucc, vParam, nowYear

	nowYear = year(date)

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
	vLink = "/admin/maechul/fusionchart"
%>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="javascript">
function goSearchChart(gubun)
{
	if (gubun == "m")
	{
		frm.yyyy1.value = frm.my1.value;
		frm.yyyy2.value = frm.my2.value;
	}
	else
	{
		frm.yyyy1.value = frm.dy1.value;
		frm.yyyy2.value = frm.dy2.value;
		frm.mm1.value = frm.dm1.value;
	}
	frm.monthday.value = gubun;
	frm.submit();
}
</script>
<!-- 그래프 시작-->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td>
			<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="#999999">
				<tr>
					<td width="400" style="padding:5; border-top:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999"  background="/images/menubar_1px.gif">
						<font color="#333333"><b>매출통계v2&gt;&gt;그래프매출통계</b></font>
					</td>
					<td align="right" style="border-bottom:1px solid #999999" bgcolor="#F4F4F4"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td style="padding:5; border-bottom:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999" bgcolor="#FFFFFF">
			<br>실금액 = 총금액 - (할인쿠폰 + 마일리지 + 기타할인) <br>순수익 = 실금액 - (매입가 + 텐배송비)<br>제휴사 수수료는 제외됩니다.
		</td>
	</tr>
</table>
<Br>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get">
<input type="hidden" name="yyyy1" value="">
<input type="hidden" name="yyyy2" value="">
<input type="hidden" name="mm1" value="">
<input type="hidden" name="mm2" value="">
<input type="hidden" name="monthday" value="">
<tr bgcolor="#FFFFFF">
	<td>
		<table cellpadding="0" cellspacing="0" border="0" width="100%" class="a">
		<tr>
			<td>
			<%
				Response.Write "<select class='select' name='my1'>"
			    for i=2002 to nowYear
					if (CStr(i)=CStr(yyyy1)) then
						Response.Write "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
					else
			    		Response.Write "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
			        end if
				next
			    Response.Write "</select>&nbsp;~&nbsp;"
				Response.Write "<select class='select' name='my2'>"
			    for i=2002 to nowYear
					if (CStr(i)=CStr(yyyy2)) then
						Response.Write "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
					else
			    		Response.Write "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
			        end if
				next
			    Response.Write "</select> 월매출 비교"
			%>
				<input type="button" class="button_s" value="검색" onClick="goSearchChart('m');">
			</td>
			<td align="right">
			<%
				Response.Write "<select class='select' name='dy1'>"
			    for i=2002 to nowYear
					if (CStr(i)=CStr(yyyy1)) then
						Response.Write "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
					else
			    		Response.Write "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
			        end if
				next
			    Response.Write "</select> 부터 "
				Response.Write "<select class='select' name='dy2'>"
			    for i=2002 to nowYear
					if (CStr(i)=CStr(yyyy2)) then
						Response.Write "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
					else
			    		Response.Write "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
			        end if
				next
			    Response.Write "</select> 까지 매년 "
			    Response.Write "<select class='select' name='dm1'>"
			    for i=1 to 12
					if (Format00(2,i)=Format00(2,mm1)) then
						Response.Write "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
					else
			    	    Response.Write "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
					end if
				next
				Response.Write "</select> 월 일매출 비교 "
			%>
				<input type="button" class="button_s" value="검색" onClick="goSearchChart('d');">
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
<table>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td style="padding:0 0 20 0"><iframe id="graph0" name="graph0" src="<%=vLink%>/iframe_graph_1year.asp?<%=vParameter%>" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td style="padding:0 0 20 0"><iframe id="graph1" name="graph1" src="<%=vLink%>/iframe_graph.asp?gubun=1&<%=vParameter%>" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td style="padding:0 0 20 0"><iframe id="graph2" name="graph2" src="<%=vLink%>/iframe_graph.asp?gubun=2&<%=vParameter%>" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td style="padding:0 0 20 0"><iframe id="graph3" name="graph3" src="<%=vLink%>/iframe_graph.asp?gubun=3&<%=vParameter%>" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe></td>
</tr>
<!--
<tr bgcolor="#FFFFFF">
	<td width="100%">
		<table border="1" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td><iframe id="graph1" name="graph1" src="<%=vLink%>/iframe_graph.asp?gubun=1&<%=vParameter%>" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe></td>
			<td><iframe id="graph2" name="graph2" src="<%=vLink%>/iframe_graph.asp?gubun=2&<%=vParameter%>" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe></td>
			<td><iframe id="graph3" name="graph3" src="<%=vLink%>/iframe_graph.asp?gubun=3&<%=vParameter%>" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe></td>
		</tr>
		</table>
	</td>
</tr>
//-->
</table>
<!-- 그래프 끝-->
</body>
</html>
