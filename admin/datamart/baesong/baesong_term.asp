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
	
	defaultdate1 = dateadd("d",-60,year(now) & "-" &TwoNumber(month(now)) & "-" & day(now))		'��¥���� ������ �⺻������ 60�������� �˻�
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

<!-- �׷��� ����-->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td>
			<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="#999999">
				<tr>
					<td width="400" style="padding:5; border-top:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999"  background="/images/menubar_1px.gif">
						<font color="#333333"><b>DATAMART&gt;&gt;��� �ҿ��� �м�</b></font>
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
					<b>��ۼҿ��Ϻм�</b>
					<br>��ǰ�� : �� ��ǰ�� ��� �ɼ� ����� ��հ�.
					<br>��ǰ+�ɼǺ� : �� ��ǰ�� �� �ɼǺ� ����� ��հ�.
					<br>��ǰ�� : �� �귣�庰 ����� ��հ�.
				</td>
				<td width="50%" valign="top">
					<b>��ǰ�׺귣���Ϻ����κм�</b>
					<br>�� ��¥���� �귣��, ��ǰ, �ɼ����� Grouping �� ��ۼҿ��� ����Ʈ.
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
<!-- �׷��� ��-->
</body>
</html>
