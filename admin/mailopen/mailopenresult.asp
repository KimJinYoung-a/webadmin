<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ���� ������ ����
' History : 2007.08.27 �ѿ�� ����
'			2012.05.09 ������ ���� �߰�
'			2012.12.04 ������ ���� ����	mailopenclass2����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/mailopenresultclass/mailopenclass2.asp"-->
<%
Dim Frealcntsum5, Fsuccesscntsum5, Fopencntsum5, Fsuccesssu5, Fopensu5, frealopensu5
Dim Frealcntsum3, Fsuccesscntsum3, Fopencntsum3, Fsuccesssu3, Fopensu3, frealopensu3, FclickSum3, FclickPer3
Dim yyyy , mm, oMailzine, oMailzine3, oMailzine5, i
	yyyy 					= requestcheckvar(request("yyyy1"),4)
	mm 						= requestcheckvar(request("mm1"),2)

If (yyyy="") Then yyyy 	= Cstr(Year(now()))
If (mm="") Then mm 		= Cstr(Month(now()))
Session("yyyy") 		= yyyy
Session("mm") 			= mm
%>
<script language="javascript" src="/admin/mailopen/daumchart/FusionCharts.js"></script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">��: <% DrawYMBox yyyy,mm %></td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();"></td>
</tr>
</form>
</table>
<br>
<%
'===========================================================������ ��ȸ�� ����==================================================================
Frealcntsum5=0
Fsuccesscntsum5=0
Fopencntsum5=0
FclickSum3=0
Fsuccesssu=0
Fopensu=0
FclickPer=0
frealopensu=0
Set oMailzine = new CMailzinelist
	oMailzine.FRectyyyy 	= yyyy
	oMailzine.FRectmm 		= mm
	oMailzine.FGubun 		= "mailzine"
	oMailzine.FMailzinelist()
If oMailzine.FTotalcount > 0 Then
%>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF"><td colspan="15">�˻���� : <b><%= oMailzine.FTotalCount %></b></td></tr>
<tr><td bgcolor="FFFFFF" colspan="15">1. Mailzine ���� �߼� ���</td></tr>
<tr bgcolor="#DDDDFF">
	<td align="center">�߼۳�¥</td>
	<td align="center">����</td>
	<td align="center">���Ϸ�</td>
	<td align="center">��������</td>
	<td align="center">�����߼����</td>
	<td align="center">�����߼����</td>
	<td align="center">�������</td>
	<td align="center">Ŭ�����</td>
</tr>
<%
	Dim Frealcntsum, Fsuccesscntsum, Fopencntsum, Fsuccesssu, Fopensu, frealopensu, FclickSum, FclickPer
	For i = 0 to oMailzine.ftotalcount -1 
%>
<tr bgcolor="FFFFFF">
	<td align="center"><%= oMailzine.FList(i).Freenddate %></td>
	<td align="center"><%= oMailzine.FList(i).Fgubun %></td>
	<td align="center"><%= oMailzine.FList(i).fmailergubun %></td>
	<td align="center"><%= oMailzine.FList(i).Ftitle %></td>
	<td align="center"><%= CurrFormat(oMailzine.FList(i).Frealcnt) %><% Frealcntsum = Frealcntsum+oMailzine.FList(i).Frealcnt %></td>
	<td align="center"><%= CurrFormat(oMailzine.FList(i).Fsuccesscnt) %><% Fsuccesscntsum = Fsuccesscntsum+oMailzine.FList(i).Fsuccesscnt %></td>
	<td align="center"><%= CurrFormat(oMailzine.FList(i).Fopencnt) %><% Fopencntsum = Fopencntsum+oMailzine.FList(i).Fopencnt %></td>
	<td align="center"><%= CurrFormat(oMailzine.FList(i).FClickCnt) %><% FclickSum = FclickSum+oMailzine.FList(i).FClickCnt %></td>
</tr>
<%	Next %>
<tr bgcolor="#DDDDFF">
	<td align="center" colspan=4>�� ��</td>
	<td align="center"><%= CurrFormat(Frealcntsum) %></td>
	<td align="center"><%= CurrFormat(Fsuccesscntsum) %></td>
	<td align="center"><%= CurrFormat(Fopencntsum) %></td>
	<td align="center"><%= CurrFormat(FclickSum) %></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td align="center" colspan=4>�� �� �� ��</td>
	<td align="center"></td>
	<td align="center"><% Fsuccesssu = (Fsuccesscntsum/Frealcntsum)*100 %><%= round(Fsuccesssu,0) %>%</td>
	<td align="center"><% Fopensu = (Fopencntsum/Fsuccesscntsum)*100 %><%= round(Fopensu,1) %>%</td>
	<td align="center"><% FclickPer = (FclickSum/Fsuccesscntsum)*100 %><%= round(FclickPer,1) %>%</td>
</tr>
<td  bgcolor="FFFFFF" colspan="15"><% frealopensu=(Fopencntsum/Frealcntsum)*100 %>
  �� �� �߼������ <%= round(Fopensu,1) %>%�� ������ �����ϸ� <%= round(FclickPer,1) %>%�� Ŭ���ϴ� ������ ���� �Ǿ����ϴ�.</td><%'= round(frealopensu,0) %>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor=FFFFFF>
	<td align="center">
		<div align="right"><input type="button" value="�׷�������Ʈ" onclick="javascript:window.print();" class="button"></div><br>
		<div id="chartdiv1" align="center"></div>
		<script type="text/javascript">
			var chart = new FusionCharts("/admin/mailopen/daumchart/MSCombiDY2D.swf", "chartdiv1", "640", "480", "0", "0");
			chart.setDataURL("/admin/mailopen/daumchart/MSCombiDY2D.asp?gubun=mailzine");
			chart.render("chartdiv1");
		</script>
	</td>
</tr>
</table><br>
<%
Else
%>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#DDDDFF"><td align=center bgcolor="#FFFFFF">Mailzine�� ���� �˻� ����� �����ϴ�.</td></tr>
</table>
<%
End If
Set oMailzine = nothing
'===========================================================������ ��ȸ�� ��==================================================================
%>

<%
'===========================================================������ ��ȸ�� ����==================================================================
Frealcntsum5=0
Fsuccesscntsum5=0
Fopencntsum5=0
FclickSum3=0
Fsuccesssu=0
Fopensu=0
FclickPer=0
frealopensu=0
Set oMailzine3 = new CMailzinelist
	oMailzine3.FRectyyyy 	= yyyy
	oMailzine3.FRectmm 		= mm
	oMailzine3.FGubun 		= "mailzine_not"
	oMailzine3.FMailzinelist()
If oMailzine3.FTotalcount > 0 Then
%>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF"><td colspan="15">�˻���� : <b><%= oMailzine3.FTotalCount %></b></td></tr>
<tr><td bgcolor="FFFFFF" colspan="15">3. ��ȸ�� Mailzine ���� �߼� ���</td></tr>
<tr bgcolor=#DDDDFF>
	<td align="center">�߼۳�¥</td>
	<td align="center">����</td>
	<td align="center">���Ϸ�</td>
	<td align="center">��������</td>
	<td align="center">�����߼����</td>
	<td align="center">�����߼����</td>
	<td align="center">�������</td>
	<td align="center">Ŭ�����</td>
</tr>
<%
	For i = 0 to oMailzine3.ftotalcount -1 
%>
<tr bgcolor=FFFFFF>
	<td align="center"><%= oMailzine3.FList(i).Freenddate %></td>
	<td align="center"><%= oMailzine3.FList(i).Fgubun %></td>
	<td align="center"><%= oMailzine3.FList(i).fmailergubun %></td>
	<td align="center"><%= oMailzine3.FList(i).Ftitle %></td>
	<td align="center"><%= CurrFormat(oMailzine3.FList(i).Frealcnt) %><% Frealcntsum3 = Frealcntsum3+oMailzine3.FList(i).Frealcnt %></td>
	<td align="center"><%= CurrFormat(oMailzine3.FList(i).Fsuccesscnt) %><% Fsuccesscntsum3 = Fsuccesscntsum3+oMailzine3.FList(i).Fsuccesscnt %></td>
	<td align="center"><%= CurrFormat(oMailzine3.FList(i).Fopencnt) %><% Fopencntsum3 = Fopencntsum3+oMailzine3.FList(i).Fopencnt %></td>
	<td align="center"><%= CurrFormat(oMailzine3.FList(i).FClickCnt) %><% FclickSum3 = FclickSum3+oMailzine3.FList(i).FClickCnt %></td>
</tr>
<%	Next %>
<tr bgcolor=#DDDDFF>
	<td align="center" colspan=4>�� ��</td>
	<td align="center"><%= CurrFormat(Frealcntsum3) %></td>
	<td align="center"><%= CurrFormat(Fsuccesscntsum3) %></td>
	<td align="center"><%= CurrFormat(Fopencntsum3) %></td>
	<td align="center"><%= CurrFormat(FclickSum3) %></td>
</tr>
<tr bgcolor=#DDDDFF>
	<td align="center" colspan=4>�� �� �� ��</td>
	<td align="center"></td>
	<td align="center"><% Fsuccesssu3 = (Fsuccesscntsum3/Frealcntsum3)*100 %><%= round(Fsuccesssu3,0) %>%</td>
	<td align="center"><% Fopensu3 = (Fopencntsum3/Fsuccesscntsum3)*100 %><%= round(Fopensu3,1) %>%</td>
	<td align="center"><% FclickPer3 = (FclickSum3/Fsuccesscntsum3)*100 %><%= round(FclickPer3,1) %>%</td>
</tr>
<td  bgcolor="FFFFFF" colspan="15"><% frealopensu3=(Fopencntsum3/Frealcntsum3)*100 %>
  �� �� �߼������ <%= round(Fopensu3,1) %>%�� ������ �����ϸ� <%= round(FclickPer3,1) %>%�� Ŭ���ϴ� ������ ���� �Ǿ����ϴ�.</td><%'= round(frealopensu,0) %>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor=FFFFFF>
	<td align="center">
		<div align="right"><input type="button" value="�׷�������Ʈ" onclick="javascript:window.print();" class="button"></div><br>
		<div id="chartdiv3" align="center"></div>
		<script type="text/javascript">
			var chart3 = new FusionCharts("/admin/mailopen/daumchart/MSCombiDY2D.swf", "chartdiv3", "640", "480", "0", "0");
			chart3.setDataURL("/admin/mailopen/daumchart/MSCombiDY2D.asp?gubun=mailzine_not");
			chart3.render("chartdiv3");
		</script>
	</td>
</tr>
</table><br>
<%
Else
%>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#DDDDFF"><td align=center bgcolor="#FFFFFF">��ȸ�� Mailzine�� ���� �˻� ����� �����ϴ�.</td></tr>
</table>
<%
End If
Set oMailzine3 = nothing
'===========================================================������ ��ȸ�� ��==================================================================
%>

<%
'===========================================================�̺�Ʈ(Ÿ��) ��ȸ�� ����==================================================================
Frealcntsum5=0
Fsuccesscntsum5=0
Fopencntsum5=0
FclickSum3=0
Fsuccesssu=0
Fopensu=0
FclickPer=0
frealopensu=0
Set oMailzine5 = new CMailzinelist
	oMailzine5.FRectyyyy 	= yyyy
	oMailzine5.FRectmm 		= mm
	oMailzine5.FGubun 		= "mailzine_event"
	oMailzine5.FMailzinelist()
If oMailzine5.FTotalcount > 0 Then
%>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF"><td colspan="15">�˻���� : <b><%= oMailzine5.FTotalCount %></b></td></tr>
<tr><td bgcolor="FFFFFF" colspan=15>5. �̺�Ʈ(Ÿ��) ���� �߼� ���</td></tr>
<tr bgcolor=#DDDDFF>
	<td align="center">�߼۳�¥</td>
	<td align="center">����</td>
	<td align="center">���Ϸ�</td>
	<td align="center">��������</td>
	<td align="center">�����߼����</td>
	<td align="center">�����߼����</td>
	<td align="center">�������</td>
	<td align="center">Ŭ�����</td>
</tr>
<%
	For i = 0 to oMailzine5.ftotalcount -1 
%>
<tr bgcolor=FFFFFF>
	<td align="center"><%= oMailzine5.FList(i).Freenddate %></td>
	<td align="center"><%= oMailzine5.FList(i).Fgubun %></td>
	<td align="center"><%= oMailzine5.FList(i).fmailergubun %></td>
	<td align="center"><%= oMailzine5.FList(i).Ftitle %></td>
	<td align="center"><%= CurrFormat(oMailzine5.FList(i).Frealcnt) %><% Frealcntsum5 = Frealcntsum5+oMailzine5.FList(i).Frealcnt %></td>
	<td align="center"><%= CurrFormat(oMailzine5.FList(i).Fsuccesscnt) %><% Fsuccesscntsum5 = Fsuccesscntsum5+oMailzine5.FList(i).Fsuccesscnt %></td>
	<td align="center"><%= CurrFormat(oMailzine5.FList(i).Fopencnt) %><% Fopencntsum5 = Fopencntsum5+oMailzine5.FList(i).Fopencnt %></td>
	<td align="center"><%= CurrFormat(oMailzine5.FList(i).FClickCnt) %><% FclickSum3 = FclickSum3+oMailzine5.FList(i).FClickCnt %></td>
</tr>
<%	Next %>
<tr bgcolor=#DDDDFF>
	<td align="center" colspan=4>�� ��</td>
	<td align="center"><%= CurrFormat(Frealcntsum5) %></td>
	<td align="center"><%= CurrFormat(Fsuccesscntsum5) %></td>
	<td align="center"><%= CurrFormat(Fopencntsum5) %></td>
	<td align="center"><%= CurrFormat(FclickSum3) %></td>
</tr>
<tr bgcolor=#DDDDFF>
	<td align="center" colspan=4>�� �� �� ��</td>
	<td align="center"></td>
	<td align="center"><% Fsuccesssu5 = (Fsuccesscntsum5/Frealcntsum5)*100 %><%= round(Fsuccesssu5,0) %>%</td>
	<td align="center"><% Fopensu5 = (Fopencntsum5/Fsuccesscntsum5)*100 %><%= round(Fopensu5,0) %>%</td>
	<td align="center"><% FclickPer3 = (FclickSum3/Fsuccesscntsum3)*100 %><%= round(FclickPer3,1) %>%</td>
</tr>
<td  bgcolor="FFFFFF" colspan=15><% frealopensu5=(Fopencntsum5/Frealcntsum5)*100 %>
  �� �� �߼������ <%= round(Fopensu5,0) %>%�� ������ �����ϴ� �ɷ� ���� �Ǿ����ϴ�.</td>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor=FFFFFF>
	<td align="center">
		<div align="right"><input type="button" value="�׷�������Ʈ" onclick="javascript:window.print();" class="button"></div><br>
		<div id="chartdiv5" align="center"></div>
		<script type="text/javascript">
			var chart = new FusionCharts("/admin/mailopen/daumchart/MSCombiDY2D.swf", "chartdiv5", "640", "480", "0", "0");
			chart.setDataURL("/admin/mailopen/daumchart/MSCombiDY2D.asp?gubun=mailzine_event");
			chart.render("chartdiv5");
		</script>
	</td>
</tr>
</table><br>
<%
Else
%>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#DDDDFF"><td align=center bgcolor="#FFFFFF">�̺�Ʈ(Ÿ��)�� ���� �˻� ����� �����ϴ�.</td></tr>
</table>
<%
End If 
Set oMailzine5 = nothing
'===========================================================�̺�Ʈ(Ÿ��) ��ȸ�� ��==================================================================
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->