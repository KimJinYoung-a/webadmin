<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/eventAppReport.asp"-->
<%
Response.CharSet = "euc-kr"
Dim offEvent, sSdate, sEdate, eventNo, arrList, i, appRunUser, appRunDay, buyprice
sSdate			= requestCheckVar(request("iSD"),10)
sEdate			= requestCheckVar(request("iED"),10)
eventNo			= requestCheckVar(request("eventNo"),10)
buyprice		= request("buyprice")
appRunUser		= requestCheckVar(request("appRunUser"),1)
appRunDay		= request("appRunDay")

SET offEvent = new COffEvent
	offEvent.FRectSdate			= sSdate
	offEvent.FRectEdate			= sEdate
	offEvent.FRectEventNo		= eventNo
	offEvent.FRectBuyprice		= buyprice
	offEvent.FRectAppRunUser	= appRunUser
	offEvent.FRectAppRunDay		= appRunDay
	arrList = offEvent.fnOffEventReportByTerm
%>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="30" bgcolor="#FFFFFF">
	<td colspan="11">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<strong>�Ⱓ�� ���</strong>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="9%"></td>
	<td width="9%">�������԰���(��)</td>
	<td width="9%">���屸��(��)</td>
	<td width="9%">���� ���(��)</td>
	<td width="9%">���԰���<br>���(%)</td>
	<td width="9%">���ŰǼ�<br>���(%)</td>
	<td width="9%">�¶��� ����<br>��ȯ�ѱݾ�(��)</td>
	<td width="9%">����ȸ��(��)</td>
	<td width="9%">ȸ������(��)</td>
	<td width="9%">Ż��(��)</td>
	<td width="9%">�¶��� ����<br />��ȯ(��)</td>
</tr>
<% If IsArray(arrList) Then %>
<% For i=0 To Ubound(arrList, 2) %>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td><%= arrList(0, i) %></td>
	<td><%= FormatNumber(arrList(7, i), 0) %></td>
	<td><%= FormatNumber(arrList(8, i), 0) %></td>
	<td><%= FormatNumber(arrList(1, i), 0) %></td>
	<td>
	<%
		If arrList(7, i) <> 0 Then
			response.write Round(arrList(1, i) / arrList(7, i) * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td>
	<%
		If arrList(8, i) <> 0 Then
			response.write Round(arrList(1, i) / arrList(8, i) * 100, 1)
		Else
			response.write 0
		End If
	%>
	</td>
	<td><%= FormatNumber(arrList(6, i), 0) %></td>
	<td><%= FormatNumber(arrList(3, i), 0) %></td>
	<td><%= FormatNumber(arrList(2, i), 0) %></td>
	<td><%= FormatNumber(arrList(4, i), 0) %></td>
	<td><%= FormatNumber(arrList(5, i), 0) %></td>
</tr>
<% Next %>
<% End If %>
</table>
<% SET offEvent = nothing %>