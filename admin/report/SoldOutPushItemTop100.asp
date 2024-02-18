<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ǰ����ǰ �԰�˸� ��û��ǰ Top100
' History : 2018.01.12 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->


<%
	Dim defaultdate1, yyyy1, mm1, dd1, yyyy2, mm2, dd2, MemberShipCardDailylist, i, strTemp, strXML, ChartViDi, strDay, strWeb, strMobile, strApp, strWebLen, strMobileLen, strAppLen, strDate, strDateLen, striOs, striOsLen, strAnd, strAndLen
	Dim vbadgeGubun, sqlstr, startDate, endDate


	defaultdate1 = dateadd("d",-6,year(now) & "-" &month(now) & "-" & day(now))		'��¥���� ������ �⺻������ 10�������� �˻�	
	yyyy1 = request("yyyy1")
	if yyyy1 = "" then yyyy1 = left(defaultdate1,4)
	mm1 = request("mm1")
	if mm1 = "" then mm1 = mid(defaultdate1,6,2)
	dd1 = request("dd1")
	if dd1 = "" then dd1 = right(defaultdate1,2)	
	yyyy2 = request("yyyy2")
	if yyyy2 = "" then yyyy2 = year(now)
	mm2 = request("mm2")
	if mm2 = "" then 
		mm2 = month(now)
		If len(mm2) = 1 Then
			mm2 = "0"&mm2
		End If
	end if
	dd2 = request("dd2")
	if dd2 = "" then dd2 = day(now)
	If Len(dd2) = 1 Then
		dd2 = "0"&dd2
	End If

	startDate = yyyy1&"-"&mm1&"-"&dd1
	enddate = yyyy2&"-"&mm2&"-"&dd2

	enddate = DateAdd("d", 1, enddate)

%>



<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form action="" name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left" width="350">
		<% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>

</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<font color="red">- ǰ����ǰ �԰�˸� ��û��ǰ Top100</font>	
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="50%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ǰ�ڵ�</td>
	<td>��ǰ��</td>
	<td>��û�Ǽ�</td>
</tr>


<%
	sqlstr = " Select top 100 AA.ItemId, i.itemname, AA.cnt From "
	sqlstr = sqlstr & "       ( "
	sqlstr = sqlstr & "       	Select itemid, count(itemid) as cnt From db_my10x10.[dbo].[tbl_SoldOutProductAlarm] "
	sqlstr = sqlstr & "       	Where RegDate between '"&startDate&"' and '"&enddate&"' "
	sqlstr = sqlstr & "       	group by itemid "
	sqlstr = sqlstr & "       )AA  "
	sqlstr = sqlstr & "       inner join db_item.dbo.tbl_item i on AA.itemid = i.itemid "
	sqlstr = sqlstr & "       order by cnt desc "
	rsget.Open sqlstr, dbget, adOpenForwardOnly, adLockReadOnly
%>
<% If Not(rsget.bof Or rsget.eof) Then %>
	<% 
		Do Until rsget.eof
	%>
		<tr bgcolor="#FFFFFF" align="center">
			<td><a href="http://www.10x10.co.kr/<%=rsget("ItemId")%>" target="_blank"><%=rsget("ItemId")%></a></td>
			<td align="left">&nbsp;<%=rsget("itemname")%></td>
			<td><%=rsget("cnt")%></td>
		</tr>
	<%
		rsget.movenext
		Loop
	%>
	</table>
<% Else %>
	<table width="50%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="4">�˻� ����� �����ϴ�.</td>
	</tr>
	</table>
<%
	End If
	rsget.close
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->