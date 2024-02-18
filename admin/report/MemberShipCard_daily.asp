<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �����ī�� ���ϵ�����
' History : 2015.06.18 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/report/MemberShipCardDailyCls.asp"-->

<%
	Dim defaultdate1, yyyy1, mm1, dd1, yyyy2, mm2, dd2, MemberShipCardDailylist, i, strTemp, strXML, ChartViDi, strDay, strWeb, strMobile, strApp, strWebLen, strMobileLen, strAppLen, strDate, strDateLen, striOs, striOsLen, strAnd, strAndLen


	defaultdate1 = dateadd("d",-10,year(now) & "-" &month(now) & "-" & day(now))		'��¥���� ������ �⺻������ 10�������� �˻�	
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
	end if
	dd2 = request("dd2")
	if dd2 = "" then dd2 = day(now)


	set MemberShipCardDailylist = new CMemberShipCardDaily
	MemberShipCardDailylist.FRectFromDate = dateserial(yyyy1,mm1,dd1)
	MemberShipCardDailylist.FRectToDate = dateserial(yyyy2,mm2,dd2)
	MemberShipCardDailylist.GetMemberShipCardDailyReport()


%>



<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form action="" name="frm" method="get">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<% drawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">			
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<font color="red">- �Ϸ��� �����ͱ����� �˻������մϴ�.<br>- �����ʹ� 2015�� 5��1�Ϻ��� �˻� �����մϴ�.</font>	
	</td>
	<td align="right">		
	</td>
</tr>
</table>
<!-- �׼� �� -->

<% if MemberShipCardDailylist.ftotalcount > 0 then %>			
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td rowspan="2">��¥</td>
		<td rowspan="2">ī��߱޼�</td>
		<td colspan="4">�����ī�� ����/����</td>
		<td colspan="2">�¶��� ���ϸ��� ��ȯ(�����+PC)</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>�����Ǽ�</td>
		<td>������ ����Ʈ</td>
		<td>���Ǽ�</td>
		<td>���� ����Ʈ</td>
		<td>��ȯ�Ǽ�</td>
		<td>��ȯ�� ����Ʈ</td>
	</tr>
	<% for i = 0 to MemberShipCardDailylist.ftotalcount -1 %>
		<tr bgcolor="#FFFFFF" align="center">
			<td><%=MemberShipCardDailylist.FItemList(i).Fregdate%></td>
			<td><%=FormatNumber(MemberShipCardDailylist.FItemList(i).FCardRegCnt,0)%></td>
			<td><%=FormatNumber(MemberShipCardDailylist.FItemList(i).FCardSavingCnt,0)%></td>
			<td><%=FormatNumber(MemberShipCardDailylist.FItemList(i).FCardSavingPoint,0)%></td>
			<td><%=FormatNumber(MemberShipCardDailylist.FItemList(i).FCardUsingCnt,0)%></td>
			<td><%=FormatNumber(MemberShipCardDailylist.FItemList(i).FCardUsingPoint,0)%></td>
			<td><%=FormatNumber(MemberShipCardDailylist.FItemList(i).FChangeOnlineCnt,0)%></td>
			<td><%=FormatNumber(MemberShipCardDailylist.FItemList(i).FChangeOnlinePoint,0)%></td>
		</tr>
	<% next %>
	</table>
<% else %>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#FFFFFF">
		<td >�˻� ����� �����ϴ�.</td>
	</tr>
	</table>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->