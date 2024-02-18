<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbEVTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/seachUserCls.asp" -->
<%

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim nowdateStr, startdateStr, nextdateStr, channel
dim i
dim research

research = request("research")

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
channel = request("channel")

if (yyyy1="") then yyyy1 = Cstr(Year(DateAdd("d", -14, Now())))
if (mm1="") then mm1 = Format00(2, Cstr(Month(DateAdd("d", -14, Now()))))
if (dd1="") then dd1 = Format00(2, Cstr(day(DateAdd("d", -14, Now()))))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Format00(2, Cstr(Month(now())))
if (dd2="") then dd2 = Format00(2, Cstr(day(now())))

startdateStr = yyyy1 + "-" + mm1 + "-" + dd1
nextdateStr = yyyy2 + "-" + mm2 + "-" + dd2

if (research = "") then
	'
end if


'// ============================================================================
dim osearchUser

set osearchUser = new CSearchUser
osearchUser.FRectStart 		= startdateStr
osearchUser.FRectEnd 		= nextdateStr
osearchUser.FRectChannel	= channel
osearchUser.FRectChannel	= channel

osearchUser.getSearchUserListEVT

function GetPercentage(v1, v2)
	if (v2 = 0) then
		GetPercentage = "-"
	else
		GetPercentage = FormatNumber(100.0 * v1 / v2, 2) & "%"
	end if
end function

%>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left" height="30" >
			�Ⱓ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;
			ä�� :
			<select class="select" name="channel">
				<option></option>
				<option value="App" <%= CHKIIF(channel="App", "selected", "") %>>��</option>
				<option value="Mob" <%= CHKIIF(channel="Mob", "selected", "") %>>�����</option>
				<option value="Web" <%= CHKIIF(channel="Web", "selected", "") %>>��</option>
			</select>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left" height="30">
			* �ִ� 100�ϱ��� �˻��˴ϴ�.
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p />

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="80" height="30" rowspan="2">��¥</td>
		<td width="40" rowspan="2">����</td>
		<td width="50" rowspan="2">ä��</td>
		<td colspan="3">�˻��Ǽ�</td>
		<td colspan="3">��ǰ��ȸ�Ǽ�</td>
		<td colspan="3">��ǰ��ȸ����</td>
		<td rowspan="2">���</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="100" height="30">�˻��Ǽ�</td>
		<td width="100">�˻��ڼ�<br />(������ ����)</td>
		<td width="100">�˻��ڼ�<br />(GGSN ����)</td>
		<td width="100">��ǰ��ȸ�Ǽ�</td>
		<td width="100">��ǰ��ȸ�ڼ�<br />(������ ����)</td>
		<td width="100">��ǰ��ȸ�ڼ�<br />(GGSN ����)</td>
		<td width="100">��ǰ��ȸ/�˻�</td>
		<td width="100">��ǰ��ȸ/�˻�<br />(������ ����)</td>
		<td width="100">��ǰ��ȸ/�˻�<br />(GGSN ����)</td>
	</tr>
	<%
	for i = 0 To osearchUser.FTotalCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="30"><%= osearchUser.FItemList(i).Fyyyymmdd %></td>
		<td><%= Left(WeekDayName(WeekDay(osearchUser.FItemList(i).Fyyyymmdd)), 1) %></td>
		<td><%= osearchUser.FItemList(i).Fchannel %></td>
		<td><%= FormatNumber(osearchUser.FItemList(i).FsearchTotCnt, 0) %></td>
		<td><%= FormatNumber(osearchUser.FItemList(i).FsearchUniqipCnt, 0) %></td>
		<td><%= FormatNumber(osearchUser.FItemList(i).FsearchGgsnCnt, 0) %></td>
		<td><%= FormatNumber(osearchUser.FItemList(i).FitemviewTotCnt, 0) %></td>
		<td><%= FormatNumber(osearchUser.FItemList(i).FitemviewUniqipCnt, 0) %></td>
		<td><%= FormatNumber(osearchUser.FItemList(i).FitemviewGgsnCnt, 0) %></td>
		<td><%= GetPercentage(osearchUser.FItemList(i).FitemviewTotCnt, osearchUser.FItemList(i).FsearchTotCnt) %></td>
		<td><%= GetPercentage(osearchUser.FItemList(i).FitemviewUniqipCnt, osearchUser.FItemList(i).FsearchUniqipCnt) %></td>
		<td><%= GetPercentage(osearchUser.FItemList(i).FitemviewGgsnCnt, osearchUser.FItemList(i).FsearchGgsnCnt) %></td>
		<td></td>
	</tr>
	<%
	next
	%>
	<% if (osearchUser.FTotalCount = 0) then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="30" colspan="4">
			�˻������ �����ϴ�.
		</td>
	</tr>
	<% end if %>
</table>
<!-- ����Ʈ ���� -->

<%

Set osearchUser = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbEVTclose.asp" -->
