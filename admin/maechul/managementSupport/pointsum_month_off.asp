<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������� ����Ʈ ���
' History : 2012.12.21 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/point/pointsum_off_cls.asp" -->

<%
Dim i, yyyy1,mm1,yyyy2,mm2, fromDate ,toDate ,csell
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")

if (yyyy1="") then yyyy1 = Cstr(Year( dateadd("m",-3,date()) ))
if (mm1="") then mm1 = Cstr(Month( dateadd("m",-3,date()) ))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))

fromDate = left(DateSerial(yyyy1, mm1,"01"),7)
toDate = left(DateSerial(yyyy2, mm2+1,"01"),7)

Set csell = New cpointsum_off_list
	csell.FRectStartdate = fromDate
	csell.FRectEndDate = toDate
	csell.FPageSize = 200
	csell.FCurrPage	= 1
	csell.FRectonoffgubun = "OFF"
	csell.fpointsum_sell_month_off()


dim item_S, item_S_60

dim item_X, item_Y, item_Z
dim item_XN, item_YN, item_ZN
dim item_ZZ

%>

<script language="javascript">

function searchSubmit()
{

	frm.submit();
}

function pop_sell_day(yyyy1, mm1, dd1, yyyy2, mm2, dd2){
	var pop_sell_day = window.open('/admin/maechul/managementsupport/pointsum_day_off.asp?yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&menupos=<%=menupos%>','pop_sell_day','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_sell_day.focus();
}

function pop_use_list(yyyy1, mm1, dd1, yyyy2, mm2, dd2, pointcode){
	var pop_use_list = window.open('/admin/maechul/managementsupport/pointsum_use_list_off.asp?yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&pointcode='+pointcode+'&menupos=<%=menupos%>','pop_use_list','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_use_list.focus();
}

function jsSaveCostpricepercent(frm) {
	if (frm.costpricepercent.value == "") {
		alert("�������� �Է��ϼ���.");
		return;
	}

	if (frm.costpricepercent.value*0 != 0) {
		alert("�������� ���ڸ� �����մϴ�.");
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻�</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				�Ⱓ : <% DrawYMBoxdynamic "yyyy1",yyyy1,"mm1",mm1,"" %> ~ <% DrawYMBoxdynamic "yyyy2",yyyy2,"mm2",mm2,"" %>
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<Br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="25">
		�˻���� : <b><%= csell.FresultCount %></b> �� �� 100�Ǳ��� �˻� �˴ϴ�.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>��¥</td>
    <td>�̿��ܾ�<br>(A)</td>
    <td>OFF������<br>(B)</td>
	<td>�¶�����ȯ<br>(C)</td>
    <td>������<br>(D)</td>
    <td>ȸ��Ż��<br>(E)</td>
    <td>�Ҹ�<br>(F)</td>
    <td>���ܾ�<br>(R)</td>

	<td>����<br>(S=D/(B+C))</td>
	<td>�ֱ�60����<br>����<br>(V)</td>
	<td>����<br>(X=R*V)</td>
	<td>������<br>(Y)</td>
	<td>��������<br>(Z=X*Y)</td>
	<td>��������</td>
	<td>���</td>
</tr>
<%
dim totbeforeremainpoint, totgainpoint, totspendpoint, totonlineshiftpoint, totuseroutpoint, totdelpoint, totremaincash
	totbeforeremainpoint = 0
	totgainpoint = 0
	totspendpoint = 0
	totonlineshiftpoint = 0
	totuseroutpoint = 0
	totdelpoint = 0
	totremaincash = 0

if csell.FresultCount > 0 then

For i = 0 To csell.FresultCount -1

totbeforeremainpoint = totbeforeremainpoint + csell.fitemlist(i).fbeforeremainpoint
totgainpoint = totgainpoint + csell.fitemlist(i).fgainpoint
totspendpoint = totspendpoint + csell.fitemlist(i).fspendpoint
totonlineshiftpoint = totonlineshiftpoint + csell.fitemlist(i).fonlineshiftpoint
totuseroutpoint = totuseroutpoint + csell.fitemlist(i).fuseroutpoint
totdelpoint = totdelpoint + csell.fitemlist(i).fdelpoint
totremaincash = totremaincash + csell.fitemlist(i).fremaincash


item_S = csell.fitemlist(i).fspendpoint * -1 * 100.0 / (csell.fitemlist(i).fgainpoint + csell.fitemlist(i).fonlineshiftpoint)
item_S_60 = csell.fitemlist(i).Fspendpoint60mon * -1 * 100.0 / (csell.fitemlist(i).Fgainpoint60mon + csell.fitemlist(i).fonlineshiftpoint60mon)

item_X = csell.fitemlist(i).fremaincash * (csell.fitemlist(i).Fspendpoint60mon * -1 / (csell.fitemlist(i).Fgainpoint60mon + csell.fitemlist(i).fonlineshiftpoint60mon))
item_Y = csell.fitemlist(i).Fcostpricepercent
item_Z = item_X * item_Y / 100

if (i = (csell.FresultCount - 1)) then
	item_ZZ = NULL
else
	item_XN = csell.fitemlist(i + 1).fremaincash * (csell.fitemlist(i + 1).Fspendpoint60mon * -1 / (csell.fitemlist(i + 1).Fgainpoint60mon + csell.fitemlist(i + 1).fonlineshiftpoint60mon))
	item_YN = csell.fitemlist(i + 1).Fcostpricepercent
	item_ZN = item_XN * item_YN / 100

	item_ZZ = item_Z - item_ZN
end if

%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
	<td>
		<a href="javascript:pop_sell_day('<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(csell.fitemlist(i).fYYYYMM,4),mid(csell.fitemlist(i).fYYYYMM,6,2)) %>');" onfocus="this.blur()">
		<%= csell.fitemlist(i).fYYYYMM %></a>
	</td>
	<td>
		<%= FormatNumber(csell.fitemlist(i).fbeforeremainpoint,0) %></a>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth(left(csell.fitemlist(i).fYYYYMM,4),mid(csell.fitemlist(i).fYYYYMM,6,2)) %>','gainpoint');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fgainpoint,0) %></a>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth(left(csell.fitemlist(i).fYYYYMM,4),mid(csell.fitemlist(i).fYYYYMM,6,2)) %>','onlineshiftpoint');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fonlineshiftpoint,0) %></a>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth(left(csell.fitemlist(i).fYYYYMM,4),mid(csell.fitemlist(i).fYYYYMM,6,2)) %>','spendpoint');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fspendpoint,0) %></a>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth(left(csell.fitemlist(i).fYYYYMM,4),mid(csell.fitemlist(i).fYYYYMM,6,2)) %>','useroutpoint');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fuseroutpoint,0) %></a>
	</td>
	<td>
		<a href="javascript:pop_use_list('<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(csell.fitemlist(i).fYYYYMM,4) %>','<%= mid(csell.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth(left(csell.fitemlist(i).fYYYYMM,4),mid(csell.fitemlist(i).fYYYYMM,6,2)) %>','delpoint');" onfocus="this.blur()">
		<%= FormatNumber(csell.fitemlist(i).fdelpoint,0) %></a>
	</td>
	<td>
		<%= FormatNumber(csell.fitemlist(i).fremaincash,0) %>
	</td>

	<td>
		<%= FormatNumber(item_S,2) %> %
	</td>
	<td>
		<%= FormatNumber(item_S_60,2) %> %
	</td>

	<td>
		<%= FormatNumber(item_X,0) %>
	</td>
	<td>
		<%= FormatNumber(item_Y,2) %> %
	</td>
	<td>
		<%= FormatNumber(item_Z,0) %>
	</td>
	<td>
		<% if Not IsNull(item_ZZ) then %>
			<%= FormatNumber(item_ZZ,0) %>
		<% end if %>
	</td>
	<form name="frm<%= i %>" method="post" action="pointsum_process.asp" onSubmit="return false">
	<input type="hidden" name="mode" value="modOffCostpricepercent">
	<input type="hidden" name="yyyymm" value="<%= csell.fitemlist(i).fYYYYMM %>">
	<td>
		<input type="text" class="text" name="costpricepercent" value="<%= item_Y %>" size="2">
		<input type="button" class="button" value="����" onClick="jsSaveCostpricepercent(frm<%= i %>)">
	</td>
	</form>
</tr>
<% next %>

<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>�հ�</td>
    <td><%= FormatNumber(totbeforeremainpoint,0) %></td>
    <td><%= FormatNumber(totgainpoint,0) %></td>
	<td><%= FormatNumber(totonlineshiftpoint,0) %></td>
    <td><%= FormatNumber(totspendpoint,0) %></td>
    <td><%= FormatNumber(totuseroutpoint,0) %></td>
    <td><%= FormatNumber(totdelpoint,0) %></td>
    <td><%= FormatNumber(totremaincash,0) %></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
	<td></td>
</tr>

<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="25">��ϵ� ������ �����ϴ�.</td>
</tr>
<% end if %>
</table>

<%
Set csell = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
