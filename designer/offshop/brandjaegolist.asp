<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<%
dim ooffsell,makerid
dim shopid

makerid = session("ssBctID")
shopid = request("shopid")

set ooffsell = new COffShopDailyStock
ooffsell.FRectMakerid = makerid
ooffsell.FRectShopId = shopid
ooffsell.GetRealJaegoList

dim i
%>
<script language='javascript'>
function inputjaego(){
    alert('��� �ý��� �����۾����� �ѽ������� ��� �Է��� �����մϴ�.');
    return;
    
	if (frm.shopid.value.length<1){
		alert('���� �����ϼ���.');
		frm.shopid.focus();
		return;
	}

	document.location = 'realjaegoinput.asp?menupos=<%= menupos %>&shopid=' + frm.shopid.value;
}

function jaegoedit(idx){
    alert('��� �ý��� �����۾����� �ѽ������� ��� �Է��� �����մϴ�.');
    return;
	document.location = 'realjaegoinput.asp?menupos=<%= menupos %>&idx=' + idx;
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			SHOP : <% drawSelectBoxOpenOffShop "shopid",shopid %>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="�ǻ���� �Է�" onClick="inputjaego();">
		</td>
		<td align="right">
			
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="40">idx</td>
		<td width="100">������ID</td>
		<td width="150">�귣��ID</td>
		<td width="150">�ǻ�����ľ��Ͻ�</td>
		<td width="150">�����</td>
		<td>����</td>
	</tr>
	<% for i=0 to ooffsell.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= ooffsell.FItemList(i).Fidx %></td>
		<td><%= ooffsell.FItemList(i).Fshopid %></td>
		<td><%= ooffsell.FItemList(i).Fmakerid %></td>
		<td><%= ooffsell.FItemList(i).Fjeagodate %></td>
		<td><%= ooffsell.FItemList(i).Fregdate %></td>
		<td><input type="button" class="button" value="����" onClick="jaegoedit('<%= ooffsell.FItemList(i).Fidx %>');"></td>
	</tr>
	<% next %>
</table>
<%
set ooffsell = Nothing
%>
<script language='javascript'>
alert('��� �ý��� �����۾����� �ѽ������� ��� �Է��� �����մϴ�.');
</script>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->