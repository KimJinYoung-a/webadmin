<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionoffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<%
dim ooffsell,makerid
dim shopid

makerid = request("makerid")
shopid = session("ssBctID")

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
	if (frm.makerid.value.length<1){
		alert('�귣�带 �����ϼ���.');
		frm.makerid.focus();
		return;
	}

	document.location = 'realjaegoinput.asp?menupos=<%= menupos %>&shopid=<%= shopid %>&makerid=' + frm.makerid.value;
}

function jaegoedit(idx){
    alert('��� �ý��� �����۾����� �ѽ������� ��� �Է��� �����մϴ�.');
    return;
	document.location = 'realjaegoinput.asp?menupos=<%= menupos %>&idx=' + idx;
}
</script>
<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
			��ü:<% drawSelectBoxDesignerwithName "makerid",makerid  %>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="800" border="0" >
<tr>
	<td align="right" class="a"><input type="button" value="�ǻ���� �Է�&gt;&gt;" onClick="inputjaego();"></td>
</tr>
</table>
<table width="800" cellspacing="1" cellpadding=2 class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF">
	<td width="40">idx</td>
	<td width="70">������ID</td>
	<td width="70">�귣��</td>
	<td width="100">�ǻ�����ľ��Ͻ�</td>
	<td width="60">�����</td>
	<td width="60">����</td>
</tr>
<% for i=0 to ooffsell.FresultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td><%= ooffsell.FItemList(i).Fidx %></td>
	<td><%= ooffsell.FItemList(i).Fshopid %></td>
	<td><%= ooffsell.FItemList(i).Fmakerid %></td>
	<td><%= ooffsell.FItemList(i).Fjeagodate %></td>
	<td><%= ooffsell.FItemList(i).Fregdate %></td>
	<td><a href="javascript:jaegoedit('<%= ooffsell.FItemList(i).Fidx %>');">edit</a></td>
</tr>
<% next %>
</table>
<%
set ooffsell = Nothing
%>
<script language='javascript'>
alert('��� �ý��� �����۾����� �ѽ������� ��� �Է��� �����մϴ�.');
</script>
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->