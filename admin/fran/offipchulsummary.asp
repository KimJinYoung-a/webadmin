<%@ language=vbscript %>
<% option explicit %>
<%
response.write " ������ - ������ ������ ���ǿ�� "
response.End

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offshopstoragecls.asp"-->

<%
dim i, shopid, startdt, enddt, makerid
dim yyyy1,mm1

shopid = request("shopid")
makerid = request("makerid")
yyyy1 = request("yyyy1")
mm1 = request("mm1")

if yyyy1="" then
	yyyy1 = Cstr(now())
	mm1 = Mid(yyyy1,6,2)
	yyyy1 = left(yyyy1,4)
	startdt = yyyy1 + "-" + mm1 + "-01"
	enddt = CStr(DateSerial(yyyy1,mm1+1,1))
else
	startdt = yyyy1 + "-" + mm1 + "-01"
	enddt = CStr(DateSerial(yyyy1,mm1+1,1))
end if

dim ooffipchul
set ooffipchul = new COffShopStorage
ooffipchul.FRectShopid = shopid
ooffipchul.FRectStartDate = startdt
ooffipchul.FRectEndDate = enddt
ooffipchul.FRectMakerid = makerid
ooffipchul.getStorageNSellList
%>
<table width="800" border="0" cellpadding="5" cellspacing="0" class=a>
<tr>
	<td>
		* ���� �Ϳ���� ������ �����ϴ�.- ������ �����ϰڽ��ϴ�.<br>
		* ���� �԰��� �������� �ۼ��Ǿ����ϴ�. - ������ ���������� �����ϰڽ��ϴ�.(�԰� ������ ������ �ȳ��ɴϴ�.)<br>
	</td>
</tr>
</table>
<br>
<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
			��� : <% DrawYMBox yyyy1,mm1 %>
			������: <% drawSelectBoxOffShop "shopid",shopid %>
			<br>
			�귣��: <% drawSelectBoxPartnerDesigner "makerid", makerid %>

		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>
<table width="800" border="0" cellpadding="5" cellspacing="0" class=a>
<tr>
	<td align=right>
		�ѰǼ� : <%=  ooffipchul.FResultCount %>
	</td>
</tr>
</table>
<br>
<table width=800 cellpadding=0 class=a cellspacing="1" bgcolor="#3d3d3d">
<tr bgcolor="#DDDDFF" align=center>
	<td width=100>�귣��</td>
	<td width=100>��ǰ��ȣ</td>
	<td width=160>��ǰ��</td>
	<td width=80>�ɼǸ�</td>
	<td width=50>�Ϳ����</td>
	<td width=50>�԰�</td>
	<td width=50>��ǰ</td>
	<td width=50>�Ǹ�</td>
	<td width=50>R</td>
</tr>
<% for i=0 to ooffipchul.FResultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td><%= ooffipchul.FItemList(i).FMakerid %></td>
	<td><%= ooffipchul.FItemList(i).GetBargode %></td>
	<td><%= ooffipchul.FItemList(i).Fitemname %></td>
	<td><%= ooffipchul.FItemList(i).Fitemoptionname %></td>
	<td><%= ooffipchul.FItemList(i).FLastrealno %></td>
	<td><%= ooffipchul.FItemList(i).Fipno %></td>
	<td><%= ooffipchul.FItemList(i).Freno %></td>
	<td><%= ooffipchul.FItemList(i).Fsellno %></td>
	<td><%= ooffipchul.FItemList(i).GetMayno %></td>
</tr>
<% next %>
</table>
<%
set ooffipchul = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->