<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ��ǰ���
' History : ������ ����
'			2017.04.12 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_confirmitemcls.asp"-->
<!-- #include virtual="/admin/offshop/shopitemmodi_top.asp"-->

<%
dim designer
dim acttype, ckonlyusing, ckonlyoff, imageview, pricediff
designer 	= requestCheckVar(request("designer"),32)
acttype 	= requestCheckVar(request("acttype"),32)
ckonlyusing	= requestCheckVar(request("ckonlyusing"),2)
ckonlyoff	= requestCheckVar(request("ckonlyoff"),2)
imageview	= requestCheckVar(request("imageview"),2)
pricediff	= requestCheckVar(request("pricediff"),10)

dim ooffitem
set ooffitem = new COffShopConfirm
ooffitem.FPageSize = 100
ooffitem.FRectDesigner = designer
ooffitem.GetOnOffDiffItemPriceList

dim i
%>
<script type='text/javascript'>

function SaveItems(frm){
	alert('������');
	//if (confirm('���� �귣�带 �����Ͻðڽ��ϱ�?')){
	//	frm.submit();
	//}
}

</script>
<table width="98%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<tr>
		<td class="a" >
		��ü:<% drawSelectBoxDesignerOffShopContract "designer",designer  %>
<!--
		<br>
		�ɼ�:

		<input type="checkbox" name="ckonlyusing" value="on" <% if ckonlyusing="on" then response.write "checked" %> >������λ�ǰ��
		<input type="checkbox" name="imageview" value="on" <% if imageview="on" then response.write "checked" %> >�̹�������
-->
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<center><br><br>�������Դϴ�.</center>
<table width="98%" border="0" cellspacing="1" cellpadding="2 bgcolor="#3d3d3d" class="a">
<% if ooffitem.FResultCount>0 then %>
<tr>
	<td colspan="8" align=right><input type=button value="����" onclick="SaveItems(frmarr);"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width=20><input type="checkbox" name="ckall" onClick="AnSelectAll(frmarr,this.checked)"></td>
	<td>��ǰ�ڵ�</td>
	<td>�귣��</td>
	<td>��ǰ��</td>
	<td>�ɼ�</td>
	<td width=30>����</td>
	<td>�Ǹſ���</td>
	<td>ON�ǸŰ�</td>
	<td>OFF�ǸŰ�</td>
</tr>
<% else %>
<tr>
	<td colspan="8" align=center> [ �˻������ �����ϴ�. ] </td>
</tr>
<% end if %>

<form name="frmarr" method=post action="shopitem_process.asp">
<input type="hidden" name="mode" value="makeridmodiarr">
<% for i=0 to ooffitem.FResultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" value="<%= ooffitem.FItemList(i).FShopItemId %>"></td>
	<td><%= ooffitem.FItemList(i).FShopItemId %></td>
	<td><%= ooffitem.FItemList(i).FMakerid %></td>
	<td><%= ooffitem.FItemList(i).FShopItemName %></td>
	<td><%= ooffitem.FItemList(i).FShopItemOptionName %></td>
	<td align=center>
		<% if ooffitem.FItemList(i).Fonlinesailyn="Y" then %>
		<font color="red">����</font>
		<% end if %>
	</td>
	<td align=right>
		<% if (ooffitem.FItemList(i).Fonlineorgprice<>ooffitem.FItemList(i).FShopItemPrice) then %>
		<%= FormatNumber(ooffitem.FItemList(i).Fonlineorgprice,0) %>
		<% end if %>
	</td>
	<td align=right><%= FormatNumber(ooffitem.FItemList(i).FOnlineSellcash,0) %></td>
	<td align=right><%= FormatNumber(ooffitem.FItemList(i).FShopItemPrice,0) %></td>
</tr>
<% next %>
</table>
</form>

<%
set ooffitem = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->