<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ��� �ڻ� ����Ʈ
' History : 2008�� 06�� 27�� �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/bscclass/equipmentcls.asp"-->
<%
dim idx			
	idx = request("idx")

if idx="" then idx=0
	
dim oequip
set oequip = new CEquipment
	oequip.FRectIdx = idx
	oequip.getOneEquipment
%>


<!-- �������Ϸ� ���� ��� �κ� -->
<%
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename="+oequip.FOneItem.getEquipCode+left(Cstr(now()),10)+".xls"
%>

<!-- ���� �������Ͽ� ����Ǵºκ�-->
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="275" align="left" cellpadding="2" cellspacing="1" class="a" bgcolor=#000000 border="0">
<form name="frmreg" method="post" action="do_equipment.asp">
<input type="hidden" name="idx" value="<%= oequip.FOneItem.Fidx %>">
<input type="hidden" name="curruserid" value="<%= session("ssBctId") %>">
<input type="hidden" name="currusername" value="<%= session("ssBctCname") %>">
<tr bgcolor="#FFFFFF">
	<td width="75" align="right"><font size=2>����ڵ� :</font></td>
	<td colspan="2" align="left" width="172"><font size=2><%= oequip.FOneItem.getEquipCode %></font></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="75" align="right"><font size=2>��ǰ�� :</font></td>
	<td colspan="2" align="left" width="172"><font size=1.5><%= oequip.FOneItem.Fequip_name %></font></td>
</tr>
</form>
</table>

<%
	set oequip = Nothing
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->