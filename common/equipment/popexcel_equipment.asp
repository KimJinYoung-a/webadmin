<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ����ڻ����
' History : 2008�� 06�� 27�� �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/common/equipment/equipment_cls.asp"-->
<%
dim idx			
	idx = requestCheckVar(request("idx"),10)

if idx="" then idx=0
	
dim oequip
set oequip = new CEquipment
	oequip.frectidx = idx
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
<tr bgcolor="#FFFFFF">
	<td width="75" align="right"><font size=2>����ڵ� :</font></td>
	<td colspan="2" align="left" width="172"><font size=2><%= oequip.FOneItem.getEquipCode %></font></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="75" align="right"><font size=2>��ǰ�� :</font></td>
	<td colspan="2" align="left" width="172"><font size=1.5><%= oequip.FOneItem.Fequip_name %></font></td>
</tr>
</table>
</body>

<%
set oequip = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->