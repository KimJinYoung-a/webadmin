<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ����ڻ����
' History : 2008�� 06�� 27�� �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/menucls.asp"-->
<!-- #include virtual="/lib/classes/common/equipment/equipment_cls.asp"-->
<%
dim gubuncd ,gubuntype ,gubunname ,boxname ,ocodelist ,parameter ,i
	gubuncd= requestCheckVar(Request("gubuncd"),10)
	gubuntype = requestCheckVar(Request("gubuntype"),10)
	gubunname= requestCheckVar(Request("gubunname"),32)
	boxname= requestCheckVar(Request("boxname"),32)

set ocodelist = new cequipmentcode
	ocodelist.FPageSize = 500
	ocodelist.FCurrPage = 1
	ocodelist.frectgubuntype = gubuntype
	ocodelist.getequipmentcodelist

parameter = "gubuncd="&gubuncd&"&gubunname="&gubunname&"&boxname="&boxname
%>

<script language="javascript">

function workerselect(wid,wname)
{
	var o_wname = opener.document.getElementsByName("<%= gubunname %>")[0];
	var o_wid = opener.document.getElementsByName("<%= boxname %>")[0];

	o_wname.value =  wname;
	o_wid.value =  wid;

	temp_workerlist_js()
	window.close();
}

function temp_workerlist_js()
{
	document.getElementById("temp_workerlist").value = '<%=gubunname%>';
}

function goPartSelect(gubuntype)
{
	document.location.href = "/common/equipment/PopequipmentList.asp?gubuntype="+gubuntype+"&<%=parameter%>"
}

</script>

<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left" style="padding-bottom:10;" colspan="2">
		<input type="hidden" name="gubuntype" value="10">
	</td>
</tr>
<tr>
	<td align="left" style="padding-bottom:3;">����� : <input type="text" name="temp_workerlist" id="temp_workerlist" value="" size="10" readonly></td>
	<td align="right"><input type="button" value="�� ��" class="button" onClick="window.close()"></td>
</tr>
<tr>
	<td colspan="2"><font color="blue">�� ���õ� ����ڸ� ���� �Ͻ÷��� �ش� ����ڸ� �ѹ� �� Ŭ���Ͻø� ������ �˴ϴ�.<br>&nbsp;&nbsp;&nbsp;&nbsp;����� �Է¶��� ������� ���ð� ä���� �� ���� �ϼ���.</font></td>
</tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="#CCCCCC">
<tr bgcolor="#EFEFEF" align="center">
	<td>��ȣ</td>
	<td>����Ÿ�Ը�</td>
	<td>���ڵ�</td>
	<td>���ڵ��</td>
	<td>���ļ���</td>
	<td>��뿩��</td>
	<td>���</td>
</tr>
<% if ocodelist.fresultcount > 0 then %>
<% for i = 0 to ocodelist.fresultcount - 1 %>
<% if ocodelist.FItemList(i).fisusing = "Y" then %>
	<tr bgcolor="#ffffff" align="center">
<% else %>
	<tr bgcolor="silver" align="center">
<% end if %>
	<td><%=ocodelist.FItemList(i).fidx%></td>
	<td><%=ocodelist.FItemList(i).ftypename%> (<%=ocodelist.FItemList(i).fgubuntype%>)</td>
	<td><%=ocodelist.FItemList(i).fgubuncd%></td>
	<td><%=ocodelist.FItemList(i).fgubunname%></td>
	<td><%=ocodelist.FItemList(i).forderno%></td>
	<td><%=ocodelist.FItemList(i).fisusing%></td>
	<td>
		<input type="button" value="����" class="button" onClick="workerselect('<%=ocodelist.FItemList(i).fgubuncd%>','<%=ocodelist.FItemList(i).fgubunname%>')">
		<input type="hidden" name="gubunname" value="<%=ocodelist.FItemList(i).fgubunname%>">
	</td>
</tr>
<% next %>

<%ELSE%>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="10">��ϵ� ������ �����ϴ�.</td>
</tr>
<%End if%>

</table>

<script>
	temp_workerlist_js()
</script>

<%
Set ocodelist = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
