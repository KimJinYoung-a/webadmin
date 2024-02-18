<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/offshop/staff/offshop_employee_managementCls.asp"-->

<%
	Dim i, cWorkCode, vWorkCode, vStartWork, vEndWork
	vWorkCode = Request("wc")
	
	If vWorkCode <> "" Then
		Set cWorkCode = New cEmployeeManagementClass_list
		cWorkCode.FRectWorkCode = vWorkCode
		cWorkCode.fWorkCodeView()
		
		vStartWork = fnChangeTimeType(cWorkCode.FOneItem.FStartWork)
		vEndWork = fnChangeTimeType(cWorkCode.FOneItem.FEndWork)
		Set cWorkCode = Nothing
	End IF
	
	Set cWorkCode = New cEmployeeManagementClass_list
	cWorkCode.fWorkCodeList()
%>

<script type="text/javascript">
<!--
function jsEditWorkCode(wc)
{
	location.href = "<%=CurrURL()%>?wc="+wc+"";
}

function goSaveWorkCode()
{
	if(frm1.workcode.value == "")
	{
		alert("�ٹ��ڵ带 �Է��ϼ���.");
		frm1.workcode.focus();
		return false;
	}
	if(frm1.startwork.value == "")
	{
		alert("��ٽð��� �Է��ϼ���.");
		frm1.startwork.focus();
		return false;
	}
	return true;
}
//-->
</script>

�� <b>��ٽð� �Է½�</b> �ð��� �Է��Ҷ��� <b>09:00 ����</b>���� �Է��ϰ�<br><b>�Ϲݱ��ڷ� �Է�</b>�Ҷ��� <b>��ٽð��� �ݵ�� ���</b>�μ���.
<form name="frm1" action="offshop_employee_workcode_proc.asp" method="post" style="margin:0px;" onSubmit="return goSaveWorkCode();">
<input type="hidden" name="action" value="<%=CHKIIF(vWorkCode<>"","update","insert")%>">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">�ٹ��ڵ�</td>
    <td align="center" width="80">��ٽð�</td>
    <td align="center" width="80">��ٽð�</td>
    <td align="center" width="70">�ٹ��ð�</td>
    <td align="center" width="60"></td>
</tr>
<tr bgcolor="#B7F0B1" height="50">
	<td align="center"><input type="text" size="5" name="workcode" value="<%= vWorkCode %>" maxlength="2" style="text-align:center;" <%=CHKIIF(vWorkCode<>"","readonly","")%>><br>�빮��, ����Ұ�</td>
	<td align="center"><input type="text" size="7" name="startwork" value="<%= vStartWork %>" style="text-align:center;"></td>
	<td align="center"><input type="text" size="7" name="endwork" value="<%= vEndWork %>" maxlength="5" style="text-align:center;"></td>
	<td align="center">�ڵ����</td>
	<td align="center"><input type="submit" value="����" class="button"></td>
</tr>
<%
	For i = 0 To cWorkCode.FTotalCount -1
%>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= cWorkCode.flist(i).FWorkCode %></td>
	<td align="center"><%= fnChangeTimeType(cWorkCode.FList(i).FStartWork) %></td>
	<td align="center"><%= fnChangeTimeType(cWorkCode.FList(i).FEndWork) %></td>
	<td align="center"><%= fnWorkTimeCalc(cWorkCode.FList(i).FStartWork, cWorkCode.FList(i).FEndWork) %></td>
	<td align="center"><input type="button" value="����" class="button" onClick="jsEditWorkCode('<%= cWorkCode.flist(i).FWorkCode %>');"></td>
</tr>
<%
	Next
%>
</table>
</form>

<% Set cWorkCode = Nothing %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->