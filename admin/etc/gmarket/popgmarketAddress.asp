<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : g����
' History : ������ ����
'			2019.07.31 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim sqlStr
Dim AddressTitle, AddressName, Phone1, Phone2, reqzipcode, reqzipaddr, reqaddress, AddressCode
sqlStr = ""
sqlStr = sqlStr & " SELECT TOP 1 * FROM db_etcmall.[dbo].[tbl_gmarket_AddressBook] "
rsget.Open sqlStr, dbget, 1
If Not(rsget.EOF or rsget.BOF) Then
	AddressCode		= rsget("AddressCode")
	AddressTitle	= rsget("AddressTitle")
	AddressName		= rsget("AddressName")
	Phone1			= rsget("Phone1")
	Phone2			= rsget("Phone2")
	reqzipcode		= rsget("reqzipcode")
	reqzipaddr		= rsget("reqzipaddr")
	reqaddress		= rsget("reqaddress")
End If
rsget.Close

%>
<script language="javascript">
function fnSaveForm() {
	var frm = document.frm;
    frm.target = "xLink2";
    frm.action = "/admin/etc/gmarket/procGmarket.asp"
    frm.submit();
}
function fngetCode() {
	var frm = document.frm;
    frm.target = "xLink2";
    frm.cmdparam.value = "AddAddressBook";
    frm.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
    frm.submit();
}
function fngetViewCode() {
	var frm = document.frm;
    frm.target = "xLink2";
    frm.cmdparam.value = "RequestAddressBook";
    frm.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
    frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" onsubmit="return false;">
<input type="hidden" value="saveAddress" name="mode">
<input type="hidden" name="cmdparam">
<tr bgcolor="#FFFFFF">
	<td>������ڵ�</td>
	<td><%= AddressCode %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>�ּҸ�</td>
	<td>
		<input type="text" name="AddressTitle" size="50" value="<%= AddressTitle %>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>�̸�</td>
	<td>
		<input type="text" name="AddressName" size="50" value="<%= AddressName %>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>��ȭ��ȣ</td>
	<td>
		<input type="text" name="Phone1" size="50" value="<%= Phone1 %>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>�ڵ�����ȣ</td>
	<td>
		<input type="text" name="Phone2" size="50" value="<%= Phone2 %>">
	</td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td rowspan="3" valign="top">�ּ�</td>
    <td>
        <input type="text" class="text" name="reqzipcode" value="<%= reqzipcode %>" size="7" readonly>
		<input type="button" class="button" value="�˻�" onClick="FnFindZipNew('frm','A')">
        <input type="button" class="button" value="�˻�(��)" onClick="TnFindZipNew('frm','A')">
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td ><input type="text" class="text" name="reqzipaddr" size="50" value="<%= reqzipaddr %>"></td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td>
        <input type="text" class="text" name="reqaddress" size="50" value="<%= reqaddress %>">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center">
		<input type="button" class="button" onclick="fngetViewCode();" value="��ȸ">
		<input type="button" class="button" onclick="fngetCode();" value="�ڵ���">
		<input type="button" class="button" onclick="fnSaveForm();" value="����">
	</td>
</tr>
</form>
</table>
<iframe name="xLink2" id="xLink2" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
