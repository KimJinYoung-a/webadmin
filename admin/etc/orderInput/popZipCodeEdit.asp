<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸�
' Hieditor : 2011.04.22 �̻� ����
'			 2012.08.24 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
Dim mode, zipcode, strSQL, xReceiveZipCode, xReceiveAddr1, xReceiveAddr2, zipaddr, useraddr
Dim xOutMallOrderSerial : xOutMallOrderSerial = requestCheckvar(request("outMallorderSerial"),30)
	zipcode = request("zipcode")
	zipaddr = request("zipaddr")
	useraddr= request("useraddr")
	mode = request("mode")

If mode = "U" Then
	strSQL = "UPDATE db_temp.dbo.tbl_XSite_TMporder SET "
	strSQL = strSQL & " ReceiveZipCode = '"&zipcode&"' "
	strSQL = strSQL & " ,ReceiveAddr1 = '"&zipaddr&"' "
	strSQL = strSQL & " ,ReceiveAddr2 = '"&useraddr&"' "
	strSQL = strSQL & " WHERE outmallorderserial = '"&xOutMallOrderSerial&"' "
	dbget.Execute strSQL
	response.write "<script>alert('����Ǿ����ϴ�');window.close();</script>"
	response.write "<script>opener.location.reload();</script>"
End If

strSQL = "SELECT TOP 1 ReceiveZipCode, ReceiveAddr1, ReceiveAddr2 "
strSQL = strSQL & " FROM db_temp.dbo.tbl_XSite_TMporder "
strSQL = strSQL & " WHERE outmallorderserial = '"&xOutMallOrderSerial&"' "
rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
If Not(rsget.EOF or rsget.BOF) Then
	xReceiveZipCode	= rsget("ReceiveZipCode")
	xReceiveAddr1	= rsget("ReceiveAddr1")
	xReceiveAddr2	= rsget("ReceiveAddr2")
End If
rsget.Close
%>

<script type="text/javascript">

function zipUpdate(){
	var frm;
	frm = document.frm;

	if(frm.zipcode.value==""){
		alert("�����ȣ�� �Է��ϼ���");
		frm.zipcode.focus();
		return false;
	}

	if (confirm('�����ȣ�� �����Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

function PopSearchZipcode(frmname) {
	var popwin = window.open("/lib/searchzip3.asp?target=" + frmname,"PopSearchZipcode","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>

<form name="frm" method="post" action="popZipCodeEdit.asp">
<input type="hidden" name="mode" value="U">
<input type="hidden" name="outMallorderSerial" value="<%= xOutMallOrderSerial %>">
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr height="50">
    <td align="center" bgcolor="#E8E8FF">�����ֹ���ȣ</td>
    <td bgcolor="#FFFFFF"><%= xOutMallOrderSerial %></td>
</tr>
<tr height="50">
    <td align="center" bgcolor="#E8E8FF">�����ȣ</td>
    <td bgcolor="#FFFFFF">
    	���� �����ȣ : <%= xReceiveZipCode %><br>
    	���� �����ȣ : <input type="text" class="text" name="zipcode" readonly maxlength="7">
        <input type="button" class="button" value="�˻�" onClick="FnFindZipNew('frm','B')">
		<input type="button" class="button" value="�˻�(���ּҺ���)" onClick="TnFindZipNew('frm','B')">
        <% '<input type="button" class="button" value="�˻�(��)" onClick="PopSearchZipcode('frm');"> %>
		<input type="hidden" name="zipaddr" size="50"><br>
		<input type="hidden" name="useraddr" size="50">
    </td>
</tr>
<tr height="50">
    <td align="center" bgcolor="#E8E8FF">�ּ�</td>
    <td bgcolor="#FFFFFF"><%= xReceiveAddr1 & "&nbsp;" & xReceiveAddr2 %></td>
</tr>
<tr align="center" height="25" bgcolor="#FFFFFF">
    <td colspan="2" align="center">
    <input type="button" value="����" class="button" onClick="zipUpdate();">
    </td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->