<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/util/base64unicode.asp" -->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<!-- #include virtual="/cscenter/lib/cs_action_mail_Function.asp"-->
<!-- #include virtual="/lib/email/MailLib2.asp"-->
<!-- #include virtual="/lib/util/DcCyberAcctUtil.asp"-->
<!-- #include virtual="/lib/classes/cscenter/sp_tenCashCls.asp" -->

<%
	Dim vQuery, vUserID, vUserName, oTenCash, vCurrentDeposit, vOrderSerial
	vUserID = Request("userid")
	
	If vUserID = "" Then
		Response.Write "<script>alert('���̵� �����ϴ�.');window.close();</script>"
		dbget.close()
		Response.End
	End IF

	Set oTenCash = New CTenCash
	oTenCash.FRectUserID = vUserID
	oTenCash.getUserCurrentTenCash
	vCurrentDeposit = oTenCash.Fcurrentdeposit
	Set oTenCash = Nothing
	
	If vCurrentDeposit = "0" Then
		Response.Write "<script>alert('��ġ���� 0�� �Դϴ�.');window.close();</script>"
		dbget.close()
		Response.End
	End IF
	
	'####### �� ������ �����Ҷ� ȸ������ �����ؾ��ϴµ� ������ ���ؿö��� ���� process �������� db ��ȸ �Է��� ���Ƽ� ���ϸ� ���̰��� �� ���������� ��ȸ.
	vQuery = "SELECT username FROM [db_user].[dbo].[tbl_user_n] WHERE userid = '" & vUserID & "'"
	rsget.Open vQuery,dbget,1
	If Not rsget.Eof Then
		vUserName = rsget(0)
	End IF
	rsget.close()
	
	
	'####### ���� ������ �ֹ���ȣ ������. ������ ��.
	vQuery = "SELECT TOP 1 orderserial FROM [db_user].[dbo].[tbl_depositlog] WHERE userid = '" & vUserID & "' AND deleteyn = 'N' AND deposit > 0 ORDER BY orderserial DESC"
	rsget.Open vQuery,dbget,1
	If Not rsget.Eof Then
		vOrderSerial = rsget(0)
	End IF
	rsget.close()
%>


<script language="javascript">
function returnToBankCash()
{
	if(isNaN(document.getElementById("returncash").value))
	{
		alert("���ڷθ� �Է��ϼ���.");
		document.getElementById("returncash").value = "<%=vCurrentDeposit%>";
		document.getElementById("returncash").focus();
		return;
	}
	if((<%=vCurrentDeposit%>-document.getElementById("returncash").value) < 0)
	{
		alert("ȯ���� ��ġ���� <%=vCurrentDeposit%>���� Ů�ϴ�.\n<%=vCurrentDeposit%>���Ϸ� �Է��� �ּ���.");
		document.getElementById("returncash").value = "<%=vCurrentDeposit%>";
		document.getElementById("returncash").focus();
		return;
	}

	if(confirm("���������� ȯ���� ������ ��Ȯ�մϱ�?") == true) {
		document.frmReturnToBankCash.submit();
	} else {
		return;
	}
}
</script>

<table class="a">
<tr height="30">
	<td style="padding-left:8px;"><img src="http://webadmin.10x10.co.kr/images/icon_arrow_link.gif"></td>
	<td style="padding-top:5px;"><b>��ġ�� ���������� ȯ��</b></td>
</tr>
</table>
<form name="frmReturnToBankCash" method="post" action="cs_popReturnToBankCash_process.asp" style="margin:0px;">
<input type="hidden" name="userid" value="<%= vUserID %>">
<input type="hidden" name="username" value="<%= vUserName %>">
<table width="380" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">���̵�</td>
  	<td bgcolor="#FFFFFF"><%=vUserID%></td>
</tr>
<tr height="25">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">���������ֹ���ȣ</td>
  	<td bgcolor="#FFFFFF"><input type="text" name="orderserial" value="<%=vOrderSerial%>"></td>
</tr>
<tr height="25">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">���¹�ȣ</td>
  	<td bgcolor="#FFFFFF">
	  	<input class="text" type="text" size="20" name="rebankaccount" value="">
	  	<input class="csbutton" type="button" value="��������" onClick="popPreReturnAcct('<%=vUserID%>','frmReturnToBankCash','rebankaccount','rebankownername','rebankname');">
  	</td>
</tr>
<tr height="25">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">�����ָ�</td>
  	<td bgcolor="#FFFFFF">
  		<input class="text" type="text" size="20" name="rebankownername" value="">
  	</td>
</tr>
<tr height="25">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">�ŷ�����</td>
  	<td bgcolor="#FFFFFF"><% DrawBankCombo "rebankname", "" %></td>
</tr>
<tr height="25">
	<td align="center" width="120" bgcolor="<%= adminColor("tabletop") %>">ȯ�Ҿ�</td>
  	<td bgcolor="#FFFFFF"><input type="text" name="returncash" id="returncash" class="input" value="<%=vCurrentDeposit%>" size="7">�� - ȯ�Ұ����Ѿ� : <%=vCurrentDeposit%>��</td>
</tr>
</table>
</form>
<table class="a" width="390">
<tr height="30">
	<td align="right"><input type="button" value="ȯ���ϱ�" class="button" onClick="returnToBankCash()"></td>
</tr>
</table>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->