<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� ����Ʈ ȸ�������� ���̵� ����
' History : 2011.03.11 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
	dim oMember, arrList, iTotCnt, i
	dim username
	username = request("unm")

	'// �̸����� �˻�
	Set oMember = new CTenByTenMember
	oMember.FPagesize 		= 10
	oMember.FCurrPage 		= 1
	oMember.FSearchType 	= "2"	'�˻�����(ȸ����)
	oMember.FSearchText 	= username
	oMember.Fstatediv 		= "Y"
	oMember.Fextparttime 	= "0"	'0:�����, 1:�����̻�
		
	arrList = oMember.fnGetMemberList
	iTotCnt = oMember.FTotCnt
	set oMember = nothing

	IF isArray(arrList) THEN
%>
<script language="javascript">
<!--
	//���� ���̵� �˻�� �̵�
	function moveTenMember(uid) {
		opener.document.location="actionTenUser.asp?uid="+uid;
		self.close();
	}
//-->
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><b>�ٹ����� ���̵���</b><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="#909090">
		<tr bgcolor="#F0F0F0" align="center">
			<td>&nbsp;</td>
			<td>���̵�</td>
			<td>�μ�</td>
			<td>����</td>
			<td>�̸�</td>
			<td>�޴���</td>
		</tr>
	<% for i=0 to iTotCnt-1 %>
		<tr bgcolor="#FFFFFF" align="center">
			<td><input type="radio" name="uid" value="<%=arrList(2,i)%>" onclick="moveTenMember(this.value)"></td>
			<td><%=arrList(2,i)%></td>
			<td><%=arrList(12,i)%></td>
			<td><%=arrList(13,i)%></td>
			<td><%=arrList(1,i)%></td>
			<td><%=arrList(17,i)%></td>
		</tr>
	<% Next %>
		</table>
	</td>
</tr>
</table>
<% else %>
<script language="javascript">
alert("�˻��� �̸��� �����ϴ�.");
self.close();
</script>
<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->