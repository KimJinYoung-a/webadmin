<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim itax_no, ierrmsg
itax_no = request("itax_no")
ierrmsg = request("ierrmsg")
%>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
  <tr valign="bottom">
    <td width="10" height="10" align="right" valign="bottom" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
    <td height="10" valign="bottom" background="/images/tbl_blue_round_02.gif" bgcolor="#F3F3FF"></td>
    <td width="10" height="10" align="left" valign="bottom" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
  </tr>
  <tr valign="top">
    <td height="20" background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
    <td height="20" background="/images/tbl_blue_round_06.gif" bgcolor="#F3F3FF"><img src="/images/icon_star.gif" align="absbottom">
    <font color="red"><strong>ERROR</strong></font></td>
    <td height="20" background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
  </tr>
  <tr valign="top">
    <td background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
    <td bgcolor="#F3F3FF">
    <% if Trim(itax_no)="-2" then %>
		<br>�׿���Ʈ�� ��ϵ� ����ڹ�ȣ�� ���ο� ��ϵ� ����ڸ��� �ٸ��ų�,
		<br>ȸ�������� �Ǿ� ���� �ʽ��ϴ�.
		<br>�ϴ� �޼��� ����.
	<% elseif Trim(itax_no)="-3" then %>
		<br><b>�׿���Ʈ Ȩ���������� �̿�� ������ ����ϼ��� (�Ǵ� 200��)</b>
		<br>�ϴ� �޼��� ����.
	<% else %>
		<br>�ϴ� �޼��� ����.
	<% end if %>
	<br>
	<b>ErrCode : [<%= itax_no %>] ErrMsg : [<%= ierrmsg %>]</b>
    </td>
    <td background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
  </tr>
  <tr>
    <td background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
    <td>
    	<br>
    	�����ڵ� ����<br>
    	-2, Sender is not valid Neoport user<br>
		: �׿���Ʈ ȸ�� ������ �ȵǽ� ��� �Դϴ�. �׿���Ʈ���� ȸ�� ������ ����ϼ���.<br>
		<a href="http://www.neoport.net" target="_blank"><font color="blue">>>�׿���Ʈ ȸ�������ϱ�</font></a><br>
		<br>
		-3, �ܿ��Ǽ����� or No Remainder<br>
		: ���ݰ�꼭 ����� �Ǵ� 200���� �ݾ��� ���ݵ˴ϴ�.<br>
		�׿���Ʈ ����Ʈ �����ʿ� ���ø� [����/��ǰ ����] ��� ��ư�� ������<br>
		�̰��� ���ż� ������ǰ,�Ǵ� ���׻�ǰ�� �����Ͻ� �� ����Ͻø� �˴ϴ�.
    </td>
    <td background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
  </tr>
  <tr valign="top">
    <td background="/images/tbl_blue_round_04.gif" bgcolor="#F3F3FF"></td>
    <td bgcolor="#F3F3FF" align="right"><a href="javascript:history.back();"><strong>&lt;&lt;�ڷΰ���</strong></a></td>
    <td background="/images/tbl_blue_round_05.gif" bgcolor="#F3F3FF"></td>
  </tr>
  <tr valign="top" bgcolor="#F3F3FF">
    <td height="10" bgcolor="#F3F3FF"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td height="10" background="/images/tbl_blue_round_08.gif"></td>
    <td height="10"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
  </tr>
</table>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
