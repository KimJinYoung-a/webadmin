<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/naverEp/epShopCls.asp"-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="5">���ݺ� ��Ī�Ұ� ��ǰ ����</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�׸�</td>
	<td>����</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td rowspan="4">���� ���¡</td>
	<td align="left">���� ��ǰ���� ��� ���ݿ� ���� ������ �ʹ� ���ų�, �������� ���� ���¡���� Ȯ�ε� ���</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="left">���� �� �����κ� �߰��� �ִ� ��ǰ</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="left">���ļ��� ��ǰ���� �� ��ۺ� ǥ�Ⱑ �Ǿ�������, ���θ� ������ ��ۺ� ��ǥ��� ���</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="left">���ļ��� ���ݺ� ������ ���� ������ �������� ��ۺ� ���̰� ��ǰ ������ ���ߴ� ����</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td rowspan="2">��ۺ� ����</td>
	<td align="left">���ļ��� ��ǰ���� �� ��ۺ�� ���θ� ������ ��ۺ� ������ ��� (*���Ǻ� ������ ���� ��ǰ ����</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="left">���ż��� �� ���� ��ۺ� �����ϴ� ��� (*��ǰ ���Ž� ��ۺ� ���� ��ǰ�� ���)</td>
</tr>

<tr align="center" bgcolor="#FFFFFF">
	<td>�߰�/��ǰ/��Ż</td>
	<td align="left">�߰�/��ǰ/����/��ũ��ġ ��ǰ, (*���ۺ�� ��ǰ�� �� �տ� [�߰�] Ű���带 ǥ���ϰ� �Ǹ��ϴ� ���)</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>�ؿܻ�ǰ</td>
	<td align="left">[�ؿ�],�ؿܼ���,���Ŵ���,OO�����,�۷ι�����,�۷ι�����, �۷ι����丮 �� �ؿܹ�� ��ǰ</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>�ɼ� �߰���</td>
	<td align="left">�ɼ� ���� ��ǰ�̸鼭 �ش� ��ǰ ���� �� �߰����� �߻��ϴ� ���</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>�ɼ� ǰ��</td>
	<td align="left">�ɼ� ���� ��ǰ�̸鼭 �ش� ��ǰ ���� �� ǰ���� ���</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>��ǰ�� ���¡</td>
	<td align="left">���������� ��ǰ���� �Ϻθ� �����Ͽ� ���ϻ�ǰ�� �뷮 ����ϴ� ���</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>��Ÿ</td>
	<td align="left">���θ� ������ ���� �� 19�� ���� ������ ����Ǵ� ���</td>
</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->

