<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �系��������
' Hieditor : �̻� ����
'			 2022.07.12 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member_board/boardCls.asp"-->
<%

Dim lBoardScmNotice
Set lBoardScmNotice = new board
	lBoardScmNotice.fnGetScmNoticeList

dim i

' ����üũ
IF Not(C_OP Or C_PSMngPart Or C_SYSTEM_Part or C_ADMIN_AUTH) Then
	Response.Write "<script type='text/javascript'>alert('�系�������� ���/������ �λ��ѹ����� �������� �����մϴ�.'); self.close();</script>"
	Response.End
End If
%>
<!-- �˻� ���� -->
<script type='text/javascript'>

function jsSubmitIns() {
	var frm = document.frmadd;

	if (frm.scheduleDate.value == '') {
		alert('������ �Է��ϼ���.');
		frm.scheduleDate.focus();
		return;
	}

	if (frm.title.value == '') {
		alert('������ �Է��ϼ���.');
		frm.title.focus();
		return;
	}

	if (frm.contents.value == '') {
		alert('������ �Է��ϼ���.');
		frm.contents.focus();
		return;
	}

	if (frm.dispno.value == '') {
		alert('ǥ�ü����� �Է��ϼ���.');
		frm.dispno.focus();
		return;
	}

	if (confirm('�����Ͻðڽ��ϱ�?') == true) {
		frm.submit();
	}
}

function jsSubmitModi(frm) {
	if (frm.scheduleDate.value == '') {
		alert('������ �Է��ϼ���.');
		frm.scheduleDate.focus();
		return;
	}

	if (frm.title.value == '') {
		alert('������ �Է��ϼ���.');
		frm.title.focus();
		return;
	}

	if (frm.contents.value == '') {
		alert('������ �Է��ϼ���.');
		frm.contents.focus();
		return;
	}

	if (frm.dispno.value == '') {
		alert('ǥ�ü����� �Է��ϼ���.');
		frm.dispno.focus();
		return;
	}

	if (confirm('�����Ͻðڽ��ϱ�?') == true) {
		frm.submit();
	}
}

function jsSubmitDel(frm) {
	if (confirm('�����Ͻðڽ��ϱ�?') == true) {
		frm.mode.value = 'del';
		frm.submit();
	}
}

</script>

<!-- ��� �� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= lBoardScmNotice.FResultCount %></b>
		</td>
	</tr>
	<tr height=30 align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">idx</td>
		<td width="100">����</td>
		<td width="120">����</td>
		<td width="210">����</td>
		<td width="100">��������</td>
		<td width="40">ǥ��<br />����</td>
		<td>���</td>
    </tr>
	<% for i = 0 to lBoardScmNotice.FResultCount - 1 %>
	<tr height=30 align="center" bgcolor="#FFFFFF">
		<form name="frmmodi<%= i %>" method="post" action="popScmNoticeModi_process.asp">
		<input type="hidden" name="mode" value="modi">
		<input type="hidden" name="idx" value="<%= lBoardScmNotice.FbrdList(i).Fidx %>">
		<td><%= lBoardScmNotice.FbrdList(i).Fidx %></td>
		<td>
			<input type="text" class="text" name="scheduleDate" value="<%= ReplaceBracket(lBoardScmNotice.FbrdList(i).FscheduleDate) %>" size="10">
		</td>
		<td>
			<input type="text" class="text" name="title" value="<%= ReplaceBracket(lBoardScmNotice.FbrdList(i).Ftitle) %>" size="15">
		</td>
		<td>
			<textarea class="textarea" name="contents" value="" cols="30" rows="3"><%= ReplaceBracket(lBoardScmNotice.FbrdList(i).Fcontents) %></textarea>
		</td>
		<td><%= lBoardScmNotice.FbrdList(i).Fmodiuserid %></td>
		<td>
			<input type="text" class="text" name="dispno" value="<%= lBoardScmNotice.FbrdList(i).Fdispno %>" size="2">
		</td>
		<td>
			<input type="button" class="button" value="�����ϱ�" onClick="jsSubmitModi(frmmodi<%= i %>)">
			&nbsp;
			<input type="button" class="button" value="�����ϱ�" onClick="jsSubmitDel(frmmodi<%= i %>)">
		</td>
		</form>
	</tr>
	<% next %>
	<tr height=30 align="center" bgcolor="#FFFFFF">
		<form name="frmadd" method="post" action="popScmNoticeModi_process.asp" style="margin:0px;">
		<input type="hidden" name="mode" value="add">
		<td>�ű�</td>
		<td>
			<input type="text" class="text" name="scheduleDate" value="" size="10">
		</td>
		<td>
			<input type="text" class="text" name="title" value="" size="15">
		</td>
		<td>
			<textarea class="textarea" name="contents" value="" cols="30" rows="3"></textarea>
		</td>
		<td><%= session("ssBctId") %></td>
		<td>
			<input type="text" class="text" name="dispno" value="" size="2">
		</td>
		<td>
			<input type="button" class="button" value="����ϱ�" onClick="jsSubmitIns()">
		</td>
		</form>
	</tr>
</table>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
