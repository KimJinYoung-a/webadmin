<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ [�ȳ�����] �⺻ ī�װ�
' History : �̻� ����
'			2021.09.10 �ѿ�� ����(�̹����̻�Կ�û �ڻ�� �ʵ��߰�, �ҽ�ǥ��ȭ, ���Ȱ�ȭ)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_replycls.asp"-->
<%
dim idx, gubunCode, currmode
	idx = requestcheckvar(getNumeric(request("idx")),10)
	gubunCode = requestcheckvar(request("gubunCode"),4)

dim oCReply
Set oCReply = new CReply
if (idx <> "") then
	currmode = "modiMaster"
	oCReply.FRectMasterIDX = idx
	oCReply.GetReplyMasterOne()
else
	currmode = "insMaster"
	oCReply.GetReplyMasterEmptyOne()
	oCReply.FOneItem.FgubunCode = gubunCode
end if

%>
<script type="text/javascript">

function fnSaveReplyMaster() {
	var frm = document.frm;

	if (frm.sitename.value == "") {
		alert("������ �Է��ϼ���.");
		frm.sitename.focus();
		return;
	}
	if (frm.title.value == "") {
		alert("ī�װ����� �Է��ϼ���.");
		frm.title.focus();
		return;
	}

	if (frm.dispOrderNo.value == "") {
		alert("ǥ�ü����� �Է��ϼ���.");
		frm.dispOrderNo.focus();
		return;
	}

	if (frm.dispOrderNo.value*0 != 0) {
		alert("ǥ�ü����� ���ڸ� �����մϴ�.");
		frm.dispOrderNo.focus();
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function fnGotoList() {
	document.location.href = "/cscenter/board/cs_replymaster_list.asp?menupos=<%= menupos %>";
}

</script>

<form name="frm" method="post" action="/cscenter/board/cs_reply_process.asp" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="<%= currmode %>">
<input type="hidden" name="masteridx" value="<%= oCReply.FOneItem.Fidx %>">
<input type="hidden" name="gubunCode" value="<%= oCReply.FOneItem.FgubunCode %>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td bgcolor="#FFFFFF" height="30" colspan=2>�� �⺻ ī�װ� <% if (currmode = "insMaster") then %>�ۼ�<% else %>����<% end if %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width=100 bgcolor="<%= adminColor("tabletop") %>" height="30">����</td>
	<td>
		<% Drawreplysitename "sitename", oCReply.FOneItem.fsitename, "" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width=100 bgcolor="<%= adminColor("tabletop") %>" height="30">�⺻ ī�װ���</td>
	<td>
		<input type="text" class="text" name="title" value="<%= oCReply.FOneItem.Ftitle %>" size="40">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" height="30">ǥ�ü���</td>
	<td>
		<input type="text" class="text" name="dispOrderNo" value="<%= oCReply.FOneItem.FdispOrderNo %>" size="4">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" height="30">��뱸��</td>
	<td>
		<select class="select" name="useYN">
			<option value="Y" <% if (oCReply.FOneItem.FuseYN = "Y") then %>selected<% end if %> >�����</option>
			<option value="N" <% if (oCReply.FOneItem.FuseYN = "N") then %>selected<% end if %> >������</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td bgcolor="<%= adminColor("tabletop") %>" height="30">�����</td>
	<td>
		<%= oCReply.FOneItem.Freguserid %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="30">
	<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td>
		<%= oCReply.FOneItem.Flastupdate %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" height="35" align="center">
		<input type="button" class="button" value="�����ϱ�" onclick="fnSaveReplyMaster()">
		&nbsp;
		<input type="button" class="button" value="�������" onclick="fnGotoList()">
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
