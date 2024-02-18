<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/MemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%

Dim idx, mode

idx		= requestCheckvar(Request("idx"),10)

mode = "modi"
if idx="" then
	idx = -1
	mode = "ins"
end if

dim oCLoginIP
Set oCLoginIP = new CLoginIP

oCLoginIP.FRectIdx = idx

oCLoginIP.GetIPOne()

dim i

%>
<script>
function jsSubmit() {
	var frm = document.frm;

	if (frm.ipaddress.value == '') {
		alert('�����Ǹ� �Է��ϼ���.');
		return;
	}

	if (frm.department_id.value == '') {
		alert('�μ��� �����ϼ���.');
		return;
	}

	if ((frm.userid.value == '') && (frm.managername.value == '') && (frm.comment.value == '')) {
		alert('���̵�/�����/�޸� �� �ϳ� �̻� ������ �Է��ϼ���.');
		return;
	}

	if (frm.usescmyn.value == '') {
		alert('���� �α��� ���θ� �Է��ϼ���.');
		return;
	}

	if (frm.uselogicsyn.value == '') {
		alert('������ �α��� ���θ� �Է��ϼ���.');
		return;
	}

	if (confirm('�����Ͻðڽ��ϱ�?') == true) {
		frm.submit();
	}
}
</script>
<%= CHKIIF(idx=-1, "[������ ����ϱ�]", "[������ ���� �����ϱ�]") %>
<form name="frm" method="POST" action="loginip_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%= mode %>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" height="30">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">idx</td>
    	<td bgcolor="#FFFFFF" align="left">
			<%= oCLoginIP.FOneItem.Fidx %>
			<input type="hidden" name="idx" value="<%= idx %>">
		</td>
    </tr>
	<tr align="center" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">IP</td>
    	<td bgcolor="#FFFFFF" align="left">
			<input type="text" class="text" name="ipaddress" value="<%= oCLoginIP.FOneItem.Fipaddress %>" size="30">
			<% if idx = -1 then %>
			* ��Ͻÿ��� 66.252.133.1,66.252.133.2 �� ���� �������� �ѹ��� ����� �� �ֽ��ϴ�.
			<% end if %>
		</td>
    </tr>
	<tr align="center" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">�μ�</td>
    	<td bgcolor="#FFFFFF" align="left">
			<%= drawSelectBoxDepartmentALL("department_id", oCLoginIP.FOneItem.Fdepartment_id) %>
		</td>
    </tr>
	<tr align="center" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">���̵�</td>
    	<td bgcolor="#FFFFFF" align="left">
			<input type="text" class="text" name="userid" value="<%= oCLoginIP.FOneItem.Fuserid %>" size="16">
		</td>
    </tr>
	<tr align="center" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
    	<td bgcolor="#FFFFFF" align="left">
			<input type="text" class="text" name="managername" value="<%= oCLoginIP.FOneItem.Fmanagername %>" size="16">
		</td>
    </tr>
	<tr align="center" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">�޸�</td>
    	<td bgcolor="#FFFFFF" align="left">
			<input type="text" class="text" name="comment" value="<%= oCLoginIP.FOneItem.Fcomment %>" size="40">
		</td>
    </tr>
	<tr align="center" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">���ηα���</td>
    	<td bgcolor="#FFFFFF" align="left">
			<select class="select" name="usescmyn" style="width:50px;">
				<option>        </option>
				<option value="Y" <%= CHKIIF(oCLoginIP.FOneItem.Fusescmyn="Y", "selected", "") %> >Y</option>
				<option value="N" <%= CHKIIF(oCLoginIP.FOneItem.Fusescmyn="N", "selected", "") %> >N</option>
			</select>
		</td>
    </tr>
	<tr align="center" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">����������ȸ����</td>
    	<td bgcolor="#FFFFFF" align="left">
			<select class="select" name="usecustomerinfoyn" style="width:50px;">
				<option>        </option>
				<option value="Y" <%= CHKIIF(oCLoginIP.FOneItem.Fusecustomerinfoyn="Y", "selected", "") %> >Y</option>
				<option value="N" <%= CHKIIF(oCLoginIP.FOneItem.Fusecustomerinfoyn="N", "selected", "") %> >N</option>
			</select>
		</td>
    </tr>
	<tr align="center" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">�������α���</td>
    	<td bgcolor="#FFFFFF" align="left">
			<select class="select" name="uselogicsyn" style="width:50px;">
				<option>        </option>
				<option value="Y" <%= CHKIIF(oCLoginIP.FOneItem.Fuselogicsyn="Y", "selected", "") %> >Y</option>
				<option value="N" <%= CHKIIF(oCLoginIP.FOneItem.Fuselogicsyn="N", "selected", "") %> >N</option>
			</select>
		</td>
    </tr>
	<tr align="center" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
    	<td bgcolor="#FFFFFF" align="left">
			<%= oCLoginIP.FOneItem.Freguserid %>
		</td>
    </tr>
	<tr align="center" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
    	<td bgcolor="#FFFFFF" align="left">
			<%= oCLoginIP.FOneItem.Fmodiuserid %>
		</td>
    </tr>
	<tr align="center" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
    	<td bgcolor="#FFFFFF" align="left">
			<select class="select" name="useyn" style="width:50px;">
				<option value="Y" <%= CHKIIF(oCLoginIP.FOneItem.Fuseyn="Y", "selected", "") %> >Y</option>
				<option value="N" <%= CHKIIF(oCLoginIP.FOneItem.Fuseyn="N", "selected", "") %> >N</option>
			</select>
		</td>
    </tr>
	<tr align="center" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">�����</td>
    	<td bgcolor="#FFFFFF" align="left">
			<%= oCLoginIP.FOneItem.Fregdate %>
		</td>
    </tr>
	<tr align="center" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
    	<td bgcolor="#FFFFFF" align="left">
			<%= oCLoginIP.FOneItem.Flastupdate %>
		</td>
    </tr>
</table>
</form>

<p />

<div align="center">
	<input type="button" class="button" value=" �� �� �� �� " onClick="jsSubmit()">
</div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
