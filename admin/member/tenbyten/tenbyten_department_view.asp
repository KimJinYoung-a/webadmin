<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%

dim mode
dim pid, cid

pid       	= requestCheckvar(request("pid"),10)
cid       	= requestCheckvar(request("cid"),10)

if (pid <> "") then
	mode = "depart_modi"
elseif (cid <> "") then
	mode = "depart_modi"
else
	'����
	response.write "����"
	dbget.close()
	response.end
end if

dim oCTenByTenDepartment
set oCTenByTenDepartment = new CTenByTenDepartment
	if (cid <> "") then
		oCTenByTenDepartment.FRectCID = cid
	else
		oCTenByTenDepartment.FRectCID = -1
	end if

	oCTenByTenDepartment.GetInfo

%>
<script language="javascript">

function fnSubmitFrm(frm) {
	if (frm.departmentName.value == "") {
		alert("�μ����� �Է��ϼ���");
		frm.departmentName.focus();
		return;
	}

	if (frm.dispOrderNo.value == "") {
		alert("ǥ�ü����� �Է��ϼ���");
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

function fnCancelFrm() {
	opener.focus();
	window.close();
}

</script>
</head>
<body leftmargin="5" topmargin="5">
<form name="frm" method="POST" action="tenbyten_department_process.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="pid" value="<%= pid %>">
<input type="hidden" name="cid" value="<%= cid %>">

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="left" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>" width="80" align="center">�μ�</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="departmentName" value="<%= oCTenByTenDepartment.FOneItem.FdepartmentName %>">
		</td>
	</tr>
    <tr align="left" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>" width="80" align="center">ǥ�ü���</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="dispOrderNo" value="<%= oCTenByTenDepartment.FOneItem.FdispOrderNo %>">
		</td>
	</tr>
    <tr align="left" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>" width="80" align="center">��뿩��</td>
		<td bgcolor="#FFFFFF">
			<select class="select" name="useYN">
				<option value="Y" <% if (oCTenByTenDepartment.FOneItem.FuseYN = "Y") then %>selected<% end if %> >�����</option>
				<option value="N" <% if (oCTenByTenDepartment.FOneItem.FuseYN = "N") then %>selected<% end if %> >������</option>
			</select>
		</td>
	</tr>
    <tr align="left" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>" width="80" align="center">�����</td>
		<td bgcolor="#FFFFFF"><%= oCTenByTenDepartment.FOneItem.Fregdate %></td>
	</tr>
    <tr align="left" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>" width="80" align="center">��������</td>
		<td bgcolor="#FFFFFF"><%= oCTenByTenDepartment.FOneItem.Flastupdate %></td>
	</tr>
    <tr align="left" height="50">
		<td bgcolor="#FFFFFF" colspan="2" align="center">
			<% if (pid <> "") then %>
			<input type="button" class="button" value=" ��� " onClick="fnSubmitFrm(frm)">
			<% else %>
			<input type="button" class="button" value=" ���� " onClick="fnSubmitFrm(frm)">
			<% end if %>
			&nbsp;
			<input type="button" class="button" value=" ��� " onClick="fnCancelFrm()">
		</td>
	</tr>
</table>

</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
