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
	'에러
	response.write "에러"
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
		alert("부서명을 입력하세요");
		frm.departmentName.focus();
		return;
	}

	if (frm.dispOrderNo.value == "") {
		alert("표시순서를 입력하세요");
		frm.dispOrderNo.focus();
		return;
	}

	if (frm.dispOrderNo.value*0 != 0) {
		alert("표시순서는 숫자만 가능합니다.");
		frm.dispOrderNo.focus();
		return;
	}

	if (confirm("저장하시겠습니까?") == true) {
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
		<td bgcolor="<%= adminColor("tabletop") %>" width="80" align="center">부서</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="departmentName" value="<%= oCTenByTenDepartment.FOneItem.FdepartmentName %>">
		</td>
	</tr>
    <tr align="left" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>" width="80" align="center">표시순서</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text" name="dispOrderNo" value="<%= oCTenByTenDepartment.FOneItem.FdispOrderNo %>">
		</td>
	</tr>
    <tr align="left" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>" width="80" align="center">사용여부</td>
		<td bgcolor="#FFFFFF">
			<select class="select" name="useYN">
				<option value="Y" <% if (oCTenByTenDepartment.FOneItem.FuseYN = "Y") then %>selected<% end if %> >사용함</option>
				<option value="N" <% if (oCTenByTenDepartment.FOneItem.FuseYN = "N") then %>selected<% end if %> >사용안함</option>
			</select>
		</td>
	</tr>
    <tr align="left" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>" width="80" align="center">등록일</td>
		<td bgcolor="#FFFFFF"><%= oCTenByTenDepartment.FOneItem.Fregdate %></td>
	</tr>
    <tr align="left" height="30">
		<td bgcolor="<%= adminColor("tabletop") %>" width="80" align="center">최종수정</td>
		<td bgcolor="#FFFFFF"><%= oCTenByTenDepartment.FOneItem.Flastupdate %></td>
	</tr>
    <tr align="left" height="50">
		<td bgcolor="#FFFFFF" colspan="2" align="center">
			<% if (pid <> "") then %>
			<input type="button" class="button" value=" 등록 " onClick="fnSubmitFrm(frm)">
			<% else %>
			<input type="button" class="button" value=" 수정 " onClick="fnSubmitFrm(frm)">
			<% end if %>
			&nbsp;
			<input type="button" class="button" value=" 취소 " onClick="fnCancelFrm()">
		</td>
	</tr>
</table>

</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
