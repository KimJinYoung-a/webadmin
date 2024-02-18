<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �α��� IP ����
' Hieditor : �̻� ����
'			 2020.07.17 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/MemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
if Not(isVPNConnect) then	' or Not(C_privacyadminuser)
	'response.write "���ε� �������� �ƴմϴ�. ������ ���ǿ�� [���ٱ���:" & C_privacyadminuser & "/VPN:" & isVPNConnect & "]"
	response.write "���ε� �������� �ƴմϴ�. ������ ���ǿ�� [VPN:" & isVPNConnect & "]"
	response.end
end if

Dim page, department_id, searchRect, searchStr, useyn, i, research
	page			= requestCheckvar(Request("page"),10)
	department_id	= requestCheckvar(Request("department_id"),10)
	searchRect		= requestCheckvar(Request("searchRect"),32)
	searchStr		= requestCheckvar(Request("searchStr"),32)
	useyn			= requestCheckvar(Request("useyn"),1)
	research			= requestCheckvar(Request("research"),2)

if page="" then page=1
if research="" and useyn="" then
	useyn = "Y"
end if
dim oCLoginIP
Set oCLoginIP = new CLoginIP

oCLoginIP.FPagesize = 20
oCLoginIP.FCurrPage = page
oCLoginIP.FRectDepartment_id = department_id
oCLoginIP.FRectSearchRect = searchRect
oCLoginIP.FRectSearchStr = searchStr
oCLoginIP.FRectuseyn = useyn
oCLoginIP.GetIPList()

%>
<script type="text/javascript">

function jsGoPage(pg)
{
	document.frm.page.value=pg;
	document.frm.submit();
}

function AddItem()
{
	var pop = window.open("loginip_write_pop.asp","loginip_write_pop","width=1400,height=800,scrollbars=yes");
	pop.focus();
}

function ModiItem(idx)
{
	var pop = window.open("loginip_write_pop.asp?idx=" + idx,"loginip_write_pop","width=1400,height=800,scrollbars=yes");
	pop.focus();
}

</script>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">

<tr height="25" align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
	    * �μ� : <%= drawSelectBoxDepartmentALL("department_id", department_id) %>
		&nbsp;
		* ��뿩�� : <% drawSelectBoxisusingYN "useyn", useyn, "" %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr height="25" align="center" bgcolor="#FFFFFF" >
    <td align="left">
		* �˻����� :
        <select class="select" name="searchRect">
			<option></option>
			<option value="ipaddress" <%= CHKIIF(searchRect="ipaddress", "selected", "") %> >������</option>
			<option value="userid" <%= CHKIIF(searchRect="userid", "selected", "") %> >���̵�</option>
			<option value="managername" <%= CHKIIF(searchRect="managername", "selected", "") %> >�����</option>
			<option value="comment" <%= CHKIIF(searchRect="comment", "selected", "") %> >�޸�</option>
		</select>
		<input type="text" class="text" name="searchStr" value="<%= searchStr %>" size="20">
    </td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<p />

<input type="button" class="button" value="����ϱ�" onClick="AddItem()">

<p />

<!-- ���� ��� ���� -->
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oCLoginIP.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= oCLoginIP.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="#E6E6E6">
	<td width="50">idx</td>
	<td width="100">IP</td>
	<td width="350">�μ�</td>
	<td width="100">���̵�</td>
	<td width="100">�����</td>
	<td>�޸�</td>
	<td width="50">SCM<br />�α���</td>
	<td width="50">��������<br />��ȸ</td>
	<td width="50">������<br />�α���</td>
	<td width="50">���<br />����</td>
	<td width="100">�����</td>
	<td width="80">�����</td>
	<td width="40">���</td>
</tr>
<%
	if oCLoginIP.FResultCount=0 then
%>
<tr>
	<td colspan="13" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� ������ �����ϴ�.</td>
</tr>
<%
	else
		for i = 0 to oCLoginIP.FResultCount - 1
%>
<tr align="center" bgcolor="<% if oCLoginIP.FitemList(i).Fuseyn="Y" then Response.Write "#FFFFFF": else Response.Write "#F0F0F0": end if %>">
	<td height="25"><%= oCLoginIP.FitemList(i).Fidx %></td>
	<td><%= oCLoginIP.FitemList(i).Fipaddress %></td>
	<td><%= oCLoginIP.FitemList(i).FdepartmentnameFull %></td>
	<td><%= oCLoginIP.FitemList(i).Fuserid %></td>
	<td><%= oCLoginIP.FitemList(i).Fmanagername %></td>
	<td><%= oCLoginIP.FitemList(i).Fcomment %></td>
	<td><%= oCLoginIP.FitemList(i).Fusescmyn %></td>
	<td><%= oCLoginIP.FitemList(i).Fusecustomerinfoyn %></td>
	<td><%= oCLoginIP.FitemList(i).Fuselogicsyn %></td>
	<td><%= oCLoginIP.FitemList(i).Fuseyn %></td>
	<td><%= oCLoginIP.FitemList(i).Fmodiuserid %></td>
	<td><%= Left(oCLoginIP.FitemList(i).Flastupdate,10) %></td>
	<td><input type="button" value="����" onclick="ModiItem(<%= oCLoginIP.FitemList(i).Fidx %>);" class="button"></td>
</tr>
<%
		next
	end if
%>
</table>
<!-- ���� ��� �� -->

<!-- ������ ���� -->
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr valign="bottom" height="25">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr valign="bottom">
			<td align="center">
			<!-- ������ ���� -->
			<%
				if oCLoginIP.HasPreScroll then
					Response.Write "<a href='javascript:jsGoPage(" & oCLoginIP.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
				else
					Response.Write "[pre] &nbsp;"
				end if

				for i=0 + oCLoginIP.StartScrollPage to oCLoginIP.FScrollCount + oCLoginIP.StartScrollPage - 1

					if i>oCLoginIP.FTotalpage then Exit for

					if CStr(page)=CStr(i) then
						Response.Write " <font color='red'>[" & i & "]</font> "
					else
						Response.Write " <a href='javascript:jsGoPage(" & i & ")'>[" & i & "]</a> "
					end if

				next

				if oCLoginIP.HasNextScroll then
					Response.Write "&nbsp; <a href='javascript:jsGoPage(" & i & ")'>[next]</a>"
				else
					Response.Write "&nbsp; [next]"
				end if
			%>
			<!-- ������ �� -->
			</td>

		</tr>
		</table>
	</td>
</tr>

</table>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
