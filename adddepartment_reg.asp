<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �����μ� �߰�
' Hieditor : 2017.08.22 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenAddDepCls.asp"-->
<%
dim omember , i , empno
dim clsAddDep, arrList, intLoop
	empno = requestcheckvar(request("empno"),32)

if empno = "" then
	response.write "<script language='javascript'>"
	response.write " 	alert('�����ȣ�� �����ϴ�');"
	response.write "	self.close();"
	response.write "</script>"
	response.end
end if

set clsAddDep = new CAddDep
  clsAddDep.Fempno = empno
  arrList = clsAddDep.fnGetAddDepList
set clsAddDep = nothing 
 
%>

<script language="javascript">
	
	//�μ��߰�
	function jsAdddep(){
		if (frm.department_id.value==''){
			alert('�μ��� �������ּ���');
			frm.department_id.focus();
			return;
		}
		
		frm.action='/common/offshop/member/adddepartment_process.asp';
		frm.mode.value='A';
		frm.submit();
	}

	//����
	function del(empno,shopid){
		if(confirm("���� �Ͻðڽ��ϱ�??") == true) {		
			location.href='/common/offshop/member/adddepartment_process.asp?empno='+empno+'&shopid='+shopid+'&mode=del';
		} else {
			return;
		}	
	}
	
	//����������
	function shopfirstchange(empno,shopid){
		if(confirm("�����Ͻ� ������ ��ǥ���������� ���� �Ͻðڽ��ϱ�??") == true) {		
			location.href='/common/offshop/member/shopuser_process.asp?empno='+empno+'&shopid='+shopid+'&mode=shopfirstchange';
		} else {
			return;
		}	
	}
		
</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frm" method="post">
<input type="hidden" name="mode">
<input type="hidden" name="empno" value="<%=empno%>">
<tr>
	<td align="left">
		�߰��� �μ�:<%= drawSelectBoxDepartment("department_id", "") %>
		<input type="button" onclick="jsAdddep();" value="�߰�" class="button">
	</td>
	<td align="right">
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->
<br>
 

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= omember.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>���</td>
	<td>���̵�</td>	
	<td>�μ�</td> 
	<td>���</td>
</tr>
<% if isArray(arrList) then %>
	
<% for intLoop=0 to arrList(2,intLoop)  %>

<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background="#ffffff";>
	<td align="center">
		<%= arrList(0,intLoop) %>
	</td>
	<td align="center">
		<%= arrList(1,intLoop) %>
	</td>	
	<td align="center">
				<%= arrList(3,intLoop) %>
	</td> 
	<td align="center">
		<input type="button" onclick=" " value="����������" class="button">
		<input type="button" onclick=" ;" value="����" class="button">
	</td>	
</tr>   
<% next %>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>


<%
set omember = nothing
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->