<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ���� ���� ���Ѽ���
' Hieditor : 2011.01.10 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/shopmaster/shopuser_cls.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<%
dim omember , i , empno
	empno = requestcheckvar(request("empno"),32)

if empno = "" then
	response.write "<script language='javascript'>"
	response.write " 	alert('�����ȣ�� �����ϴ�');"
	response.write "	self.close();"
	response.write "</script>"
	response.end
end if

set omember = new cshopuser_list
	omember.frectempno = empno
	omember.getshopusermember_list()
%>

<script language="javascript">
	
	//�����߰�
	function shopmemberadd(){
		if (frm.shopid.value==''){
			alert('������ �������ּ���');
			frm.shopid.focus();
			return;
		}
		
		frm.action='/common/offshop/member/shopuser_process.asp';
		frm.mode.value='shopmemberadd';
		frm.submit();
	}

	//����
	function del(empno,shopid){
		if(confirm("���� �Ͻðڽ��ϱ�??") == true) {		
			location.href='/common/offshop/member/shopuser_process.asp?empno='+empno+'&shopid='+shopid+'&mode=del';
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
		�߰��� ����:<% drawSelectBoxOffShopdiv_off "shopid" , "", "1,5,11","","" %>
		<input type="button" onclick="shopmemberadd();" value="�߰�" class="button">
	</td>
	<td align="right">
	</td>
</tr>
</form>
</table>
<!-- �׼� �� -->
<br>

<% if omember.FTotalCount > 0 then %>
	<% if (C_ADMIN_AUTH) then %>
		(�����ں�) : <%= omember.FItemList(0).fid %> / <%= omember.FItemList(0).fpassword %>
	<% end if %>
<% end if %>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= omember.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�����ȣ</td>
	<td>ID</td>	
	<td>����</td>
	<td>��ǥ������</td>
	<td>���</td>
</tr>
<% if omember.ftotalcount > 0 then %>
	
<% for i=0 to omember.ftotalcount - 1 %>

<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='#ffffff';>
	<td align="center">
		<%= omember.FItemList(i).fempno %>
	</td>
	<td align="center">
		<%= omember.FItemList(i).fid %>
	</td>	
	<td align="center">
		<%= omember.FItemList(i).fshopid %>/<%= omember.FItemList(i).fshopname %>
	</td>
	<td align="center">
		<%
		if omember.FItemList(i).firstisusing = "" or isnull(omember.FItemList(i).firstisusing) then
			response.write "��������"
		else
			response.write omember.FItemList(i).firstisusing
		end if
		%>
	</td>
	<td align="center">
		<input type="button" onclick="shopfirstchange('<%= omember.FItemList(i).fempno %>','<%= omember.FItemList(i).fshopid %>');" value="����������" class="button">
		<input type="button" onclick="del('<%= omember.FItemList(i).fempno %>','<%= omember.FItemList(i).fshopid %>');" value="����" class="button">
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