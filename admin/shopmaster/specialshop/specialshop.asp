<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���ȸ����
' Hieditor : 2009.12.28 �ѿ�� ����
'			 2022.07.06 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/specialshop/specialshop_cls.asp"-->

<%
dim ospecialshop ,i,page , id , status , isusing
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
	page = requestCheckVar(getNumeric(request("page")),10)
	id = requestCheckVar(getNumeric(request("id")),10)
	status = requestCheckVar(request("status"),1)
	isusing = requestCheckVar(request("isusing"),1)
	if page = "" then page = 1

set ospecialshop = new cspecialshop_list
	ospecialshop.FPageSize = 20
	ospecialshop.FCurrPage = page
	ospecialshop.frectid = id
	ospecialshop.frectisusing = isusing
	ospecialshop.frectstatus = status
	ospecialshop.fspecialshop_list()
	
%>

<script type='text/javascript'>

// ���&����
function reg(id){
	var reg = window.open('/admin/shopmaster/specialshop/specialshop_edit.asp?id='+id,'reg','width=1200,height=600,scrollbars=yes,resizable=yes');
	reg.focus();
}

//��ǰ ���&����
function regitem(id){
	var regitem = window.open('/admin/shopmaster/specialshop/specialshop_edititem.asp?id='+id,'regitem','width=1400,height=700,scrollbars=yes,resizable=yes');
	regitem.focus();
}

//�̺�Ʈ���� �Ǽ��� ����
function statuschange(){
	var statuschange = window.open('/admin/shopmaster/specialshop/specialshop_process.asp?mode=statuschange','statuschange','width=50,height=50,scrollbars=yes,resizable=yes');
	statuschange.focus();
}

//��ǰ �Ǽ��� ����
function itemupdate(){
	var itemupdate = window.open('/admin/shopmaster/specialshop/specialshop_process.asp?mode=itemupdate','itemupdate','width=50,height=50,scrollbars=yes,resizable=yes');
	itemupdate.focus();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method=get action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			ID:<input type="text" name="id" value="<%=id%>" size=5>
			&nbsp;����:<% drawstatus "status" , status ,"" %>
			&nbsp;��뿩��:<select name="isusing" value="<%=isusing%>">
				<option value="" <% if isusing = "" then response.write " selected" %>>����</option>
				<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
			</select>						
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			
		</td>
	</tr>
</table>
</form>
<!-- �˻� �� -->

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="�������λ�ǰ�Ǽ�������" onclick="itemupdate();">
		<% if C_ADMIN_AUTH then %>
			<input type="button" class="button" value="[�����ڱ���]�Ǽ������°��ʱ�ȭ(���糯¥���������ʱ�ȭ.�ǵ���ȭ���Ͽ���������)" onclick="statuschange();"><br>
		<% end if %>
	</td>
	<td align="right">	
		<input type="button" class="button" value="�űԵ��" onclick="reg('');">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= ospecialshop.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= ospecialshop.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">   		
		<td align="center">ID</td>
		<td align="center">�׸�</td>
		<td align="center">������</td>
		<td align="center">������</td>
		<td align="center">����</td>			
		<td align="center" width=150>���</td>	
    </tr>
	<% if ospecialshop.FresultCount>0 then %>
		<% for i=0 to ospecialshop.FresultCount-1 %>
		<form action="" name="frmBuyPrc<%=i%>" method="get" style="margin:0px;">			
		
		<% if ospecialshop.FItemList(i).fisusing = "N" then %>    
		<tr align="center" bgcolor="#FFFFaa">
		<% else %>    
		<tr align="center" bgcolor="#FFFFFF">
		<% end if %>
			<td align="center">
				<%= ospecialshop.FItemList(i).fid %>			
			</td>
			<td align="center">
				<%= ReplaceBracket(ospecialshop.FItemList(i).Ftitle) %>			
			</td>
			<td align="center">
				<%= FormatDate(ospecialshop.FItemList(i).fopenDate,"0000-00-00") %>
			</td>
			<td align="center">
				<%
					If ospecialshop.FItemList(i).FendDate <> "" then
						Response.Write FormatDate(ospecialshop.FItemList(i).FendDate,"0000-00-00")
					end if
				%>
			</td>
			<td align="center">
				<%= ospecialshop.FItemList(i).fstatusstr %>
			</td>
			<td align="center">
				<input type="button" onclick="reg(<%= ospecialshop.FItemList(i).fid %>)" value="����" class="button">
				<input type="button" onclick="regitem(<%= ospecialshop.FItemList(i).fid %>)" value="��ǰ���[<%= ospecialshop.FItemList(i).fitemcount %>��]" class="button">
			</td>
		</tr>   
		</form>
		<% next %>

		<tr height="25" bgcolor="FFFFFF">
			<td colspan="15" align="center">
				<% if ospecialshop.HasPreScroll then %>
					<span class="list_link"><a href="?page=<%= ospecialshop.StartScrollPage-1 %>&id=<%=id%>&status=<%=status%>">[pre]</a></span>
				<% else %>
				[pre]
				<% end if %>
				<% for i = 0 + ospecialshop.StartScrollPage to ospecialshop.StartScrollPage + ospecialshop.FScrollCount - 1 %>
					<% if (i > ospecialshop.FTotalpage) then Exit for %>
					<% if CStr(i) = CStr(ospecialshop.FCurrPage) then %>
					<span class="page_link"><font color="red"><b><%= i %></b></font></span>
					<% else %>
					<a href="?page=<%= i %>&id=<%=id%>&status=<%=status%>" class="list_link"><font color="#000000"><%= i %></font></a>
					<% end if %>
				<% next %>
				<% if ospecialshop.HasNextScroll then %>
					<span class="list_link"><a href="?page=<%= i %>&id=<%=id%>&status=<%=status%>">[next]</a></span>
				<% else %>
				[next]
				<% end if %>
			</td>
		</tr>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
</table>

<%
set ospecialshop = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
