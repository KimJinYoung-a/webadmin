<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �귣�� ��� ����
' Hieditor : 2010.05.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls.asp"-->
<%
dim page, makerid, catecode ,crect, mrect, contractNo, ContractState , i
	page    = request("page")
	makerid = request("makerid")
	catecode= request("catecode")
	crect   = request("crect")
	mrect   = request("mrect")
	contractNo  = request("contractNo")
	ContractState = request("ContractState")
	if (page="") then page=1

dim ocontract
set ocontract = new CPartnerContract
	ocontract.FPageSize=50
	ocontract.FCurrPage = page
	ocontract.FRectCateCode = catecode
	ocontract.FRectMakerid = makerid
	ocontract.FRectCompanyName = crect
	ocontract.FRectManagerName = mrect
	ocontract.FRectContractno  = contractNo
	ocontract.FRectContractState = ContractState
	ocontract.GetContractList
%>

<script language='javascript'>

function RegNewContract(){
    var popwin = window.open('contractReg.asp','contractReg','width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function RegContractProtoType(){
    var popwin = window.open('contractPrototypeReg.asp','contractPrototypeReg','width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function EditContract(makerid,contractID){
    var popwin = window.open('contractReg.asp?makerid=' + makerid + '&contractID=' + contractID,'contractReg','width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function NextPage(page){
	frm.page.value = page;
	frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		ī�װ� : <% SelectBoxBrandCategory "catecode", catecode %>
		&nbsp;
		&nbsp;
		�������
		<select name="ContractState">
		<option value="">��ü
		<option value="0" <% if ContractState="0" then response.write "selected" %> >������
		<option value="1" <% if ContractState="1" then response.write "selected" %> >��ü����
		<option value="3" <% if ContractState="3" then response.write "selected" %> >��üȮ��
		<option value="7" <% if ContractState="7" then response.write "selected" %> >���Ϸ�
		<option value="-2" <% if ContractState="-2" then response.write "selected" %> >-2
		<option value="-1" <% if ContractState="-1" then response.write "selected" %> >����
		</select>
		<br>
		�귣����̵� : <input type="text" name="makerid" value="<%= makerid %>" Maxlength="32" size="10">
		&nbsp;&nbsp;
		ȸ��� : <input type="text" name="crect" value="<%= crect %>" Maxlength="32" size="10">
		&nbsp;&nbsp;
		����ڸ� : <input type="text" name="mrect" value="<%= mrect %>" Maxlength="32" size="10">
		&nbsp;&nbsp;
		��༭��ȣ : <input type="text" name="contractNo" value="<%= contractNo %>" Maxlength="32" size="10">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
        	<a href="javascript:RegNewContract();"><img src="/images/icon_new_registration.gif" width="75" height="20" border="0"></a>
        	(�������޴��Դϴ�. <a href="/admin/member/contract/ctrList.asp?menupos=1619">��Ʈ�ʰ���&gt;&gt;��ü������</a> �޴��� ����ϼ���)
		</td>
		<td align="right">
        	<% if (C_ADMIN_AUTH) then %>
        	<input type="button" value="��༭ ���� ���" onClick="RegContractProtoType()" class="button">
        	<% end if %>
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="7" align="right"><%= FormatNumber(ocontract.FTotalCount,0) %></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="120" >��༭��ȣ</td>
    <td width="120" >�귣��ID</td>
    <td >��༭ ��</td>
    <td width="120" >����</td>
    <td width="90" >�����</td>
    <td width="90" >�����ID</td>
</tr>
<% for i=0 to ocontract.FResultCount - 1 %>
<tr bgcolor="#FFFFFF">
    <td><a href="javascript:EditContract('<%= ocontract.FITemList(i).FMakerid %>','<%= ocontract.FITemList(i).FContractId %>');"><%= ocontract.FITemList(i).FContractNo %></a></td>
    <td><%= ocontract.FITemList(i).FMakerid %></td>
    <td><a href="javascript:EditContract('<%= ocontract.FITemList(i).FMakerid %>','<%= ocontract.FITemList(i).FContractId %>');"><%= ocontract.FITemList(i).FContractName %></a></td>
    <td><font color="<%= ocontract.FITemList(i).GetContractStateColor %>"><%= ocontract.FITemList(i).GetContractStateName %></font></td>
    <td><%= ocontract.FITemList(i).FregDate %></td>
    <td><%= ocontract.FITemList(i).FregUserId %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="6" align="center">
        <% if ocontract.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ocontract.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ocontract.StartScrollPage to ocontract.FScrollCount + ocontract.StartScrollPage - 1 %>
			<% if i>ocontract.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ocontract.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
</tr>
</table>

<%
	set ocontract = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->