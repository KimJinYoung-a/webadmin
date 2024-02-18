<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ü ��� ����
' Hieditor : 2013.11.20 ������ ����
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<%
dim makerid : makerid = session("ssBctID")
dim groupid : groupid = getPartnerId2GroupID(makerid)

dim page, contractNo, ContractState, selmakerid
dim reqCtr, grpType
dim i
	page    = requestCheckVar(request("page"),10)
	contractNo  = requestCheckVar(request("contractNo"),20)
	ContractState = requestCheckVar(request("ContractState"),10)
    grpType = requestCheckvar(request("grpType"),10)
    selmakerid  = requestCheckvar(request("selmakerid"),32)

	if (page="") then page=1
dim ocontract
set ocontract = new CPartnerContract
	ocontract.FPageSize=50
	ocontract.FCurrPage = page
	ocontract.FRectMakerid = selmakerid
	ocontract.FRectGroupID = groupid
	ocontract.FRectContractState = ContractState
    ocontract.FRectGrpType  = grpType

	ocontract.GetNewContractListUpcheView

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<!--script type="text/javascript" src="contract.js"></script-->
<script language='javascript'>

function NextPage(page){
	frm.page.value = page;
	frm.submit();
}

function modiContract(ctrkey){
    var popwin = window.open('editContract.asp?ctrkey=' + ctrkey,'editContract','width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function chgBrand(comp){
    var imakerid=comp.value;

    document.frm.submit();
}

function dnPdf(iUri,ctrKey){
    var popwin = window.open(iUri,'dnPdf'+ctrKey,'width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function pop_PreTypeContract(ContractID){
    var popwin = window.open('/designer/company/popContract.asp?ContractID=' + ContractID,'popContract','width=650,height=800,scrollbars=yes,resizable=yes')
    popwin.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
        �귣�� ���� :
        <% CALL DrawSameGroupBrandUpche(groupid,selmakerid,"selmakerid","onChange='chgBrand(this)'") %>

        ��� ������� :
    	<select name="ContractState" >
    	<option value="">��ü
    	<option value="1" <% if ContractState="1" then response.write "selected" %> >��ü����
    	<option value="3" <% if ContractState="3" then response.write "selected" %> >��üȮ��
    	<option value="7" <% if ContractState="7" then response.write "selected" %> >���Ϸ�
    	</select>

		&nbsp;&nbsp;
		<input type="radio" name="grpType" value="" <%=CHKIIF(grpType="","checked","")%> >��༭�� ���� &nbsp;<input type="radio" name="grpType" value="M" <%=CHKIIF(grpType="M","checked","")%>>�Ǹ�ó/������º� ���ĺ���

	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="13" align="right">�� <%= FormatNumber(ocontract.FTotalCount,0) %> �� <%=page%>/<%=ocontract.FTotalPage%> page</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="100" >��༭ ��</td>
    <td width="100" >��༭��ȣ</td>
    <td width="80" >��ü��</td>
    <td width="120" >�귣��ID</td>
    <td width="100" >�Ǹ�ó</td>
    <td width="100" >����</td>
    <td width="100" >�����</td>
    <td width="100" >����</td>
    <td >��༭�ٿ�ε�</td>
</tr>
<% if ocontract.FResultCount>0 then %>
<% for i=0 to ocontract.FResultCount - 1 %>
<tr bgcolor="#FFFFFF">
    <td><%= ocontract.FITemList(i).FContractName %></td>
    <td align="center"><%= CHKIIF(isNULL(ocontract.FITemList(i).FctrNo) or ocontract.FITemList(i).FctrNo="","-",ocontract.FITemList(i).FctrNo) %></td>
    <td><%= ocontract.FITemList(i).FcompanyName %></td>
    <td><%= ocontract.FITemList(i).FMakerid %></td>
    <td align="center"><%= ocontract.FITemList(i).getMajorSellplaceName %></td>
    <td align="center"><%= ocontract.FITemList(i).getMajorMarginStr %></td>
    <td align="center"><%= ocontract.FITemList(i).FcontractDate %></td>
    <td align="center"><font color="<%= ocontract.FITemList(i).GetContractStateColor %>" title="<%= ocontract.FITemList(i).GetStateActiondate %>"><%= ocontract.FITemList(i).GetContractStateName %></font></td>
    <td align="center"><img src="/images/pdficon.gif" style="cursor:pointer" onClick="dnPdf('<%=ocontract.FITemList(i).getPdfDownLinkUrl %>','<%=ocontract.FITemList(i).FctrKey %>');"></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="13" align="center">
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
<% else %>
    <%
    dim isNewContractTypeExists
    call isNotFinishNewContractExists(makerid, groupid, isNewContractTypeExists)  ''�ű� ��༭ ���°�츸 ���� ��༭ ��ũ
    %>
<tr bgcolor="#FFFFFF">
    <td colspan="13" align="center">��ϵ� �ű� ��༭�� �����ϴ�.
    <% if (not isNewContractTypeExists) then %>
    <% if (isPreTypeContractExists(makerid)) then %> <br><br><a href="javascript:pop_PreTypeContract(0)"><font color="blue">(���� ��༭ ����)</font></a> <% end if %>
    <% end if %>
    </td>
</tr>
<% end if %>
</table>

<%
	set ocontract = nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->