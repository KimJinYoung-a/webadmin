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
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/partners/fingersUpcheAgreeCls.asp"-->
<%
dim makerid : makerid = session("ssBctID")
dim groupid : groupid = getPartnerId2GroupID(makerid)
dim agreestate : agreestate = requestCheckvar(request("agreestate"),10)

dim page, contractNo, ContractState, selmakerid
dim reqCtr
dim i
	page    = requestCheckVar(request("page"),10)
	contractNo  = requestCheckVar(request("contractNo"),20)
	ContractState = requestCheckVar(request("ContractState"),10)
    ''selmakerid  = requestCheckvar(request("selmakerid"),32)


''���/��� ���� ���� üũ
dim isAgreeReq 
isAgreeReq = IsFingersUpcheAgreeNotiRequire(groupid,makerid)

dim retMakeYakgan, iagreeIdx1, retMakeContract, iagreeIdx2
if (isAgreeReq) then
    retMakeContract = checkUpcheContractMake(groupid,makerid,iagreeIdx2)
   '' response.write retMakeContract
   '' response.write iagreeIdx1
    
    retMakeYakgan = checkUpcheYakganAgreeMake(groupid,makerid,iagreeIdx1)
   '' response.write retMakeYakgan
   '' response.write iagreeIdx1
end if

	if (page="") then page=1
dim ocontract
set ocontract = new CFingersUpcheAgree
	ocontract.FPageSize=20
	ocontract.FCurrPage = page
	ocontract.FRectMakerid = makerid ''selmakerid
	ocontract.FRectGroupID = groupid
	ocontract.FRectagreeState = agreeState

	ocontract.GetFingersUpcheAgreeHistList_UpcheView

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="contract.js"></script>
<script language='javascript'>
$( document ).ready(function() {
    <% if (isAgreeReq) then %>
    alert('�Ǹ��� �̿� ��� �� ��༭ ���� �� �̿� �����մϴ�.\r\n(������ �� ���� �� �ּ���.)');
    <% end if %>
});

function NextPage(page){
	frm.page.value = page;
	frm.submit();
}


function chgBrand(comp){
    var imakerid=comp.value;

    document.frm.submit();
}

function dnPdfFingers(iUri,ctrKey){
    var popwin = window.open(iUri,'dnPdf'+ctrKey,'width=1024,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function confirmContract(agreeIdx){
    var iUri = "confirmContract.asp?agreeIdx="+agreeIdx;
    var popwin = window.open(iUri,'confirmContract'+agreeIdx,'width=1024,height=768,scrollbars=yes,resizable=yes');
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
	    <% if (FALSE) then %>
        �귣�� ���� :
        <% CALL DrawSameGroupBrandUpche(groupid,selmakerid,"selmakerid","onChange='chgBrand(this)'") %>
        <% end if %>
    
        ��� ������� :
        <% Call DrawAgreeStateCombo("agreestate",agreestate) %>
    	


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
    <% if(FALSE) then%>
    <td width="100" >�Ǹ�ó</td>
    <td width="100" >����</td>
    <% end if %>
    <td width="100" >������</td>
    <td width="100" >����</td>
    <td >��༭����</td>
    <td >�ٿ�ε�</td>
</tr>
<% if ocontract.FResultCount>0 then %>
<% for i=0 to ocontract.FResultCount - 1 %>
<tr bgcolor="#FFFFFF">
    <td><%= ocontract.FITemList(i).FContractName %></td>
    <td align="center"><%= ocontract.FITemList(i).FContractNo %></td>
    <td><%= ocontract.FITemList(i).FcompanyName %></td>
    <td><%= ocontract.FITemList(i).FMakerid %></td>
    <% if(FALSE) then%>
    <td align="center"><%= "" %></td>
    <td align="center"><%= "" %></td>
    <% end if %>
    <td align="center"><%= ocontract.FITemList(i).Fagreedate  %></td>
    <td align="center"><%= ocontract.FITemList(i).getAgreeStateName  %></td>
    <td align="center"><img src="/images/iexplorer.gif" style="cursor:pointer" onClick="confirmContract('<%=ocontract.FITemList(i).FagreeIdx %>');"></td>
    <td align="center">
        <% if ocontract.FITemList(i).IsAgreeFinished then %>
        <img src="/images/pdficon.gif" style="cursor:pointer" onClick="dnPdfFingers('<%=ocontract.FITemList(i).getPdfDownLinkUrl %>','<%=ocontract.FITemList(i).FagreeIdx %>');">
        <% end if %>
    </td>
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
<tr bgcolor="#FFFFFF">
    <td colspan="13" align="center">��ϵ� �ű� ��༭�� �����ϴ�.
    </td>
</tr>
<% end if %>
</table>

<%
	set ocontract = nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->