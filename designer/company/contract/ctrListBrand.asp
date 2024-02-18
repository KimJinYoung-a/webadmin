<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체 계약 관리
' Hieditor : 2013.11.20 서동석 생성
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

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
        브랜드 선택 :
        <% CALL DrawSameGroupBrandUpche(groupid,selmakerid,"selmakerid","onChange='chgBrand(this)'") %>

        계약 진행상태 :
    	<select name="ContractState" >
    	<option value="">전체
    	<option value="1" <% if ContractState="1" then response.write "selected" %> >업체오픈
    	<option value="3" <% if ContractState="3" then response.write "selected" %> >업체확인
    	<option value="7" <% if ContractState="7" then response.write "selected" %> >계약완료
    	</select>

		&nbsp;&nbsp;
		<input type="radio" name="grpType" value="" <%=CHKIIF(grpType="","checked","")%> >계약서별 보기 &nbsp;<input type="radio" name="grpType" value="M" <%=CHKIIF(grpType="M","checked","")%>>판매처/계약형태별 펼쳐보기

	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="13" align="right">총 <%= FormatNumber(ocontract.FTotalCount,0) %> 건 <%=page%>/<%=ocontract.FTotalPage%> page</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="100" >계약서 명</td>
    <td width="100" >계약서번호</td>
    <td width="80" >업체명</td>
    <td width="120" >브랜드ID</td>
    <td width="100" >판매처</td>
    <td width="100" >마진</td>
    <td width="100" >계약일</td>
    <td width="100" >상태</td>
    <td >계약서다운로드</td>
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
    call isNotFinishNewContractExists(makerid, groupid, isNewContractTypeExists)  ''신규 계약서 없는경우만 기존 계약서 링크
    %>
<tr bgcolor="#FFFFFF">
    <td colspan="13" align="center">등록된 신규 계약서가 없습니다.
    <% if (not isNewContractTypeExists) then %>
    <% if (isPreTypeContractExists(makerid)) then %> <br><br><a href="javascript:pop_PreTypeContract(0)"><font color="blue">(기존 계약서 보기)</font></a> <% end if %>
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