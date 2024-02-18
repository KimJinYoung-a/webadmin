<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  cs 메모
' History : 2007.01.01 이상구 생성
'           2016.12.07 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/3pl/partnerCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim i, useyn
Dim page
	useyn    = requestCheckVar(request("useyn"),32)
	page     = requestCheckVar(request("page"),10)

If page = "" Then page = 1

if (request("research") = "")	 then
	useyn = "Y"
end if


dim oCTPLPartner
set oCTPLPartner = New CTPLPartner
	oCTPLPartner.FCurrPage					= page
	oCTPLPartner.FRectUseYN					= useyn
	oCTPLPartner.FPageSize					= 100

oCTPLPartner.GetTPLPartnerCompanyList
%>

<script type="text/javascript">
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function jsPopModi(companyid, partnercompanyid) {
	var popwin = window.open("pop_partnercompany_modify.asp?companyid=" + companyid + "&partnercompanyid=" + partnercompanyid,"jsPopModi","width=400 height=170 scrollbars=auto resizable=yes");
	popwin.focus();
}

function jsSubmit() {
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" height="30" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		사용여부 : <% Call drawSelectBoxUsingYN("useyn", useyn) %>
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:jsSubmit(frm);">
	</td>
</tr>
</table>
</form>

<p />

<div align="right">
	<input type="button" class="button" value="등록하기" onClick="jsPopModi('', '-1')">
</div>

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		검색결과 : <b><%= FormatNumber(oCTPLPartner.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCTPLPartner.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width="120">아이디</td>
	<td width="300">고객사</td>
	<td width="300">제휴사</td>
	<td width="40">API</td>
	<td width="40">사용<br />여부</td>
	<td width="180">등록일</td>
	<td width="180">최종수정</td>
    <td>비고</td>
</tr>
<% if (oCTPLPartner.FResultCount > 0) then %>
	<% for i = 0 to (oCTPLPartner.FResultCount - 1) %>
<tr align="center" bgcolor="<%= CHKIIF(oCTPLPartner.FItemList(i).Fuseyn<>"Y", "#DDDDDD", "#FFFFFF")%>" height="25">
  		<td><a href="javascript:jsPopModi('<%= oCTPLPartner.FItemList(i).Fcompanyid %>', '<%= oCTPLPartner.FItemList(i).Fpartnercompanyid %>')"><%= oCTPLPartner.FItemList(i).Fcompanyid %></a></td>
		<td><a href="javascript:jsPopModi('<%= oCTPLPartner.FItemList(i).Fcompanyid %>', '<%= oCTPLPartner.FItemList(i).Fpartnercompanyid %>')"><%= oCTPLPartner.FItemList(i).Fcompanyname %></a></td>
		<td><a href="javascript:jsPopModi('<%= oCTPLPartner.FItemList(i).Fcompanyid %>', '<%= oCTPLPartner.FItemList(i).Fpartnercompanyid %>')"><%= oCTPLPartner.FItemList(i).Fpartnercompanyname %></a></td>
		<td><%= oCTPLPartner.FItemList(i).FapiAvail %></td>
		<td><%= oCTPLPartner.FItemList(i).Fuseyn %></td>
		<td><%= oCTPLPartner.FItemList(i).Fregdate %></td>
		<td><%= oCTPLPartner.FItemList(i).Flastupdt %></td>
		<td></td>
    </tr>
	<% next %>
	<tr height="20">
	    <td colspan="8" align="center" bgcolor="#FFFFFF">
	        <% if oCTPLPartner.HasPreScroll then %>
			<a href="javascript:goPage('<%= oCTPLPartner.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + oCTPLPartner.StartScrollPage to oCTPLPartner.FScrollCount + oCTPLPartner.StartScrollPage - 1 %>
	    		<% if i>oCTPLPartner.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if oCTPLPartner.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
<% else %>
    <tr height="25" bgcolor="#FFFFFF" align="center">
        <td colspan="8">검색결과가 없습니다.</td>
    </tr>
<% end if %>

</table>

<%
set oCTPLPartner = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
