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
<!-- #include virtual="/lib/classes/3pl/brandCls.asp" -->
<!-- #include virtual="/lib/classes/3pl/common.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim i, useyn, companyid
Dim page
	useyn    	= requestCheckVar(request("useyn"),32)
	companyid	= requestCheckVar(request("companyid"),32)
	page     	= requestCheckVar(request("page"),10)

If page = "" Then page = 1

if (request("research") = "")	 then
	useyn = "Y"
end if


dim oCTPLBrand
set oCTPLBrand = New CTPLBrand
	oCTPLBrand.FCurrPage					= page
	oCTPLBrand.FRectUseYN					= useyn
	oCTPLBrand.FRectCompanyID				= companyid
	oCTPLBrand.FPageSize					= 20

oCTPLBrand.GetTPLBrandList
%>

<script type="text/javascript">
function goPage(pg){
    frm.page.value = pg;
    frm.submit();
}

function jsPopModi(companyid, brandid) {
	var popwin = window.open("pop_brand_modify.asp?companyid=" + companyid + "&brandid=" + brandid,"jsPopModi","width=400 height=250 scrollbars=auto resizable=yes");
	popwin.focus();
}

function jsSubmit(frm) {
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
		고객사 : <% Call SelectBoxCompanyID("companyid", companyid, CHKIIF(useyn="Y", "Y", "")) %>
		&nbsp;
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
	<input type="button" class="button" value="등록하기" onClick="jsPopModi('', '')">
</div>

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="11">
		검색결과 : <b><%= FormatNumber(oCTPLBrand.FTotalCount,0) %></b>
		&nbsp;
		페이지 : <b> <%= FormatNumber(page,0) %> / <%= FormatNumber(oCTPLBrand.FTotalPage,0) %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width="60">IDX</td>
	<td>고객사</td>
	<td>브랜드</td>
	<td width="200">브랜드명</td>
	<td width="200">고객사<br />브랜드코드</td>
	<td width="40">사용<br />여부</td>
	<td width="40">고객사<br />사용<br />여부</td>
	<td width="180">등록일</td>
	<td width="180">최종수정</td>
    <td>비고</td>
</tr>
<% if (oCTPLBrand.FResultCount > 0) then %>
	<% for i = 0 to (oCTPLBrand.FResultCount - 1) %>
	<tr align="center" bgcolor="<%= CHKIIF(oCTPLBrand.FItemList(i).Fuseyn<>"Y" or oCTPLBrand.FItemList(i).Fcompanyuseyn<>"Y", "#DDDDDD", "#FFFFFF")%>" height="25">
  		<td><a href="javascript:jsPopModi('<%= oCTPLBrand.FItemList(i).Fcompanyid %>', '<%= oCTPLBrand.FItemList(i).Fbrandid %>')"><%= oCTPLBrand.FItemList(i).FbrandSeq %></a></td>
		<td><a href="javascript:jsPopModi('<%= oCTPLBrand.FItemList(i).Fcompanyid %>', '<%= oCTPLBrand.FItemList(i).Fbrandid %>')"><%= oCTPLBrand.FItemList(i).Fcompanyid %></a></td>
		<td><a href="javascript:jsPopModi('<%= oCTPLBrand.FItemList(i).Fcompanyid %>', '<%= oCTPLBrand.FItemList(i).Fbrandid %>')"><%= oCTPLBrand.FItemList(i).FbrandnameEng %></a></td>
		<td><a href="javascript:jsPopModi('<%= oCTPLBrand.FItemList(i).Fcompanyid %>', '<%= oCTPLBrand.FItemList(i).Fbrandid %>')"><%= oCTPLBrand.FItemList(i).Fbrandname %></a></td>
		<td><a href="javascript:jsPopModi('<%= oCTPLBrand.FItemList(i).Fcompanyid %>', '<%= oCTPLBrand.FItemList(i).Fbrandid %>')"><%= oCTPLBrand.FItemList(i).FcompanyBrandId %></a></td>
		<td><%= oCTPLBrand.FItemList(i).Fuseyn %></td>
		<td><%= oCTPLBrand.FItemList(i).Fcompanyuseyn %></td>
		<td><%= oCTPLBrand.FItemList(i).Fregdate %></td>
		<td><%= oCTPLBrand.FItemList(i).Flastupdt %></td>
		<td></td>
    </tr>
	<% next %>
	<tr height="20">
	    <td colspan="11" align="center" bgcolor="#FFFFFF">
	        <% if oCTPLBrand.HasPreScroll then %>
			<a href="javascript:goPage('<%= oCTPLBrand.StartScrollPage-1 %>');">[pre]</a>
	    	<% else %>
	    		[pre]
	    	<% end if %>

	    	<% for i=0 + oCTPLBrand.StartScrollPage to oCTPLBrand.FScrollCount + oCTPLBrand.StartScrollPage - 1 %>
	    		<% if i>oCTPLBrand.FTotalpage then Exit for %>
	    		<% if CStr(page)=CStr(i) then %>
	    		<font color="red">[<%= i %>]</font>
	    		<% else %>
	    		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
	    		<% end if %>
	    	<% next %>

	    	<% if oCTPLBrand.HasNextScroll then %>
	    		<a href="javascript:goPage('<%= i %>');">[next]</a>
	    	<% else %>
	    		[next]
	    	<% end if %>
	    </td>
	</tr>
<% else %>
    <tr height="25" bgcolor="#FFFFFF" align="center">
        <td colspan="11">검색결과가 없습니다.</td>
    </tr>
<% end if %>

</table>

<%
set oCTPLBrand = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
