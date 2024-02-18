<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/PlusSaleCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim i

Dim page : page			= requestCheckvar(request("page"),8)
Dim useyn : useyn		= requestCheckvar(request("useyn"),1)
Dim research : research	= requestCheckvar(request("research"),2)
Dim keyword : keyword	= requestCheckvar(request("keyword"),32)
Dim idx : idx	= requestCheckvar(request("idx"),8)
Dim viewDate : viewDate	= requestCheckvar(request("viewDate"),10)

if page = "" then
    page = "1"
end if

if (research = "") then
    useyn = "Y"
end if


dim oCBuyBenefit
set oCBuyBenefit = new CBuyBenefit
	oCBuyBenefit.FCurrPage = page
	oCBuyBenefit.Fpagesize = 20
    oCBuyBenefit.FRectUseYN = useyn
	oCBuyBenefit.FRectIdx = idx
	oCBuyBenefit.FRectKeyword = keyword
	oCBuyBenefit.FRectViewDate = viewDate

	oCBuyBenefit.GetBuyBenefitList

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>
function SubmitFrm() {
	document.frm.submit();
}

function jsPopModify(idx) {
    var popwin = window.open("PlusSaleModify.asp?buy_benefit_idx=" + idx,"jsPopModify" + idx,"width=1200 height=600 scrollbars=yes resizable=yes");
    popwin.focus();

}
// 페이지 이동
function NextPage(ipage)
{
	document.frm.page.value= ipage;
	document.frm.submit();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
        <label style="margin-right:10px;">IDX : <input type="text" name="idx" value="<%=idx%>" class="text" size="4" autocomplete="off" /></label>
		<label style="margin-right:10px;">제목 : <input type="text" name="keyword" value="<%=keyword%>" class="text" size="12" autocomplete="off" /></label>
		<label style="margin-right:10px;">기준일 : <input type="date" name="viewDate" value="<%=viewDate%>" class="text" size="10" /></label>
		<label><input type="checkbox" name="useyn" value="Y" <%= CHKIIF(useyn="Y", "checked", "") %>> 삭제내역 제외</label>
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="SubmitFrm();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<p />

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="구매혜택 등록" onclick="jsPopModify('');">
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- 액션 끝 -->

<p />

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		검색결과 : <b><%= oCBuyBenefit.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= oCBuyBenefit.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width=60 rowspan="2">IDX</td>
	<td width=100 rowspan="2">구분</td>
    <td rowspan="2">제목</td>
    <td rowspan="2">부제목</td>
    <td width=80 rowspan="2">시작일시</td>
    <td width=80 rowspan="2">종료일시</td>
    <td width=40 rowspan="2">전체<br />대상</td>
    <td width=150 colspan="3">채널</td>
    <td width=40 rowspan="2">노출<br />순위</td>
    <td width=180 rowspan="2">최종수정</td>
    <td rowspan="2">비고</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
    <td width=40>WWW</td>
    <td width=40>모바일</td>
    <td width=40>APP</td>
</tr>
<% for i=0 to oCBuyBenefit.FResultcount-1 %>
<tr bgcolor="<%= CHKIIF(oCBuyBenefit.FItemList(i).Fuse_yn="Y", "#FFFFFF", "#EEEEEE") %>" align="center" height="40">
    <td><a href="javascript:jsPopModify(<%= oCBuyBenefit.FItemList(i).Fbuy_benefit_idx %>)"><%= oCBuyBenefit.FItemList(i).Fbuy_benefit_idx %></a></td>
    <td><%= oCBuyBenefit.FItemList(i).Fbenefit_type_name %></td>
    <td><a href="javascript:jsPopModify(<%= oCBuyBenefit.FItemList(i).Fbuy_benefit_idx %>)"><%= oCBuyBenefit.FItemList(i).Fbenefit_title %></a></td>
    <td><a href="javascript:jsPopModify(<%= oCBuyBenefit.FItemList(i).Fbuy_benefit_idx %>)"><%= oCBuyBenefit.FItemList(i).Fbenefit_subtitle %></a></td>
    <td><%= oCBuyBenefit.FItemList(i).Fbenefit_start_dt %></td>
    <td><%= oCBuyBenefit.FItemList(i).Fbenefit_end_dt %></td>
    <td><%= oCBuyBenefit.FItemList(i).Fwhole_target_yn %></td>
    <td><%= oCBuyBenefit.FItemList(i).Fchannel_www_yn %></td>
    <td><%= oCBuyBenefit.FItemList(i).Fchannel_mob_yn %></td>
    <td><%= oCBuyBenefit.FItemList(i).Fchannel_app_yn %></td>
    <td><%= oCBuyBenefit.FItemList(i).Fshow_rank %></td>
    <td><%= oCBuyBenefit.FItemList(i).Flast_update_dt %></td>
    <td></td>

</tr>
<% next %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
        	<% if oCBuyBenefit.HasPreScroll then %>
				<a href="javascript:NextPage('<%= oCBuyBenefit.StartScrollPage-1 %>');">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + oCBuyBenefit.StartScrollPage to oCBuyBenefit.FScrollCount + oCBuyBenefit.StartScrollPage - 1 %>
				<% if i>oCBuyBenefit.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if oCBuyBenefit.HasNextScroll then %>
				<a href="javascript:NextPage('<%= i %>');">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
