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
// ������ �̵�
function NextPage(ipage)
{
	document.frm.page.value= ipage;
	document.frm.submit();
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
        <label style="margin-right:10px;">IDX : <input type="text" name="idx" value="<%=idx%>" class="text" size="4" autocomplete="off" /></label>
		<label style="margin-right:10px;">���� : <input type="text" name="keyword" value="<%=keyword%>" class="text" size="12" autocomplete="off" /></label>
		<label style="margin-right:10px;">������ : <input type="date" name="viewDate" value="<%=viewDate%>" class="text" size="10" /></label>
		<label><input type="checkbox" name="useyn" value="Y" <%= CHKIIF(useyn="Y", "checked", "") %>> �������� ����</label>
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitFrm();">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<p />

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="�������� ���" onclick="jsPopModify('');">
	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- �׼� �� -->

<p />

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		�˻���� : <b><%= oCBuyBenefit.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCBuyBenefit.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width=60 rowspan="2">IDX</td>
	<td width=100 rowspan="2">����</td>
    <td rowspan="2">����</td>
    <td rowspan="2">������</td>
    <td width=80 rowspan="2">�����Ͻ�</td>
    <td width=80 rowspan="2">�����Ͻ�</td>
    <td width=40 rowspan="2">��ü<br />���</td>
    <td width=150 colspan="3">ä��</td>
    <td width=40 rowspan="2">����<br />����</td>
    <td width=180 rowspan="2">��������</td>
    <td rowspan="2">���</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
    <td width=40>WWW</td>
    <td width=40>�����</td>
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
