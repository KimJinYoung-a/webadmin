<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/PlusSaleCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim i

Dim page : page			= requestCheckvar(request("page"),10)
Dim useyn : useyn		= requestCheckvar(request("useyn"),1)
Dim research : research	= requestCheckvar(request("research"),2)
Dim keyword : keyword	= requestCheckvar(request("keyword"),32)
Dim idx : idx	= requestCheckvar(request("idx"),8)
Dim exzyn : exzyn		= requestCheckvar(request("exzyn"),1)

if page = "" then
    page = "1"
end if

if (research = "") then
    useyn = "Y"
    exzyn = "Y"
end if

dim oCBuyBenefit
set oCBuyBenefit = new CBuyBenefit
	oCBuyBenefit.FCurrPage = page
	oCBuyBenefit.Fpagesize = 50
    oCBuyBenefit.FRectUseYN = useyn
    oCBuyBenefit.FRectExistYN = exzyn
	oCBuyBenefit.FRectIdx = idx
	oCBuyBenefit.FRectKeyword = keyword
	oCBuyBenefit.GetBuyBenefitStat

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script type='text/javascript'>
function SubmitFrm() {
	document.frm.page.value=1;
    document.frm.submit();
}

function jsUpdateOrderCount(benefit_group_no) {
    var frm = document.frmAct;
    frm.mode.value = "updOrderCount";
    frm.benefit_group_no.value = benefit_group_no;
    frm.submit();
}

function jsPopBenifit(buy_benefit_idx) {
    var popwin = window.open("PlusSaleModify.asp?buy_benefit_idx=" + buy_benefit_idx,"jsPopBenifit" + buy_benefit_idx,"width=1200 height=600 scrollbars=yes resizable=yes");
    popwin.focus();
}

function jsViewSold(buy_benefit_idx) {
    var popwin = window.open("PlusSaleSoldItemList.asp?buy_benefit_idx=" + buy_benefit_idx,"jsViewSold" + buy_benefit_idx,"width=800 height=600 scrollbars=yes resizable=yes");
    popwin.focus();
}

function jsPopBenifitGroup(buy_benefit_idx, benefit_group_no) {
    //PlusSaleGroupModify.asp?buy_benefit_idx=1&benefit_group_no=1
    var popwin = window.open("PlusSaleGroupModify.asp?buy_benefit_idx=" + buy_benefit_idx + "&benefit_group_no=" + benefit_group_no,"jsPopBenifitGroup" + benefit_group_no,"width=1200 height=600 scrollbars=yes resizable=yes");
    popwin.focus();
}

function NextPage(pg) {
	document.frm.page.value=pg;
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
		<label style="margin-right:10px;"><input type="checkbox" name="useyn" value="Y" <%= CHKIIF(useyn="Y", "checked", "") %>> �������� ����</label>
        <label><input type="checkbox" name="exzyn" value="Y" <%= CHKIIF(exzyn="Y", "checked", "") %>> ������� ����</label>
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitFrm();">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

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
	<td width=120>������ IDX</td>
    <td>���ø�</td>
    <td width=120>���ñ׷� IDX</td>
    <td>���ñ׷��</td>
    <td width=120>���<br />�ֹ��Ǽ�</td>
    <td width=120>�÷�������<br />�ֹ��Ǽ�</td>
    <td width=120>���ź���</td>
    <td width=120>�÷�������<br />�ֹ���ǰ��</td>
    <td>���</td>
</tr>
<% for i=0 to oCBuyBenefit.FResultcount-1 %>
<tr bgcolor="#FFFFFF" align="center" height="45">
    <td>
        <a href="javascript:jsPopBenifit(<%= oCBuyBenefit.FItemList(i).Fbuy_benefit_idx %>)">
            <%= oCBuyBenefit.FItemList(i).Fbuy_benefit_idx %>
        </a>
    </td>
    <td><%= oCBuyBenefit.FItemList(i).Fbenefit_title %></td>
    <td>
        <a href="javascript:jsPopBenifitGroup(<%= oCBuyBenefit.FItemList(i).Fbuy_benefit_idx %>, <%= oCBuyBenefit.FItemList(i).Fbenefit_group_no %>)">
            <%= oCBuyBenefit.FItemList(i).Fbenefit_group_no %>
        </a>
    </td>
    <td><%= oCBuyBenefit.FItemList(i).Fgroup_name %></td>
    <td><%= FormatNumber(oCBuyBenefit.FItemList(i).FtargetOrderCount, 0) %></td>
    <td><%= FormatNumber(oCBuyBenefit.FItemList(i).ForderCnt, 0) %></td>
    <td>
        <% if oCBuyBenefit.FItemList(i).FtargetOrderCount = 0 then %>
        --
        <% elseif oCBuyBenefit.FItemList(i).FtargetOrderCount > 0 then %>
        <%= Round(100 * oCBuyBenefit.FItemList(i).ForderCnt / oCBuyBenefit.FItemList(i).FtargetOrderCount, 1) %> %
        <% end if %>
    </td>
    <td><%= FormatNumber(oCBuyBenefit.FItemList(i).FItemCnt, 0) %></td>
    <td>
        <% if oCBuyBenefit.FItemList(i).ForderCnt > 0 then %>
        <input type="button" class="button" value=" <%=chkIIF(oCBuyBenefit.FItemList(i).FtargetOrderCount>0,"������Ʈ","�ڷ����")%> " onClick="jsUpdateOrderCount(<%= oCBuyBenefit.FItemList(i).Fbenefit_group_no %>)">
        <% end if %>
        <input type="button" class="button" value=" �Ǹų��� " onClick="jsViewSold(<%= oCBuyBenefit.FItemList(i).Fbuy_benefit_idx %>)">
    </td>
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

<form name="frmAct" methos="post" action="PlusSale_process.asp">
    <input type="hidden" name="mode" value="">
    <input type="hidden" name="benefit_group_no" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
