<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/PlusSaleCls.asp"-->
<%

dim INSERT_MODE : INSERT_MODE = True
dim buy_benefit_idx
dim i, j, k
dim page :page = 1

buy_benefit_idx = request("buy_benefit_idx")
if (buy_benefit_idx <> "") then
    if Not IsNumeric(buy_benefit_idx) then
        buy_benefit_idx = ""
    end if
end if

if (buy_benefit_idx <> "") then
    INSERT_MODE = False
end if


'// ===============================================
'// ��������
'// ===============================================
dim oCBuyBenefit
set oCBuyBenefit = new CBuyBenefit
oCBuyBenefit.FRectBuyBenefitIdx = CHKIIF(buy_benefit_idx="", "-1", buy_benefit_idx)

oCBuyBenefit.GetCBuyBenefitMasterOne


'// ===============================================
'// ���ñ׷�
'// ===============================================
dim oCBuyBenefitGroup
set oCBuyBenefitGroup = new CBuyBenefit
oCBuyBenefitGroup.FRectBuyBenefitIdx = CHKIIF(buy_benefit_idx="", "-1", buy_benefit_idx)
oCBuyBenefitGroup.FRectUseYN = "Y"
oCBuyBenefitGroup.FPageSize = 100
oCBuyBenefitGroup.FCurrPage = page

oCBuyBenefitGroup.GetBuyBenefitGroupList

%>
<script language='javascript'>
function jsCheckMaster(frm) {
    if (frm.benefit_type.value == '') {
        alert('���ñ����� �����ϼ���.');
        return false;
    }

    if (frm.benefit_title.value == '') {
        alert('���������� �����ϼ���.');
        return false;
    }

    if (frm.benefit_start_dt.value == '') {
        alert('���ý����Ͻø� �����ϼ���.');
        return false;
    }

    if (frm.benefit_end_dt.value == '') {
        alert('���������Ͻø� �����ϼ���.');
        return false;
    }

    if (frm.whole_target_yn.value == '') {
        alert('��ü��󿩺θ� �����ϼ���.');
        return false;
    }

    if ((frm.channel_www_yn.value == '') || (frm.channel_mob_yn.value == '') || (frm.channel_app_yn.value == '')) {
        alert('����ä���� �����ϼ���.');
        return false;
    }

    if ((frm.channel_www_yn.value == 'N') && (frm.channel_mob_yn.value == 'N') && (frm.channel_app_yn.value == 'N')) {
        alert('��� �ϳ��� ����ä���� �����ϼ���.');
        return false;
    }

    if (frm.show_rank.value == '') {
        alert('��������� �����ϼ���.');
        return false;
    }

    return true;
}
function ModiMaster(frm) {
    if (jsCheckMaster(frm) == false) {
        return;
    }
	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.submit();
	}
}

function jsPopViewItems(idx) {
    var popwin = window.open("PlusSaleItemList.asp?benefit_group_no=" + idx,"jsPopViewItems" + idx,"width=1200 height=600 scrollbars=yes resizable=yes");
    popwin.focus();
}

function ModiGroup(buy_benefit_idx, idx) {
    var popwin = window.open("PlusSaleGroupModify.asp?buy_benefit_idx=<%= buy_benefit_idx %>&benefit_group_no=" + idx,"ModiGroup" + idx,"width=1200 height=600 scrollbars=yes resizable=yes");
    popwin.focus();
}

function jsPopInfo(buy_benefit_idx) {
    var popwin = window.open("PlusSaleInfo.asp?buy_benefit_idx=" + buy_benefit_idx,"jsPopInfo" + buy_benefit_idx,"width=600 height=800 scrollbars=yes resizable=yes");
    popwin.focus();
}

</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<!-- ��ܹ� ���� -->
	<tr height="25" bgcolor="<%= adminColor("gray") %>">
		<td colspan="4">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
	        <font color="red"><strong>�������� <%= CHKIIF(INSERT_MODE, "�ۼ�", "����") %></strong></font>
		</td>
	</tr>
	<!-- ��ܹ� �� -->
	<form name="frmMaster" method="post" action="PlusSale_process.asp">
    <input type="hidden" name="mode" value="<%= CHKIIF(INSERT_MODE, "insmaster", "modimaster") %>">
    <tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">IDX</td>
        <td colspan="3">
            <%= buy_benefit_idx %>
            <input type="hidden" name="buy_benefit_idx" value="<%= buy_benefit_idx %>">
        </td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">���ñ���</td>
        <td colspan="3">
            <% drawPartnerCommCodeBox True,"PSBenefitType","benefit_type", oCBuyBenefit.FOneItem.Fbenefit_type,"" %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">��������</td>
        <td colspan="3">
            <input type="text" class="text" name="benefit_title" value="<%= oCBuyBenefit.FOneItem.Fbenefit_title %>" size="80">
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">���ú�����</td>
        <td colspan="3">
            <input type="text" class="text" name="benefit_subtitle" value="<%= oCBuyBenefit.FOneItem.Fbenefit_subtitle %>" size="40">
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">���ý����Ͻ�</td>
        <td>
            <input type="text" class="text" name="benefit_start_dt" value="<%= oCBuyBenefit.FOneItem.Fbenefit_start_dt %>" size="12">
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">���������Ͻ�</td>
        <td>
            <input type="text" class="text" name="benefit_end_dt" value="<%= left(oCBuyBenefit.FOneItem.Fbenefit_end_dt,10) %>" size="10">
            <input type="text" class="text" name="benefit_end_dt_time" value="<%= CHKIIF(INSERT_MODE, "23:59:59", right(FormatDate(oCBuyBenefit.FOneItem.Fbenefit_end_dt,"0000-00-00 00:00:00"),8)) %>" size="8">
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">��ü��󿩺�</td>
        <td>
            <% drawPartnerCommCodeBox True,"yn","whole_target_yn", oCBuyBenefit.FOneItem.Fwhole_target_yn,"" %>
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">��뿩��</td>
        <td>
            <% drawPartnerCommCodeBox False,"useyn","use_yn", oCBuyBenefit.FOneItem.Fuse_yn,"" %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">����ä��</td>
        <td colspan="3">
            WWW : <% drawPartnerCommCodeBox True,"yn","channel_www_yn", oCBuyBenefit.FOneItem.Fchannel_www_yn,"" %>
            &nbsp;
            ����� : <% drawPartnerCommCodeBox True,"yn","channel_mob_yn", oCBuyBenefit.FOneItem.Fchannel_mob_yn,"" %>
            &nbsp;
            APP : <% drawPartnerCommCodeBox True,"yn","channel_app_yn", oCBuyBenefit.FOneItem.Fchannel_app_yn,"" %>
        </td>
	</tr>
    <tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">�ȳ�����</td>
        <td colspan="3">
            <% if INSERT_MODE then %>
            * ���� �ۼ� �� �ȳ������� ����� �� �ֽ��ϴ�.
            <% else %>
            <input type="button" class="button" value="�ۼ��ϱ�" onClick="jsPopInfo(<%= buy_benefit_idx %>)">
            <% end if %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">�������</td>
        <td>
            <input type="text" class="text" name="show_rank" value="<%= oCBuyBenefit.FOneItem.Fshow_rank %>" size="8">
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100"></td>
        <td>

        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">����Ͻ�</td>
        <td>
            <%= oCBuyBenefit.FOneItem.Freg_dt %>
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">��ϰ�����ID</td>
        <td>
            <%= oCBuyBenefit.FOneItem.Freg_admin_id %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">���������Ͻ�</td>
        <td>
            <%= oCBuyBenefit.FOneItem.Flast_update_dt %>
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">��������ID</td>
        <td>
            <%= oCBuyBenefit.FOneItem.Flast_update_admin_id %>
        </td>
	</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            &nbsp;
        	<input type="button" class="button" value=" �����ϱ� " onclick="ModiMaster(frmMaster)">
	    </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>

<% if (INSERT_MODE = False) then %>

<p />

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
        <img src="/images/icon_arrow_down.gif" align="absbottom">
        <font color="red"><strong>���ñ׷� ���</strong></font>
        &nbsp;
		�˻���� : <b><%= oCBuyBenefitGroup.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCBuyBenefitGroup.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width=60 rowspan="2">IDX</td>
	<td width=100 rowspan="2">����</td>
    <td rowspan="2">�׷��</td>
    <td width=80 rowspan="2">�귣��ID</td>
    <td width=80 rowspan="2">���Ǳݾ�</td>
    <td width=80 rowspan="2">��۱���</td>
    <td width=80 rowspan="2">ī�װ�</td>
    <td colspan="2">�̺�Ʈ</td>
    <td rowspan="2">���</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
    <td width=80>�ڵ�</td>
    <td width=100>��������</td>
</tr>
<% for i=0 to oCBuyBenefitGroup.FResultcount-1 %>
<tr bgcolor="<%= CHKIIF(oCBuyBenefitGroup.FItemList(i).Fuse_yn="Y", "#FFFFFF", "#EEEEEE") %>" align="center" height="25">
    <td><%= oCBuyBenefitGroup.FItemList(i).Fbenefit_group_no %></td>
    <td><%= oCBuyBenefitGroup.FItemList(i).Fgroup_type_name %></td>
    <td><%= oCBuyBenefitGroup.FItemList(i).Fgroup_name %></td>
    <td><%= oCBuyBenefitGroup.FItemList(i).Fmakerid %></td>
    <td><%= FormatNumber(oCBuyBenefitGroup.FItemList(i).Fcondition_amount, 0) %></td>
    <td><%= oCBuyBenefitGroup.FItemList(i).Fdelivery_type_name %></td>
    <td><%= oCBuyBenefitGroup.FItemList(i).Fcatecode %></td>
    <td><%= oCBuyBenefitGroup.FItemList(i).Fevtcode %></td>
    <td><%= oCBuyBenefitGroup.FItemList(i).Fevt_buy_condition_name %></td>
    <td>
        <input type="button" class="button" value="��ǰ����" onclick="jsPopViewItems(<%= oCBuyBenefitGroup.FItemList(i).Fbenefit_group_no %>)">
        &nbsp;
        <input type="button" class="button" value="�׷����" onclick="ModiGroup(<%= buy_benefit_idx %>, <%= oCBuyBenefitGroup.FItemList(i).Fbenefit_group_no %>)">
    </td>
</tr>
<% next %>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            &nbsp;
        	<input type="button" class="button" value=" �׷��߰� " onclick="ModiGroup(<%= buy_benefit_idx %>, '')">
	    </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>

<% end if %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
