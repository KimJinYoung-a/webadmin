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
'// 구매혜택
'// ===============================================
dim oCBuyBenefit
set oCBuyBenefit = new CBuyBenefit
oCBuyBenefit.FRectBuyBenefitIdx = CHKIIF(buy_benefit_idx="", "-1", buy_benefit_idx)

oCBuyBenefit.GetCBuyBenefitMasterOne


'// ===============================================
'// 혜택그룹
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
        alert('혜택구분을 선택하세요.');
        return false;
    }

    if (frm.benefit_title.value == '') {
        alert('혜택제목을 선택하세요.');
        return false;
    }

    if (frm.benefit_start_dt.value == '') {
        alert('혜택시작일시를 선택하세요.');
        return false;
    }

    if (frm.benefit_end_dt.value == '') {
        alert('혜택종료일시를 선택하세요.');
        return false;
    }

    if (frm.whole_target_yn.value == '') {
        alert('전체대상여부를 선택하세요.');
        return false;
    }

    if ((frm.channel_www_yn.value == '') || (frm.channel_mob_yn.value == '') || (frm.channel_app_yn.value == '')) {
        alert('적용채널을 선택하세요.');
        return false;
    }

    if ((frm.channel_www_yn.value == 'N') && (frm.channel_mob_yn.value == 'N') && (frm.channel_app_yn.value == 'N')) {
        alert('적어도 하나의 적용채널을 선택하세요.');
        return false;
    }

    if (frm.show_rank.value == '') {
        alert('노출순위를 선택하세요.');
        return false;
    }

    return true;
}
function ModiMaster(frm) {
    if (jsCheckMaster(frm) == false) {
        return;
    }
	var ret = confirm('저장 하시겠습니까?');

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
	<!-- 상단바 시작 -->
	<tr height="25" bgcolor="<%= adminColor("gray") %>">
		<td colspan="4">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
	        <font color="red"><strong>구매혜택 <%= CHKIIF(INSERT_MODE, "작성", "수정") %></strong></font>
		</td>
	</tr>
	<!-- 상단바 끝 -->
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
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">혜택구분</td>
        <td colspan="3">
            <% drawPartnerCommCodeBox True,"PSBenefitType","benefit_type", oCBuyBenefit.FOneItem.Fbenefit_type,"" %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">혜택제목</td>
        <td colspan="3">
            <input type="text" class="text" name="benefit_title" value="<%= oCBuyBenefit.FOneItem.Fbenefit_title %>" size="80">
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">혜택부제목</td>
        <td colspan="3">
            <input type="text" class="text" name="benefit_subtitle" value="<%= oCBuyBenefit.FOneItem.Fbenefit_subtitle %>" size="40">
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">혜택시작일시</td>
        <td>
            <input type="text" class="text" name="benefit_start_dt" value="<%= oCBuyBenefit.FOneItem.Fbenefit_start_dt %>" size="12">
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">혜택종료일시</td>
        <td>
            <input type="text" class="text" name="benefit_end_dt" value="<%= left(oCBuyBenefit.FOneItem.Fbenefit_end_dt,10) %>" size="10">
            <input type="text" class="text" name="benefit_end_dt_time" value="<%= CHKIIF(INSERT_MODE, "23:59:59", right(FormatDate(oCBuyBenefit.FOneItem.Fbenefit_end_dt,"0000-00-00 00:00:00"),8)) %>" size="8">
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">전체대상여부</td>
        <td>
            <% drawPartnerCommCodeBox True,"yn","whole_target_yn", oCBuyBenefit.FOneItem.Fwhole_target_yn,"" %>
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">사용여부</td>
        <td>
            <% drawPartnerCommCodeBox False,"useyn","use_yn", oCBuyBenefit.FOneItem.Fuse_yn,"" %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">적용채널</td>
        <td colspan="3">
            WWW : <% drawPartnerCommCodeBox True,"yn","channel_www_yn", oCBuyBenefit.FOneItem.Fchannel_www_yn,"" %>
            &nbsp;
            모바일 : <% drawPartnerCommCodeBox True,"yn","channel_mob_yn", oCBuyBenefit.FOneItem.Fchannel_mob_yn,"" %>
            &nbsp;
            APP : <% drawPartnerCommCodeBox True,"yn","channel_app_yn", oCBuyBenefit.FOneItem.Fchannel_app_yn,"" %>
        </td>
	</tr>
    <tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">안내내용</td>
        <td colspan="3">
            <% if INSERT_MODE then %>
            * 내역 작성 후 안내문구를 등록할 수 있습니다.
            <% else %>
            <input type="button" class="button" value="작성하기" onClick="jsPopInfo(<%= buy_benefit_idx %>)">
            <% end if %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">노출순위</td>
        <td>
            <input type="text" class="text" name="show_rank" value="<%= oCBuyBenefit.FOneItem.Fshow_rank %>" size="8">
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100"></td>
        <td>

        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">등록일시</td>
        <td>
            <%= oCBuyBenefit.FOneItem.Freg_dt %>
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">등록관리자ID</td>
        <td>
            <%= oCBuyBenefit.FOneItem.Freg_admin_id %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">최종수정일시</td>
        <td>
            <%= oCBuyBenefit.FOneItem.Flast_update_dt %>
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">최종수정ID</td>
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
        	<input type="button" class="button" value=" 저장하기 " onclick="ModiMaster(frmMaster)">
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

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
        <img src="/images/icon_arrow_down.gif" align="absbottom">
        <font color="red"><strong>혜택그룹 목록</strong></font>
        &nbsp;
		검색결과 : <b><%= oCBuyBenefitGroup.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= oCBuyBenefitGroup.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width=60 rowspan="2">IDX</td>
	<td width=100 rowspan="2">구분</td>
    <td rowspan="2">그룹명</td>
    <td width=80 rowspan="2">브랜드ID</td>
    <td width=80 rowspan="2">조건금액</td>
    <td width=80 rowspan="2">배송구분</td>
    <td width=80 rowspan="2">카테고리</td>
    <td colspan="2">이벤트</td>
    <td rowspan="2">비고</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
    <td width=80>코드</td>
    <td width=100>구매조건</td>
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
        <input type="button" class="button" value="상품보기" onclick="jsPopViewItems(<%= oCBuyBenefitGroup.FItemList(i).Fbenefit_group_no %>)">
        &nbsp;
        <input type="button" class="button" value="그룹수정" onclick="ModiGroup(<%= buy_benefit_idx %>, <%= oCBuyBenefitGroup.FItemList(i).Fbenefit_group_no %>)">
    </td>
</tr>
<% next %>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            &nbsp;
        	<input type="button" class="button" value=" 그룹추가 " onclick="ModiGroup(<%= buy_benefit_idx %>, '')">
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
