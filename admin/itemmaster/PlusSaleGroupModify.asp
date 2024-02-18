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
dim buy_benefit_idx, benefit_group_no
dim i, j, k
dim page :page = 1

buy_benefit_idx = request("buy_benefit_idx")
benefit_group_no = request("benefit_group_no")
if (benefit_group_no <> "") then
    if Not IsNumeric(benefit_group_no) then
        benefit_group_no = ""
    end if
end if

if (benefit_group_no <> "") then
    INSERT_MODE = False
end if


'// ===============================================
'// 혜택그룹
'// ===============================================
dim oCBuyBenefitGroup
set oCBuyBenefitGroup = new CBuyBenefit
oCBuyBenefitGroup.FRectBenefitGroupNo = CHKIIF(benefit_group_no="", "-1", benefit_group_no)
oCBuyBenefitGroup.FRectUseYN = "Y"
oCBuyBenefitGroup.FPageSize = 100
oCBuyBenefitGroup.FCurrPage = page

oCBuyBenefitGroup.GetCBuyBenefitGroupOne

%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script language='javascript'>

function jsCheckGroup(frm) {
    if ((frm.group_type.value != '1') && (frm.group_type.value != '5')) {
        alert('그룹없음 또는 금액별-브랜드 만 선택 가능합니다.');
        return false;
    }

    if (frm.sort_no.value == '') {
        alert('정렬번호를 선택하세요.');
        return false;
    }

    if (frm.condition_amount.value == '') {
        alert('조건금액을 입력하세요.');
        return false;
    }

    return true;
}

function ModiGroup(frm) {
    if (jsCheckGroup(frm) == false) {
        return;
    }
	var ret = confirm('저장 하시겠습니까?');

	if (ret) {
		frm.submit();
	}
}

// 아직 지원안하는 기능 이용불가 처리
$(function(){
    //inputBox disable (카테고리, 이벤트 제한)
    $("#catecode,#evtcode").removeClass("text").addClass("text_ro").attr("readonly",true);

    //selectBox disable (그룹없음, 브랜드만 허용)
    $("select[name=group_type] option").each(function(){
        if($(this).val()!="1" && $(this).val()!=="5"){
            $(this).attr("disabled",true);
        }
    });
});
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<!-- 상단바 시작 -->
	<tr height="25" bgcolor="<%= adminColor("gray") %>">
		<td colspan="4">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
	        <font color="red"><strong>혜택그룹 <%= CHKIIF(INSERT_MODE, "추가", "수정") %></strong></font>
		</td>
	</tr>
	<!-- 상단바 끝 -->
	<form name="frmGroup" method="post" action="PlusSale_process.asp">
    <input type="hidden" name="mode" value="<%= CHKIIF(INSERT_MODE, "insgroup", "modigroup") %>">
    <input type="hidden" name="buy_benefit_idx" value="<%= buy_benefit_idx %>">
    <tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">IDX</td>
        <td colspan="3">
            <%= benefit_group_no %>
            <input type="hidden" name="benefit_group_no" value="<%= benefit_group_no %>">
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">그룹구분</td>
        <td>
            <% drawPartnerCommCodeBox False,"PSGroupType","group_type", oCBuyBenefitGroup.FOneItem.Fgroup_type,"" %>
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">그룹명</td>
        <td>
            <input type="text" class="text" name="group_name" value="<%= oCBuyBenefitGroup.FOneItem.Fgroup_name %>" size="40">
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">정렬번호</td>
        <td>
            <input type="text" class="text" name="sort_no" value="<%= oCBuyBenefitGroup.FOneItem.Fsort_no %>" size="8">
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">사용여부</td>
        <td>
            <% drawPartnerCommCodeBox False,"useyn","use_yn", oCBuyBenefitGroup.FOneItem.Fuse_yn,"" %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">조건금액</td>
        <td>
            <input type="text" class="text" name="condition_amount" value="<%= oCBuyBenefitGroup.FOneItem.Fcondition_amount %>" size="8">
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">배송구분</td>
        <td>
            <% drawPartnerCommCodeBox False,"PSDeliveryType","delivery_type", oCBuyBenefitGroup.FOneItem.Fdelivery_type,"" %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">카테고리코드</td>
        <td>
            <input type="text" class="text" id="catecode" name="catecode" value="<%= oCBuyBenefitGroup.FOneItem.Fcatecode %>" size="8">
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">브랜드ID</td>
        <td>
            <%	drawSelectBoxDesignerWithName "makerid", oCBuyBenefitGroup.FOneItem.Fmakerid %>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">이벤트코드</td>
        <td>
            <input type="text" class="text" id="evtcode" name="evtcode" value="<%= oCBuyBenefitGroup.FOneItem.Fevtcode %>" size="8">
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">이벤트구매조건</td>
        <td>
            <% drawPartnerCommCodeBox True,"PSBuyCondition","evt_buy_condition", oCBuyBenefitGroup.FOneItem.Fevt_buy_condition,"" %>
        </td>
	</tr>
    </form>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            &nbsp;
        	<input type="button" class="button" value=" 저장하기 " onclick="ModiGroup(frmGroup)">
	    </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
