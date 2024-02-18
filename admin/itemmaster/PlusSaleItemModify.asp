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
dim benefit_group_no, plus_sale_item_idx
dim i, j, k
dim page :page = 1

benefit_group_no = request("benefit_group_no")
plus_sale_item_idx = request("plus_sale_item_idx")
if (benefit_group_no <> "") then
    if Not IsNumeric(benefit_group_no) then
        benefit_group_no = ""
    end if
end if

if (benefit_group_no = "") then
    response.write "잘못된 접근입니다."
    dbget.close : response.end
end if

if (plus_sale_item_idx <> "") then
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


'// ===============================================
'// 혜택상품
'// ===============================================
dim oCBuyBenefitGroupItem
set oCBuyBenefitGroupItem = new CBuyBenefit
oCBuyBenefitGroupItem.FRectPlusSaleItemIdx = CHKIIF(plus_sale_item_idx="", "-1", plus_sale_item_idx)

oCBuyBenefitGroupItem.GetBuyBenefitGroupItemOne

if INSERT_MODE then
    oCBuyBenefitGroupItem.FOneItem.Fsort_no = 1
end if

%>
<script src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>

function jsCheckItem(frm) {
    if (frm.itemid.value == '') {
        alert('상품코드를 입력하세요.');
        return false;
    }

    <% if (INSERT_MODE = True) then %>
    if (frm.itemid.value != frm.valid.value) {
        alert('상품체크하기 버튼을 누르세요.');
        return false;
    }
    <% end if %>

    if (frm.plus_sale_price.value == '') {
        alert('할인가격을 입력하세요.');
        return false;
    }

    if (frm.plus_sale_pct.value == '') {
        alert('할인율를 입력하세요.');
        return false;
    }

    if (frm.plus_sale_buyprice.value == '') {
        alert('할인매입가를 입력하세요.');
        return false;
    }

    if (frm.sale_burden_type.value == '') {
        alert('할인부담구분을 선택하세요.');
        return false;
    }

    if (parseInt(frm.plus_sale_buyprice.value)>=parseInt(frm.plus_sale_price.value)){
        alert('할인매입가 이상!\n가격을 확인해주세요.\n\n※ 만약 역마진으로 상품등록을 해야한다면\nCEO 품의 진행 후 개발운영팀에 별도 등록 요청해주세요.');
        return false;
    }

    //if (frm.limit_yn.value == '') {
    //    alert('한정여부를 선택하세요.');
    //    return false;
    //}

    if (frm.sort_no.value == '') {
        alert('정렬순서를 입력하세요.');
        return false;
    }

    return true;
}

function ModifyItem(frm) {
    if (jsCheckItem(frm) == false) {
        return;
    }
	var ret = confirm('저장 하시겠습니까?');

	if (ret) {
		frm.submit();
	}
}

function jsCheckItemID() {
    var frm = document.frmItem;
    var itemid = frm.itemid.value;
    var makerid = frm.makerid.value;
    var url = 'PlusSaleItemAjax.asp?itemid=' + itemid + '&makerid=' + makerid;

    if (itemid == '') {
        alert('상품코드를 입력하세요.');
        return;
    }

    if (itemid*0 != 0) {
        alert('상품코드는 숫자만 가능합니다.');
        return;
    }

    $.ajax({
		type: "get",
		url: url,
		cache: false,
		success: function(data) {
            // alert(data);
            try {
                var json = $.parseJSON(data);
                if (json.result == 'err') {
                    alert(json.message);
                } else {
                    alert('OK');
                    document.frmItem.valid.value = document.frmItem.itemid.value;
                    document.frmItem.sellcash.value = json.sellcash;
                }
            } catch(e) {
                alert('일시적인 오류가 발생했습니다.');
                document.frmItem.valid.value = '';
            }
		}
		,error: function(err) {
			alert(err.responseText);
            document.frmItem.valid.value = '';
		}
	});
}

function jsCalcPct() {
    var frm = document.frmItem;

    var sellcash = frm.sellcash.value;
    var plus_sale_price = frm.plus_sale_price.value;

    if (sellcash == '' || plus_sale_price == '') { return; }
    if (sellcash*0 != 0 || plus_sale_price*0 != 0) { return; }

    frm.plus_sale_pct.value = parseInt((sellcash - plus_sale_price) / sellcash * 100);
}
$(function() {
$("select[name=limit_yn]").not(":selected").attr("disabled", "disabled");
});
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<!-- 상단바 시작 -->
	<tr height="25" bgcolor="<%= adminColor("gray") %>">
		<td colspan="4">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
	        <font color="red"><strong>혜택상품 <%= CHKIIF(INSERT_MODE, "추가", "수정") %></strong></font>
		</td>
	</tr>
	<!-- 상단바 끝 -->
	<form name="frmItem" method="post" action="PlusSale_process.asp">
    <input type="hidden" name="mode" value="<%= CHKIIF(INSERT_MODE, "insitem", "modiitem") %>">
    <input type="hidden" name="benefit_group_no" value="<%= benefit_group_no %>">
    <tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">IDX</td>
        <td>
            <%= plus_sale_item_idx %>
            <input type="hidden" name="plus_sale_item_idx" value="<%= plus_sale_item_idx %>">
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">브랜드</td>
        <td>
            <%= oCBuyBenefitGroup.FOneItem.Fmakerid %>
            <input type="hidden" name="makerid" value="<%= oCBuyBenefitGroup.FOneItem.Fmakerid %>">
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">상품코드</td>
        <td>
            <input type="text" class="text" name="itemid" value="<%= oCBuyBenefitGroupItem.FOneItem.Fitemid %>" size="8">
            <% if (INSERT_MODE = True) then %>
            <input type="button" class="button" value="체크하기" onClick="jsCheckItemID()">
            * 옵션가 있는 상품은 등록불가합니다.
            <input type="hidden" name="valid" value="">
            <% end if %>
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">소비자가</td>
        <td>
            <input type="text" class="text_ro" name="sellcash" value="<%= oCBuyBenefitGroupItem.FOneItem.Fsellcash %>" size="8" readOnly>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">할인가격</td>
        <td>
            <input type="text" class="text" name="plus_sale_price" value="<%= oCBuyBenefitGroupItem.FOneItem.Fplus_sale_price %>" size="8" onFocusOut="jsCalcPct()">
            &nbsp;
            <input type="text" class="text_ro" name="plus_sale_pct" value="<%= oCBuyBenefitGroupItem.FOneItem.Fplus_sale_pct %>" size="2" readOnly> %
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">할인매입가</td>
        <td>
            <input type="text" class="text" name="plus_sale_buyprice" value="<%= oCBuyBenefitGroupItem.FOneItem.Fplus_sale_buyprice %>" size="8">
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">할인부담구분</td>
        <td>
            <% drawPartnerCommCodeBox True,"PSBurdenType","sale_burden_type", oCBuyBenefitGroupItem.FOneItem.Fsale_burden_type,"" %>
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100"></td>
        <td>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">한정여부</td>
        <td>
            <% drawPartnerCommCodeBox True,"yn","limit_yn", oCBuyBenefitGroupItem.FOneItem.Flimit_yn,"" %> <font color="#808080">(한정 셋팅 개발대기)</ont>
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">한정수량</td>
        <td>
            <input type="text" class="text" name="limit_cnt" value="<%= oCBuyBenefitGroupItem.FOneItem.Flimit_cnt %>" size="8">
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">최대구매수</td>
        <td>
            <input type="text" class="text" name="max_buy_cnt" value="<%= oCBuyBenefitGroupItem.FOneItem.Fmax_buy_cnt %>" size="8">
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">한정수량</td>
        <td>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">뱃지내용</td>
        <td>
            <input type="text" class="text" name="badge_contents" value="<%= oCBuyBenefitGroupItem.FOneItem.Fbadge_contents %>" size="20">
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">유의사항</td>
        <td>
            <input type="text" class="text" name="notice" value="<%= oCBuyBenefitGroupItem.FOneItem.Fnotice %>" size="20">
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">정렬순서</td>
        <td>
            <input type="text" class="text" name="sort_no" value="<%= oCBuyBenefitGroupItem.FOneItem.Fsort_no %>" size="8">
        </td>
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">사용여부</td>
        <td>
            <% drawPartnerCommCodeBox False,"yn","use_yn", oCBuyBenefitGroupItem.FOneItem.Fuse_yn,"" %>
        </td>
	</tr>
    </form>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            &nbsp;
        	<input type="button" class="button" value=" 저장하기 " onclick="ModifyItem(frmItem)">
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
