<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/PlusSaleCls.asp"-->
<%

dim buy_benefit_idx

buy_benefit_idx = request("buy_benefit_idx")
if (buy_benefit_idx <> "") then
    if Not IsNumeric(buy_benefit_idx) then
        buy_benefit_idx = ""
    end if
end if

if (buy_benefit_idx = "") then
    response.write "잘못된 접근입니다."
    dbget.close : response.end
end if

dim oCBuyBenefit
set oCBuyBenefit = new CBuyBenefit
oCBuyBenefit.FRectBuyBenefitIdx = CHKIIF(buy_benefit_idx="", "-1", buy_benefit_idx)

oCBuyBenefit.GetCBuyBenefitMasterOne

%>
<script language='javascript'>

function ModiInfo() {
    var frm = document.frmMaster;
    if (confirm('저장하시겠습니까?')) {
        frm.submit();
    }
}

</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<!-- 상단바 시작 -->
	<tr height="25" bgcolor="<%= adminColor("gray") %>">
		<td colspan="4">
			<img src="/images/icon_arrow_down.gif" align="absbottom">
	        <font color="red"><strong>안내문구 수정</strong></font>
		</td>
	</tr>
	<!-- 상단바 끝 -->
	<form name="frmMaster" method="post" action="PlusSale_process.asp">
    <input type="hidden" name="mode" value="modiinfo">
    <input type="hidden" name="buy_benefit_idx" value="<%= buy_benefit_idx %>">
    <tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">모바일</td>
        <td colspan="3">
            <textarea cols="60" rows="20" name="info_contents_mobile"><%= oCBuyBenefit.FOneItem.Finfo_contents_mobile %></textarea>
        </td>
	</tr>
    <tr bgcolor="#FFFFFF" height="25">
		<td align="center" bgcolor="<%= adminColor("tabletop") %>" width="100">WWW</td>
        <td colspan="3">
            <textarea cols="60" rows="20" name="info_contents_www"><%= oCBuyBenefit.FOneItem.Finfo_contents_www %></textarea>
        </td>
	</tr>
    </form>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            &nbsp;
        	<input type="button" class="button" value=" 저장하기 " onclick="ModiInfo(frmMaster)">
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
