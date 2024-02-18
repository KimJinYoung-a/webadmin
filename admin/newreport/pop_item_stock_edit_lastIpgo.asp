<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 오프라인 상품 등록 통합
' Hieditor : 2011.10.18 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->

<%

dim itemgubun, itemid, itemoption
dim yyyy1, mm1
dim i


itemgubun = requestCheckVar(request("itemgubun"),32)
itemid = requestCheckVar(request("itemid"),32)
itemoption = requestCheckVar(request("itemoption"),32)

yyyy1 = requestCheckVar(request("yyyy1"),32)
mm1 = requestCheckVar(request("mm1"),32)


'// ===========================================================================
dim ostockerrorinfo

set ostockerrorinfo = new CMonthlyStock
    ostockerrorinfo.FRectItemGubun = itemgubun
	ostockerrorinfo.FRectItemId = itemid
	ostockerrorinfo.FRectItemOption = itemoption

	ostockerrorinfo.GetMonthlyNullIpgoInfo


	Response.write "폐기메뉴"
	Response.end

%>

<script language='javascript'>

//저장
function UpdateStockMWDiv(frm) {
	if (frm.yyyymm.value.length != 7) {
		alert("최종입고일을 지정하세요");
		return;
	}


	if (confirm("업데이트 하시겠습니까?\n\n최종입고일은 재고월과 같거나 이후인경우만 업데이트 됩니다.(기 입력 내역은 업데이트 안됨)") == true) {
		frm.submit();
	}
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="pop_item_stock_edit_process.asp">
<input type="hidden" name="mode" value="updatelastipgo">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="<%= itemoption %>">

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width="100" height="25">상품코드</td>
	<td bgcolor="#FFFFFF">
		<%= itemgubun %>-<%= itemid %>-<%= itemoption %>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">시작월</td>
	<td bgcolor="#FFFFFF">
		<%= ostockerrorinfo.FOneItem.FMIN_YYYYMM %>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">종료월</td>
	<td bgcolor="#FFFFFF">
		<%= ostockerrorinfo.FOneItem.FMAX_YYYYMM %>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">최초입고</td>
	<td bgcolor="#FFFFFF">
		<%= ostockerrorinfo.FOneItem.FminIpgodate %>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">최종입고</td>
	<td bgcolor="#FFFFFF">
		<%= ostockerrorinfo.FOneItem.FmaxIpgodate %>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">미지정갯수</td>
	<td bgcolor="#FFFFFF">
		<%= ostockerrorinfo.FOneItem.FnullCNT %>
	</td>
</tr>


<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">최종입고일 지정</td>
	<td bgcolor="#FFFFFF">

	<input type="text" name="yyyymm" value="<%=ostockerrorinfo.FOneItem.FMIN_YYYYMM%>" size="8" maxlength="7">
	<!--
		<select class="select" name="yyyymm">
			<option value="">-선택-</option>
			<option value="<%= yyyy1 %>-<%= mm1 %>"><%= yyyy1 %>-<%= mm1 %></option>
		</select>
	-->
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td colspan=2 height="35" align="center">
		<input type="button" class="button" value="재고정보 업데이트" onClick="UpdateStockMWDiv(frm)">
	</td>
</tr>

</form>
</table>
<%
SET ostockerrorinfo = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
