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

	ostockerrorinfo.GetMonthlyErrorInfo
%>

<script language='javascript'>

//저장
function UpdateStockMWDiv(frm) {
	if (frm.yyyymm.value == "") {
		alert("대상년월을 지정하세요");
		return;
	}

	if (frm.lastmwdiv.value == "") {
		alert("매입구분을 지정하세요");
		return;
	}

	if (confirm("업데이트 하시겠습니까?") == true) {
		frm.submit();
	}
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="pop_item_stock_edit_process.asp">
<input type="hidden" name="mode" value="updatelastmwdiv">
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
	<td height="25">매입지정</td>
	<td bgcolor="#FFFFFF">
		<%= ostockerrorinfo.FOneItem.FMaeipCount %>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">위탁지정</td>
	<td bgcolor="#FFFFFF">
		<%= ostockerrorinfo.FOneItem.FWitakCount %>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">미지정</td>
	<td bgcolor="#FFFFFF">
		<%= ostockerrorinfo.FOneItem.FErrorCount %>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">매입구분</td>
	<td bgcolor="#FFFFFF">
		<select class="select" name="lastmwdiv">
			<option value="">-선택-</option>
			<option value="M">매입</option>
			<option value="W">위탁</option>
		</select>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">대상년월</td>
	<td bgcolor="#FFFFFF">
		<select class="select" name="yyyymm">
			<option value="">-선택-</option>
			<option value="<%= yyyy1 %>-<%= mm1 %>"><%= yyyy1 %>-<%= mm1 %></option>
			<option value="all">미지정 전체내역</option>
		</select>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td colspan=2 height="35" align="center">
		<input type="button" class="button" value="재고정보 업데이트" onClick="UpdateStockMWDiv(frm)">
	</td>
</tr>

</form>
</table>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->