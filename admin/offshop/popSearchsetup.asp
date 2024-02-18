<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/eventAppReport.asp"-->
<%
Dim buyprice, appRunUser, appRunDay
appRunUser		= requestCheckVar(request("appRunUser"),1)
buyprice		= request("buyprice")
appRunDay		= request("appRunDay")
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function fnSearch(){
	opener.$("#appRunUser").val($("#appRunUser").val());
	opener.$("#buyprice").val($("#buyprice").val());
	if (!$.isNumeric($("#buyprice").val())) {
		alert('숫자만 입력하세요');
		$("#buyprice").val("");
		$("#buyprice").focus();
		return false;
	}
	if (!$.isNumeric($("#appRunDay").val())) {
		alert('숫자만 입력하세요');
		$("#appRunDay").val("");
		$("#appRunDay").focus();
		return false;
	}
	opener.$("#appRunDay").val($("#appRunDay").val());
	alert("적용이 완료되었습니다.");
	opener.document.frm.submit();
	self.close();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="POST">
<tr height="30" bgcolor="#FFFFFF" id="BuyTr">
	<td colspan="11">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<strong>구매 전환 금액 설정</strong>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" id="BuyTr2">
	<td width="20%">구매 금액</td>
	<td align="LEFT" bgcolor="#FFFFFF"><input type="text" name="buyprice" id="buyprice" value="<%=buyprice%>"> 원 이상 기준</td>
</tr>
<tr height="30" bgcolor="#FFFFFF">
	<td colspan="11">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<strong>유저 조건 설정</strong>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20%">구매/활성화 유저</td>
	<td align="LEFT" bgcolor="#FFFFFF">
		<Select name="appRunUser" id="appRunUser" class="select">
			<option value="0" <%= Chkiif(appRunUser="0", "selected", "") %> >App 최종 접속일</option>
			<option value="1" <%= Chkiif(appRunUser="1", "selected", "") %> >장바구니 최종 결제일</option>
		</Select>
		<input type="text" size="5" name="appRunDay" id="appRunDay" value="<%= appRunDay %>"> 일 이내 기준
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="2">
		<input type="button" value="적용" class="button" onclick="fnSearch();" />
		<input type="button" value="취소" class="button" onclick="self.close();"/>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->