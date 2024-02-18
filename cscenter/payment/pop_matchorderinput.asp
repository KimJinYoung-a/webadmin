<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim idx, bankdate
idx = request("idx")
bankdate = request("bankdate")

%>

<script language="javascript">

function jsSubmitMatch(frm) {
	if (frm.finishstr.value == "") {
		alert("내용을 입력하세요.");
		frm.finishstr.focus();
		return;
	}

	if (getByteLength(frm.finishstr.value) > 32) {
		alert("내용이 너무 길어서 입력할 수 없습니다.\n(한글기준 16자까지 가능)");
		frm.finishstr.focus();
		return;
	}

	if (frm.ipkumCause.value == "") {
		alert("입금사유를 선택하세요.");
		frm.ipkumCause.focus();
		return;
	}

	if ((frm.ipkumCause.value == "직접입력") && (frm.ipkumCauseText.value == "")) {
		alert("입금사유를 입력하세요.");
		frm.ipkumCauseText.focus();
		return;
	}

	if (frm.ipkumCause.value == "직접입력") {
		if (getByteLength(frm.ipkumCauseText.value) > 32) {
			alert("입금사유가 너무 길어서 입력할 수 없습니다.\n(한글기준 16자까지 가능)");
			frm.ipkumCauseText.focus();
			return;
		}
	}

	if (confirm("매칭하시겠습니까?") == true) {
		frm.submit();
	}
}

function Change_ipkumCause(comp) {
    if (comp.value=="직접입력") {
        document.all.span_ipkumCauseText.style.display = "inline";
    }else{
        document.all.span_ipkumCauseText.style.display = "none";
    }
}

function getByteLength(str) {
	var ret;

	ret = 0;
	for (var i = 0; i <= str.length - 1; i++) {
		var ch = str.charAt(i);
		if (escape(ch).length > 4) {
			ret = ret + 2;
		} else {
			ret = ret + 1;
		}
	}

    return ret;
}

</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="pop_matchorderlist_Process.asp">
	<input type="hidden" name="mode" value="matchByHand">
	<input type="hidden" name="ipkumidx" value="<%= idx %>">
	<input type="hidden" name="bankdate" value="<%= bankdate %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td colspan="2">수기매칭 내용입력(주문번호 등)</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" height="30">
		<td width="80">내용</td>
    	<td align="left">
			<input type="text" class="text" name="finishstr" size="25" value="">
		</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" height="30">
		<td width="80">입금사유</td>
    	<td align="left">
			<select class="select" name="ipkumCause" onChange="Change_ipkumCause(this);">
				<option value=""></option>
				<option value="추가 배송비">추가 배송비</option>
				<option value="추가 상품대금">추가 상품대금</option>
				<option value="주문결제">주문결제</option>
				<option value="은행 이자">은행 이자</option>
				<option value="직접입력">직접입력</option>
			</select>
			<span name="span_ipkumCauseText" id="span_ipkumCauseText" style='display:none'>
			<input type="text" class="text" name="ipkumCauseText" size="15" value="">
		</td>
    </tr>
	</form>
</table>

<br>

<div align="center">
<input type="button" class="button" value="매칭하기" onClick="jsSubmitMatch(frm)">
</div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
