<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 수기 매출
' History : 2012.12.11 이상구 생성
'			2013.04.23 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offmanualmeachulcls.asp"-->

<%

dim ErrStr

%>

<script language='javascript'>

String.prototype.trim = function() {
    return this.replace(/(^\s*)|(\s*$)/gi, "");
}

function checkClick() {
	var frm = document.frm;
	var orgdata = frm.orgdata.value;
	var oneline, onelineItems, blankLineCount;

	orgdata = orgdata.split("\n");
	blankLineCount = 0;

	for (var i = 0; i < orgdata.length; i++) {
		oneline = orgdata[i];

		if (oneline.trim() == "") {
			blankLineCount = blankLineCount + 1;
			continue;
		}

		onelineItems = oneline.split("\t");

		if (onelineItems.length != 14) {
			alert("각각의 라인은 14개의 컬럼이 되어야 합니다.\n\n" + oneline);
			return false;
		}

		// 검증
		if (onelineItems[1] != "90") {
			alert("90 상품만 등록가능합니다.");
			return false;
		}

		if (onelineItems[1] != "90") {
			alert("90 상품만 등록가능합니다. [" + onelineItems[1] + "]");
			return false;
		}

		if ((onelineItems[4] == "") || (onelineItems[5] == "") || (onelineItems[6] == "")) {
			alert("카테고리가 누락되었습니다.");
			return false;
		}

		if (onelineItems[7].replace(",", "")*0 != 0) {
			alert("금액이 숫자가 아닙니다. [" + onelineItems[7] + "]");
			return false;
		}

		if ((onelineItems[12] != "과세") && (onelineItems[12] != "면세")) {
			alert("잘못된 과세구분입니다.");
			return false;
		}
	}

	if ((orgdata.length - blankLineCount) > 100) {
		alert("한번에 100개의 상품까지 등록가능합니다.");
		return false;
	}

	alert("OK : 상품수 " + (orgdata.length - blankLineCount));

	return true;
}

function uploadClick() {
	var frm = document.frm;

	if (checkClick() != true) {
		return;
	}

	if (confirm('자료를 업로드합니다.\n\n진행하시겠습니까?')) {
		frm.mode.value="uploaddata";
		frm.submit();
	}
}

function saveClick() {
	var frm = document.frm;
	var checkeditemexist = false;
	var dataarr = "";

	for (var i = 0; ; i++) {
		var v = document.getElementById("chk_" + i);
		if (v == undefined) {
			break;
		}

		if (v.checked == true) {
			checkeditemexist = true;
			break;
		}
	}

	if (checkeditemexist == false) {
		alert("저장할 매출이 없습니다.");
		return;
	}

	if (confirm("매출등록 하시겠습니까?") == true) {
		dataarr = "-1";
		for (var i = 0; ; i++) {
			var v = document.getElementById("chk_" + i);
			if (v == undefined) {
				break;
			}

			if (v.checked == true) {
				dataarr = dataarr + "," + v.value
			}
		}

		frm.orgdata.value = dataarr;

		frm.mode.value="regtemporder";
		frm.submit();
	}
}

function delClick() {
	var frm = document.frm;
	var checkeditemexist = false;
	var dataarr = "";

	for (var i = 0; ; i++) {
		var v = document.getElementById("chk_" + i);
		if (v == undefined) {
			break;
		}

		if (v.checked == true) {
			checkeditemexist = true;
			break;
		}
	}

	if (checkeditemexist == false) {
		alert("삭제할 대상이 없습니다.");
		return;
	}

	if (confirm("삭제 하시겠습니까?") == true) {
		dataarr = "-1";
		for (var i = 0; ; i++) {
			var v = document.getElementById("chk_" + i);
			if (v == undefined) {
				break;
			}

			if (v.checked == true) {
				dataarr = dataarr + "," + v.value
			}
		}

		frm.orgdata.value = dataarr;

		frm.mode.value="deltemporder";
		frm.submit();
	}
}

function CheckAll(chk) {
	for (var i = 0; ; i++) {
		var v = document.getElementById("chk_" + i);
		if (v == undefined) {
			return;
		}

		if (v.disabled != true) {
			v.checked = chk.checked;
		}
	}
}

function clearData() {
	var frm = document.frm;
	frm.orgdata.value = "";
}

</script>

<table border=0 cellspacing=0 cellpadding=0 class="a">
<form name="frm" method="post" action="<%=uploadImgUrl%>/linkweb/offshop/item/item_add_multi_off.asp" onSubmit="return false;">
<input type="hidden" name="mode" value="">
<tr>
	<td>
		<font color="red">탭으로 분리</font><br>
		브랜드ID, 상품구분, 상품명, 옵션명, 카테고리(대), 카테고리(중), 카테고리(소), 판매가, 매입가, 매장공급가, 사용유무, 센터매입구분, 과세구분, 범용바코드<br>
		TOMS001&lt;TAB>90&lt;TAB>NAVY CANVAS CLASSICS TINY&lt;TAB>120/5&lt;TAB>070&lt;TAB>030&lt;TAB>060&lt;TAB>25,000&lt;TAB>0&lt;TAB>0&lt;TAB>사용함&lt;TAB>특정&lt;TAB>과세&lt;TAB>TSAA90000000296<br>
		<font color="red">옵션을 제외한 모든값에는 공란이 있으면 등록이 안됩니다.</font>

		<% if (ErrStr <> "") then %>
			<br><br><b><font color="red"><%= Replace(ErrStr, "\n", "<br>") %></font></b><br><br>
		<% end if %>
	</td>
	<td align="right" valign="bottom">
	</td>
</tr>
<tr>
	<td colspan=2>
	<textarea name="orgdata" cols=120 rows=5></textarea>
	</td>
</tr>
<tr>
	<td>
	<input type= button class="button" value="Clear" onClick="clearData();">
	</td>
	<td>
		<input type= button class="button" value=" 체 크 " onclick="checkClick()">
		<input type= button class="button" value="업로드" onclick="uploadClick()">
	</td>
</tr>
</form>
</table>

<p>

</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
