<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 메일진
' History : 2018.04.27 이상구 생성(메일러 연동 생성 메일러로 발송 내역 전송. 메일 가져오기 생성.)
'			2019.06.24 정태훈 수정(템플릿 기능 신규 추가)
'			2020.05.28 한용민 수정(TMS 메일러 추가)
'###########################################################
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mailzinecls.asp"-->
<%

'// ============================================================================
dim classDate, mode

classDate = requestCheckVar(request("dt"),32)

if (classDate <> "") then
	mode = "modicls"
else
	mode = "inscls"
end if


'// ============================================================================
dim oClass
set oClass = new CMailzineList
oClass.FRectDate = classDate
oClass.frectmailergubun = "EMS"
oClass.MailzineClassOne()

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script language="JavaScript">
function jsSubmit(frm) {
	if (frm.classDate.value == '') {
		alert('발송일자를 입력하세요.');
		return;
	}

	var day = new Date(frm.classDate.value);
	var today = new Date();

	/*
	if (frm.classDate.value <= today.yyyymmdd()) {
		alert('발송일자는 내일자부터 입력가능합니다.\n(오늘 또는 어제날짜 입력불가!)');
		return;
	}
	*/

	if (isInt(frm.itemid1.value, false) != true) {
		alert('클래스 01 의 상품코드를 입력하세요.');
		return;
	}

	if (isInt(frm.salePer1.value, false) != true) {
		alert('클래스 01 의 할인율을 입력하세요.');
		return;
	}

	if (frm.classDesc1.value == '') {
		alert('클래스 01 의 강좌설명을 입력하세요.');
		return;
	}

	if (frm.classSubDesc1.value == '') {
		alert('클래스 01 의 강좌서브설명을 입력하세요.');
		return;
	}

	if (day.getDay() == 5) {
		// 금요일은 클래스가 3개이다.
		if (isInt(frm.itemid2.value, false) != true) {
			alert('클래스 02 의 상품코드를 입력하세요.');
			return;
		}

		if (isInt(frm.salePer2.value, false) != true) {
			alert('클래스 02 의 할인율을 입력하세요.');
			return;
		}

		if (frm.classDesc2.value == '') {
			alert('클래스 02 의 강좌설명을 입력하세요.');
			return;
		}

		if (frm.classSubDesc2.value == '') {
			alert('클래스 02 의 강좌서브설명을 입력하세요.');
			return;
		}

		if (isInt(frm.itemid3.value, false) != true) {
			alert('클래스 03 의 상품코드를 입력하세요.');
			return;
		}

		if (isInt(frm.salePer3.value, false) != true) {
			alert('클래스 03 의 할인율을 입력하세요.');
			return;
		}

		if (frm.classDesc3.value == '') {
			alert('클래스 03 의 강좌설명을 입력하세요.');
			return;
		}

		if (frm.classSubDesc3.value == '') {
			alert('클래스 03 의 강좌서브설명을 입력하세요.');
			return;
		}
	} else {
		if ((frm.itemid2.value != '') || (frm.itemid3.value != '')) {
			if (confirm('금요일만 3개의 클래스 등록이 가능합니다.\n\n입력된 클래스 02/03 의 정보는 저장되지 않습니다.\n\n진행하시겠습니까?') != true) {
				return;
			}
		}
	}

	if (confirm('저장하시겠습니까?') == true) {
		frm.submit();
	}
}

Date.prototype.yyyymmdd = function() {
  var mm = this.getMonth() + 1; // getMonth() is zero-based
  var dd = this.getDate();

  return [this.getFullYear(),
          (mm>9 ? '' : '0') + mm,
          (dd>9 ? '' : '0') + dd
         ].join('-');
};

function isInt(value, allowBlank) {
	if ((value == '') && (allowBlank == true)) { return true; }
	return !isNaN(value) && parseInt(value, 10) == value;
}

function jsSetDisabledObj(obj, disabled) {
	obj.disabled = disabled;
	if (obj.type != 'textarea') {
		obj.style.background = disabled ? '#DDDDDD' : '#FFFFFF';
	}
}

function jsSetItemState() {
	var frm = document.frm;
	if (frm.classDate.value == '') { return; }
	var day = new Date(frm.classDate.value);

	if (day.getDay() == 5) {
		jsSetDisabledObj(frm.itemid2, false);
		jsSetDisabledObj(frm.salePer2, false);
		jsSetDisabledObj(frm.classDesc2, false);
		jsSetDisabledObj(frm.classSubDesc2, false);
		jsSetDisabledObj(frm.itemid3, false);
		jsSetDisabledObj(frm.salePer3, false);
		jsSetDisabledObj(frm.classDesc3, false);
		jsSetDisabledObj(frm.classSubDesc3, false);
	} else {
		jsSetDisabledObj(frm.itemid2, true);
		jsSetDisabledObj(frm.salePer2, true);
		jsSetDisabledObj(frm.classDesc2, true);
		jsSetDisabledObj(frm.classSubDesc2, true);
		jsSetDisabledObj(frm.itemid3, true);
		jsSetDisabledObj(frm.salePer3, true);
		jsSetDisabledObj(frm.classDesc3, true);
		jsSetDisabledObj(frm.classSubDesc3, true);
	}
}

$(document).ready(function(){
	jsSetItemState();
});

</script>
<table width="95%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
	<form name="frm" method="post" action="mailzine_process.asp">
	<input type="hidden" name="mode" value="<%= mode %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" width="150">발송일자</td>
		<td colspan="2">
			<input id="classDate" name="classDate" value="<%= oClass.FOneItem.FclassDate %>" class="text_ro" size="10" maxlength="10" readonly />
			<% if (mode = "inscls") then %>
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="classDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script language="javascript">
			var classDate = new Calendar({
				inputField : "classDate", trigger    : "classDate_trigger",
				onSelect: function() {
					this.hide();
					jsSetItemState();
				}, bottomBar: true, dateFormat: "%Y-%m-%d", fdow: 0
			});
			</script>
			<% end if %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" width="150" rowspan="4">클래스 01</td>
		<td align="center" width="100">상품코드</td>
		<td align="left"><input type="text" class="text" name="itemid1" value="<%= oClass.FOneItem.Fitemid1 %>" size="10"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">할인율</td>
		<td align="left"><input type="text" class="text" name="salePer1" value="<%= oClass.FOneItem.FsalePer1 %>" size="10"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">강좌설명</td>
		<td align="left"><input type="text" class="text" name="classDesc1" value="<%= oClass.FOneItem.FclassDesc1 %>" size="50"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">강좌서브설명</td>
		<td align="left"><input type="text" class="text" name="classSubDesc1" value="<%= oClass.FOneItem.FclassSubDesc1 %>" size="50"></td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" width="150" rowspan="4">클래스 02</td>
		<td align="center" width="100">상품코드</td>
		<td align="left"><input type="text" class="text" name="itemid2" value="<%= oClass.FOneItem.Fitemid2 %>" size="10"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">할인율</td>
		<td align="left"><input type="text" class="text" name="salePer2" value="<%= oClass.FOneItem.FsalePer2 %>" size="10"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">강좌설명</td>
		<td align="left"><input type="text" class="text" name="classDesc2" value="<%= oClass.FOneItem.FclassDesc2 %>" size="50"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">강좌서브설명</td>
		<td align="left"><input type="text" class="text" name="classSubDesc2" value="<%= oClass.FOneItem.FclassSubDesc2 %>" size="50"></td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" width="150" rowspan="4">클래스 03</td>
		<td align="center" width="100">상품코드</td>
		<td align="left"><input type="text" class="text" name="itemid3" value="<%= oClass.FOneItem.Fitemid3 %>" size="10"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">할인율</td>
		<td align="left"><input type="text" class="text" name="salePer3" value="<%= oClass.FOneItem.FsalePer3 %>" size="10"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">강좌설명</td>
		<td align="left"><input type="text" class="text" name="classDesc3" value="<%= oClass.FOneItem.FclassDesc3 %>" size="50"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">강좌서브설명</td>
		<td align="left"><input type="text" class="text" name="classSubDesc3" value="<%= oClass.FOneItem.FclassSubDesc3 %>" size="50"></td>
	</tr>
	</form>

	<tr bgcolor="#FFFFFF" height="50">
		<td align="center" colspan="3">
			<input type="button" class="button" value=" 저 장 하 기 " onClick="jsSubmit(document.frm)">
		</td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
