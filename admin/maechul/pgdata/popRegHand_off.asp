<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 판매 등록 관리
' Hieditor : 2011.04.22 이상구 생성
'			 2012.08.24 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script language="javascript">

function jsSubmit(){
	var frm = document.frm;

	if ($("input[name=orgpgkey]").val() == '') {
		alert('PG사KEY 를 입력하세요.');
		return;
	}

	if ($("input[name=gubun]:checked").val() == undefined) {
		alert('구분을 선택하세요.');
		return;
	}

	if ($("input[name=gubun]:checked").val() == 'cancel') {
		if ($("input[name=canceldate]").val() == '') {
			alert('취소일시를 입력하세요.');
			return;
		}

		var fromDate = new Date('<%= Left(DateAdd("m", -1, Now()), 7) + "-01" %>');
		var toDate = new Date('<%= Left(DateAdd("m", 1, Now()), 7) + "-01" %>');
		var cancelDate = new Date($("input[name=canceldate]").val());

		if (isNaN(cancelDate)) {
			alert('잘못된 취소일자입니다.');
			return;
		} else if ((cancelDate < fromDate) || (cancelDate >= toDate)) {
			alert('잘못된 취소일자입니다.(' + formatDate(cancelDate) + ')');
			return;
		}

		/*
		if ($("input[name=ipkumdate]").val() == '') {
			alert('입금예정일을 입력하세요.');
			return;
		}
		*/
	}

	frm.submit();
}

function formatDate(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;

    return [year, month, day].join('-');
}

</script>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>수기데이타 등록(OFF)</b>
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frm" method="post" action="pgdata_process.asp" style="margin: 0px;">
<input type="hidden" name="mode" value="addhand" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">원PG사KEY:</td>
	<td align="left">
		<input type="text" class="text" name="orgpgkey" size="32">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">구분:</td>
	<td align="left">
		<input type="radio" name="gubun" value="cancel"> 카드사취소
		&nbsp;
		<input type="radio" name="gubun" value="dup"> 금액 0원 승인건
		&nbsp;
		<input type="radio" name="gubun" value="del"> 수기등록 삭제
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">승인(취소)일시:</td>
	<td align="left">
		<input type="text" class="text" name="canceldate" size="32"> * 예: 2019-02-19 17:46:52
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">입금예정일:</td>
	<td align="left">
		<input type="text" class="text" name="ipkumdate" size="10">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">주문번호:</td>
	<td align="left">
		<input type="text" class="text" name="orderserial" size="10">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan="2" height="35">
	    <input type="button" class="button" value="등록" onClick="jsSubmit();">
	    <input type="button" class="button" value="취소" onClick="self.close();">
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
