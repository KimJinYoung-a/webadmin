<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 이니렌탈 판매/취소 등록 관리
' Hieditor : 2021.05.10 원승현 생성
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

	if ($("input[name=inirentalpgkey]").val() == '') {
		alert('PG사KEY 를 입력하세요.');
		return;
	}

	if ($("input[name=inirentalgubun]:checked").val() == undefined) {
		alert('구분을 선택하세요.');
		return;
	}

	if ($("input[name=inirentalgubun]:checked").val() == 'inirentalcancel') {
		if ($("input[name=inirentalconfirmdate]").val() == '') {
			alert('취소일시를 입력하세요.');
			return;
		}

		var fromDate = new Date('<%= Left(DateAdd("m", -1, Now()), 7) + "-01" %>');
		var toDate = new Date('<%= Left(DateAdd("m", 1, Now()), 7) + "-01" %>');
		var cancelDate = new Date($("input[name=inirentalconfirmdate]").val());

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

	if ($("input[name=inirentalipkumdate]:checked").val() == '') {
		alert('입금예정(정산)일을 입력하세요');
		return;
	}    
	if ($("input[name=inirentalappprice]:checked").val() == '') {
		alert('구매금액을 입력하세요');
		return;
	}
	if ($("input[name=inirentalcommprice]:checked").val() == '') {
		alert('수수료를 입력하세요');
		return;
	}    
	if ($("input[name=inirentalcommvatprice]:checked").val() == '') {
		alert('부가세를 입력하세요');
		return;
	}
	if ($("input[name=inirentaljungsanprice]:checked").val() == '') {
		alert('정산예정액을 입력하세요');
		return;
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
		<b>수기데이타 등록(ON)</b>
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frm" method="post" action="pgdata_process.asp" style="margin: 0px;">
<input type="hidden" name="mode" value="addIniRentalManualWrite" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">
    PG사KEY:</td>
	<td align="left">
		<input type="text" class="text" name="inirentalpgkey" size="64">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">구분:</td>
	<td align="left">
		<!--
		<input type="radio" name="gubun" value="cancel"> 카드사취소
		&nbsp;
		<input type="radio" name="gubun" value="dup"> 금액 0원 승인건
		-->
		<input type="radio" name="inirentalgubun" value="inirentalbuy" checked> 렌탈매출등록
		&nbsp;
		<input type="radio" name="inirentalgubun" value="inirentalcancel"> 렌탈취소등록
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">승인(취소)일시:</td>
	<td align="left">
		<input type="text" class="text" name="inirentalconfirmdate" size="32"> * 예: 2019-02-19 17:46:52
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">입금예정(정산)일:</td>
	<td align="left">
		<input type="text" class="text" name="inirentalipkumdate" size="32"> * 예: 2019-02-19 17:46:52
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">구매금액:</td>
	<td align="left">
		<input type="text" class="text" name="inirentalappprice" size="50">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">수수료:</td>
	<td align="left">
		<input type="text" class="text" name="inirentalcommprice" size="50">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">부가세:</td>
	<td align="left">
		<input type="text" class="text" name="inirentalcommvatprice" size="50">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">정산예정액:</td>
	<td align="left">
		<input type="text" class="text" name="inirentaljungsanprice" size="50">
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
