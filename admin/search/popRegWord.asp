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
<%

%>
<script language="javascript">
function jsSubmit() {
	var frm = document.frm;

	if (frm.aliasWord.value == '') {
		alert('동의어를 입력하세요.');
		return;
	}

	if (frm.mainWord.value == '') {
		alert('메인키워드를 입력하세요.');
		return;
	}

	if (confirm('저장하시겠습니까?') == true) {
		frm.mode.value = 'ins';
		frm.submit();
	}
}
</script>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>상품명 키워드 등록</b>
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="popRegWord_process.asp">
<input type="hidden" name="mode" value="">
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">동의어</td>
	<td align="left">
		<input type="text" class="text" name="aliasWord" value="" size="20">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">메인키워드</td>
	<td align="left">
		<input type="text" class="text" name="mainWord" value="" size="20">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan="2" height="35">
	    <input type="button" class="button" value="등록" onClick="jsSubmit();">
	    <input type="button" class="button" value="취소" onClick="self.close();">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
