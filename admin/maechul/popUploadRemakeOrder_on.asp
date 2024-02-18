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
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script language="javascript">

function jsSubmit(){
	var frm = document.frm;

    if (confirm('등록하시겠습니까?') == true) {
        frm.submit();
    }
}

</script>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>매출로그 재작성큐 등록(ON)</b>
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frm" method="post" action="popUploadRemakeOrder_process.asp" style="margin: 0px;">
<input type="hidden" name="mode" value="reMakeOrdrQueOn" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">주문번호 :</td>
	<td align="left">
		<textarea class="textarea" name="orderserial" cols="32" rows="16"></textarea>
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
