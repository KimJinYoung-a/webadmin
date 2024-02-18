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

''우선 예치금, 기프트카드만 수정 가능

dim idx, yyyymmdd, srcGbn
dim mode

idx     	= requestcheckvar(request("idx"),32)
yyyymmdd    = requestcheckvar(request("yyyymmdd"),32)
srcGbn      = requestcheckvar(request("srcGbn"),32)

if (srcGbn = "G") then
    mode = "modiGiftDate"
else
    mode = "modiDepositDate"
end if

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script language="javascript">

function jsSubmit(){
	var frm = document.frm;

    if (confirm('저장하시겠습니까?') == true) {
        frm.submit();
    }
}

</script>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>일자변경</b>
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frm" method="post" action="pointsum_process.asp" style="margin: 0px;">
<input type="hidden" name="mode" value="<%= mode %>" />
<input type="hidden" name="idx" value="<%= idx %>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">날짜:</td>
	<td align="left">
		<input type="text" class="text" name="yyyymmdd" value="<%= yyyymmdd %>" size="10">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan="2" height="35">
	    <input type="button" class="button" value="저장" onClick="jsSubmit();">
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
