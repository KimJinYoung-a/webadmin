<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터 회원탈퇴
' History : 2019.01.08 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim userid, userseq, i
	userid = requestCheckvar(request("userid"),32)
	userseq = requestCheckvar(request("userseq"),32)

%>
<script type="text/javascript">

function DelonUser() {
    if (frm.complaintext.value==""){
        alert("탈퇴사유를 입력해 주세요.");
        return;
    }

	if (confirm('온라인 고객을 탈퇴처리 합니다.\n탈퇴후에는 개발팀에서도 복구가 절대 불가능 합니다.\n진행하시겠습니까?') == true) {
		frm.mode.value = "delonuser";
		frm.action = "/cscenter/member/domodifyuserinfo.asp";
		frm.submit();
	}
}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			※ 고객 탈퇴 처리
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<form name="frm" method="post" action="" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="userid" value="<%= userid %>">
<input type="hidden" name="userseq" value="<%= userseq %>">
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
    <td height="30" width="120" bgcolor="#DDDDFF">탈퇴사유 :</td>
    <td bgcolor="#FFFFFF" >
        <textarea cols="100" rows="5" name="complaintext"></textarea>
    </td>
</tr>
<tr>
	<td align="center" colspan=2 bgcolor="#FFFFFF">
		<input type="button" class="button" value="탈퇴처리" onClick="DelonUser();">
		<input type="button" class="button" value=" 창닫기 " onClick="self.close()">
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
