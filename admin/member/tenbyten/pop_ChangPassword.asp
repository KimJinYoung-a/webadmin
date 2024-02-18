<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  사원권한등록
' History : 2011.01.19 정윤정 생성
'			2017.09.25 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPwithLog.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim userid
	userid = requestCheckVar(request("userid"),32)

If not(C_ADMIN_AUTH or C_PSMngPart) Then
	response.write "<script  type='text/javascript'>"
	response.write "	alert('권한이 없습니다.');"
	response.write "</script>"
	dbget.close() : response.end
end if
%>

<script type="text/javascript">
	function jsSubmit(){
		if (jsChkBlank(document.frmPW.sPW.value)){
			alert("변경할 비밀번호를 입력하세요.");
			document.frmPW.sPW.focus();
			return;
		}

		if (document.frmPW.sPW.value.replace(/\s/g, "").length < 6 || document.frmPW.sPW.value.replace(/\s/g, "").length > 16){
			alert("비밀번호는 공백없이 6~16자입니다.");
			document.frmPW.sPW.focus();
			return ;
		}

		if ((document.frmPW.sPW.value)!=(document.frmPW.sPW1.value)){
			alert("비밀번호가 일치하지 않습니다.");
			document.frmPW.sPW1.focus();
			return;
		}

		if (!fnChkComplexPassword(frmPW.sPW.value)) {
			alert('새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)');
			frmPW.sPW.focus();
			return;
		}

		if(confirm("비밀변호를 변경하시겠습니까?")){
			document.frmPW.submit();
		}

	}

	//로드시 포커스
	window.onload = function(){
		document.frmPW.sPW.focus();
	}
</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<br>※ 패스워드 변경시 초기화 되는 기능
		<br>1. 계정이 사용안함 상태 인경우, 사용함으로 변경됨
		<br>2. 장기간 미사용으로 인해 계정이 잠긴경우, 잠김이 해제됨.
		<br>3. 패스워드를 틀려서 잠긴경우, 잠김이 해제 됩니다.
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frmPW" method="post" action="/admin/member/tenbyten/procUseridChangedPw.asp">
	<input type="hidden" name="uid" value="<%=userid%>">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="28">
		<td width="100" bgcolor="#E6E6E6" align="center">텐바이텐ID</td>
		<td bgcolor="#ffffff"><%=userid%></td>
	</tr>
	<tr>
		<td bgcolor="#E6E6E6"  align="center">비밀번호</td><td bgcolor="#ffffff"><input type="password" name="sPW" size="16">
			<div style="font-size:8pt;padding:1px;">새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)</div>
			</td>
	</tr>
	<tr>
		<td bgcolor="#E6E6E6"  align="center">비밀번호 확인</td><td bgcolor="#ffffff"><input type="password" name="sPW1" size="16"></td>
	</tr>
</table>
<div style="width:100%;text-align:center;padding:10"><input type="button" class="button" value="확인" onClick="jsSubmit();"></div>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->