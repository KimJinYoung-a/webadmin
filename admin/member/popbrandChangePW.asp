<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 비밀번호 변경
' History : 2014.02.03 정윤정 생성
'			2021.07.16 한용민 수정(2차패스워드 제거)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPwithLog.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim brandid, sType
	brandid = requestCheckVar(request("bid"),32)
	sType		= requestCheckVar(request("sT"),1)
	'관리자 여부 확인
	if not (C_ADMIN_AUTH or C_SYSTEM_Part or C_CSUser or C_MD or C_OFF_part or C_logics_Part) then
			Call Alert_close ("관리자만 변경가능합니다.권한을 확인해주세요")
	end if
%>

<script type="text/javascript">
	function jsSubmit(){
		if (jsChkBlank(document.frmPW.sPW.value)){
			alert("변경할 비밀번호를 입력하세요.");
			document.frmPW.sPW.focus();
			return;
		}

		if (document.frmPW.sPW.value.replace(/\s/g, "").length < 8 || document.frmPW.sPW.value.replace(/\s/g, "").length > 16){
			alert("비밀번호는 공백없이 8~16자입니다.");
			document.frmPW.sPW.focus();
			return ;
		}

		if ((document.frmPW.sPW.value)!=(document.frmPW.sPW1.value)){
			alert("비밀번호가 일치하지 않습니다.");
			document.frmPW.sPW1.focus();
			return;
		}

		//if (jsChkBlank(document.frmPW.sPWS1.value)){
		//	alert("변경할 2차 비밀번호를 입력하세요.");
		//	document.frmPW.sPWS1.focus();
		//	return;
		//}

		//if (document.frmPW.sPWS1.value.replace(/\s/g, "").length < 8 || document.frmPW.sPWS1.value.replace(/\s/g, "").length > 16){
		//	alert("비밀번호는 공백없이 8~16자입니다.");
		//	document.frmPW.sPWS1.focus();
		//	return ;
		//}

		//if ((document.frmPW.sPWS1.value)!=(document.frmPW.sPWS2.value)){
		//	alert("비밀번호가 일치하지 않습니다.");
		//	document.frmPW.sPWS1.focus();
		//	return;
		//}

		if (!fnChkComplexPassword(frmPW.sPW.value)) {
			alert('패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)');
			frmPW.sPW.focus();
			return;
		}
		//if (!fnChkComplexPassword(frmPW.sPWS1.value)) {
		//	alert('2차 새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)');
		//	frmPW.sPWS1.focus();
		//	return;
		//}

		if(confirm("비밀변호를 변경하시겠습니까?")){
			document.frmPW.submit();
		}

	}

	//로드시 포커스
	window.onload = function(){
		document.frmPW.sPW.focus();
	}
</script>

<form name="frmPW" method="post" action="/admin/member/procbrandChangePW.asp" style="margin:0px;">
<input type="hidden" name="bid" value="<%=brandid%>">
<input type="hidden" name="sT" value="<%=sType%>">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="28">
		<td width="100" bgcolor="#E6E6E6" align="center">브랜드ID</td>
		<td bgcolor="#ffffff"><%=brandid%></td>
	</tr>
	<tr>
		<td bgcolor="#E6E6E6"  align="center">비밀번호</td><td bgcolor="#ffffff"><input type="password" name="sPW" size="16">
			<div style="font-size:8pt;padding:1px;">새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)</div>
			</td>
	</tr>
	<tr>
		<td bgcolor="#E6E6E6"  align="center">비밀번호 확인</td><td bgcolor="#ffffff"><input type="password" name="sPW1" size="16"></td>
	</tr>
	<!--<tr>
		<td bgcolor="#E6E6E6"  align="center">2차 비밀번호</td><td bgcolor="#ffffff"><input type="password" name="sPWS1" size="16"> 
			</td>
	</tr>-->
	<!--<tr>
		<td bgcolor="#E6E6E6"  align="center">2차 비밀번호 확인</td><td bgcolor="#ffffff"><input type="password" name="sPWS2" size="16"></td>
	</tr>-->
</table>
<div style="width:100%;text-align:center;padding:10"><input type="button" class="button" value="확인" onClick="jsSubmit();"></div>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->