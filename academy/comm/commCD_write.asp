<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/CommCd_cls.asp"-->
<%
	'// 변수 선언 //
	dim oComm, i, lp, searchDiv

	searchDiv = RequestCheckvar(request("searchDiv"),16)

	'// 클래스 선언
	set oComm = new CComm
%>
<script language='javascript'>
<!--
	// 입력폼 검사
	function chk_form(frm)
	{
		if(!frm.groupCd.value)
		{
			alert("그룹을 선택해주십시오.");
			frm.groupCd.focus();
			return false;
		}

		if(frm.commCd.value.length<4)
		{
			alert("공통코드를 입력해주십시오.\n\n※코드는 4자리입니다.");
			frm.commCd.focus();
			return false;
		}

		if(!frm.commNm.value)
		{
			alert("코드명을 입력해주십시오.");
			frm.commNm.focus();
			return false;
		}

		// 폼 전송
		return true;
	}


	// 코드 기본값 지정
	function chgGrpCd(gcd)
	{
		document.frm_write.commCd.value= gcd.substring(0,1);
	}


	// 코드 중복 검사
	function chkDuple(ccd)
	{
		if(ccd.length<4)
		{
			alert("공통코드를 입력해주십시오.\n\n※코드는 4자리입니다.");
			return;
		}
		else
		{
			FrameCHK.location = "inc_chk_commCd.asp?commCd=" + ccd;
		}
	}
//-->
</script>
<!-- 쓰기 화면 시작 -->
<table width="750" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doCommCd.asp">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>공통코드 신규등록</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">그룹</td>
	<td width="630" bgcolor="#FFFFFF">
		<select name="groupCd" onChange="chgGrpCd(frm_write.groupCd.value)">
			<option value="">전체</option>
			<%= oComm.optGroupCd(searchDiv)%>
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">공통코드</td>
	<td width="630" bgcolor="#FFFFFF">
		<input type="text" name="commCd" size="4" maxlength="4" value="<%=left(searchDiv,1)%>">
		<img src="/images/icon_1.gif" width="55" height="21" border="0" onClick="chkDuple(frm_write.commCd.value)" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">코드명</td>
	<td bgcolor="#FDFDFD"><input type="text" name="commNm" size="20" maxlength="30"></td>
</tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<iframe name="FrameCHK" src="" frameborder="0" width="0" height="0"></iframe>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
