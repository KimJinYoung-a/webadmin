<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	'// 변수 선언 //
	dim CateDiv
	CateDiv = RequestCheckvar(request("CateDiv"),16)
%>
<script language='javascript'>
<!--
	// 입력폼 검사
	function chk_form(frm)
	{
		if(!frm.CateDiv.value)
		{
			alert("카테고리 구분을 선택해주십시오.");
			frm.CateDiv.focus();
			return false;
		}

		if(frm.CateCd.value.length<2)
		{
			alert("코드를 입력해주십시오.\n\n※코드는 2자리입니다.");
			frm.CateCd.focus();
			return false;
		}

		if(!frm.Cate_Name.value)
		{
			alert("코드명을 입력해주십시오.");
			frm.Cate_Name.focus();
			return false;
		}

		// 폼 전송
		return true;
	}


	// 코드 기본값 지정
	function chgDiv(cdv)
	{
		if(cdv=='CateCD2') {
			document.all.lyEngFrm.style.display="";
		} else {
			document.all.lyEngFrm.style.display="none";
		}
	}

//-->
</script>
<!-- 쓰기 화면 시작 -->
<table width="750" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doCategory.asp">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>카테고리 코드 신규등록</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">구분</td>
	<td width="630" bgcolor="#FFFFFF">
		<select name="CateDiv" onChange="chgDiv(frm_write.CateDiv.value)">
			<option value="">선택</option>
			<option value="CateCD1" <% if CateDiv="CateCD1" then Response.Write "selected" %>>클래스</option>
			<option value="CateCD2" <% if CateDiv="CateCD2" then Response.Write "selected" %>>강좌분야</option>
			<option value="CateCD3" <% if CateDiv="CateCD3" then Response.Write "selected" %>>장소구분</option>
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">카테고리 코드</td>
	<td width="630" bgcolor="#FFFFFF">
		<input type="text" name="CateCd" size="2" maxlength="2" value="">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">코드명</td>
	<td bgcolor="#FDFDFD"><input type="text" name="Cate_Name" size="20" maxlength="30"></td>
</tr>
<tr id="lyEngFrm" <% if CateDiv<>"CateCD2" then Response.Write "style='display:none'" %>>
	<td align="center" width="120" bgcolor="#E8E8F1">코드명(영문)</td>
	<td bgcolor="#FDFDFD"><input type="text" name="Cate_NameEng" size="30" maxlength="40"></td>
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
