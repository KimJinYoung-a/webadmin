<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/Compliment_cls.asp"-->
<%
	'// 변수 선언 //
	dim cplId
	dim page, param, searchDiv, searchString

	dim oCompliment, i, lp

	'// 파라메터 접수 //
	cplId = RequestCheckvar(request("cplId"),10)
	page = RequestCheckvar(request("page"),10)
	searchDiv = RequestCheckvar(request("searchDiv"),16)
	searchString = RequestCheckvar(request("searchString"),128)

	param = "&menupos=" & menupos & "&page=" & page & "&searchDiv=" & searchDiv & "&searchString=" & server.URLencode(searchString)

	'// 내용 접수
	set oCompliment = new Ccpl
	oCompliment.FRectcplId = cplId

	oCompliment.GetcplRead

%>
<script language='javascript'>
<!--
	// 입력폼 검사
	function chk_form(frm)
	{
		if(!frm.commCd.value)
		{
			alert("구분을 선택해주십시오.");
			frm.commCd.focus();
			return false;
		}

		if(!frm.cplCont.value)
		{
			alert("내용을 작성해주십시오.");
			frm.cplCont.focus();
			return false;
		}

		// 폼 전송
		return true;
	}

//-->
</script>
<!-- 쓰기 화면 시작 -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doCompliment.asp">
<input type="hidden" name="cplId" value="<%=cplId%>">
<input type="hidden" name="mode" value="modify">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="2" height="26" align="left"><b>인사말 내용 수정</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>상태</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" <% if oCompliment.FcplList(0).Fisusing="Y" then Response.Write "checked"%>> 사용 &nbsp;
		<input type="radio" name="isusing" value="N" <% if oCompliment.FcplList(0).Fisusing="N" then Response.Write "checked"%>> 삭제
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>구분</td>
	<td bgcolor="#FFFFFF">
		<select name="commCd">
			<option value="">선택</option>
			<%=oCompliment.optCommCd("'E000'", oCompliment.FcplList(0).FcommCd)%>">
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>내용</td>
	<td bgcolor="#FFFFFF"><textarea name="cplCont" rows="14" cols="80"><%=oCompliment.FcplList(0).FcplCont%></textarea></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<a href="Compliment_list.asp?cplId=<%=cplId & param%>"><img src="/images/icon_cancel.gif" border="0" align="absmiddle"></a>
	</td>
</tr>
</form>
</table>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->