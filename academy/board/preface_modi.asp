<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/preface_cls.asp"-->
<%
	'// 변수 선언 //
	dim prfId, groupCd, commCd
	dim page, param, searchDiv, searchString

	dim oPreface, i, lp

	'// 파라메터 접수 //
	prfId = RequestCheckvar(request("prfId"),10)
	groupCd = RequestCheckvar(request("groupCd"),16)
	page = RequestCheckvar(request("page"),10)
	searchDiv = RequestCheckvar(request("searchDiv"),32)
	searchString = RequestCheckvar(request("searchString"),128)

	param = "&menupos=" & menupos & "&page=" & page & "&searchDiv=" & searchDiv & "&searchString=" & server.URLencode(searchString)

	'// 내용 접수
	set oPreface = new Cprf
	oPreface.FRectprfId = prfId

	oPreface.GetprfRead

	if groupCd="" then
		groupCd = oPreface.FprfList(0).FgroupCd
		commCd = oPreface.FprfList(0).FcommCd
	else
		commCd = ""
	end if
%>
<script language='javascript'>
<!--
	// 입력폼 검사
	function chk_form(frm)
	{
		if(!frm.groupCd.value)
		{
			alert("분류를 선택해주십시오.");
			frm.groupCd.focus();
			return false;
		}

		if(!frm.commCd.value)
		{
			alert("구분을 선택해주십시오.");
			frm.commCd.focus();
			return false;
		}

		if(!frm.prfCont.value)
		{
			alert("내용을 작성해주십시오.");
			frm.prfCont.focus();
			return false;
		}

		// 폼 전송
		return true;
	}

//-->
</script>
<!-- 쓰기 화면 시작 -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doPreface.asp">
<input type="hidden" name="prfId" value="<%=prfId%>">
<input type="hidden" name="mode" value="modify">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="2" height="26" align="left"><b>머릿말 내용 수정</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>상태</td>
	<td bgcolor="#FFFFFF">
		<input type="radio" name="isusing" value="Y" <% if oPreface.FprfList(0).Fisusing="Y" then Response.Write "checked"%>> 사용 &nbsp;
		<input type="radio" name="isusing" value="N" <% if oPreface.FprfList(0).Fisusing="N" then Response.Write "checked"%>> 삭제
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>구분</td>
	<td bgcolor="#FFFFFF">
		분류 <select name="groupCd">
			<option value="">선택</option>
			<%=oPreface.optgroupCd(groupCd)%>">
		</select>
		/ 구분
		<select name="commCd">
			<option value="">선택</option>
			<%=oPreface.optCommCd("'H000'", commCd)%>">
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>내용</td>
	<td bgcolor="#FFFFFF"><textarea name="prfCont" rows="14" cols="80"><%=db2html(oPreface.FprfList(0).FprfCont)%></textarea></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<a href="Preface_list.asp?prfId=<%=prfId & param%>"><img src="/images/icon_cancel.gif" border="0" align="absmiddle"></a>
	</td>
</tr>
</form>
</table>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
