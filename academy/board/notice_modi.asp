<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/notice_cls.asp"-->
<%
	'// 변수 선언 //
	dim ntcId
	dim page, searchDiv, searchKey, searchString

	dim oNotice, i, lp

	'// 파라메터 접수 //
	ntcId = RequestCheckvar(request("ntcId"),10)
	page = RequestCheckvar(request("page"),10)
	searchDiv = RequestCheckvar(request("searchDiv"),16)
	searchKey = RequestCheckvar(request("searchKey"),32)
	searchString = RequestCheckvar(request("searchString"),128)

	if page="" then page=1
	if searchKey="" then searchKey="titleLong"

	'// 내용 접수
	set oNotice = new CNotice
	oNotice.FRectntcId = ntcId

	oNotice.GetNoitceRead
%>
<script language='javascript'>
<!--
	// 입력폼 검사
	function chk_form(frm)
	{
		if(!frm.title.value)
		{
			alert("제목을 입력해주십시오.");
			frm.title.focus();
			return false;
		}

		if(!frm.contents.value)
		{
			alert("내용을 작성해주십시오.");
			frm.contents.focus();
			return false;
		}

		// 폼 전송
		return true;
	}
//-->
</script>
<!-- 쓰기 화면 시작 -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doNotice.asp">
<input type="hidden" name="ntcId" value="<%=ntcId%>">
<input type="hidden" name="mode" value="modify">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="2" height="26" align="left"><b>공지사항 내용 수정</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>구분</td>
	<td bgcolor="#FFFFFF">
		<select name="commCd">
		<%=oNotice.optCommCd("F000", oNotice.FNoticeList(0).FcommCd)%>
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>제목</td>
	<td bgcolor="#FFFFFF"><input type="text" name="title" size="40" maxlength="40" value="<%=db2html(oNotice.FNoticeList(0).Ftitle)%>"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">내용</td>
	<td bgcolor="#FFFFFF"><textarea name="contents" rows="14" cols="80"><%=db2html(oNotice.FNoticeList(0).Fcontents)%></textarea></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
