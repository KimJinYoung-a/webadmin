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
	dim CommCd
	dim page, searchDiv, searchKey, searchString, isusing, param

	dim oComm, i, lp

	'// 파라메터 접수 //
	CommCd = RequestCheckvar(request("CommCd"),10)
	page = RequestCheckvar(request("page"),10)
	searchDiv = RequestCheckvar(request("searchDiv"),16)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = RequestCheckvar(request("searchString"),128)
	isusing = RequestCheckvar(request("isusing"),2)

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey &_
			"&searchString=" & server.URLencode(searchString) & "&isusing=" & isusing	'페이지 변수

	'// 내용 접수
	set oComm = new CComm
	oComm.FRectCommCd = CommCd

	oComm.GetCommRead

	if (oComm.FResultCount = 0) then
	    response.write "<script>alert('존재하지 않는 코드입니다.'); history.back();</script>"
	    dbget.close()	:	response.End
	end if
%>
<script language='javascript'>
<!--
	// 입력폼 검사
	function chk_form(frm)
	{
		if(!frm.commNm.value)
		{
			alert("코드명을 입력해주십시오.");
			frm.commNm.focus();
			return false;
		}

		// 폼 전송
		return true;
	}
//-->
</script>
<!-- 쓰기 화면 시작 -->
<table width="750" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doCommCd.asp">
<input type="hidden" name="mode" value="modify">
<input type="hidden" name="CommCd" value="<%=CommCd%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>공통코드 상세 내용 / 수정</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">그룹</td>
	<td width="630" bgcolor="#FFFFFF"><%=oComm.FCommList(0).FgroupNm%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">공통코드</td>
	<td width="630" bgcolor="#FFFFFF"><b><%=oComm.FCommList(0).FcommCd%></b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">코드명</td>
	<td bgcolor="#FDFDFD"><input type="text" name="commNm" value="<%=db2html(oComm.FCommList(0).FcommNm)%>" size="20" maxlength="30"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">사용여부</td>
	<td bgcolor="#FDFDFD">
		<input type="radio" name="isUsing" value="Y" <% if oComm.FCommList(0).Fisusing="사용" then Response.Write "checked"%>> 사용 &nbsp; &nbsp;
		<input type="radio" name="isUsing" value="N" <% if oComm.FCommList(0).Fisusing="삭제" then Response.Write "checked"%>> 삭제
	</td>
</tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="self.location='CommCd_List.asp?menupos=<%=menupos & param%>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->