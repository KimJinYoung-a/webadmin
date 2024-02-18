<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/categoryCls.asp"-->
<%
	'// 변수 선언 //
	dim CateCd
	dim page, searchKey, searchString, isusing, param, CateDiv
	dim oCate, i, lp

	'// 파라메터 접수 //
	CateDiv = RequestCheckvar(request("CateDiv"),16)
	CateCd = RequestCheckvar(request("CateCd"),3)
	page = RequestCheckvar(request("page"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = request("searchString")
	isusing = RequestCheckvar(request("isusing"),1)
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	param = "&page=" & page & "&searchKey=" & searchKey & "&CateDiv=" & CateDiv &_
			"&searchString=" & server.URLencode(searchString) & "&isusing=" & isusing	'페이지 변수

	'// 내용 접수
	set oCate = new CCate
	oCate.FCateDiv = CateDiv
	oCate.FRectCateCd = CateCd

	oCate.GetCateRead

	if (oCate.FResultCount = 0) then
	    response.write "<script>alert('존재하지 않는 코드입니다.'); history.back();</script>"
	    dbget.close()	:	response.End
	end if

	function getCateDivName(cdv)
		Select Case cdv
			Case "CateCD1"
				getCateDivName = "클래스"
			Case "CateCD2"
				getCateDivName = "강좌분야"
			Case "CateCD3"
				getCateDivName = "장소구분"
		End Select
	end function
%>
<script language='javascript'>
<!--
	// 입력폼 검사
	function chk_form(frm)
	{
		if(!frm.Cate_Name.value)
		{
			alert("코드명을 입력해주십시오.");
			frm.Cate_Name.focus();
			return false;
		}

		// 폼 전송
		return true;
	}
//-->
</script>
<!-- 쓰기 화면 시작 -->
<table width="750" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doCategory.asp">
<input type="hidden" name="mode" value="modify">
<input type="hidden" name="CateDiv" value="<%=CateDiv%>">
<input type="hidden" name="CateCd" value="<%=CateCd%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>카테고리 코드 상세 내용 / 수정</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">구분</td>
	<td width="630" bgcolor="#FFFFFF"><%=getCateDivName(CateDiv)%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">카테고리 코드</td>
	<td width="630" bgcolor="#FFFFFF"><b><%=oCate.FCateList(0).FCateCd%></b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">코드명</td>
	<td bgcolor="#FDFDFD"><input type="text" name="Cate_Name" value="<%=db2html(oCate.FCateList(0).FCateCD_Name)%>" size="20" maxlength="30"></td>
</tr>
<% if CateDiv="CateCD2" then %>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">코드명(영문)</td>
	<td bgcolor="#FDFDFD"><input type="text" name="Cate_NameEng" value="<%=db2html(oCate.FCateList(0).FCateCD_NameEng)%>" size="30" maxlength="40"></td>
</tr>
<% end if %>
<% if CateDiv<>"CateCD1" then %>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">표시순서</td>
	<td bgcolor="#FDFDFD"><input type="text" name="sortNo" value="<%=db2html(oCate.FCateList(0).FsortNo)%>" size="3"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">사용여부</td>
	<td bgcolor="#FDFDFD">
		<input type="radio" name="isUsing" value="Y" <% if oCate.FCateList(0).Fisusing="사용" then Response.Write "checked"%>> 사용 &nbsp; &nbsp;
		<input type="radio" name="isUsing" value="N" <% if oCate.FCateList(0).Fisusing="삭제" then Response.Write "checked"%>> 삭제
	</td>
</tr>
<% end if %>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="self.location='CategoryList.asp?menupos=<%=menupos & param%>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->