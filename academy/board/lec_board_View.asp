<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/board_cls.asp"-->
<%
	'// 변수 선언 //
	dim brdId, lecUserId
	dim page, searchDiv, searchKey, searchString, isanswer, param

	dim oBoard, i, lp

	'// 파라메터 접수 //
	brdId = RequestCheckvar(request("brdId"),10)
	page = RequestCheckvar(request("page"),10)
	searchDiv = RequestCheckvar(request("searchDiv"),16)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = RequestCheckvar(request("searchString"),128)
	isanswer = RequestCheckvar(request("isanswer"),1)

	if page="" then page=1
	if searchKey="" then searchKey="qstTitle"

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey &_
			"&searchString=" & server.URLencode(searchString) & "&isanswer=" & isanswer	'페이지 변수

	'// 내용 접수
	set oBoard = new CBoard
	oBoard.FRectbrdId = brdId

	oBoard.GetBoardRead

	if (oBoard.FResultCount = 0) then
	    response.write "<script>alert('존재하지 않는 글이거나, 탈퇴한 고객입니다.'); history.back();</script>"
	    dbget.close()	:	response.End
	end if
%>
<script language='javascript'>
<!--
	// 입력폼 검사
	function chk_form(frm)
	{
		if(!frm.ansTitle.value)
		{
			alert("답변 제목을 입력해주십시오.");
			frm.ansTitle.focus();
			return false;
		}

		if(!frm.ansCont.value)
		{
			alert("답변 내용을 작성해주십시오.");
			frm.ansCont.focus();
			return false;
		}

		// 폼 전송
		return true;
	}


	// 답변 머릿말 넣기
	function chgCont(qcd, ccd)
	{
		FrameCHK.location="inc_board_cont.asp?brdId=<%=brdId%>&qcd=" + qcd + "&ccd=" + ccd;
	}

	// 구분 변경
	function GotoBoardChange(){
		if (confirm('구분을 변경하시겠습니까?')){
			document.frm_write.mode.value="change";
			document.frm_write.submit();
		}
	}

	// 글삭제
	function GotoBoardDel(){
		if (confirm('삭제 하시겠습니까?')){
			document.frm_write.mode.value="delete";
			document.frm_write.submit();
		}
	}

//-->
</script>
<!-- 쓰기 화면 시작 -->
<table width="760" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doLecBoard.asp">
<input type="hidden" name="mode" value="answer">
<input type="hidden" name="brdId" value="<%=brdId%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="isanswer" value="<%=isanswer%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="4"><b>강사게시판 상세 내용 / 답변 작성</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">구분</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<select name="commCd">
		<%=oBoard.optCommCd("'G000'", oBoard.FBoardList(0).FcommCd)%>
		</select>
		<img src="/images/icon_change.gif" onClick="GotoBoardChange()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">작성자</td>
	<td bgcolor="#FDFDFD" width="260"><%=oBoard.FBoardList(0).FlecUserId%></td>
	<td align="center" width="120" bgcolor="#E8E8F1">작성일시</td>
	<td bgcolor="#FDFDFD" width="260"><%=oBoard.FBoardList(0).Fregdate%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">문의 제목</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=db2html(oBoard.FBoardList(0).FqstTitle)%></td>
</tr>
<tr>
	<td colspan="4" align="center" bgcolor="#E8E8F1">문의 내용</td>
</tr>
<tr>
	<td bgcolor="#FDFDFD" colspan="4" style="padding:10px"><%=nl2br(db2html(oBoard.FBoardList(0).FqstCont))%></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>답변 제목</td>
	<td bgcolor="#FFFFFF" colspan="3"><input type="text" name="ansTitle" value="<%=db2html(oBoard.FBoardList(0).FansTitle)%>" size="40" maxlength="40"></td>
</tr>
<% if oBoard.FBoardList(0).Fisanswer="대기" then %>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">머릿말/인사말</td>
	<td bgcolor="#FFFFFF" colspan="3">
		머릿말
		<select name="preface" onchange="chgCont(this.value, compliment.value)">
			<%= oBoard.optCommCd("'G000'", oBoard.FBoardList(0).FcommCd)%>
		</select>
		/ 인사말
		<select name="compliment" onchange="chgCont(preface.value, this.value)">
			<option value="">선택</option>
			<%= oBoard.optCommCd("'E000'", "")%>
		</select>
	</td>
</tr>
<% end if %>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>답변 내용</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="ansCont" rows="14" cols="80"><%=db2html(oBoard.inputAnswerCont(oBoard.FBoardList(0).FbrdId,"",""))%></textarea></td>
</tr>
<tr>
	<td colspan="4" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_reply.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoBoardDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="self.location='lec_board_List.asp?menupos=<%=menupos & param%>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<iframe name="FrameCHK" src="" frameborder="0" width="0" height="0"></iframe>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
