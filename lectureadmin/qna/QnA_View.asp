<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/QnA_cls.asp"-->
<%
	'// 변수 선언 //
	dim qnaId, qstUserId
	dim page, searchDiv, searchKey, searchString, isanswer, param

	dim oQnA, oQnAList, oLec, i, lp

	'// 파라메터 접수 //
	qnaId = RequestCheckvar(request("qnaId"),10)
	page = RequestCheckvar(request("page"),10)
	searchDiv = RequestCheckvar(request("searchDiv"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = RequestCheckvar(request("searchString"),128)
	isanswer = RequestCheckvar(request("isanswer"),1)
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end if
	if page="" then page=1
	if searchKey="" then searchKey="qstTitle"

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey &_
			"&searchString=" & server.URLencode(searchString) & "&isanswer=" & isanswer	'페이지 변수

	'// 내용 접수
	set oQnA = new CQnA_Lecture
	oQnA.FRectqnaId = qnaId

	oQnA.GetQnARead

	if (oQnA.FResultCount = 0) then
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

		if(!frm.ansContents.value)
		{
			alert("답변 내용을 작성해주십시오.");
			frm.ansContents.focus();
			return false;
		}

		// 폼 전송
		return true;
	}


	// 답변 머릿말 넣기
	function chgCont(qcd, ccd)
	{
		FrameCHK.location="inc_qna_cont.asp?qnaId=<%=qnaId%>&qcd=" + qcd + "&ccd=" + ccd;
	}

	// 글삭제
	function GotoqnaDel(){
		if (confirm('삭제 하시겠습니까?')){
			document.frm_write.mode.value="delete";
			document.frm_write.submit();
		}
	}

//-->
</script>
<!-- 쓰기 화면 시작 -->
<table width="760" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doQnA.asp">
<input type="hidden" name="mode" value="answer">
<input type="hidden" name="qnaId" value="<%=qnaId%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="isanswer" value="<%=isanswer%>">
<input type="hidden" name="qstUserName" value="<%=oQnA.FQnAList(0).Fusername%>">
<input type="hidden" name="regdate" value="<%=oQnA.FQnAList(0).Fregdate%>">
<input type="hidden" name="qstContents" value="<%=db2html(oQnA.FQnAList(0).FqstContents)%>">
<input type="hidden" name="qstTitle" value="<%=db2html(oQnA.FQnAList(0).FqstTitle)%>">

<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="4"><b>QnA 상세 내용 / 답변 작성</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">분류</td>
	<td width="260" bgcolor="#FFFFFF"><%=oQnA.FQnAList(lp).FgroupNm%></td>
	<td align="center" width="120" bgcolor="#E8E8F1">구분</td>
	<td width="260" bgcolor="#FFFFFF">
		<%=oQnA.FQnAList(0).FcommNm%>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">작성자</td>
	<td bgcolor="#FDFDFD" width="260"><%=oQnA.FQnAList(0).Fusername & "(" & oQnA.FQnAList(0).FqstUserid & ")"%></td>
	<td align="center" width="120" bgcolor="#E8E8F1">작성일시</td>
	<td bgcolor="#FDFDFD" width="260"><%=oQnA.FQnAList(0).Fregdate%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">수신 이메일</td>
	<td bgcolor="#FDFDFD">
		<%=db2html(oQnA.FQnAList(0).FqstUserMail)%>
		<input type="hidden" name="qstUserMail" value="<%=oQnA.FQnAList(0).FqstUserMail%>">
	</td>
	<td align="center" width="120" bgcolor="#E8E8F1">메일 수신여부</td>
	<td bgcolor="#FDFDFD">
		<%=oQnA.FQnAList(0).FmailOk%>
		<input type="hidden" name="mailOk" value="<%=oQnA.FQnAList(0).FmailOk%>">
	</td>
</tr>
<%
	if oQna.FQnAList(0).FlecIdx<>"" then
		set oLec = new CQnA
		oLec.FRectlecIdx = oQna.FQnAList(0).FlecIdx

		oLec.GetLecRead

		if oLec.FlecList(0).FcateName<>"" then
%>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">강좌정보</td>
	<td bgcolor="#FDFDFD" width="640" colspan="3"><%= "[" & oLec.FlecList(0).FcateName & "] " & db2html(oLec.FlecList(0).FlecTitle)%></td>
</tr>
<%
		end if
	end if
%>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">문의 제목</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=db2html(oQnA.FQnAList(0).FqstTitle)%></td>
</tr>
<tr>
	<td colspan="4" align="center" bgcolor="#E8E8F1">문의 내용</td>
</tr>
<tr>
	<td bgcolor="#FDFDFD" colspan="4" style="padding:10px"><%=nl2br(db2html(oQnA.FQnAList(0).FqstContents))%></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>답변 제목</td>
	<td bgcolor="#FFFFFF" colspan="3"><input type="text" name="ansTitle" value="<%=db2html(oQnA.FQnAList(0).FansTitle)%>" size="40" maxlength="40"></td>
</tr>
<% if oQnA.FQnAList(0).Fisanswer="대기" then %>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">인사말</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="hidden" name="preface" value="A999">
		<select name="compliment" onchange="chgCont(preface.value, this.value)">
			<option value="">선택</option>
			<%= oQnA.optCommCd("'E000'", "")%>
		</select>
	</td>
</tr>
<% end if %>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>답변 내용</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="ansContents" rows="14" cols="80"><%=db2html(oQnA.inputAnswerCont(oQnA.FQnAList(0).FqnaId,"A999",""))%></textarea></td>
</tr>
<tr>
	<td colspan="4" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_reply.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoqnaDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_list.gif" onClick="self.location='QnA_List.asp?menupos=<%=menupos & param%>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<iframe name="FrameCHK" src="" frameborder="0" width="0" height="0"></iframe>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
