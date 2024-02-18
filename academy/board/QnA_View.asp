<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.09.16 한용민 수정
'	Description : 1:1상담
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
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

	if page="" then page=1
	if searchKey="" then searchKey="qstTitle"

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey &_
			"&searchString=" & server.URLencode(searchString) & "&isanswer=" & isanswer	'페이지 변수

	'// 내용 접수
	set oQnA = new CQnA
	oQnA.FRectqnaId = qnaId

	oQnA.GetQnARead

	if (oQnA.FResultCount = 0) then
	    response.write "<script>alert('존재하지 않는 글이거나, 탈퇴한 고객입니다.'); history.back();</script>"
	    dbget.close()	:	response.End
	end if
%>
<script language='javascript'>

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

	// 구분 변경
	function GotoqnaChange(){
		if (confirm('구분을 변경하시겠습니까?')){
			document.frm_write.mode.value="change";
			document.frm_write.submit();
		}
	}

	// 글삭제
	function GotoqnaDel(){
		if (confirm('삭제 하시겠습니까?')){
			document.frm_write.mode.value="delete";
			document.frm_write.submit();
		}
	}

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
		<select name="commCd">
		<%=oQnA.optCommCd("'" & Left(oQnA.FQnAList(0).FcommCd,1) & "000'", oQnA.FQnAList(0).FcommCd)%>
		</select>
		<img src="/images/icon_change.gif" onClick="GotoqnaChange()" style="cursor:pointer" align="absmiddle">
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
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">만족도(1-5개)</td>
	<td bgcolor="#FDFDFD" width="260">
		<% 
		if oQnA.FQnAList(0).fbestviewcount <> 0 then
		for i = 1 to oQnA.FQnAList(0).fbestviewcount 
		%>
		<img src="http://image.thefingers.co.kr/academy2009/lecture/star_on_gray.gif">
		<% 
		next 
		else
		%>
		평가 없음
		<% end if %>
	</td>
	<td align="center" width="120" bgcolor="#E8E8F1"></td>
	<td bgcolor="#FDFDFD" width="260"></td>
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
	<td align="center" width="120" bgcolor="#E8E8F1">관련주문번호</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=db2html(oQnA.FQnAList(0).Forderserial)%></td>
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
	<td align="center" width="120" bgcolor="#DDDDFF">머릿말/인사말</td>
	<td bgcolor="#FFFFFF" colspan="3">
		머릿말
		<select name="preface" onchange="chgCont(this.value, compliment.value)">
			<%= oQnA.optPrfCd("'A000'", "H999")%>
		</select>
		/ 인사말
		<select name="compliment" onchange="chgCont(preface.value, this.value)">
			<option value="">선택</option>
			<%= oQnA.optCommCd("'E000'", "")%>
		</select>
	</td>
</tr>
<% end if %>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>답변 내용</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="ansContents" rows="14" cols="80"><%=db2html(oQnA.inputAnswerCont(oQnA.FQnAList(0).FqnaId,"",""))%></textarea></td>
</tr>
<tr>
	<td colspan="4" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_reply.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoqnaDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="self.location='QnA_List.asp?menupos=<%=menupos & param%>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<iframe name="FrameCHK" src="" frameborder="0" width="0" height="0"></iframe>
<!-- 쓰기 화면 끝 -->
<%
		'문의자 아이디 저장
		qstUserId = oQnA.FQnAList(0).FqstUserid
	set oQnA = Nothing
%>
<!-- 관련 리스트 시작  -->
<%
	'// 클래스 선언
	set oQnAList = new CQnA
	oQnAList.FCurrPage = 1 ''page <- 이페이지가 그 페이지가 아님.
	oQnAList.FPageSize = 200
	oQnAList.FRectuserid = qstUserId

	oQnAList.GetQnAList
%>
<br><br>
<table width="760" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#F0F0FD">
		<td colspan="6" align="center"><%= qstUserId %> 회원의 지난 문의 목록</td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="40">번호</td>
		<td align="center" width="120">구분</td>
		<td align="center">제목</td>
		<td align="center" width="70">등록자</td>
		<td align="center" width="50">상태</td>
		<td align="center" width="80">등록일</td>
	</tr>
	<%
		for lp=0 to oQnAList.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oQnAList.FQnAList(lp).FqnaId %></td>
		<td><%= oQnAList.FQnAList(lp).FcommNm %></td>
		<td align="left"><a href="QnA_view.asp?qnaId=<%= oQnAList.FQnAList(lp).FqnaId & param%>"><%= db2html(oQnAList.FQnAList(lp).FqstTitle) %></a></td>
		<td><%= oQnAList.FQnAList(lp).FqstUserId %></td>
		<td><%= oQnAList.FQnAList(lp).Fisanswer %></td>
		<td><%= FormatDate(oQnAList.FQnAList(lp).Fregdate,"0000.00.00") %></td>
	</tr>
	<%
		next
	%>
</table>
<!-- 관련 리스트 끝  -->
<%
	set oQnAList = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->