<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 qna
' Hieditor : 2009.11.27 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->
<%
'// 변수 선언 //
dim qnaId, qstUserId , searchDiv
dim oQnA, oQnAList, oLec, i, lp
	'// 파라메터 접수 //
	qnaId = request("qnaId")
	searchDiv = request("searchDiv")

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

	// 글삭제
	function GotoqnaDel(){
		if (confirm('삭제 하시겠습니까?')){
			document.frm_write.mode.value="delete";
			document.frm_write.submit();
		}
	}

</script>

<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doQnA.asp">
<input type="hidden" name="mode" value="answer">
<input type="hidden" name="qnaId" value="<%=qnaId%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="qstUserName" value="<%=oQnA.FQnAList(0).Fusername%>">
<input type="hidden" name="regdate" value="<%=oQnA.FQnAList(0).Fregdate%>">
<input type="hidden" name="qstContents" value="<%=db2html(oQnA.FQnAList(0).FqstContents)%>">
<input type="hidden" name="qstTitle" value="<%=db2html(oQnA.FQnAList(0).FqstTitle)%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="4"><b>QnA 상세 내용 / 답변 작성</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#FFFFFF">수신 이메일</td>
	<td bgcolor="#FFFFFF">
		<%=db2html(oQnA.FQnAList(0).FqstUserMail)%>
		<input type="hidden" name="qstUserMail" value="<%=oQnA.FQnAList(0).FqstUserMail%>">
	</td>
	<td align="center" width="120" bgcolor="#FFFFFF">상태</td>
	<td width="260" bgcolor="#FFFFFF">
		<%=oQnA.FQnAList(0).Fisanswer%>
	</td>	
</tr>
<tr>
	<td align="center" width="120" bgcolor="#FFFFFF">작성자</td>
	<td bgcolor="#FDFDFD" width="260">
		<%=oQnA.FQnAList(0).Fusername & "(" & oQnA.FQnAList(0).FqstUserid & ")"%>
	</td>
	<td align="center" width="120" bgcolor="#FFFFFF">작성일시</td>
	<td bgcolor="#FDFDFD" width="260"><%=oQnA.FQnAList(0).Fregdate%></td>
</tr>
<%
	if oQna.FQnAList(0).FlecIdx<>"" then
		set oLec = new CQnA
		oLec.FRectlecIdx = oQna.FQnAList(0).FlecIdx

		oLec.GetLecRead

		if oLec.FlecList(0).FcateName<>"" then
%>
<tr>
	<td align="center" width="120" bgcolor="#FFFFFF">강좌정보</td>
	<td bgcolor="#FDFDFD" width="640" colspan="3"><%= "[" & oLec.FlecList(0).FcateName & "] " & db2html(oLec.FlecList(0).FlecTitle)%></td>
</tr>
<%
		end if
	end if
%>
<tr>
	<td align="center" width="120" bgcolor="#FFFFFF">문의 제목</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=db2html(oQnA.FQnAList(0).FqstTitle)%></td>
</tr>
<tr>
	<td colspan="4" align="center" bgcolor="#FFFFFF">문의 내용</td>
</tr>
<tr>
	<td bgcolor="#FDFDFD" colspan="4" style="padding:10px"><%=nl2br(db2html(oQnA.FQnAList(0).FqstContents))%></td>
</tr>
</table>
<br>
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>답변 제목</td>
	<td bgcolor="#FFFFFF" colspan="3"><input type="text" name="ansTitle" value="<%=db2html(oQnA.FQnAList(0).FansTitle)%>" size="40" maxlength="40"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>답변 내용</td>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="ansContents" rows="14" cols="80"><%=db2html(oQnA.FQnAList(0).FansContents)%></textarea></td>
</tr>
<tr>
	<td colspan="4" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_reply.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoqnaDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="self.location='QnA_List.asp'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>

<!-- 관련 리스트 시작  -->
<%

'문의자 아이디 저장
qstUserId = oQnA.FQnAList(0).FqstUserid
set oQnAList = Nothing

'//회원일 경우
if qstUserId <> "" then
	
	'// 다시 클래스 선언
	set oQnAList = new CQnA
		oQnAList.FCurrPage = 1 
		oQnAList.FPageSize = 50
		oQnAList.FRectuserid = qstUserId
		oQnAList.GetQnAList
%>
	<br>
	<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
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
			<td><%= oQnAList.FQnAList(lp).Fcommcd %></td>
			<td align="left"><a href="QnA_view.asp?qnaId=<%= oQnAList.FQnAList(lp).FqnaId %>"><%= db2html(oQnAList.FQnAList(lp).FqstTitle) %></a></td>
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
end if	
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->