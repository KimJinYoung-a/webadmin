<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 사내일정공지
' Hieditor : 이상구 생성
'			 2022.07.12 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member_board/boardCls.asp"-->
<%

Dim lBoardScmNotice
Set lBoardScmNotice = new board
	lBoardScmNotice.fnGetScmNoticeList

dim i

' 권한체크
IF Not(C_OP Or C_PSMngPart Or C_SYSTEM_Part or C_ADMIN_AUTH) Then
	Response.Write "<script type='text/javascript'>alert('사내일정공지 등록/수정은 인사총무팀과 개발팀만 가능합니다.'); self.close();</script>"
	Response.End
End If
%>
<!-- 검색 시작 -->
<script type='text/javascript'>

function jsSubmitIns() {
	var frm = document.frmadd;

	if (frm.scheduleDate.value == '') {
		alert('일정을 입력하세요.');
		frm.scheduleDate.focus();
		return;
	}

	if (frm.title.value == '') {
		alert('제목을 입력하세요.');
		frm.title.focus();
		return;
	}

	if (frm.contents.value == '') {
		alert('내용을 입력하세요.');
		frm.contents.focus();
		return;
	}

	if (frm.dispno.value == '') {
		alert('표시순서를 입력하세요.');
		frm.dispno.focus();
		return;
	}

	if (confirm('저장하시겠습니까?') == true) {
		frm.submit();
	}
}

function jsSubmitModi(frm) {
	if (frm.scheduleDate.value == '') {
		alert('일정을 입력하세요.');
		frm.scheduleDate.focus();
		return;
	}

	if (frm.title.value == '') {
		alert('제목을 입력하세요.');
		frm.title.focus();
		return;
	}

	if (frm.contents.value == '') {
		alert('내용을 입력하세요.');
		frm.contents.focus();
		return;
	}

	if (frm.dispno.value == '') {
		alert('표시순서를 입력하세요.');
		frm.dispno.focus();
		return;
	}

	if (confirm('저장하시겠습니까?') == true) {
		frm.submit();
	}
}

function jsSubmitDel(frm) {
	if (confirm('삭제하시겠습니까?') == true) {
		frm.mode.value = 'del';
		frm.submit();
	}
}

</script>

<!-- 상단 띠 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= lBoardScmNotice.FResultCount %></b>
		</td>
	</tr>
	<tr height=30 align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">idx</td>
		<td width="100">일정</td>
		<td width="120">제목</td>
		<td width="210">내용</td>
		<td width="100">최종수정</td>
		<td width="40">표시<br />순서</td>
		<td>비고</td>
    </tr>
	<% for i = 0 to lBoardScmNotice.FResultCount - 1 %>
	<tr height=30 align="center" bgcolor="#FFFFFF">
		<form name="frmmodi<%= i %>" method="post" action="popScmNoticeModi_process.asp">
		<input type="hidden" name="mode" value="modi">
		<input type="hidden" name="idx" value="<%= lBoardScmNotice.FbrdList(i).Fidx %>">
		<td><%= lBoardScmNotice.FbrdList(i).Fidx %></td>
		<td>
			<input type="text" class="text" name="scheduleDate" value="<%= ReplaceBracket(lBoardScmNotice.FbrdList(i).FscheduleDate) %>" size="10">
		</td>
		<td>
			<input type="text" class="text" name="title" value="<%= ReplaceBracket(lBoardScmNotice.FbrdList(i).Ftitle) %>" size="15">
		</td>
		<td>
			<textarea class="textarea" name="contents" value="" cols="30" rows="3"><%= ReplaceBracket(lBoardScmNotice.FbrdList(i).Fcontents) %></textarea>
		</td>
		<td><%= lBoardScmNotice.FbrdList(i).Fmodiuserid %></td>
		<td>
			<input type="text" class="text" name="dispno" value="<%= lBoardScmNotice.FbrdList(i).Fdispno %>" size="2">
		</td>
		<td>
			<input type="button" class="button" value="수정하기" onClick="jsSubmitModi(frmmodi<%= i %>)">
			&nbsp;
			<input type="button" class="button" value="삭제하기" onClick="jsSubmitDel(frmmodi<%= i %>)">
		</td>
		</form>
	</tr>
	<% next %>
	<tr height=30 align="center" bgcolor="#FFFFFF">
		<form name="frmadd" method="post" action="popScmNoticeModi_process.asp" style="margin:0px;">
		<input type="hidden" name="mode" value="add">
		<td>신규</td>
		<td>
			<input type="text" class="text" name="scheduleDate" value="" size="10">
		</td>
		<td>
			<input type="text" class="text" name="title" value="" size="15">
		</td>
		<td>
			<textarea class="textarea" name="contents" value="" cols="30" rows="3"></textarea>
		</td>
		<td><%= session("ssBctId") %></td>
		<td>
			<input type="text" class="text" name="dispno" value="" size="2">
		</td>
		<td>
			<input type="button" class="button" value="등록하기" onClick="jsSubmitIns()">
		</td>
		</form>
	</tr>
</table>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
