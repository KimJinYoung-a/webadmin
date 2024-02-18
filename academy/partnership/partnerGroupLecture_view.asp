<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.09.10 한용민 수정/추가
'	Description : 파트너쉽
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/partner_lecturecls.asp"-->
<%
	'// 변수 선언 //
	dim idx
	dim page, searchKey, searchString, searchConfirm, param

	dim oLecture, i, lp

	'// 파라메터 접수 //
	idx = RequestCheckvar(request("idx"),10)
	page = RequestCheckvar(request("page"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = request("searchString")
	searchConfirm = RequestCheckvar(request("searchConfirm"),1)
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
		response.write "</script>"
		response.End
		end if
	end if

	param = "&page=" & page & "&searchKey=" & searchKey  &_
			"&searchString=" & server.URLencode(searchString) & "&searchConfirm=" & searchConfirm	'페이지 변수

	'// 내용 접수
	set oLecture = new CPartnerGroupLecture
	oLecture.FRectidx = idx

	oLecture.GetPartnerGroupView
%>

<script language='javascript'>

	// 입력폼 검사
	function chk_form(frm)
	{
		if(!frm.confirmMemo.value)
		{
			alert("답변 내용을 작성해주십시오.");
			frm.confirmMemo.focus();
			return false;
		}

		// 폼 전송
		return true;
	}


	// 글삭제
	function GotoLectureDel(){
		if (confirm('삭제 하시겠습니까?')){
			document.frm_write.mode.value="DelGroup";
			document.frm_write.submit();
		}
	}


</script>

<!-- 쓰기 화면 시작 -->
<table width="760" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doPartnerLecture.asp">
<input type="hidden" name="mode" value="AnsGroup">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="searchConfirm" value="<%=searchConfirm%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="4"><b>단체수강 문의 상세 내용 / 답변 작성</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">강좌명</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Flecturename%></td>
	<td align="center" width="120" bgcolor="#DDDDFF">희망강의일</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Flecturedate%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">단체명</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Fpartyname%></td>
	<td align="center" width="120" bgcolor="#DDDDFF">수강인원수</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Fpartymannumber%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">신청자명</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Fpartymastername%></td>
	<td align="center" width="120" bgcolor="#DDDDFF">휴대전화</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).Fpartymasterhp%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">전화</td>
	<td bgcolor="#FDFDFD" width="260"><%=oLecture.FItemList(0).fpartymastertel%></td>
	<td align="center" width="120" bgcolor="#DDDDFF">수강방식</td>
	<td bgcolor="#FDFDFD" width="260">
		<%
		if oLecture.FItemList(0).flecturearea = 0 then
			response.write "내부진행강좌"
		else
			response.write "외부진행강좌"
		end if
		%>	
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">이메일</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=oLecture.FItemList(0).Fpartymastermail%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">강좌취재 / 사진촬영</td>
	<td bgcolor="#FDFDFD" colspan="3"><%=oLecture.FItemList(0).Fchoiceyn%></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>답변 내용</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<textarea name="confirmMemo" rows="10" cols="80"><%=oLecture.FItemList(0).FconfirmMemo%></textarea><br>
		※ 답변 내용은 기록을 위한 것입니다. 고객에게 전달 되지 않으므로 참고용으로 사용해주십시오.
	</td>
</tr>
<tr>
	<td colspan="4" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_reply.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_delete.gif" onClick="GotoLectureDel()" style="cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="self.location='partnerGroupLecture_List.asp?menupos=<%=menupos & param%>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- 쓰기 화면 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->