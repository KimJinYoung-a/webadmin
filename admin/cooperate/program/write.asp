<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/programchangeCls.asp"-->

<%
	Dim vIdx, vTitle, vContent, iCurrentpage, vRegUserID, vSign1, vSign2, vSign1Date, vSign2Date, vFileName, vRegdate, FUsername, vDoc_Idx
	Dim vParam, vChkList, vSign1Chk, vSign2Chk
	vIdx 			= requestCheckVar(Request("pidx"),10)
	vDoc_Idx		= requestCheckVar(Request("didx"),10)
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)

	Dim cPrCh
	Set cPrCh = New CProgramChange
	cPrCh.FPIdx = vIdx
	cPrCh.fnGetPrChView

	vTitle = cPrCh.FTitle
	vContent = cPrCh.FContent
	vRegUserID = cPrCh.FReguserid
	FUsername = cPrCh.FUsername
	vFileName = cPrCh.FFileName
	vSign1 = cPrCh.FSign1
	vSign2 = cPrCh.FSign2
	vSign1Date = cPrCh.FSign1date
	vSign2Date = cPrCh.FSign2date
	vRegdate = cPrCh.FRegdate
	If vIdx <> "" Then
		vDoc_Idx = cPrCh.FDocIdx
		If vDoc_Idx = "0" Then vDoc_Idx = "" End If
	End If
	vChkList = cPrCh.FChkList
	vSign1Chk = cPrCh.FSign1Chk
	vSign2Chk = cPrCh.FSign2Chk
	Set cPrCh = Nothing

	vParam = "&menupos="&request("menupos")&"&reguserid="&Request("reguserid")&"&title="&Request("title")&"&1check="&Request("1check")&"&2check="&Request("2check")&""
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language="Javascript">
function checkform()
{
	if (frm.title.value == "")
	{
		alert("제목을 입력하세요!");
		frm.title.focus();
		return;
	}
	if (frm.contents.value == "")
	{
		alert("내용을 입력하세요!");
		return;
	}
	if($('input[name="programchk"]').is(":checked") == false)
	{
		alert("체크리스트의 내용을 체크하세요!");
		return;
	}
	<% If session("ssBctId") = "tozzinet" Then %>
	if($('input[name="sign1chk"]').is(":checked") == false)
	{
		alert("1차 결제 확인을 체크하세요!");
		return;
	}
	<% ElseIf session("ssBctId") = "kobula" Then %>
	if($('input[name="sign2chk"]').is(":checked") == false)
	{
		alert("2차 결제 확인을 체크하세요!");
		return;
	}
	<% End If %>
	frm.submit();
}

function goSign(){
	frm.gubun.value = "sign";
	frm.submit();
}
</script>
<form name="frm" action="proc.asp" method="post" style="margin:0px;">
<input type="hidden" name="pidx" value="<%=vIdx%>">
<input type="hidden" name="gubun" value="">
<input type="hidden" name="menupos" value="<%=request("menupos")%>">
<input type="hidden" name="iC" value="<%=request("iCurrentpage")%>">
<input type="hidden" name="didx" value="<%=vDoc_Idx%>">
<table border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding-bottom:10">
		<table border="0" align="left" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
		<% If vIdx <> "" Then %>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">번 호</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">No. <%=vIdx%></td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">작성자</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5"><%=FUsername%> (등록일 : <%=vRegdate%>)</td>
		</tr>
		<% End If %>
		<% If vDoc_Idx <> "" Then %>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">업무협조번호</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">No. <%=vDoc_Idx%></td>
		</tr>
		<% End If %>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">제 목</td>
			<td bgcolor="#FFFFFF" style="padding: 0 0 0 5">
				<input type="text" class="text" name="title" value="<%=vTitle%>" size="110" maxlength="74">
			</td>
		</tr>
		<tr height="30">
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">파일명</td>
			<td bgcolor="#FFFFFF" style="padding: 5 0 5 5">
				<textarea class="textarea" name="filename" cols="110" rows="6"><%=vFileName%></textarea>
			</td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">간략내용</td>
			<td bgcolor="#FFFFFF" style="padding: 5 5 5 5">
				<input type="text" class="text" name="contents" value="<%=vContent%>" size="110" maxlength="198">
			</td>
		</tr>
		<tr>
			<td width="100" align="center"  bgcolor="<%= adminColor("tabletop") %>">
				<% If session("ssBctId") = "tozzinet" OR session("ssBctId") = "kobula" Then %>
					작성자<br>체크리스트
				<% Else %>
					결제자<br>체크리스트
				<% End If %>
			</td>
			<td bgcolor="#FFFFFF" style="padding: 5 5 5 5">
				<table class="a" width="100%">
				<tr>
					<td style="padding:3px;"><label id="programchk1" style="cursor:pointer;"><input type="checkbox" name="programchk" value="1" id="programchk1" <%=fnCheckBoxCheck(vChkList,"1")%>> 파라메터 체크(길이, 속성(문자형,숫자형), 값의 유무, 불완전한 값의 처리 등)</label></td>
				</tr>
				<tr>
					<td style="padding:3px;"><label id="programchk2" style="cursor:pointer;"><input type="checkbox" name="programchk" value="2" id="programchk2" <%=fnCheckBoxCheck(vChkList,"2")%>> 폼값에 개인정보(ID, PW 등의 중요정보)가 담겨있는지 체크(회원가입, 로그인 제외)</label></td>
				</tr>
				<tr>
					<td style="padding:3px;"><label id="programchk3" style="cursor:pointer;"><input type="checkbox" name="programchk" value="3" id="programchk3" <%=fnCheckBoxCheck(vChkList,"3")%>> 로그인 정보가 반드시 필요한 페이지에 ID체크 include 파일이 있는지 체크</label></td>
				</tr>
				<tr>
					<td style="padding:3px;"><label id="programchk4" style="cursor:pointer;"><input type="checkbox" name="programchk" value="4" id="programchk4" <%=fnCheckBoxCheck(vChkList,"4")%>> "(주)텐바이텐 개발 표준 및 보안 코딩 가이드" 에 맞는 코딩인지 체크</label></td>
				</tr>
				<tr>
					<td style="padding:3px;"><label id="programchk5" style="cursor:pointer;"><input type="checkbox" name="programchk" value="5" id="programchk5" <%=fnCheckBoxCheck(vChkList,"5")%>> 업로드 파일이 있는 경우 MIME TYPE, 용량, 무결성 등의 체크가 되었는지 체크</label></td>
				</tr>
				<tr>
					<td style="padding:3px;"><label id="programchk6" style="cursor:pointer;"><input type="checkbox" name="programchk" value="6" id="programchk6" <%=fnCheckBoxCheck(vChkList,"6")%>> 기획에 따른 모든 경우의 수로 테스트를 하였는지 체크</label></td>
				</tr>
				<tr>
					<td style="padding:3px;"><label id="programchk7" style="cursor:pointer;"><input type="checkbox" name="programchk" value="7" id="programchk7" <%=fnCheckBoxCheck(vChkList,"7")%>> 실서버에 올리기 위한 준비를 하였는지 체크(테스트 데이터 삭제, 소스정리 등)</label></td>
				</tr>
				<tr>
					<td style="padding:3px;"><label id="programchk8" style="cursor:pointer;"><input type="checkbox" name="programchk" value="8" id="programchk8" <%=fnCheckBoxCheck(vChkList,"8")%>> 개발 상급자가 있는 경우 최종적으로 상급자의 검증이 이루어졌는지 체크</label></td>
				</tr>
				<% If session("ssBctId") = "tozzinet" Then %>
				<tr>
					<td height="70" style="padding:15px;" bgcolor="<%= adminColor("tabletop") %>"><label id="sign1chk" style="cursor:pointer;"><input type="checkbox" name="sign1chk" value="1" id="sign1chk" <%=CHKIIF(vSign1Chk=True,"checked","")%>> 1차 결제 확인</label></td>
				</tr>
				<% ElseIf session("ssBctId") = "kobula" Then %>
				<tr>
					<td height="70" style="padding:15px;" bgcolor="<%= adminColor("tabletop") %>"><label id="sign2chk" style="cursor:pointer;"><input type="checkbox" name="sign2chk" value="1" id="sign2chk" <%=CHKIIF(vSign2Chk=True,"checked","")%>> 2차 결제 확인</label></td>
				</tr>
				<% End If %>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<table width="810" border="0" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td width="50%" align="left">
	<% If vDoc_Idx = "" Then %>
		<input type="button" value="리스트" onClick="location.href='index.asp?iC=<%=iCurrentpage%><%=vParam%>';">
	<% End If %>
	</td>
	<td width="50%" align="right">
		<% If vRegUserID = session("ssBctId") OR vIdx = "" Then %>
			<input type="button" value="저 장" onClick="checkform();">
		<% End If %>
	</td>
</tr>
</table>
</form>

<br><br>

<% If vIdx <> "" Then %>
<table border="0" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<tr height="30">
	<td align="left" bgcolor="#FFFFFF">
		1차 파트장 결제 :
		<% If session("ssBctId") = "tozzinet" Then %>
			<% If vSign1 = "" Then %>
				<input type="button" value="결제하기" onClick="goSign()">
			<% Else %>
				결제 완료. <%=vSign1Date%>
			<% End If %>
		<% Else %>
			<% If vSign1 = "" Then %>
				결제 전.
			<% Else %>
				결제 완료. <%=vSign1Date%>
			<% End If %>
		<% End If %>
	</td>
</tr>
<tr height="30">
	<td align="left" bgcolor="#FFFFFF">
		2차 팀장 결제 :
		<% If session("ssBctId") = "kobula" Then %>
			<% If vSign2 = "" Then %>
				<input type="button" value="결제하기" onClick="goSign()">
			<% Else %>
				결제 완료. <%=vSign2Date%>
			<% End If %>
		<% Else %>
			<% If vSign2 = "" Then %>
				결제 전.
			<% Else %>
				결제 완료. <%=vSign2Date%>
			<% End If %>
		<% End If %>
	</td>
</tr>
</table>
<% End If %>

<% If vIdx <> "" Then %>
<!-- ####### 답변쓰기 ####### //-->
<br>
<iframe src="iframe_program_ans.asp?pidx=<%=vIdx%>" name="iframeDB1" width="814" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
<!-- ####### 답변쓰기 ####### //-->
<% End If %>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
