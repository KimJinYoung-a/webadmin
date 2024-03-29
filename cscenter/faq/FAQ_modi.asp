<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [CS]각종설정>>[FAQ]관리 
' Hieditor : 2009.03.02 이영진 생성
'			 2021.07.30 한용민 수정(사용여부 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/faq_cls.asp"-->
<%
	'// 변수 선언 //
	dim faqid, isusing
	dim page, searchDiv, searchKey, searchString, param

	dim ofaq, i, lp

	'// 파라메터 접수 //
	faqid = request("faqid")
	page = request("page")
	searchDiv = request("searchDiv")
	searchKey = request("searchKey")
	searchString = request("searchString")
	isusing = requestcheckvar(request("isusing"),1)

	if page="" then page=1
	if searchKey="" then searchKey="title"

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'페이지 변수

	'// 내용 접수
	set ofaq = new Cfaq
	ofaq.FRectfaqid = faqid

	ofaq.GetFAQRead

isusing = ofaq.FfaqList(0).fisusing
if isusing="" or isnull(isusing) then isusing="Y"
%>
<script type='text/javascript'>
<!--
	// 입력폼 검사
	function chk_form(frm)
	{
		if(!frm.commCd.value)
		{
			alert("구분을 선택해주십시오.");
			frm.commCd.focus();
			return false;
		}

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
		if(!frm.isusing.value){
			alert("사용여부를 선택해 주세요.");
			frm.isusing.focus();
			return false;
		}

		// 폼 전송
		return true;
	}
//-->
</script>
<!-- 쓰기 화면 시작 -->
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="faq_process.asp">
<input type="hidden" name="mode" value="UPD">
<input type="hidden" name="faqId" value="<%=faqId%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#F0F0FD">
	<td colspan="2"><b>faq 정보 수정</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>구분</td>
	<td bgcolor="#FFFFFF">
		<select class="select" name="commCd">
			<option value="">선택</option>
			<%= db2html(ofaq.optCommCd("Z200", ofaq.FfaqList(0).FcommCd)) %>
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>정렬순서</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="disporder" value="<%= ofaq.FfaqList(0).Fdisporder %>" size="3" maxlength="3">&nbsp;&nbsp;숫자입력(0-999)사이값</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>제목</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="title" value="<%= ReplaceBracket(db2html(ofaq.FfaqList(0).Ftitle)) %>" size="80" maxlength="80"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>내용</td>
	<td bgcolor="#FFFFFF"><textarea name="contents" class="textarea" rows="14" cols="80"><%= ReplaceBracket(db2html(ofaq.FfaqList(0).Fcontents)) %></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>Link명</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="linkname" value="<%= ReplaceBracket(db2html(ofaq.FfaqList(0).Flinkname)) %>" size="30" maxlength="30"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>LinkURL</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="linkurl" value="<%= ReplaceBracket(db2html(ofaq.FfaqList(0).Flinkurl)) %>" size="80" maxlength="80"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>사용여부</td>
	<td bgcolor="#FFFFFF"><% drawSelectBoxUsingYN "isusing", isusing %></td>
</tr>
<tr align="center" height="25" bgcolor="#F0F0FD">
	<td colspan="2">
		<input type="submit" class="button" value="저장하기">
		<input type="button" class="button" value="취소하기" onClick="history.back()">
	</td>
</tr>
</table>
</form>
<!-- 쓰기 화면 끝 -->
<%
	set ofaq = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->