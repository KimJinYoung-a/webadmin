<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp"-->
<%
'###############################################
' PageName : category_left_topKeyword_write.asp
' Discription : 카테고리 탑키워드 등록/수정
' History : 2008.03.31 허진원 생성
'         : 2008.10.27 중카테고리 처리 추가(허진원)
'         : 2009.04.15 이미지 추가(허진원)
'###############################################

	Dim idx, cdl, cdm

	idx = Request("idx")
	cdl = Request("cdl")
	cdm = Request("cdm")

	'// 내용 접수
	dim oKeyword
	Set oKeyword = new CCategoryKeyWord
	oKeyword.FRectIdx = idx

	if idx<>"" then
		oKeyword.GetCaFavKeyWord
	end if
%>
<script language="javascript">
<!--
	// 아이템 저장 실행
	function goSubmit()
	{
		// 카테고리 선택여부 검사
		if(!(document.frm.cdl.value)) {
			alert("카테고리를 선택해주세요.");
			return;
		}
		// 키워드 입력여부 검사
		if(!document.frm.keyword.value) {
			alert("관련 키워드를 입력해주세요.");
			document.frm.keyword.focus();
			return;
		}
		// 링크 입력여부 검사
		if(!document.frm.linkinfo.value) {
			alert("키워드 클릭시 이동할 링크를 입력해주세요.");
			document.frm.linkinfo.focus();
			return;
		}
		// 순서 입력여부 검사
		if(!document.frm.sortNo.value) {
			alert("표시 순서를 입력해주세요.\n※ 순서는 숫자이며 적을수록 순번이 높습니다.");
			document.frm.sortNo.focus();
			return;
		}

		<% if idx="" then %>
		if(confirm("작성하신 내용을 등록하시겠습니까?")) {
			document.frm.mode.value="add";
			document.frm.action="DoCateTopKeyword.asp";
			document.frm.submit();
		}
		<% else %>
		if(confirm("수정하신 내용을 저장하시겠습니까?")) {
			document.frm.mode.value="modify";
			document.frm.action="DoCateTopKeyword.asp";
			document.frm.submit();
		}
		<% end if %>
	}

	// 카테고리 변경시 명령
	function changecontent() {
		if(frm.cdl.value=="110") {
			frm.action="category_left_topKeyword_write.asp";
			frm.submit();
		}
	}

	function putLinkText(key) {
		switch(key) {
			case 'search':
				document.frm.linkinfo.value='/search/search_result.asp?rect=' + document.frm.keyword.value;
				break;
			case 'cate':
				document.frm.linkinfo.value='/shopping/category_list.asp?cdl=대코드&cdm=중코드&cds=소코드';
				break;
			case 'event':
				document.frm.linkinfo.value='/event/eventmain.asp?eventid=이벤트번호';
				break;
		}
	}
//-->
</script>
<!-- 폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="DoCateTopKeyword.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="mode" value="">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<% if idx="" then %>
		<font color="red"><b>탑키워드 등록</b></font>
		<% else %>
		<font color="red"><b>탑키워드 수정</b></font>
		<% end if%>
	</td>
</tr>
<% if idx<>"" then %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">일련번호</td>
	<td align="left"><input type="text" name="idx" value="<%=idx%>" readonly size="10" class="text_ro"></td>
</tr>
<% end if %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">카테고리</td>
	<td align="left">
	<%
		if idx<>"" then
			DrawSelectBoxCategoryLarge "cdl", oKeyword.FitemList(0).FCDL
			if oKeyword.FitemList(0).FCDL="110" then
				if cdm="" then
					DrawSelectBoxCategoryMid "cdm", oKeyword.FitemList(0).FCDL, oKeyword.FitemList(0).FCDM
				else
					DrawSelectBoxCategoryMid "cdm", oKeyword.FitemList(0).FCDL, cdm
				end if
			end if
		else
			DrawSelectBoxCategoryLarge "cdl", cdl
			if cdl="110" then
				DrawSelectBoxCategoryMid "cdm", cdl, cdm
			end if
		end if
	%>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">키워드</td>
	<td align="left"><input type="text" name="keyword" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).FKeyword%>" size="32" maxlength="32" class="text"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">링크</td>
	<td align="left">
		<table cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td colspan="2"><input type="text" name="linkinfo" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).Flinkinfo%>" size="80" maxlength="128" class="text"></td>
		<tr>
		<tr>
			<td valign="top"><font color="#707080">예)</font></td>
			<td valign="top">
				<font color="#707070">
				- <span style="cursor:pointer" onClick="putLinkText('search')">검색결과 링크 : /search/search_result.asp?rect=<font color="darkred">검색어</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('cate')">카테고리 링크 : /shopping/category_list.asp?cdl=<font color="darkred">대코드</font>&cdm=<font color="darkred">중코드</font>&cds=<font color="darkred">소코드</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('event')">이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font></span>
				</font>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">관련상품</td>
	<td align="left">
		<input type="text" name="itemid" size="10" maxlength="8" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).Fitemid%>">
		<font color="#707070">※ 상품번호를 입력, 키워드 관련 이미지용이므로 키워드와 연관이 없어도 무관</font>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">표시순서</td>
	<td align="left"><input type="text" name="sortNo" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).FsortNo: else Response.Write "0" %>" size="3" class="text"></td></td>
</tr>
<tr>
	<td align="center" colspan="2" bgcolor="#FFFFFF">
		<input type="button" class="button" value="저장" onClick="goSubmit()"> &nbsp;
		<input type="button" class="button" value="취소" onClick="self.history.back()">
	</td>
</tr>
</form>
<!-- 폼 끝 -->
</table>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
