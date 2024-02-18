<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : RelateKeywordLink_Edit.asp
' Discription : 카테고리 관련 키워드 등록/수정
' History : 2008.03.28 허진원 생성
'			2022.07.05 한용민 수정(isms취약점조치, 표준코딩으로변경)
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp"-->
<%
	Dim LinkCode
	LinkCode = requestcheckvar(getNumeric(request("rid")),10)

	'// 내용 접수
	dim oRelate
	Set oRelate = new CRelateList
	oRelate.FRectLinkCode = LinkCode

	if LinkCode<>"" then
		oRelate.GetRelateLinkItem
	end if
%>
<script type='text/javascript'>
<!--
	// 아이템 저장 실행
	function goSubmit()
	{
		// 카테고리 중분류까지 입력했는지 검사
		if(!(document.frm.cdl.value&&document.frm.cdm.value)) {
			alert("카테고리를 선택해주세요.\n\n※ 관련 키워드는 카테고리 중분류까지 선택하셔야합니다.");
			return;
		}
		// 키워드 입력여부 검사
		if(!document.frm.linkKeyword.value) {
			alert("관련 키워드를 입력해주세요.");
			document.frm.linkKeyword.focus();
			return;
		}
		// 링크 입력여부 검사
		if(!document.frm.linkURL.value) {
			alert("키워드 클릭시 이동할 링크를 입력해주세요.");
			document.frm.linkURL.focus();
			return;
		}

		<% if LinkCode="" then %>
		if(confirm("작성하신 내용을 등록하시겠습니까?")) {
			document.frm.mode.value="add";
			document.frm.action="DoRelate_Process.asp";
			document.frm.submit();
		}
		<% else %>
		if(confirm("수정하신 내용을 저장하시겠습니까?")) {
			document.frm.mode.value="modify";
			document.frm.action="DoRelate_Process.asp";
			document.frm.submit();
		}
		<% end if %>
	}
//-->
</script>
<!-- 폼 시작 -->
<form name="frm" method="get" action="" action="DoRelate_Process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="mode" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<% if LinkCode="" then %>
		<font color="red"><b>관련 키워드 등록</b></font>
		<% else %>
		<font color="red"><b>관련 키워드 수정</b></font>
		<% end if%>
	</td>
</tr>
<% if LinkCode<>"" then %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">링크코드</td>
	<td align="left"><input type="text" name="rid" value="<%=LinkCode%>" readonly size="10" class="text_ro"></td>
</tr>
<% end if %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">카테고리</td>
	<td align="left">
		<%
			'카테고리 지정
			if LinkCode<>"" then
				tmp_cdl = oRelate.FitemList(1).FcdL
				tmp_cdm = oRelate.FitemList(1).FcdM
				tmp_cds = oRelate.FitemList(1).FcdS
			end if
		%>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">키워드</td>
	<td align="left"><input type="text" name="linkKeyword" value="<% if LinkCode<>"" then Response.Write ReplaceBracket(oRelate.FitemList(1).FlinkKeyword) %>" size="32" maxlength="32" class="text"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">링크</td>
	<td align="left">
		<table cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td colspan="2"><input type="text" name="linkURL" value="<% if LinkCode<>"" then Response.Write ReplaceBracket(oRelate.FitemList(1).FlinkURL) %>" size="80" maxlength="128" class="text"></td>
		<tr>
		<tr>
			<td valign="top"><font color="#707080">예)</font></td>
			<td valign="top">
				<font color="#707070">
				- 카테고리 링크 : /shopping/category_list.asp?cdl=<font color="darkred">대코드</font>&cdm=<font color="darkred">중코드</font>&cds=<font color="darkred">소코드</font><br>
				- 이벤트 링크 : /event/eventmain.asp?eventid=<font color="darkred">이벤트코드</font>
				</font>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" colspan="2" bgcolor="#FFFFFF">
		<input type="button" class="button" value="저장" onClick="goSubmit()"> &nbsp;
		<input type="button" class="button" value="취소" onClick="self.history.back()">
	</td>
</tr>
<!-- 폼 끝 -->
</table>
</form>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->